# GenerateFileWorker.py

import io, logging, email, subprocess, os, tempfile, shutil, sys
from contextlib import contextmanager
from PyQt5.QtCore import QObject, pyqtSignal
from lhafile import LhaFile
import img2pdf
import libjpeg
from PIL import Image, ImageFile
import PIL
from PyPDF2 import PdfFileMerger
from subprocess import PIPE
from .configreader import ConfigReader

LOGGER = logging.getLogger(__name__)

@contextmanager
def tempdir(prefix='tmp'):
    """A context manager for creating and then deleting a temporary directory."""
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    try:
        yield tmpdir
    finally:
        shutil.rmtree(tmpdir)

class ExportCancelledError(Exception):
    pass

class GenerateFileWorker(QObject):
    initGenerate = pyqtSignal(int, bool)
    initExport = pyqtSignal(int)
    progress = pyqtSignal(bool)
    progressExport = pyqtSignal()
    completed = pyqtSignal(str, dict, list, bool)
    exportCompleted = pyqtSignal(str, int, int, list, bool)
    exportCancelled = pyqtSignal()
    fileProgressStatus = pyqtSignal(str)
    exportProgressStatus = pyqtSignal(str)
    kill = pyqtSignal()
        
    def __init__(self, tmpdir, files, con, arccon, librepath, gimppath, gspath, exportDestination = None, parent = None):
        super(GenerateFileWorker, self).__init__(parent)
        self._tmpdir = tmpdir
        self._files = files
        self._gimppath = gimppath
        self._librepath = librepath
        self._con = con
        self._arccon = arccon
        self._destination = exportDestination
        self._cancelled = False
        self._config = ConfigReader.get_instance()
        self._gspath = gspath
        
    def work(self):
        try:
            if self._destination is None:
                filename, errors = self.generateFile(self._files[0])
                self.exportCompleted.emit(filename, 1, 1 if filename is None else 0, errors, False)
            else:
                failed = 0
                counter = 0
                errorMessages = []
                merger = PdfFileMerger()
                for f in self._files:
                    self._raiseIfCancelled()
                    counter += 1
                    filename, errors = self.generateFile(f)
                    errorMessages.extend(errors)
                    self._raiseIfCancelled()
                    if filename is None:
                        failed += 1
                    else:
                        bmtext = " ".join([f["beschreibung"], f["datum"].strftime('%d.%m.%Y %H:%M')])
                        merger.append(filename, bookmark=bmtext)
                    self.progressExport.emit()
                
                tmpdest = self._destination
                if self._gspath is not None and self._config.getValue("shrinkPDF", True):
                    tmpdest = os.path.join(self._tmpdir, 'export.pdf')
                
                try:
                    self.exportProgressStatus.emit('Schreibe Exportdatei...')
                    if failed < counter:
                        merger.write(tmpdest)
                except IOError as e:
                    failed = counter
                    errorMessages.append("Fehler beim Schreiben der PDF-Datei: {}".format(e))
                finally:
                    merger.close()
                
                if self._gspath is not None and self._config.getValue("shrinkPDF", True):
                    self.exportProgressStatus.emit('Optimiere PDF-Datei mit Ghostscript...')
                    result = subprocess.run([self._gspath, '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4', '-dPDFSETTINGS=/printer',
                        '-dNOPAUSE', '-dQUIET', '-dBATCH', '-sOutputFile={}'.format(self._destination), tmpdest], check=True, stdout=PIPE, stderr=PIPE)
                
                self.exportCompleted.emit(self._destination, counter, failed, errorMessages, True)
        except ExportCancelledError:
            self.exportCancelled.emit()
        
    def cancel(self):
        self._cancelled = True
    
    def _raiseIfCancelled(self):
        if self._cancelled:
            raise ExportCancelledError('Export cancelled by user')
    
    def generateFile(self, file):
        filename = os.sep.join([self._tmpdir, '{}.pdf'.format(file["id"])])
        collectedErrors = []
        cleanupfiles = []
        isExport = self._destination is not None
        
        try:
            if not os.path.isfile(filename):
                self.fileProgressStatus.emit("Hole Blob aus Datenbank...")
                self._raiseIfCancelled()
                selectStm = "SELECT a.FDATEI FROM ARCHIV a WHERE a.FSUROGAT = ?"
                if 'medoffarc' in file:
                    cur = self._arccon.cursor()
                else:
                    cur = self._con.cursor()
                cur.execute(selectStm, (file["id"],))
                (datei,) = cur.fetchone()
                
                merger = PdfFileMerger()
                ios = None
                try:
                    contents = datei.read()
                    ios = io.BytesIO(contents)
                except:
                    ios = io.BytesIO(datei)
                
                self._raiseIfCancelled()
                self.fileProgressStatus.emit("Öffne Archivdatei...")
                lf = LhaFile(ios)
                
                appended = False
                loextensions = ('.odt', '.ods')
                
                attcounter = 0
                
                self.initGenerate.emit(len(lf.namelist()), isExport)
                
                for name in lf.namelist():
                    self._raiseIfCancelled()
                    self.progress.emit(isExport)
                    content = lf.read(name)
                    _, extension = os.path.splitext(name)
                    if content[0:5] == b'%PDF-':                  
                        merger.append(io.BytesIO(content))
                        appended = True
                    elif content[0:5] == b'{\\rtf' or (content[0:4] == bytes.fromhex('504B0304') and extension in loextensions):
                        if self._librepath is not None:
                            with tempdir() as tmpdir:
                                tmpfile = os.sep.join([tmpdir, "temp" + extension])
                                pdffile = os.sep.join([tmpdir, "temp.pdf"])
                                with open(tmpfile, "wb") as f:
                                    f.write(content)
                                command = '"'+" ".join(['"'+self._librepath+'"', "--convert-to pdf", "--outdir", '"'+tmpdir+'"', '"'+tmpfile+'"'])+'"'
                                if os.system(command) == 0:
                                    try:
                                        with open(pdffile, "rb") as f:
                                            merger.append(io.BytesIO(f.read()))
                                            
                                        appended = True
                                    except:
                                        err = "%s: Fehler beim Öffnen der konvertierten PDF-Datei '%s' (konvertiert aus '%s')" % (file['beschreibung'], pdffile, tmpfile)
                                        collectedErrors.append(err)
                                else:
                                    err = "%s: Fehler beim Ausführen des Kommandos: '%s'" % (file['beschreibung'], command)
                                    collectedErrors.append(err)
                        else:
                            err = "%s: Die Konvertierung nach PDF ist nicht möglich, da keine LibreOffice-Installation gefunden wurde"
                            collectedErrors.append(err)
                    elif name == "message.eml":
                        # eArztbrief
                        eml = email.message_from_bytes(content)
                        errors = []                    
                        for part in eml.get_payload():
                            self._raiseIfCancelled()
                            fnam = part.get_filename()
                            partcont = part.get_payload(decode=True)
                            if partcont[0:5] == b'%PDF-':
                                merger.append(io.BytesIO(partcont))
                                appended = True
                            else:
                                errors.append("%s: eArztbrief: nicht unterstütztes Anhangsformat in Anhang '%s'" % (file["beschreibung"], fnam))
                        
                        if not appended and len(errors) > 0:
                            err = '\n'.join(errors)
                            collectedErrors.append(err)
                    else:
                        try:
                            if self._gimppath is not None and content[0:3] == b'II*' and self._config.getValue('useGimpForTiff', False):                                
                                LOGGER.debug("{}: {}: Export via GIMP unter '{}'".format(file["beschreibung"], name, self._gimppath))
                                tiffile = os.sep.join([self._tmpdir, '{}.{}.tif'.format(file["id"], attcounter)])
                                outfile = os.sep.join([self._tmpdir, '{}.{}.pdf'.format(file["id"], attcounter)])
                                cleanupfiles.append(tiffile)
                                cleanupfiles.append(outfile)
                                with open(tiffile, 'wb') as f:
                                    f.write(content)
                                batchscript = '(let* ((image (car (gimp-file-load RUN-NONINTERACTIVE "{infile}" "{infile}")))(drawable (car (gimp-image-get-active-layer image))))\
                                    (file-pdf-save2 RUN-NONINTERACTIVE image drawable "{outfile}" "{outfile}" FALSE TRUE TRUE TRUE FALSE)(gimp-image-delete image) (gimp-quit 0))'.format(infile=tiffile.replace('\\', '\\\\'), outfile=outfile.replace('\\', '\\\\'))
                                gimp_path = self._gimppath
                                result = subprocess.run([gimp_path, '-i', '-b', batchscript], check=True, stdout=PIPE, stderr=PIPE)
                                if result.stdout is not None or result.stderr is not None:
                                    LOGGER.debug("GIMP output: {} {}".format(result.stdout, result.stderr))
                                merger.append(outfile)
                                appended = True
                                attcounter += 1
                            elif self._config.getValue('useImg2pdf', False) or content[0:4] == bytes.fromhex('FFD8FFE0'):
                                LOGGER.debug("Using PIL for file conversion")
                                inbuffer = io.BytesIO(content)
                                try:
                                    img = Image.open(inbuffer)
                                    img.load()
                                except OSError as e:
                                    LOGGER.debug("Failed: {}. Trying libjpeg now for lossless compressed JPEG file.".format(e))
                                    img = Image.fromarray(libjpeg.decode(content))
                                if self._config.getValue('fitToA4', False):
                                    twidth = 2480
                                    theight = 3508
                                    cwidth, cheight = img.size
                                    if cwidth > twidth or cheight > theight:
                                        img.resize((twidth, theight), resample = PIL.Image.LANCZOS)
                                    elif cwidth < twidth and cheight < theight:
                                        nimg = Image.new('RGB', (twidth, theight), color = 'white')
                                        nimg.paste(img, (round((twidth-cwidth)/2), round((theight-cheight)/2)))
                                        img = nimg
                                outbuffer = io.BytesIO()
                                img.save(outbuffer, 'PDF', resolution=300)
                                merger.append(outbuffer)
                                appended = True
                                del inbuffer
                            else:
                                LOGGER.debug("Using img2pdf for file conversion")
                                if self._config.getValue('fitToA4', False):
                                    a4inpt = (img2pdf.mm_to_pt(210),img2pdf.mm_to_pt(297))
                                    layout_fun = img2pdf.get_layout_fun(a4inpt)
                                else:
                                    layout_fun = None
                                merger.append(io.BytesIO(img2pdf.convert(content), layout_fun=layout_fun))
                                appended = True
                        except Exception as e:
                            err = "%s: Dateiinhalt '%s' ist kein unterstützter Dateityp -> wird nicht an PDF angehängt (%s)" % (file["beschreibung"], name, e)
                            LOGGER.debug(err)
                            collectedErrors.append(err)
                
                self._raiseIfCancelled()
                if appended:
                    try:
                        merger.write(filename)
                    except Exception as e:
                        err = "{}: Fehler beim Schreiben der Ausgabedatei '{}': {}".format(file["beschreibung"], filename, e)
                        self._av.displayErrorMessage(err)
                else:
                    filename = None
                    blobfilename = os.sep.join([self._tmpdir, '{}.blb'.format(file["id"])])
                    ios.seek(0)
                    with open(blobfilename, 'wb') as f:
                        f.write(ios.read())
                    err = "Ein Abzug des Blobinhalts wurde nach '{}' geschrieben. Er ist dort bis zum Programmende verfügbar.".format(blobfilename)
                    collectedErrors.append(err)
                    
                merger.close()
                
                try:
                    datei.close()
                except:
                    pass
                ios.close()
                
                self.completed.emit(filename, file, collectedErrors, isExport)
        except Exception as e:
            LOGGER.debug("Exception on generating file: {}".format(e))
            self.completed.emit(None, file, collectedErrors, isExport)
            raise e
        finally:
            for f in cleanupfiles:
                try:
                    os.unlink(f)
                except:
                    pass
            LOGGER.debug("File generation completed")
            
        
        return (filename, collectedErrors)