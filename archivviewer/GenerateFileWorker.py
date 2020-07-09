# GenerateFileWorker.py

import io, logging, email, subprocess, os, tempfile, shutil
from contextlib import contextmanager
from PyQt5.QtCore import QObject, pyqtSignal
from lhafile import LhaFile
import img2pdf
import libjpeg
from PIL import Image, ImageFile
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
        
    def __init__(self, tmpdir, files, con, arccon, librepath, gimppath, exportDestination = None, parent = None):
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
        
    def work(self):
        try:
            if self._destination is None:
                filename, errors = self.generateFile(self._files[0])
                self.exportCompleted.emit(filename, 1, 0 if filename is None else 1, errors, False)
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
                
                try:
                    self.exportProgressStatus.emit('Schreibe Exportdatei...')
                    if failed < counter:
                        merger.write(self._destination)
                except IOError as e:
                    failed = counter
                    errorMessages.append("Fehler beim Schreiben der PDF-Datei: {}".format(e))
                finally:
                    merger.close()
                    
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
                            if content[0:4] == bytes.fromhex('FFD8FFE0'):
                                LOGGER.debug("{}: {}: Export via libjpeg".format(file["beschreibung"], name))
                                img = Image.fromarray(libjpeg.decode(content))
                                outbuffer = io.BytesIO()
                                img.save(outbuffer, 'PDF')
                                merger.append(outbuffer)
                                appended = True
                            elif self._gimppath is not None and content[0:3] == b'II*' and self._config.getValue('useGimpForTiff', False):                                
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
                            elif self._config.getValue('useImg2pdf', True):
                                LOGGER.debug("Using img2pdf for file conversion")
                                merger.append(io.BytesIO(img2pdf.convert(content)))
                                appended = True
                            else:
                                LOGGER.debug("Using PIL for file conversion")
                                inbuffer = io.BytesIO(content)
                                img = Image.open(inbuffer)
                                outbuffer = io.BytesIO()
                                img.save(outbuffer, 'PDF')
                                merger.append(outbuffer)
                                appended = True
                                del inbuffer
                        except Exception as e:
                            err = "%s: Dateiinhalt '%s' ist kein unterstützter Dateityp -> wird nicht an PDF angehängt (%s)" % (file["beschreibung"], name, e)
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
        except:
            raise
        finally:
            for f in cleanupfiles:
                try:
                    os.unlink(f)
                except:
                    pass
        
            self.completed.emit(filename, file, collectedErrors, isExport)
        
        return (filename, collectedErrors)