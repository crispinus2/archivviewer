import sys, os, io, email, traceback, subprocess
from PyPDF2 import PdfFileMerger
from lhafile import LhaFile
import pylibjpeg
from PIL import Image
import img2pdf
import logging
from libtiff import TIFF
from subprocess import PIPE

logging.basicConfig(level=logging.DEBUG)

def generateFile(file):
    collectedErrors = []
    merger = PdfFileMerger()
    ios = file
    lf = LhaFile(ios)
    appended = False
    loextensions = ('.odt', '.ods')
    for name in lf.namelist():
        content = lf.read(name)
        _, extension = os.path.splitext(name)
        if content[0:5] == b'%PDF-':                  
            merger.append(io.BytesIO(content))
            appended = True
        elif content[0:5] == b'{\\rtf' or (content[0:4] == bytes.fromhex('504B0304') and extension in loextensions):
            with tempdir() as tmpdir:
                tmpfile = os.sep.join([tmpdir, "temp" + extension])
                pdffile = tmpdir + os.sep + "temp.pdf"
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
        elif name == "message.eml":
            # eArztbrief
            eml = email.message_from_bytes(content)
            errors = []                    
            for part in eml.get_payload():
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
                    img = Image.fromarray(pylibjpeg.decode(content))
                    outbuffer = io.BytesIO()
                    img.save(outbuffer, 'PDF')
                    merger.append(outbuffer)
                    appended = True
                #elif self._config.getValue('useImg2pdf', True):
                #    LOGGER.debug("Using img2pdf for file conversion")
                #    merger.append(io.BytesIO(img2pdf.convert(content)))
                #    appended = True
                else:
                    if content[0:3] == b'II*':
                        #with open('tmp.tif', 'wb') as f:
                        #    f.write(content)
                        infile = 'tmp.tif'
                        outfile = 'tmp.pdf'
                        batchscript = '(let* ((image (car (gimp-file-load RUN-NONINTERACTIVE "{infile}" "{infile}")))(drawable (car (gimp-image-get-active-layer image))))\
                            (file-pdf-save2 RUN-NONINTERACTIVE image drawable "{outfile}" "{outfile}" FALSE TRUE TRUE TRUE FALSE)(gimp-image-delete image) (gimp-quit 0))'.format(infile=infile, outfile=outfile)
                        gimp_path = 'C:\\Program Files\\GIMP 2\\bin\\gimp-2.10.exe'
                        subprocess.run([gimp_path, '-i', '-b', batchscript], check=True, stdout=PIPE, stderr=PIPE)
                        merger.append(outfile)
                        appended = True
            except Exception as e:
                err = "Dateiinhalt '%s' ist kein unterstützter Dateityp -> wird nicht an PDF angehängt (%s)" % (name, type(e))
                collectedErrors.append(err)
                traceback.print_exc(e)
    
    #if appended:
    #    try:
    #        merger.write(filename)
    #    except Exception as e:
    #        err = "{}: Fehler beim Schreiben der Ausgabedatei '{}': {}".format(file["beschreibung"], filename, e)
    #        if errorSlot:
    #            errorSlot.emit(err)
    #        else:
    #            self._av.displayErrorMessage(err)
    print('\n'.join(collectedErrors))
        
    merger.close()
    
if __name__ == '__main__':
    with open(sys.argv[1], 'rb') as f:
        generateFile(f)