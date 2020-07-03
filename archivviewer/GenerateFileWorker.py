# GenerateFileWorker.py

class GenerateFileWorker(QObject):
    initGenerate = pyqtSignal(int)
    progress = pyqtSignal()
    completed = pyqtSignal()
    errors = pyqtSignal()
    
    def __init__(self, tmpdir, file, parent = None):
        super(GenerateFileThread, self).__init__(parent)
        self._tmpdir = tmpdir
        self._file = file
        
    def work(self):
        self.generateFile(self._file, self.errors)
        
    def generateFile(self, file, errorSlot = None):
        filename = self._tmpdir + os.sep + '{}.pdf'.format(file["id"])
        collectedErrors = []
        with self.lock("generateFile"):
            if not os.path.isfile(filename):
                selectStm = "SELECT a.FDATEI FROM ARCHIV a WHERE a.FSUROGAT = ?"
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
                        
                lf = LhaFile(ios)
                
                appended = False
                loextensions = ('.odt', '.ods')
                
                attcounter = 0
                cleanupfiles = []
                
                for name in lf.namelist():
                    content = lf.read(name)
                    _, extension = os.path.splitext(name)
                    if content[0:5] == b'%PDF-':                  
                        merger.append(io.BytesIO(content))
                        appended = True
                    elif content[0:5] == b'{\\rtf' or (content[0:4] == bytes.fromhex('504B0304') and extension in loextensions):
                        if self._librepath is not None:
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
                                        if errorSlot:
                                            errorSlot.emit(err)
                                        else:
                                            collectedErrors.append(err)
                                else:
                                    err = "%s: Fehler beim Ausführen des Kommandos: '%s'" % (file['beschreibung'], command)
                                    if errorSlot:
                                        errorSlot.emit(err)
                                    else:
                                        collectedErrors.append(err)
                        else:
                            err = "%s: Die Konvertierung nach PDF ist nicht möglich, da keine LibreOffice-Installation gefunden wurde"
                            if errorSlot:
                                errorSlot.emit(err)
                            else:
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
                            if errorSlot:
                                errorSlot.emit(err)
                            else:
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
                                    (gimp-file-save RUN-NONINTERACTIVE image drawable "{outfile}" "{outfile}")(gimp-image-delete image) (gimp-quit 0))'.format(infile=tiffile.replace('\\', '\\\\'), outfile=outfile.replace('\\', '\\\\'))
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
                            if errorSlot:
                                errorSlot.emit(err)
                            else:
                                collectedErrors.append(err)
                
                if appended:
                    try:
                        merger.write(filename)
                    except Exception as e:
                        err = "{}: Fehler beim Schreiben der Ausgabedatei '{}': {}".format(file["beschreibung"], filename, e)
                        if errorSlot:
                            errorSlot.emit(err)
                        else:
                            self._av.displayErrorMessage(err)
                else:
                    filename = None
                    blobfilename = os.sep.join([self._tmpdir, '{}.blb'.format(file["id"])])
                    ios.seek(0)
                    with open(blobfilename, 'wb') as f:
                        f.write(ios.read())
                    err = "Ein Abzug des Blobinhalts wurde nach '{}' geschrieben. Er ist dort bis zum Programmende verfügbar.".format(blobfilename)
                    if errorSlot:
                        errorSlot.emit(err)
                    else:
                        collectedErrors.append(err)
                        self._av.displayErrorMessage('\n'.join(collectedErrors))
                    
                merger.close()
                
                try:
                    datei.close()
                except:
                    pass
                ios.close()
                
                for f in cleanupfiles:
                    try:
                        os.unlink(f)
                    except:
                        pass
                
        return filename