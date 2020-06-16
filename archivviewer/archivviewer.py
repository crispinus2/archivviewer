# Archivviewer.py

import sys, codecs, os, fdb, json, tempfile, shutil, subprocess, io
from datetime import datetime, timedelta
from collections import OrderedDict
from contextlib import contextmanager
from pathlib import Path
from lhafile import LhaFile
from PyPDF2 import PdfFileMerger
import img2pdf
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot, QObject, QMutex, QTranslator, QLocale, QLibraryInfo
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from archivviewer.forms import ArchivViewer

exportThread = None

def parseBlobs(blobs):
    entries = {}
    for blob in blobs:
        offset = 0
        totalLength =  int.from_bytes(blob[offset:offset+2], 'little')
        offset += 4
        entryCount = int.from_bytes(blob[offset:offset+2], 'little')
        offset += 2
        
        while offset < len(blob):
            entryLength = int.from_bytes(blob[offset:offset+2], 'little')
            offset += 2
            result = parseBlobEntry(blob[offset:offset+entryLength])
            entries[result['categoryId']] = result['name']
            offset += entryLength
    
    return OrderedDict(sorted(entries.items()))

def parseBlobEntry(blob):
    offset = 6
    name_length = int.from_bytes(blob[offset:offset+2], 'little')
    offset += 2
    name = blob[offset:offset+name_length-1].decode('cp1252')
    offset += name_length
    catid = int.from_bytes(blob[-2:],  'little')
    
    return { 'name': name, 'categoryId': catid }

def displayErrorMessage(msg):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText(msg)
    msg.setStandardButtons(QMessageBox.Ok)

class ArchivViewer(QMainWindow, ArchivviewerUi.Ui_MainWindow):
    def __init__(self, parent = None):
        super(ArchivViewer, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle("Archiv Viewer")

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, gdtfile, model):
        super().__init__()
        self.gdtfile = gdtfile
        self.model = model
    
    def on_modified(self, event):
        print(f'event type: {event.event_type}  path : {event.src_path}')
        if self.gdtfile == event.src_path:
            infos = readGDT(self.gdtfile)
            try:
                patid = int(infos["id"])
                av.patientName.setText('{name}, {surname} [{id}]'.format(**infos))
                self.model.setActivePatient(infos)
                
            except TypeError:
                pass

class PdfExporter(QObject):
    progress = pyqtSignal(int)
    completed = pyqtSignal(int, str)
    error = pyqtSignal(str)
    kill = pyqtSignal()

    def __init__(self, model, filelist, destination, parent=None):
        super(PdfExporter, self).__init__(parent)
        self.model = model
        self.files = filelist
        self.destination = destination
     
    def work(self):
        self.model.exportAsPdfThread(self, self.files, self.destination)
        self.kill.emit()
    

class ArchivTableModel(QAbstractTableModel):    
    _startDate = datetime(1890, 1, 1)

    def __init__(self, con, tmpdir, librepath, table, application):
        super(ArchivTableModel, self).__init__()
        self._files = []
        self._categories = []
        self._con = con
        self._tmpdir = tmpdir
        self._librepath = librepath
        self._table = table
        self._application = application
        self._infos = {}
        self.exportThread = None
        self.mutex = QMutex(mode=QMutex.Recursive)
    
    def __del__(self):
        if self.exportThread:
            self.exportThread.wait()
    
    @contextmanager
    def lock(self, msg=None):
        if msg:
            #print("Acquiring lock for {}.".format(msg))
            pass
        self.mutex.lock()
        yield
        if msg:
            #print("Releasing lock for {}.".format(msg))
            pass
        self.mutex.unlock()
    
    def setActivePatient(self, infos):
        with self.lock("setActivePatient"):        
            self.infos = infos
            self.reloadData(int(infos["id"]))
    
    def data(self, index, role):
        if role == Qt.DisplayRole:
            with self.lock("setActivePatient"):
                file = self._files[index.row()]
                col = index.column()
                
                if col == 0:
                    return file["datum"].strftime('%d.%m.%Y')
                elif col == 1:
                    return file["datum"].strftime('%H:%M')
                elif col == 2:
                    try:
                        return self._categories[file["category"]]
                    except KeyError:
                        return file["category"]
                elif col == 3:
                    return file["beschreibung"]
    
    def rowCount(self, index):
        with self.lock("rowCount"):
            return len(self._files)
        
    def columnCount(self, index):
        return 4
        
    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if section == 0:
                    return "Datum"
                elif section == 1:
                    return "Zeit"
                elif section == 2:
                    return "Kategorie"
                elif section == 3:
                    return "Beschreibung"
    
    def reloadData(self, patnr):
        self._application.setOverrideCursor(Qt.WaitCursor)
        self.beginResetModel()
        self.reloadCategories()
        self._files = []
        
        selectStm = "SELECT a.FSUROGAT, a.FTEXT, a.FEINTRAGSART, a.FZEIT, a.FDATUM FROM ARCHIV a WHERE a.FPATNR = ? ORDER BY a.FDATUM DESC, a.FZEIT DESC"
        cur = con.cursor()
        cur.execute(selectStm, (patnr,))
        
        for (surogat, beschreibung, eintragsart, zeit, datum) in cur:
            self._files.append({
                'id': surogat,
                'datum': self._startDate + timedelta(days = datum, seconds = zeit),
                'beschreibung': beschreibung,
                'category': eintragsart
            })
        
        del cur
        
        print("Found %d files." % (len(self._files)))
        
        self.endResetModel()
        self._table.resizeColumnsToContents()
        self._table.horizontalHeader().setStretchLastSection(True)
        
        self._application.restoreOverrideCursor()
    
    def reloadCategories(self):
        cur = self._con.cursor()
        cur.execute("SELECT s.FKATEGORIELISTE, s.FABLAGELISTE, s.FBRIEFKATEGORIELISTE FROM MOSYSTEM s")
        for blobs in cur:
            self._categories = parseBlobs(blobs)
            break
        print("Found %d categories" % (len(self._categories)))
        del cur
    
    def generateFile(self, rowIndex, errorSlot = None):
        with self.lock("generateFile"):
            file = self._files[rowIndex]
            filename = self._tmpdir + os.sep + '{}.pdf'.format(file["id"])
        if not os.path.isfile(filename):
            selectStm = "SELECT a.FDATEI FROM ARCHIV a WHERE a.FSUROGAT = ?"
            cur = con.cursor()
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
            
            for name in lf.namelist():
                content = lf.read(name)
                if content[0:5] == b'%PDF-':                  
                    merger.append(io.BytesIO(content))
                elif content[0:5] == b'{\\rtf':
                    with tempdir() as tmpdir:
                        tmpfile = tmpdir + os.sep + "temp.rtf"
                        pdffile = tmpdir + os.sep + "temp.pdf"
                        with open(tmpfile, "wb") as f:
                            f.write(content)
                        command = '"'+" ".join(['"'+self._librepath+'"', "--convert-to pdf", "--outdir", '"'+tmpdir+'"', '"'+tmpfile+'"'])+'"'
                        if os.system(command) == 0:
                            try:
                                with open(pdffile, "rb") as f:
                                    merger.append(io.BytesIO(f.read()))
                            except:
                                if errorSlot:
                                    errorSlot.emit("Fehler beim Öffnen der konvertierten PDF-Datei '%s'" % (pdffile))
                                else:
                                    displayErrorMessage("Fehler beim Öffnen der konvertierten PDF-Datei '%s'" % (pdffile))
                        else:
                            if errorSlot:
                                errorSlot.emit("Fehler beim Ausführen des Kommandos: '%s'" % (command))
                            else:
                                displayErrorMessage("Fehler beim Ausführen des Kommandos: '%s'" % (command))
                elif name == "message.eml":
                    # eArztbrief
                    eml = email.message_from_bytes(content)
                    
                    for part in eml.get_payload():
                        fnam = part.get_filename()
                        partcont = part.get_payload(decode=True)
                        if partcont[0:5] == b'%PDF-':
                            print("eArztbrief: hänge Anhang '%s' an den Export an" % (fnam))
                            merger.append(io.BytesIO(partcont))
                        else:
                            print("eArztbrief: nicht unterstütztes Anhangsformat in Anhang '%s'" % (fnam))
                else:
                    try:
                        merger.append(io.BytesIO(img2pdf.convert(content)))
                    except Exception as e:
                        print("Dateiinhalt '%s' ist kein unterstützter Dateityp -> wird nicht an PDF angehängt (%s)" % (name, e))
            
            merger.write(filename)
            merger.close()
            
            try:
                datei.close()
            except:
                pass
            ios.close()
        return filename
    
    def displayFile(self, rowIndex):
        self._application.setOverrideCursor(Qt.WaitCursor)
        filename = self.generateFile(rowIndex)
        self._application.restoreOverrideCursor()
        subprocess.run(['start', filename], shell=True)
    
    def exportAsPdfThread(self, thread, filelist, destination):
        self._application.setOverrideCursor(Qt.WaitCursor)
        with self.lock("exportAsPdfThread"):
            files = list(self._files)
        try:
            merger = PdfFileMerger()
            counter = 0
            filelist = sorted(filelist)
            for file in filelist:
                counter += 1
                thread.progress.emit(counter)
                
                filename = self.generateFile(file, errorSlot = thread.error)
                bmtext = " ".join([files[file]["beschreibung"], files[file]["datum"].strftime('%d.%m.%Y %H:%M')])
                merger.append(filename, bookmark=bmtext)
            
            
            merger.write(destination)
            merger.close()
        except IOError as e:    
            thread.error.emit("Fehler beim Schreiben der PDF-Datei: {}".format(e))
            
        thread.progress.emit(0)
        self._application.restoreOverrideCursor()
        thread.completed.emit(counter, destination)
        
        return counter
    
    def updateProgress(self, value):
        av.exportProgress.setFormat('%d von %d' % (value, av.exportProgress.maximum()))
        av.exportProgress.setValue(value)
                
    def exportCompleted(self, counter, destination):
        av.exportProgress.setFormat('')
        av.exportProgress.setEnabled(False)
        av.exportPdf.setEnabled(True)
        av.documentView.setEnabled(True)
        QMessageBox.information(None, "Export abgeschlossen", "%d Dokumente wurden nach '%s' exportiert" % (counter, destination))
    
    def handleError(self, msg):
        displayErrorMessage(msg)
    
    def exportAsPdf(self, filelist):   
        if len(filelist) == 0:
            buttonReply = QMessageBox.question(av, 'PDF-Export', "Kein Dokument ausgewählt. Export aus allen Dokumenten des Patienten erzeugen?", QMessageBox.Yes | QMessageBox.No)
            if(buttonReply == QMessageBox.Yes):
                with self.lock("exportAsPdf (filelist)"):
                    filelist = range(len(self._files))
            else:
                return
                
        dirconfpath = os.sep.join([os.environ["AppData"], "ArchivViewer", "config.json"])
        outfiledir = str(Path.home())
        try:
            with open(dirconfpath, "r") as f:
                conf = json.load(f)
                outfiledir = conf["outfiledir"]
        except:
            pass
            
        outfilename = os.sep.join([outfiledir, 'Patientenakte_%d_%s_%s_%s-%s.pdf' % (int(infos["id"]), infos["name"], infos["surname"], infos["birthdate"], datetime.now().strftime('%Y%m%d%H%M%S'))])
        destination, _ = QFileDialog.getSaveFileName(av, "Auswahl als PDF exportieren", outfilename, "PDF-Datei (*.pdf)")
        if len(destination) > 0:
            try:
                os.makedirs(os.path.dirname(dirconfpath), exist_ok = True)
                with open(dirconfpath, "w") as f:
                    json.dump({ 'outfiledir': os.path.dirname(destination) }, f, indent = 1)
            except:
                pass
                
            av.exportPdf.setEnabled(False)
            av.documentView.setEnabled(False)
            av.exportProgress.setEnabled(True)
            av.exportProgress.setRange(0, len(filelist))
            av.exportProgress.setFormat('0 von %d' % (len(filelist)))
            self.pdfExporter = PdfExporter(self, filelist, destination)
            self.exportThread = QThread()
            self.pdfExporter.moveToThread(self.exportThread)
            self.pdfExporter.kill.connect(self.exportThread.quit)
            self.pdfExporter.progress.connect(self.updateProgress)
            self.pdfExporter.completed.connect(self.exportCompleted)
            self.pdfExporter.error.connect(self.handleError)
            self.exportThread.started.connect(self.pdfExporter.work)
            self.exportThread.start()
            print("Export thread started")
            

def readGDT(gdtfile):
    grabinfo = {
        3000: "id",
        3101: "name",
        3102: "surname",
        3103: "birthdate"
    }
    infos = {
        "id": None,
        "name": None,
        "surname": None
    }
    with codecs.open(gdtfile, encoding="iso-8859-15", mode="r") as f:
        for line in f:
            linelen = int(line[:3])
            feldkennung = int(line[3:7])
            inhalt = line[7:linelen - 2]
            if feldkennung in grabinfo:
                infos[grabinfo[feldkennung]] = inhalt
                
    
    return infos

@contextmanager
def tempdir(prefix='tmp'):
    """A context manager for creating and then deleting a temporary directory."""
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    try:
        yield tmpdir
    finally:
        shutil.rmtree(tmpdir)

def tableDoubleClicked(table, model):
    row = table.currentIndex().row()
    if row > -1:
        model.displayFile(row)

def exportSelectionAsPdf(table, model):
    global exportThread
    
    indexes = table.selectionModel().selectedRows()
    files = []
    for idx in indexes:
        files.append(idx.row())
    
    model.exportAsPdf(files)

def main():
    app = QApplication(sys.argv)
    qt_translator = QTranslator()
    qt_translator.load("qt_" + QLocale.system().name(),
        QLibraryInfo.location(QLibraryInfo.TranslationsPath))
    app.installTranslator(qt_translator)
    
    qtbase_translator = QTranslator()
    qt_translator.load("qtbase_" + QLocale.system().name(),
        QLibraryInfo.location(QLibraryInfo.TranslationsPath))
    app.installTranslator(qtbase_translator)
    
    conffile = os.getcwd() + os.sep + "Patientenakte.cnf"
    conffile2 = os.path.dirname(os.path.realpath(__file__)) + os.sep + "Patientenakte.cnf"
    writeconf = True
    defaultHost = "localhost"
    defaultDb = "D:\\INDAMED\\DAT\\MEDOFF.GDB"
    defaultDbUser = "sysdba"
    defaultDbPassword = "masterkey"
    defaultClientLib = None
    defaultOutput = os.getcwd()
    defaultLibrePath = "C:\Program Files\LibreOffice\program\soffice.exe"

    for cfg in [conffile, conffile2]:
        try:
            print("Attempting config %s" % (cfg))
            with open(cfg, 'r') as f:
                conf = json.load(f)
                if "hostname" in conf:
                    defaultHost = conf["hostname"]
                if "database" in conf:
                    defaultDb = conf["database"]
                if "dbuser" in conf:
                    defaultDbUser = conf["dbuser"]
                if "dbpassword" in conf:
                    defaultDbPassword = conf["dbpassword"]
                if "clientlib" in conf:
                    defaultClientLib = conf["clientlib"]
                if "output" in conf:
                    defaultOutput = conf["output"]
                if "libreoffice" in conf:
                    defaultLibrePath = conf["libreoffice"]
                if cfg != conffile:
                    writeconf = False
            break
        except Exception as e:
            print("Failed to load config: %s." % (e))
    
    print("Client lib is %s" % (defaultClientLib))
    try:
        con = fdb.connect(host=defaultHost, database=defaultDb, 
            user=defaultDbUser, password=defaultDbPassword, fb_library_name=defaultClientLib)
    except Exception as e:
        displayErrorMessage(e)
        sys.exit(app.exec_())
    with tempdir() as myTemp:
        gdtfile = 'C:\\EAS\\aktpat.bdt'
        av = ArchivViewer()
        tm = ArchivTableModel(con, myTemp, defaultLibrePath, av.documentView, app)
        av.documentView.doubleClicked.connect(lambda: tableDoubleClicked(av.documentView, tm))
        av.documentView.setModel(tm)
        av.exportPdf.clicked.connect(lambda: exportSelectionAsPdf(av.documentView, tm))
        event_handler = FileChangeHandler(gdtfile, tm)
        av.action_quit.triggered.connect(lambda: app.quit())
        av.action_about.triggered.connect(lambda: QMessageBox.about(av, "Über Archiv Viewer", 
            """<p><b>Archiv Viewer</b> ist eine zur Verwendung mit Medical Office der Fa. Indamed entwickelte
             Software, die synchron zur Medical Office-Anwendung die gespeicherten Dokumente eines Patienten im Archiv
             anzeigen kann. Zusätzlich können ausgewählte Dokumente auch als PDF-Datei exportiert werden.</p><p>(c) 2020 Julian Hartig - Lizensiert unter den Bedingungen der GPLv3</p>"""))
        
        observer = Observer()
        observer.schedule(event_handler, path=os.path.dirname(gdtfile), recursive=False)
        observer.start()
        
        try:
            infos = readGDT(gdtfile)
            av.patientName.setText('{name}, {surname} [{id}]'.format(**infos))
            tm.reloadData(int(infos["id"]))
        except Exception as e:
            print("While loading GDT file: %s" % (e))
        
        av.show()
        ret = app.exec_()
        observer.stop()
        if exportThread:
            exportThread.stop()
            exportThread.join()
        observer.join()
    sys.exit(ret)