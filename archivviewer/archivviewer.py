# Archivviewer.py

import sys, codecs, os, fdb, json, tempfile, shutil, subprocess, io, winreg, configparser, email, logging
from datetime import datetime, timedelta
from collections import OrderedDict
from contextlib import contextmanager
from pathlib import Path
from lhafile import LhaFile
from PyPDF2 import PdfFileMerger
import img2pdf
import pylibjpeg
from PIL import Image, ImageFile
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog, QStyle
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot, QObject, QMutex, QTranslator, QLocale, QLibraryInfo, QEvent, QSettings
from PyQt5.QtGui import QColor, QBrush, QIcon
from PyQt5.QtWinExtras import QWinTaskbarProgress, QWinTaskbarButton
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from functools import reduce

from archivviewer.forms import ArchivviewerUi
from .configreader import ConfigReader
from .CategoryModel import CategoryModel
from .FilesTableDelegate import FilesTableDelegate

exportThread = None

logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger(__name__)

def ilen(iterable):
    return reduce(lambda sum, element: sum + 1, iterable, 0)

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
    QMessageBox.critical(None, "Fehler", str(msg))

class ArchivViewer(QMainWindow, ArchivviewerUi):
    def __init__(self, parent = None):
        super(ArchivViewer, self).__init__(parent)
        self._config = ConfigReader.get_instance()
        self.taskbar_button = None
        self.taskbar_progress = None
        self.setupUi(self)
        iconpath = os.sep.join([os.path.dirname(os.path.realpath(__file__)), 'icon128.png'])
        self.setWindowIcon(QIcon(iconpath))
        self.setWindowTitle("Archiv Viewer")
        self.refreshFiles.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self.exportPdf.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.readSettings()
        self.delegate = FilesTableDelegate()
        self.documentView.setItemDelegate(self.delegate)
        self.actionStayOnTop.changed.connect(self.stayOnTopChanged)
        self.actionShowPDFAfterExport.changed.connect(self.showPDFAfterExportChanged)
        self.actionUseImg2pdf.changed.connect(self.useImg2pdfChanged)
    
    def displayErrorMessage(self, msg):
        QMessageBox.critical(self, "Fehler", str(msg))    
        
    def stayOnTopChanged(self):
        ontop = self.actionStayOnTop.isChecked()
        self._config.setValue('stayontop', ontop)
        if ontop:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & (~Qt.WindowStaysOnTopHint))
        self.show()
    
    def showPDFAfterExportChanged(self):
        show = self.actionShowPDFAfterExport.isChecked()
        self._config.setValue('showPDFAfterExport', show)
    
    def useImg2pdfChanged(self):
        use = self.actionUseImg2pdf.isChecked()
        self._config.setValue('useImg2pdf', use)
    
    def event(self, evt):
        ontop = self._config.getValue('stayontop')
        if evt.type() == QEvent.WindowActivate or evt.type() == QEvent.HoverEnter:
            self.setWindowOpacity(1.0)
        elif (evt.type() == QEvent.WindowDeactivate or evt.type() == QEvent.HoverLeave) and not self.isActiveWindow() and ontop:
            self.setWindowOpacity(0.6)
        return QMainWindow.event(self, evt)
    
    def showEvent(self, evt):
        self.taskbar_button = QWinTaskbarButton()
        self.taskbar_progress = self.taskbar_button.progress()
        self.taskbar_button.setWindow(self.windowHandle())
    
    def closeEvent(self, evt):
        settings = QSettings(QSettings.UserScope, "cortex", "ArchivViewer")
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("windowState", self.saveState())
        settings.setValue("splitter", self.splitter.saveState())
        settings.sync()
        QMainWindow.closeEvent(self, evt)
    
    def readSettings(self):
        settings = QSettings(QSettings.UserScope, "cortex", "ArchivViewer")
        try:
            self.restoreGeometry(settings.value("geometry"))
            self.restoreState(settings.value("windowState"))
            self.splitter.restoreState(settings.value("splitter"))
        except Exception as e:
            pass
        

class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, gdtfile, model):
        super().__init__()
        self.gdtfile = gdtfile
        self.model = model
    
    def on_modified(self, event):
        if self.gdtfile == event.src_path:
            infos = readGDT(self.gdtfile)
            try:
                patid = int(infos["id"])
                self.model.activePatientChanged.emit(infos)
                
            except TypeError:
                pass

class PdfExporter(QObject):
    progress = pyqtSignal(int)
    completed = pyqtSignal(int, int, str)
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
    _dataReloaded = pyqtSignal()
    activePatientChanged = pyqtSignal(dict)

    def __init__(self, con, tmpdir, librepath, mainwindow, application, categoryModel):
        super(ArchivTableModel, self).__init__()
        self._unfilteredFiles = []
        self._files = []
        self._con = con
        self._tmpdir = tmpdir
        self._librepath = librepath
        self._table = mainwindow.documentView
        self._av = mainwindow
        self._application = application
        self._infos = {}
        self._exportErrors = []
        self._categoryModel = categoryModel
        self._categoryFilter = set()
        self._av.categoryList.selectionModel().selectionChanged.connect(self.categorySelectionChanged)
        self._av.filterDescription.textEdited.connect(self.filterTextChanged)
        self._config = ConfigReader.get_instance()
        self._dataReloaded.connect(self._av.filterDescription.clear)
        self._dataReloaded.connect(self.updateLabel)
        self._av.refreshFiles.clicked.connect(lambda: self.activePatientChanged.emit(self._infos))
        self.activePatientChanged.connect(self.setActivePatient)
        self.exportThread = None
        self.mutex = QMutex(mode=QMutex.Recursive)
        
    
    def __del__(self):
        if self.exportThread:
            self.exportThread.wait()
    
    @contextmanager
    def lock(self, msg=None):
        if msg:
            #print("Entering {}".format(msg))
            pass
        self.mutex.lock()
        try:
            yield
        except:
            raise
        finally:
            if msg:
                #print("Leaving {}".format(msg))
                pass
            self.mutex.unlock()
    
    def categorySelectionChanged(self, selected, deselected):
        for idx in selected.indexes():
            id = self._categoryModel.idAtRow(idx.row())
            self._categoryFilter.add(id)
        for idx in deselected.indexes():
            id = self._categoryModel.idAtRow(idx.row())
            self._categoryFilter.discard(id)
                
        self.applyFilters()
    
    def filterTextChanged(self, _):
        self.applyFilters()
    
    def applyFilters(self):
        self.beginResetModel()
        self._applyFilters()
        self.endResetModel()
                
    def _applyFilters(self):
        
        self._files = list(filter(lambda x: 
                                 (len(self._categoryFilter) == 0 or x['category'] in self._categoryFilter)
                                 and (len(self._av.filterDescription.text()) == 0 or self._av.filterDescription.text().lower() in x['beschreibung'].lower()), self._unfilteredFiles))
    
    def endResetModel(self):
        QAbstractTableModel.endResetModel(self)
        self._table.resizeColumnsToContents()
        self._table.horizontalHeader().setStretchLastSection(True)
        hasFiles = len(self._files) > 0
        self._table.horizontalHeader().setVisible(hasFiles)
        self._av.exportPdf.setEnabled(hasFiles)
    
    def updateLabel(self):
        
        unb = self._infos["birthdate"]
        newinfos =  { **self._infos, 'birthdate': '{}.{}.{}'.format(unb[0:2], unb[2:4], unb[4:8]) }
        labeltext = '{id}, {name}, {surname}, *{birthdate}'.format(**newinfos)
        self._av.patientName.setText(labeltext)
        self._av.setWindowTitle('Archiv Viewer - {}'.format(labeltext))
        
    def setActivePatient(self, infos):
        self.beginResetModel()
        
        self._infos = infos
        unb = infos["birthdate"]
        newinfos =  { **infos, 'birthdate': '{}.{}.{}'.format(unb[0:2], unb[2:4], unb[4:8]) }
            
        self.reloadData(int(infos["id"]))
        self.endResetModel()
        self._av.refreshFiles.setEnabled(True)
        self._dataReloaded.emit()
    
    def data(self, index, role):
        if role == Qt.DisplayRole:
            file = self._files[index.row()]
            col = index.column()
            
            if col == 0:
                return file["datum"].strftime('%d.%m.%Y')
            elif col == 1:
                return file["datum"].strftime('%H:%M')
            elif col == 2:
                try:
                    return self._categoryModel.categoryById(file["category"])['krankenblatt']
                except KeyError:
                    return file["category"]
            elif col == 3:
                return file["beschreibung"]
        elif role == Qt.BackgroundRole:
            file = self._files[index.row()]
            col = index.column()
            if col == 2:
                try:
                    colors = self._categoryModel.colorById(file["category"]) 
                    if colors['red'] is not None:
                        return QBrush(QColor.fromRgb(colors['red'], colors['green'], colors['blue']))
                except KeyError:
                    pass
        elif role == Qt.ToolTipRole:
            col = index.column()
            
            file = self._files[index.row()]
            if col == 2:
                return self._categoryModel.categoryById(file["category"])['name']
            elif col == 3:
                return file["beschreibung"]
        elif role == Qt.TextAlignmentRole:
            if index.column() == 2:
                return Qt.AlignCenter
            
                
    
    def rowCount(self, index):
        rc = len(self._files)
        return rc
        
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
        
        self._unfilteredFiles = []
        
        selectStm = "SELECT a.FSUROGAT, a.FTEXT, a.FEINTRAGSART, a.FZEIT, a.FDATUM FROM ARCHIV a WHERE a.FPATNR = ? AND a.FEINTRAGSART > 0 ORDER BY a.FDATUM DESC, a.FZEIT DESC"
        cur = self._con.cursor()
        cur.execute(selectStm, (patnr,))
        
        for (surogat, beschreibung, eintragsart, zeit, datum) in cur:
            self._unfilteredFiles.append({
                'id': surogat,
                'datum': self._startDate + timedelta(days = datum, seconds = zeit),
                'beschreibung': beschreibung,
                'category': eintragsart
            })
        
        del cur
        self._applyFilters()
        
        self._application.restoreOverrideCursor()
    
    def generateFile(self, file, errorSlot = None):
        filename = self._tmpdir + os.sep + '{}.pdf'.format(file["id"])
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
                                    if errorSlot:
                                        errorSlot.emit(err)
                                    else:
                                        self._av.displayErrorMessage(err)
                            else:
                                err = "%s: Fehler beim Ausführen des Kommandos: '%s'" % (file['beschreibung'], command)
                                if errorSlot:
                                    errorSlot.emit(err)
                                else:
                                    self._av.displayErrorMessage(err)
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
                            errmsg = '\n'.join(errors)
                            if errorSlot:
                                errorSlot.emit(err)
                            else:
                                self._av.displayErrorMessage(err)
                    else:
                        try:
                            if content[0:4] == bytes.fromhex('FFD8FFE0'):
                                img = Image.fromarray(pylibjpeg.decode(content))
                                outbuffer = io.BytesIO()
                                img.save(outbuffer, 'PDF')
                                merger.append(outbuffer)
                                appended = True
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
                                self._av.displayErrorMessage(err)
                
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
                        self._av.displayErrorMessage(err)
                    
                merger.close()
                
                try:
                    datei.close()
                except:
                    pass
                ios.close()
        return filename
    
    def displayFile(self, rowIndex):
        self._application.setOverrideCursor(Qt.WaitCursor)
        file = self._files[rowIndex]
        filename = self.generateFile(file)
        self._application.restoreOverrideCursor()
        if filename is not None:
            subprocess.run(['start', filename], shell=True)
    
    def exportAsPdfThread(self, thread, filelist, destination):       
        try:
            merger = PdfFileMerger()
            counter = 0
            failed = 0
            for file in filelist:
                counter += 1
                thread.progress.emit(counter)
                
                filename = self.generateFile(file, errorSlot = thread.error)
                if filename is None:
                    failed += 1
                else:
                    bmtext = " ".join([file["beschreibung"], file["datum"].strftime('%d.%m.%Y %H:%M')])
                    merger.append(filename, bookmark=bmtext)
            
            if failed < counter:
                merger.write(destination)
            merger.close()
        except IOError as e:
            failed = counter
            thread.error.emit("Fehler beim Schreiben der PDF-Datei: {}".format(e))
            
        thread.progress.emit(0)
        thread.completed.emit(counter, failed, destination)
        
        return counter
    
    def updateProgress(self, value):
        self._av.exportProgress.setFormat('%d von %d' % (value, self._av.exportProgress.maximum()))
        self._av.exportProgress.setValue(value)
        self._av.taskbar_progress.setValue(value)
        
    def exportCompleted(self, counter, failed, destination):
        self._application.restoreOverrideCursor()
        self._av.exportProgress.setFormat('')
        self._av.exportProgress.setEnabled(False)
        self._av.exportPdf.setEnabled(len(self._files)>0)
        self._av.documentView.setEnabled(True)
        self._av.categoryList.setEnabled(True)
        self._av.filterDescription.setEnabled(True)
        self._av.taskbar_progress.hide()
        success = False
        if failed == 0:
            QMessageBox.information(self._av, "Export abgeschlossen", "%d Dokumente wurden nach '%s' exportiert" % (counter, destination))
            success = True
        elif failed < counter:
            message = '\n'.join([ "%d von %d Dokumenten wurden nach '%s' exportiert\n\nWährend des Exports sind Fehler aufgetreten:\n" % (counter-failed, counter, destination), *self._exportErrors ])
            QMessageBox.warning(self._av, "Export abgeschlossen", message)
            success = True
        else:
            message = '\n'.join([ "Es konnten keine Dokumente exportiert werden:", *self._exportErrors ])
            QMessageBox.critical(self._av, "Export fehlgeschlagen", message)
            
        if success and self._config.getValue('showPDFAfterExport', False):
            subprocess.run(['start', destination.replace('/', '\\')], shell=True)
    
    def handleError(self, msg):
        self._exportErrors.append(msg)
    
    def exportAsPdf(self, filelist):   
        if len(filelist) == 0:
            buttonReply = QMessageBox.question(self._av, 'PDF-Export', "Kein Dokument ausgewählt. Export aus allen angezeigten Dokumenten des Patienten erzeugen?", QMessageBox.Yes | QMessageBox.No)
            if(buttonReply == QMessageBox.Yes):
                filelist = range(len(self._files))
            else:
                return
        
        filelist = sorted(filelist)
        files = []
        for f in filelist:
            files.append(self._files[f])
        
        conf = ConfigReader.get_instance()
        outfiledir = conf.getValue('outfiledir', '')
                    
        outfilename = os.sep.join([outfiledir, 'Patientenakte_%d_%s_%s_%s-%s.pdf' % (int(self._infos["id"]), 
            self._infos["name"], self._infos["surname"], self._infos["birthdate"], datetime.now().strftime('%Y%m%d%H%M%S'))])
        destination, _ = QFileDialog.getSaveFileName(self._av, "Auswahl als PDF exportieren", outfilename, "PDF-Datei (*.pdf)")
        if len(destination) > 0:
            try:
                conf.setValue('outfiledir', os.path.dirname(destination))
            except:
                pass
            self._application.setOverrideCursor(Qt.WaitCursor)
            self._exportErrors = []    
            self._av.exportPdf.setEnabled(False)
            self._av.documentView.setEnabled(False)
            self._av.categoryList.setEnabled(False)
            self._av.filterDescription.setEnabled(False)
            self._av.exportProgress.setEnabled(True)
            self._av.exportProgress.setRange(0, len(filelist))
            self._av.exportProgress.setFormat('0 von %d' % (len(filelist)))
            self._av.taskbar_progress.setRange(0, len(filelist))
            self._av.taskbar_progress.setValue(0)
            self._av.taskbar_progress.show()
            self.pdfExporter = PdfExporter(self, files, destination)
            self.exportThread = QThread()
            self.pdfExporter.moveToThread(self.exportThread)
            self.pdfExporter.kill.connect(self.exportThread.quit)
            self.pdfExporter.progress.connect(self.updateProgress)
            self.pdfExporter.completed.connect(self.exportCompleted)
            self.pdfExporter.error.connect(self.handleError)
            self.exportThread.started.connect(self.pdfExporter.work)
            self.exportThread.start()

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
    config = ConfigReader.get_instance()
    if config.getValue('loglevel', 'info') == 'debug':
        logging.getLogger().setLevel(logging.DEBUG)
    qt_translator = QTranslator()
    qt_translator.load("qt_" + QLocale.system().name(),
        QLibraryInfo.location(QLibraryInfo.TranslationsPath))
    app.installTranslator(qt_translator)
    
    qtbase_translator = QTranslator()
    qt_translator.load("qtbase_" + QLocale.system().name(),
        QLibraryInfo.location(QLibraryInfo.TranslationsPath))
    app.installTranslator(qtbase_translator)
    
    try:
        mokey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\WOW6432Node\\INDAMED')
        is64_bits = sys.maxsize > 2**32
        if is64_bits:
            defaultClientLib = os.sep.join([winreg.QueryValueEx(mokey, 'DataPath')[0], '..', 'Firebird', 'bin', 'fbclient.dll'])
        else:
            defaultClientLib = os.sep.join([winreg.QueryValueEx(mokey, 'LocalPath')[0], 'gds32.dll'])
    except OSError as e:
        displayErrorMessage("Failed to open Medical Office registry key: {}".format(e))
        sys.exit()
        
    conffile = os.getcwd() + os.sep + "Patientenakte.cnf"
    conffile2 = os.path.dirname(os.path.realpath(__file__)) + os.sep + "Patientenakte.cnf"
        
    try:
        rstservini = configparser.ConfigParser()
        rstservini.read(os.sep.join([os.environ["SYSTEMROOT"], 'rstserv.ini']))
        defaultHost = rstservini["SYSTEM"]["Computername"]
        defaultDb = os.sep.join([rstservini["MED95"]["DataPath"], "MEDOFF.GDB"])
    except Exception as e:
        displayErrorMessage("Failed to open rstserv.ini: {}".format(e))
        sys.exit()
    
    defaultDbUser = "sysdba"
    defaultDbPassword = "masterkey" 
    defaultLibrePath = "C:\Program Files\LibreOffice\program\soffice.exe"

    for cfg in [conffile2]:
        try:
            LOGGER.debug("Attempting config %s" % (cfg))
            with open(cfg, 'r') as f:
                conf = json.load(f)
                if "dbuser" in conf:
                    defaultDbUser = conf["dbuser"]
                if "dbpassword" in conf:
                    defaultDbPassword = conf["dbpassword"]
                if "libreoffice" in conf:
                    defaultLibrePath = conf["libreoffice"]
                if "clientlib" in conf:
                    defaultClientLib = conf["clientlib"]
            break
        except Exception as e:
            LOGGER.info("Failed to load config: %s." % (e))
    
    LOGGER.debug("Client lib is %s" % (defaultClientLib))
    LOGGER.debug("DB Path is %s on %s" % (defaultDb, defaultHost))
    try:
        LOGGER.debug("Connecting db '{}' at '{}', port 2013 as user {}".format(defaultDb, defaultHost, defaultDbUser))
        con = fdb.connect(host=defaultHost, database=defaultDb, port=2013,
            user=defaultDbUser, password=defaultDbPassword, fb_library_name=defaultClientLib)
        LOGGER.debug("Connection established.")
    except Exception as e:
        displayErrorMessage('Fehler beim Verbinden mit der Datenbank: {}. Pfad zur DLL-Datei: {}'.format(e, defaultClientLib))
        sys.exit()
        
    try:
        cur = con.cursor()
        stm = "SELECT FVARVALUE FROM MED95INI WHERE FCLIENTNAME=? AND FVARNAME='PatexportDatei'"
        cur.execute(stm, (os.environ["COMPUTERNAME"],))
        res = cur.fetchone()
        if res is None:
            raise Exception("Keine Konfiguration für den Namen '{}' hinterlegt!".format(os.environ["COMPUTERNAME"]))
            
        gdtfile = res[0][:-1].decode('windows-1252')
        del cur
        if not os.path.isfile(gdtfile):
            raise Exception("Ungültiger Pfad: '{}'. Bitte korrekten Pfad für Patientenexportdatei im Datenpflegesystem konfigurieren. \
                Es muss in Medical Office anschließend mindestens ein Patient aufgerufen werden, um die Datei zu initialisieren.".format(gdtfile))
    except Exception as e:
        displayErrorMessage("Fehler beim Feststellen des Exportpfades: {}".format(e))
        sys.exit()
        
    with tempdir() as myTemp:
        av = ArchivViewer()
        try:
            cm = CategoryModel(con)
        except Exception as e:
            displayErrorMessage("Fehler beim Laden der Kategorien: {}".format(e))
            sys.exit()
        av.categoryList.setModel(cm)
        tm = ArchivTableModel(con, myTemp, defaultLibrePath, av, app, cm)
        av.documentView.doubleClicked.connect(lambda: tableDoubleClicked(av.documentView, tm))
        av.documentView.setModel(tm)
        av.actionStayOnTop.setChecked(config.getValue('stayontop', False))
        av.actionShowPDFAfterExport.setChecked(config.getValue('showPDFAfterExport', False))
        av.actionUseImg2pdf.setChecked(config.getValue('useImg2pdf', True))
        if config.getValue('stayontop', False):
            av.setWindowFlags(av.windowFlags() | Qt.WindowStaysOnTopHint)
        av.exportPdf.clicked.connect(lambda: exportSelectionAsPdf(av.documentView, tm))
        event_handler = FileChangeHandler(gdtfile, tm)
        av.action_quit.triggered.connect(lambda: app.quit())
        av.action_about.triggered.connect(lambda: QMessageBox.about(av, "Über Archiv Viewer", 
            """<p><b>Archiv Viewer</b> ist eine zur Verwendung mit Medical Office der Fa. Indamed entwickelte
             Software, die synchron zur Medical Office-Anwendung die gespeicherten Dokumente eines Patienten im Archiv
             anzeigen kann. Zusätzlich können ausgewählte Dokumente auch als PDF-Datei exportiert werden.</p>
             <p><a href=\"https://github.com/crispinus2/archivviewer\">https://github.com/crispinus2/archivviewer</a></p>
             <p>(c) 2020 Julian Hartig - Lizensiert unter den Bedingungen der GPLv3</p>"""))
        
        observer = Observer()
        observer.schedule(event_handler, path=os.path.dirname(gdtfile), recursive=False)
        observer.start()
        
        try:
            infos = readGDT(gdtfile)
            tm.activePatientChanged.emit(infos)
        except Exception as e:
            displayErrorMessage("While loading GDT file: %s" % (e))
        
        av.show()
        ret = app.exec_()
        observer.stop()
        if exportThread:
            exportThread.stop()
            exportThread.join()
        observer.join()
    sys.exit(ret)
