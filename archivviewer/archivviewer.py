# Archivviewer.py

import sys, codecs, os, fdb, json, tempfile, shutil, subprocess, io, winreg, configparser, email, logging
from subprocess import PIPE
from datetime import datetime, timedelta
from collections import OrderedDict
import collections
from contextlib import contextmanager
from pathlib import Path
from PyPDF2 import PdfFileMerger
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
from .GenerateFileWorker import GenerateFileWorker

exportThread = None

logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger(__name__)

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
        self.actionUseGimpForTiff.changed.connect(self.useGimpForTiffChanged)
    
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
    
    def useGimpForTiffChanged(self):
        use = self.actionUseGimpForTiff.isChecked()
        self._config.setValue('useGimpForTiff', use)
    
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

class ArchivTableModel(QAbstractTableModel):    
    _startDate = datetime(1890, 1, 1)
    _dataReloaded = pyqtSignal()
    activePatientChanged = pyqtSignal(dict)

    def __init__(self, con, tmpdir, librepath, mainwindow, application, categoryModel, gimppath):
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
        self._exportMerger = None
        self._exportFiles = []
        self._exportDestination = None
        self._exportFailed = 0
        self._exportSuccessful = 0
        self._generateFileWorker = None
        self._categoryModel = categoryModel
        self._gimppath = gimppath
        self._categoryFilter = set()
        self._av.categoryList.selectionModel().selectionChanged.connect(self.categorySelectionChanged)
        self._av.filterDescription.textEdited.connect(self.filterTextChanged)
        self._config = ConfigReader.get_instance()
        self._dataReloaded.connect(self._av.filterDescription.clear)
        self._dataReloaded.connect(self.updateLabel)
        self._av.refreshFiles.clicked.connect(lambda: self.activePatientChanged.emit(self._infos))
        self._av.actionShowRemovedItems.changed.connect(self.showRemovedItemsChanged)
        self.activePatientChanged.connect(self.setActivePatient)
        self.exportThread = None
        self.generateFileThread = None
        self.mutex = QMutex(mode=QMutex.Recursive)
        
    
    def __del__(self):
        if self.exportThread:
            self.exportThread.wait()
        if self.generateFileThread:
            self.generateFileThread.wait()
    
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
    
    def showRemovedItemsChanged(self):
        self._config.setValue("showRemovedItems", self._av.actionShowRemovedItems.isChecked())
        self.beginResetModel()
        self.reloadData()
        self.endResetModel()
    
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
            
        self.reloadData()
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
                try:
                    return self._categoryModel.categoryById(file["category"])['name']
                except KeyError:
                    return "(unbekannte Kategorie)"
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
    
    def reloadData(self):
        patnr = None
        try:
            patnr = int(self._infos["id"])
        except KeyError:
            pass
        
        if patnr is not None:
            self._application.setOverrideCursor(Qt.WaitCursor)
        
            self._unfilteredFiles = []
            
            
            
            stmParts = [ "SELECT a.FSUROGAT, a.FTEXT, a.FEINTRAGSART, a.FZEIT, a.FDATUM FROM ARCHIV a WHERE" ]
            if not self._config.getValue("showRemovedItems", False):
                stmParts.append("EXISTS (SELECT 1 FROM LTAG l WHERE a.FSUROGAT = l.FEINTRAGSNR AND a.FEINTRAGSART = l.FEINTRAGSART) AND")
            stmParts.append("a.FPATNR = ? AND a.FEINTRAGSART > 0 ORDER BY a.FDATUM DESC, a.FZEIT DESC")
            
            selectStm = ' '.join(stmParts)
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
        
    def displayFile(self, rowIndex):
        self.exportAsPdf([ rowIndex ], False)
        
    def exportNextFile(self):       
        if self.generateFileThread is not None:
            self.generateFileThread.quit()
            self.generateFileThread.wait()
        try:
            fil = self._exportFiles.popleft()
            self._exportSuccessful += 1
            if self._exportMerger is not None:
                self.updateProgress(self._exportSuccessful)
            self._generateFileWorker = GenerateFileWorker(self._tmpdir, fil, self._con, self._librepath, self._gimppath)
            self.generateFileThread = QThread()
            self._generateFileWorker.moveToThread(self.generateFileThread)
            self._generateFileWorker.kill.connect(self.generateFileThread.quit)
            self._generateFileWorker.progress.connect(self.generateFileProgress)
            self._generateFileWorker.completed.connect(self.generateFileComplete)
            self._generateFileWorker.errors.connect(self.handleError)
            self._generateFileWorker.initGenerate.connect(self.generateFileStarted)
            self.generateFileThread.started.connect(self._generateFileWorker.work)
            self.generateFileThread.start()
        except IndexError:
            self.generateFileThread = None
            self._generateFileWorker = None
            
            if self._exportMerger is not None:
                try:
                    self._av.exportProgress.setFormat('Schreibe Exportdatei...')
                    if self._exportFailed < self._exportSuccessful:
                        self._exportMerger.write(self._exportDestination)
                    self._exportMerger.close()
                except IOError as e:
                    self._exportFailed = self._exportSuccessful
                    self.handleError("Fehler beim Schreiben der PDF-Datei: {}".format(e))
            
            self._application.restoreOverrideCursor()
            self._av.exportProgress.setFormat('')
            self._av.exportProgress.setValue(0)
            self._av.exportProgress.setEnabled(False)
            self._av.exportPdf.setEnabled(len(self._files)>0)
            self._av.documentView.setEnabled(True)
            self._av.categoryList.setEnabled(True)
            self._av.filterDescription.setEnabled(True)
            self._av.refreshFiles.setEnabled(True)
            self._av.taskbar_progress.hide()
            
            if self._exportMerger is not None:
                self._exportMerger = None
                self.exportCompleted()
        
    def generateFileStarted(self, maxFiles):
        self._av.exportFileProgress.setRange(0, maxFiles)
        self._av.exportFileProgress.setValue(0)
        self._av.exportFileProgress.setEnabled(True)
        if self._exportMerger is None:
            self._av.taskbar_progress.setRange(0, maxFiles)
            self._av.taskbar_progress.setValue(0)
            self._av.taskbar_progress.show()
    
    def generateFileProgress(self, value = None):
        if value is None:
            value = self._av.exportFileProgress.value() + 1
        self._av.exportFileProgress.setValue(value)
        self._av.exportFileProgress.setFormat('%d von %d Unterdokumenten' % (value, self._av.exportFileProgress.maximum()))
        if self._exportMerger is None:
            self._av.taskbar_progress.setValue(value)
    
    def generateFileComplete(self, filename, fileinfo):
        if self._exportMerger is not None:
            if filename is None:
                self._exportFailed += 1
            else:
                bmtext = " ".join([fileinfo["beschreibung"], fileinfo["datum"].strftime('%d.%m.%Y %H:%M')])
                self._exportMerger.append(filename, bookmark=bmtext)
        
        self._av.exportFileProgress.setFormat('')
        self._av.exportFileProgress.setValue(0)
        self._av.exportFileProgress.setEnabled(False)
        
        if self._exportMerger is None:
            if filename is not None:
                subprocess.run(['start', filename], shell=True)
            else:
                message = '\n'.join([ "Das Konvertieren in PDF ist fehlgeschlagen:", *self._exportErrors ])
                QMessageBox.critical(self._av, "Export fehlgeschlagen", message)
                
        self.exportNextFile()
    
    def updateProgress(self, value):
        self._av.exportProgress.setFormat('%d von %d Dokumenten' % (value, self._av.exportProgress.maximum()))
        self._av.exportProgress.setValue(value)
        self._av.taskbar_progress.setValue(value)
        
    def exportCompleted(self):
        counter = self._exportSuccessful
        failed = self._exportFailed
        destination = self._exportDestination
        self._exportDestination = None
        self._exportSuccessful = 0
        self._exportFailed = 0
        
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
    
    def exportAsPdf(self, filelist, doExport = True):   
        if len(filelist) == 0 and doExport:
            buttonReply = QMessageBox.question(self._av, 'PDF-Export', "Kein Dokument ausgewählt. Export aus allen angezeigten Dokumenten des Patienten erzeugen?", QMessageBox.Yes | QMessageBox.No)
            if(buttonReply == QMessageBox.Yes):
                filelist = range(len(self._files))
            else:
                return
        
        filelist = sorted(filelist)
        files = []
        for f in filelist:
            files.append(self._files[f])
        
        if doExport:
            conf = ConfigReader.get_instance()
            outfiledir = conf.getValue('outfiledir', '')
                        
            outfilename = os.sep.join([outfiledir, 'Patientenakte_%d_%s_%s_%s-%s.pdf' % (int(self._infos["id"]), 
                self._infos["name"], self._infos["surname"], self._infos["birthdate"], datetime.now().strftime('%Y%m%d%H%M%S'))])
            destination, _ = QFileDialog.getSaveFileName(self._av, "Auswahl als PDF exportieren", outfilename, "PDF-Datei (*.pdf)")
        if not doExport or len(destination) > 0:
            if doExport:
                try:
                    conf.setValue('outfiledir', os.path.dirname(destination))
                except:
                    pass
                self._exportDestination = destination
                self._exportMerger = PdfFileMerger()
                self._av.exportProgress.setEnabled(True)
                self._av.exportProgress.setRange(0, len(filelist))
                self._av.exportProgress.setFormat('0 von %d Dokumenten' % (len(filelist)))
                self._av.taskbar_progress.setRange(0, len(filelist))
                self._av.taskbar_progress.setValue(0)
                self._av.taskbar_progress.show()
            self._application.setOverrideCursor(Qt.WaitCursor)
            self._exportErrors = []  
            self._exportFiles = collections.deque(files)
            self._exportFailed = 0
            self._exportSuccessful = 0
            self._av.exportPdf.setEnabled(False)
            self._av.documentView.setEnabled(False)
            self._av.categoryList.setEnabled(False)
            self._av.filterDescription.setEnabled(False)
            self._av.refreshFiles.setEnabled(False)
            self.exportNextFile()

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
    try:
        sokey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\soffice.exe')
        defaultLibrePath = winreg.QueryValueEx(sokey, '')[0]
        LOGGER.debug("LibreOffice soffice.exe found in '{}'".format(defaultLibrePath))
    except OSError as e:
        LOGGER.debug('Failed to open soffice.exe-Key: {}'.format(e))
        defaultLibrePath = None
        
    gimppath = None
    try:
        try:
            gimpkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\GIMP-2_is1', 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        except OSError as e:
            gimpkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\GIMP-2_is1', 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        gimppath = winreg.QueryValueEx(gimpkey, 'DisplayIcon')[0]
        LOGGER.debug("GIMP executable found in '{}'".format(gimppath))
    except OSError as e:
        LOGGER.debug('Failed to open GIMP-Key: {}'.format(e))
        gimppath = None

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
                if "gimppath" in conf:
                    gimppath = conf["gimppath"]
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
        tm = ArchivTableModel(con, myTemp, defaultLibrePath, av, app, cm, gimppath)
        av.documentView.doubleClicked.connect(lambda: tableDoubleClicked(av.documentView, tm))
        av.documentView.setModel(tm)
        av.actionStayOnTop.setChecked(config.getValue('stayontop', False))
        av.actionShowPDFAfterExport.setChecked(config.getValue('showPDFAfterExport', False))
        av.actionShowRemovedItems.setChecked(config.getValue("showRemovedItems", False))
        av.actionUseImg2pdf.setChecked(config.getValue('useImg2pdf', True))
        av.actionUseGimpForTiff.setEnabled(gimppath is not None)
        if av.actionUseGimpForTiff.isEnabled():
            av.actionUseGimpForTiff.setChecked(config.getValue('useGimpForTiff', False))
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
