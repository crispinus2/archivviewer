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
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot, QObject, QTranslator, QLocale, QLibraryInfo, QEvent, QSettings, QItemSelectionModel, QItemSelection, QItemSelectionRange
from PyQt5.QtGui import QColor, QBrush, QIcon
from PyQt5.QtWinExtras import QWinTaskbarProgress, QWinTaskbarButton
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from archivviewer.forms import ArchivviewerUi
from .configreader import ConfigReader
from .CategoryModel import CategoryModel
from .PresetModel import PresetModel
from .FilesTableDelegate import FilesTableDelegate
from .GenerateFileWorker import GenerateFileWorker

logging.basicConfig(level=logging.INFO)
LOGGER = logging.getLogger(__name__)

def displayErrorMessage(msg):
    QMessageBox.critical(None, "Fehler", str(msg))

class ArchivViewer(QMainWindow, ArchivviewerUi):
    def __init__(self, con, parent = None):
        super(ArchivViewer, self).__init__(parent)
        self._config = ConfigReader.get_instance()
        self._con = con
        self.taskbar_button = None
        self.taskbar_progress = None
        self.setupUi(self)
        iconpath = os.sep.join([os.path.dirname(os.path.realpath(__file__)), 'icon128.png'])
        self.setWindowIcon(QIcon(iconpath))
        self.setWindowTitle("Archiv Viewer")
        self.refreshFiles.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self.exportPdf.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.savePreset.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.savePreset.setText('')
        self.clearPreset.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        self.clearPreset.setText('')
        self.readSettings()
        self.delegate = FilesTableDelegate()
        self.documentView.setItemDelegate(self.delegate)
        self.actionStayOnTop.changed.connect(self.stayOnTopChanged)
        self.actionShowPDFAfterExport.changed.connect(self.showPDFAfterExportChanged)
        self.actionUseImg2pdf.changed.connect(self.useImg2pdfChanged)
        self.actionUseGimpForTiff.changed.connect(self.useGimpForTiffChanged)
        self.actionOptimizeExport.changed.connect(self.optimizeExportChanged)
        self.actionFitToA4.changed.connect(self.fitToA4Changed)
        self.presetModel = PresetModel(self.categoryList)
        self.presets.setModel(self.presetModel)
        self.presets.editTextChanged.connect(self.presetsEditTextChanged)
        self.clearPreset.clicked.connect(self.clearPresetClicked)
        self.savePreset.clicked.connect(self.savePresetClicked)
        self.presets.currentIndexChanged.connect(self.presetsIndexChanged)
        try:
            self.categoryListModel = CategoryModel(self._con)
        except Exception as e:
            displayErrorMessage("Fehler beim Laden der Kategorien: {}".format(e))
            sys.exit()
        self.categoryList.setModel(self.categoryListModel)
        self.categoryListModel.dataChanged.connect(self.categoryListModelDataChanged)
        
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
    
    def clearPresetClicked(self):
        buttonReply = QMessageBox.question(self, 'Voreinstellung löschen', "Soll die aktive Voreinstellung wirklich gelöscht werden?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if(buttonReply == QMessageBox.Yes):
            curidx = self.presets.findText(self.presets.currentText())
            self.presets.removeItem(curidx)
            self.presets.setCurrentIndex(0)
        else:
            return
        
    def savePresetClicked(self):
        curidx = self.presets.findText(self.presets.currentText())
        if curidx != 0:
            ids = [ self.categoryList.model().idAtRow(idx.row()) for idx in self.categoryList.selectionModel().selectedRows() ]
            
            if curidx > 0:
                self.presetModel.setSelectedCategories(curidx, ids)
            else:
                curidx = self.presetModel.insertPreset(self.presets.currentText(), ids)
            self.presets.setCurrentIndex(curidx)
    
    def presetsEditTextChanged(self, text):
        idx = self.presets.findText(text)
        self.savePreset.setEnabled(idx != 0 and len(text)>0)
        self.clearPreset.setEnabled(idx > 0)
    
    def presetsIndexChanged(self, idx):
        self.savePreset.setEnabled(idx > 0)
        self.clearPreset.setEnabled(idx > 0)
        categories = self.presetModel.categoriesAtIndex(idx)
        selm = self.categoryList.selectionModel()
        cmodel = self.categoryList.model()
        catidxs = [ cmodel.createIndex(row, 0, None) for row in range(cmodel.rowCount(None)) ]
        
        qis = QItemSelection()
        for cidx in catidxs:
            if cmodel.idAtRow(cidx.row()) in categories:
                qir = QItemSelectionRange(cidx)
                qis.append(qir)
        selm.select(qis, QItemSelectionModel.ClearAndSelect)
    
    def categoryListModelDataChanged(self, begin, end):
        self.presetsIndexChanged(self.presetList.currentIndex())
    
    def showPDFAfterExportChanged(self):
        show = self.actionShowPDFAfterExport.isChecked()
        self._config.setValue('showPDFAfterExport', show)
    
    def useImg2pdfChanged(self):
        use = self.actionUseImg2pdf.isChecked()
        self._config.setValue('useImg2pdf', use)
        
    def optimizeExportChanged(self):
        optimize = self.actionOptimizeExport.isChecked()
        self._config.setValue('shrinkPDF', optimize)
    
    def useGimpForTiffChanged(self):
        use = self.actionUseGimpForTiff.isChecked()
        self._config.setValue('useGimpForTiff', use)
    
    def fitToA4Changed(self):
        self._config.setValue('fitToA4', self.actionFitToA4.isChecked())
    
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

    def __init__(self, con, arccon, tmpdir, librepath, mainwindow, application, gimppath, gspath):
        super(ArchivTableModel, self).__init__()
        self._unfilteredFiles = []
        self._files = []
        self._con = con
        self._arccon = arccon
        self._tmpdir = tmpdir
        self._librepath = librepath
        self._table = mainwindow.documentView
        self._av = mainwindow
        self._application = application
        self._infos = {}
        self._generateFileWorker = None
        self._categoryModel = self._av.categoryListModel
        self._gimppath = gimppath
        self._categoryFilter = set()
        self._av.categoryList.selectionModel().selectionChanged.connect(self.categorySelectionChanged)
        self._av.filterDescription.textEdited.connect(self.filterTextChanged)
        self._config = ConfigReader.get_instance()
        self._dataReloaded.connect(self._av.filterDescription.clear)
        self._dataReloaded.connect(self.updateLabel)
        self._av.refreshFiles.clicked.connect(lambda: self.activePatientChanged.emit(self._infos))
        self._av.actionShowRemovedItems.changed.connect(self.showRemovedItemsChanged)
        self._av.cancelExport.hide()
        self._av.cancelExport.clicked.connect(self.cancelExport)
        self._gspath = gspath
        self.activePatientChanged.connect(self.setActivePatient)
        self.generateFileThread = None      
    
    def __del__(self):
        if self.generateFileThread:
            self.generateFileThread.wait()
        
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
            
            if self._arccon is not None:
                selectStm = "SELECT a.FSUROGAT, a.FTEXT, a.FEINTRAGSART, a.FZEIT, a.FDATUM FROM ARCHIV a WHERE a.FPATNR = ? AND a.FEINTRAGSART > 0 ORDER BY a.FDATUM DESC, a.FZEIT DESC"
                cur = self._arccon.cursor()
                cur.execute(selectStm, (patnr,))
                
                added = False
                for (surogat, beschreibung, eintragsart, zeit, datum) in cur:
                    add = True
                    if not self._config.getValue("showRemovedItems", False):
                        select2 = "SELECT COUNT(*) FROM LTAG l WHERE l.FEINTRAGSNR = ? AND l.FEINTRAGSART = ?"
                        cur2 = self._con.cursor()
                        cur2.execute(select2, (surogat, eintragsart))
                        if cur2.fetchone()[0] == 0:
                            add = False
                    
                    if add:
                        added = True
                        self._unfilteredFiles.append({
                            'id': surogat,
                            'datum': self._startDate + timedelta(days = datum, seconds = zeit),
                            'beschreibung': beschreibung,
                            'category': eintragsart,
                            'medoffarc': True
                        })
            
            self._unfilteredFiles.sort(key = lambda x: x['datum'], reverse = True)
            
            self._applyFilters()
            
            self._application.restoreOverrideCursor()
        
    def displayFile(self, rowIndex):
        self.exportAsPdf([ rowIndex ], False)
          
    def generateFileStarted(self, maxFiles, isExport):
        self._av.exportFileProgress.setRange(0, maxFiles)
        self._av.exportFileProgress.setValue(0)
        self._av.exportFileProgress.setEnabled(True)
        if not isExport:
            self._av.taskbar_progress.setRange(0, maxFiles)
            self._av.taskbar_progress.setValue(0)
            self._av.taskbar_progress.show()
    
    def generateFileProgress(self, isExport, value = None):
        if value is None:
            value = self._av.exportFileProgress.value() + 1
        self._av.exportFileProgress.setValue(value)
        self._av.exportFileProgress.setFormat('%d von %d Unterdokumenten' % (value, self._av.exportFileProgress.maximum()))
        if not isExport:
            self._av.taskbar_progress.setValue(value)
    
    def generateFileComplete(self, filename, fileinfo, errors, isExport):
        self._av.exportFileProgress.setFormat('')
        self._av.exportFileProgress.setValue(0)
        self._av.exportFileProgress.setEnabled(False)
    
    def exportStarted(self, maxDocuments):
        self._av.exportFileProgress.setRange(0, maxFiles)
        self._av.exportFileProgress.setValue(0)
        self._av.exportFileProgress.setEnabled(True)
        if self._exportMerger is None:
            self._av.taskbar_progress.setRange(0, maxFiles)
            self._av.taskbar_progress.setValue(0)
            self._av.taskbar_progress.show()
            
    def updateProgress(self, value = None):
        if value is None:
            value = self._av.exportProgress.value() + 1
        self._av.exportProgress.setFormat('%d von %d Dokumenten' % (value, self._av.exportProgress.maximum()))
        self._av.exportProgress.setValue(value)
        self._av.taskbar_progress.setValue(value)
    
    def exportProgressStatus(self, message):
        self._av.exportProgress.setFormat(message)
    
    def fileProgressStatus(self, message):
        self._av.exportFileProgress.setFormat(message)
    
    def exportCancelled(self):
        self._application.restoreOverrideCursor()
        self._av.exportProgress.setFormat('')
        self._av.exportProgress.setValue(0)
        self._av.exportProgress.setEnabled(False)
        self._av.exportPdf.show()
        self._av.cancelExport.hide()
        self._av.exportPdf.setEnabled(len(self._files)>0)
        self._av.documentView.setEnabled(True)
        self._av.categoryList.setEnabled(True)
        self._av.filterDescription.setEnabled(True)
        self._av.refreshFiles.setEnabled(True)
        self._av.groupBox.setEnabled(True)
        self._av.taskbar_progress.hide()
        self.generateFileThread.quit()
        self.generateFileThread.wait()
        self.generateFileThread = None
        self._generateFileWorker = None
    
    def exportCompleted(self, filename, counter, failed, errors, isExport):
        destination = filename
        
        self.exportCancelled()
        
        success = False
        if isExport:
            if failed == 0:
                QMessageBox.information(self._av, "Export abgeschlossen", "%d Dokumente wurden nach '%s' exportiert" % (counter, destination))
                success = True
            elif failed < counter:
                message = '\n'.join([ "%d von %d Dokumenten wurden nach '%s' exportiert\n\nWährend des Exports sind Fehler aufgetreten:\n" % (counter-failed, counter, destination), *errors ])
                QMessageBox.warning(self._av, "Export abgeschlossen", message)
                success = True
            else:
                message = '\n'.join([ "Es konnten keine Dokumente exportiert werden:", *errors ])
                QMessageBox.critical(self._av, "Export fehlgeschlagen", message)
        else:
            if filename is None or failed == 1:
                message = '\n'.join([ "Das Konvertieren in PDF ist fehlgeschlagen:", *errors ])
                QMessageBox.critical(self._av, "Konvertierung fehlgeschlagen", message)
            else:
                success = True
        
        if success and (self._config.getValue('showPDFAfterExport', False) or not isExport):
            subprocess.run(['start', destination.replace('/', '\\')], shell=True)
    
    def cancelExport(self):
        self._av.cancelExport.setEnabled(False)
        if self._generateFileWorker is not None:
            self._generateFileWorker.cancel()
    
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
        
        destination = None
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
            self._av.exportPdf.hide()
            self._av.cancelExport.show()
            self._av.cancelExport.setEnabled(True)
            self._av.documentView.setEnabled(False)
            self._av.categoryList.setEnabled(False)
            self._av.filterDescription.setEnabled(False)
            self._av.refreshFiles.setEnabled(False)
            self._av.exportProgress.setEnabled(True)
            self._av.exportFileProgress.setEnabled(True)
            self._av.groupBox.setEnabled(False)
            
            self._generateFileWorker = GenerateFileWorker(self._tmpdir, files, self._con, self._arccon, self._librepath, self._gimppath, self._gspath, exportDestination = destination)
            self.generateFileThread = QThread()
            self._generateFileWorker.moveToThread(self.generateFileThread)
            self._generateFileWorker.kill.connect(self.generateFileThread.quit)
            self._generateFileWorker.progress.connect(self.generateFileProgress)
            self._generateFileWorker.completed.connect(self.generateFileComplete)
            self._generateFileWorker.initGenerate.connect(self.generateFileStarted)
            self._generateFileWorker.initExport.connect(self.exportStarted)
            self._generateFileWorker.progressExport.connect(self.updateProgress)
            self._generateFileWorker.exportCompleted.connect(self.exportCompleted)
            self._generateFileWorker.exportProgressStatus.connect(self.exportProgressStatus)
            self._generateFileWorker.fileProgressStatus.connect(self.fileProgressStatus)
            self._generateFileWorker.exportCancelled.connect(self.exportCancelled)
            self.generateFileThread.started.connect(self._generateFileWorker.work)
            self.generateFileThread.start()

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
        
    conffile2 = os.path.dirname(os.path.realpath(__file__)) + os.sep + "Patientenakte.cnf"
        
    try:
        rstservini = configparser.ConfigParser()
        rstservini.read(os.sep.join([os.environ["SYSTEMROOT"], 'rstserv.ini']))
        defaultHost = rstservini["SYSTEM"]["Computername"]
        defaultDb = os.sep.join([rstservini["MED95"]["DataPath"], "MEDOFF.GDB"])
        defaultArcDb = os.sep.join([rstservini["MED95"]["DataPath"], "MEDOFFARC.GDB"])
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

    gspath = None
    try:
        try:
            gskey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Artifex\\GPL Ghostscript', 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        except OSError as e:
            gskey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Artifex\\GPL Ghostscript', 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        try:
            gsverkey = winreg.OpenKey(gskey, winreg.EnumKey(gskey, 0))
            gsbasepath = os.path.join(winreg.QueryValueEx(gsverkey, '')[0], 'bin')
            gspath = os.path.join(gsbasepath, 'gswin64c.exe')
            if not os.path.exists(gspath):
                gspath = os.path.join(gsbasepath, 'gswin32c.exe')
                if not os.path.exists(gspath):
                    raise Exception("No suitable ghostscript executable found in {}".format(gsbasepath))
                
            LOGGER.debug("Ghostscript version {} found in '{}'".format(winreg.EnumKey(gskey, 0), gspath))
        except Exception as e:
            LOGGER.debug("No usable ghostscript installation found. Continuing without shrink option: {}".format(e))
    except OSError as e:
        LOGGER.debug('Failed to open ghostscript key: {}'.format(e))

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
                if "ghostscript" in conf:
                    gspath = conf["ghostscript"]
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
        LOGGER.debug("Connecting archive db '{}' at '{}', port 2013 as user {}".format(defaultArcDb, defaultHost, defaultDbUser))
        arccon = fdb.connect(host=defaultHost, database=defaultArcDb, port=2013,
            user=defaultDbUser, password=defaultDbPassword, fb_library_name=defaultClientLib)
        LOGGER.debug("Connection established.")
    except Exception as e:
        displayErrorMessage('Fehler beim Verbinden mit der Archivdatenbank: {}. Dieser Fehler kann bedenkenlos ignoriert werden, falls keine Archivdatenbank verfügbar ist (z.B. auf Mobilsystemen).'.format(e))
        arccon = None    
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
        av = ArchivViewer(con)        
        tm = ArchivTableModel(con, arccon, myTemp, defaultLibrePath, av, app, gimppath, gspath)
        av.documentView.doubleClicked.connect(lambda: tableDoubleClicked(av.documentView, tm))
        av.documentView.setModel(tm)
        av.actionStayOnTop.setChecked(config.getValue('stayontop', False))
        av.actionShowPDFAfterExport.setChecked(config.getValue('showPDFAfterExport', False))
        av.actionShowRemovedItems.setChecked(config.getValue("showRemovedItems", False))
        av.actionUseImg2pdf.setChecked(config.getValue('useImg2pdf', True))
        av.actionOptimizeExport.setEnabled(gspath is not None)
        if av.actionOptimizeExport.isEnabled():
            av.actionOptimizeExport.setChecked(config.getValue('shrinkPDF', True))
        av.actionUseGimpForTiff.setEnabled(gimppath is not None)
        av.actionFitToA4.setChecked(config.getValue('fitToA4', False))
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
        observer.join()
    sys.exit(ret)
