# CategoryModel.py

import io, struct, os, sys, logging
from PyQt5.QtCore import QAbstractItemModel, Qt
from collections import OrderedDict
from itertools import islice
from .configreader import ConfigReader

LOGGER = logging.getLogger(__name__)

class PresetModel(QAbstractItemModel):
    def __init__(self, categoryList):
        super(PresetModel, self).__init__()
        self._categoryList = categoryList
        self._config = ConfigReader.get_instance()
        self._presets = self._config.getValue("categoryPresets", [])
        self._presets.sort(key = lambda x: x["name"])
        self._archivecategories = {}
        self.reloadCategories()

    def rowCount(self, _):
        return len(self._presets)
    
    def columnCount(self, _):
        return 1
    
    def insertRows(int row, int count, _):
        self.beginInsertRows()
        selcats = [ self._categoryList().model().idAtRow(idx.row()) for idx in self._categoryList.selectedIndexes() ]
            
        insert_elem = [ { 'name': None, 'categories': selcats } for _ in range(count) ]
        if row == len(self._presets):
            self._presets.extend(insert_elem)
        else:
            self._presets[row:row] = insert_elem
        self.endInsertRows()
    
    def data(self, index, role):
        if role == Qt.DisplayRole:
            name = self._presets[index.row()]["name"]
            if name is None:
                return "(Unbenannt)"
            return name
        
    def setData(self, index, value):
        self._presets[index.row()]["name"] = value
        self._config.setValue("categoryPresets", self._presets)
        self.dataChanged.emit(index, index, [Qt.EditRole, Qt.DisplayRole])
