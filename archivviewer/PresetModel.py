# CategoryModel.py

import io, struct, os, sys, logging
from PyQt5.QtCore import QAbstractItemModel, QModelIndex, Qt
from collections import OrderedDict
from itertools import islice
from bisect import bisect_left
from .configreader import ConfigReader

LOGGER = logging.getLogger(__name__)

class PresetModel(QAbstractItemModel):
    def __init__(self, categoryList):
        super(PresetModel, self).__init__()
        self._categoryList = categoryList
        self._config = ConfigReader.get_instance()
        self._presets = [ {'name': '(alle)', 'categories': [] }, *self._config.getValue("categoryPresets", []) ]
        self._presets.sort(key = lambda x: x["name"]) 
        self._presetNames = [ k["name"] for k in self._presets ]

    def rowCount(self, _):
        return len(self._presets)
    
    def columnCount(self, _):
        return 1
    
    def categoriesAtIndex(self, idx):
        return self._presets[idx]['categories']
    
    def setSelectedCategories(self, idx, categories):
        self._presets[idx]['categories'] = categories
        self._savePresets()
    
    def _savePresets(self):
        self._config.setValue("categoryPresets", tuple(filter(lambda x: len(x["name"]) > 0, self._presets[1:])))
    
    def index(self, row, column, parent):
        if row < self.rowCount(parent) and column == 0:
            return self.createIndex(row, column)
        else:
            return QModelIndex(parent)
    
    def insertRows(self, row, count, parent):
        self.beginInsertRows(parent, row, row+count-1)
        selcats = [ self._categoryList.model().idAtRow(idx.row()) for idx in self._categoryList.selectedIndexes() ]
            
        insert_elem = [ { 'name': '(unbenannt)', 'categories': selcats } for _ in range(count) ]
        if row == len(self._presets):
            self._presets.extend(insert_elem)
        else:
            self._presets[row:row] = insert_elem
        self.endInsertRows()
        
        return True
    
    def insertPreset(self, name, ids):
        pos = bisect_left(self._presetNames, name)
        self._presetNames.insert(pos, name)
        self._presets.insert(pos, { 'name': name, 'categories': ids })
        
        return pos
    
    def data(self, index, role):
        if role == Qt.DisplayRole or role == Qt.EditRole:
            name = self._presets[index.row()]["name"]
            if name is None:
                return "(Unbenannt)"
            return name
            
        return None

    def removeRows(self, row, count, parent):
        self.beginRemoveRows(parent, row, row+count-1)
        del self._presets[row:row+count]
        self._savePresets()
        self.endRemoveRows()
        
        return True
        
    def setData(self, index, value, role):
        self._presets[index.row()]["name"] = value
        self._savePresets()
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        
        return True
        
    def parent(self, child):
        return QModelIndex()
