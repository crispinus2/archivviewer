# FilesTableDelegate.py

from PyQt5 import QtWidgets, QtGui, QtCore

class FilesTableDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self):
        QtWidgets.QStyledItemDelegate.__init__(self)
        
    def paint(self, painter, option, index):
        if index.column() == 2 and option.backgroundBrush != QtCore.Qt.NoBrush:
            painter.save()
            painter.setPen(QtCore.Qt.NoPen)
            painter.setBrush(option.backgroundBrush)
            painter.drawRoundedRect(option.rect.left(), option.rect.top(), option.rect.width()-0.5, option.rect.height()-0.5, 10, 10)
            painter.restore()
            adj_option = option
            adj_option.backgroundBrush = QtGui.QBrush(QtCore.Qt.NoBrush)
        else:
            adj_option = option
            
        QtWidgets.QStyledItemDelegate.paint(self, painter, adj_option, index)
        
