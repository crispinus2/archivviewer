# FilesTableDelegate.py

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtGui import QPalette, QBrush
from PyQt5.QtWidgets import QStyle

class FilesTableDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self):
        QtWidgets.QStyledItemDelegate.__init__(self)
    
    def initStyleOption(self, option, index, myPainter = False):
        QtWidgets.QStyledItemDelegate.initStyleOption(self, option, index)
        
        if not myPainter and index.column() == 2:
            normalText = option.palette.brush(QPalette.ColorGroup.Normal, QPalette.ColorRole.Text)
            option.palette.setBrush(QPalette.ColorGroup.Normal, QPalette.ColorRole.Highlight, QBrush(QtCore.Qt.NoBrush))
            if option.backgroundBrush.style() != QtCore.Qt.NoBrush:
                option.palette.setBrush(QPalette.ColorGroup.Normal, QPalette.ColorRole.HighlightedText, normalText)
            option.backgroundBrush = QtGui.QBrush(QtCore.Qt.NoBrush)
            
    
    def paint(self, painter, option, index):
        adj_option = option
        if index.column() == 2:
            adj_option.backgroundBrush = QtGui.QBrush(QtCore.Qt.NoBrush)
        
        self.initStyleOption(option, index, True)
        if index.column() == 2: # and option.backgroundBrush != QtCore.Qt.NoBrush:
            painter.save()
            painter.setPen(QtCore.Qt.NoPen)
            if option.state & QStyle.State_Selected:
                painter.fillRect(option.rect, option.palette.brush(QPalette.ColorGroup.Normal, QPalette.ColorRole.Highlight))
            brush = index.data(QtCore.Qt.BackgroundRole)
            if brush is None:
                brush = QBrush(QtCore.Qt.NoBrush)
            painter.setBrush(brush)
            painter.drawRoundedRect(option.rect.left(), option.rect.top(), option.rect.width()-0.5, option.rect.height()-0.5, 10, 10)
            painter.restore()
            
        QtWidgets.QStyledItemDelegate.paint(self, painter, adj_option, index)
            
            
        
        
