pyinstaller -w --add-binary archivviewer/icon128.png;archivviewer --add-data LICENSE.txt;. --icon resource/icon.ico --hidden-import libjpeg --hidden-import pylibjpeg ArchivViewer.py