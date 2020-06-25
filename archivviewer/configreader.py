# configreader.py

import os, json
from contextlib import contextmanager
import threading
from PyQt5.QtCore import QMutex

class ConfigReader:
    __instance = None

    @staticmethod
    def get_instance():
        if ConfigReader.__instance == None:
            with threading.Lock():
                if ConfigReader.__instance == None:
                    ConfigReader()
        return ConfigReader.__instance

    def __init__(self):
        if ConfigReader.__instance != None:
            raise Exception("This is a singleton class, don't instantiate it directly!")
        else:
            ConfigReader.__instance = self
    
        self.dirconfpath = os.sep.join([os.environ["AppData"], "ArchivViewer", "config.json"])
        self._config = {}
        self._mutex = QMutex(mode=QMutex.Recursive)
        self.readConfig()
    
    @contextmanager
    def lock(self):
        self._mutex.lock()
        yield
        self._mutex.unlock()
    
    def readConfig(self):
        with self.lock():
            try:
                with open(self.dirconfpath, "r") as f:
                    self._config = json.load(f)
            except:
                pass
            
    def getValue(self, valName, default = None):
        with self.lock():
            try:
                return self._config[valName]
            except KeyError:
                return default
            
    def setValue(self, valName, val):
        with self.lock():
            self._config[valName] = val
            self._writeConfig()
        
    def _writeConfig(self):
        os.makedirs(os.path.dirname(self.dirconfpath), exist_ok = True)
        with open(self.dirconfpath, "w") as f:
            json.dump(self._config, f, indent = 1)
            
    def writeConfig(self):
        with self.lock():
            self._writeConfig()