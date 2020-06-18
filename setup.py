# setup.py

from setuptools import setup, find_packages

try:
    from pyqt_distutils.build_ui import build_ui
    cmdclass = {'build_ui': build_ui}
except ImportError:
    build_ui = None  # user won't have pyqt_distutils when deploying
    cmdclass = {}

setup(
    name='ArchivViewer',
    version='0.1.0',
    author='Julian Hartig',
    author_email='julian@whisper-net.de',
    packages=['archivviewer', 'archivviewer.forms'],
    url='http://pypi.python.org/pypi/archivviewer/',
    license='GPLv3',
    description='Archive viewer for use with Medical Office AIS by Indamed',
    install_requires=[
        "lhafile >= 0.2.2",
        "PyPDF2 >= 1.26.0",
        "img2pdf >= 0.3.6",
        "PyQt5 >= 5.15.0",
        "watchdog >= 0.10.2",
        "pyqt-distutils"
    ],
    entry_points = {
        "console_scripts": ['archivviewer = archivviewer.archivviewer:main']
    },
    cmdclass=cmdclass
)