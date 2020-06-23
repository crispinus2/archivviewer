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
    version='03',
    author='Julian Hartig',
    author_email='julian@whisper-net.de',
    packages=['archivviewer', 'archivviewer.forms'],
    url='https://github.com/crispinus2/archivviewer/',
    download_url = 'https://github.com/crispinus2/archivviewer/archive/v_03.tar.gz',
    license='GPLv3',
    description='Archive viewer for use with Medical Office AIS by Indamed',
    install_requires=[
        "lhafile >= 0.2.2",
        "PyPDF2 >= 1.26.0",
        "img2pdf >= 0.3.6",
        "PyQt5 >= 5.15.0",
        "watchdog >= 0.10.2",
        "pyqt-distutils",
        "fdb"
    ],
    entry_points = {
        "gui_scripts": ['archivviewer = archivviewer.archivviewer:main']
    },
    cmdclass=cmdclass,
    classifiers = [
        'Development Status :: 3 - Alpha',      
        'Intended Audience :: Healthcare Industry',      
        'Topic :: Utilities',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',   
        'Programming Language :: Python :: 3',      
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
    ]
)