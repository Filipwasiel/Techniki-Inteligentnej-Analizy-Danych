import sys
import os
try:
    import six
    for importer in sys.meta_path:
        if type(importer).__name__ == "_SixMetaPathImporter":
            importer._path = None
except ImportError:
    pass
if sys.stdout is None or sys.stderr is None:
    class DummyStream:
        def write(self, data):
            pass
        def flush(self):
            pass
    sys.stdout = DummyStream()
    sys.stderr = DummyStream()
# WAZNE NIE USUWAJ TEGO
import pandas
from PySide6.QtWidgets import QApplication
from UI.main_window import MainWindow

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()

if __name__ == '__main__':
    main()
