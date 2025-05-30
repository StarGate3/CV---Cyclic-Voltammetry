"""
Plik main.py
------------
Punkt wejścia do aplikacji. Inicjuje QApplication i wyświetla główne okno.
"""

import sys
from PyQt6 import QtWidgets
from main_window import MainWindow

def main():
    """Główna funkcja uruchamiająca aplikację."""
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
