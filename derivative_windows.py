"""
Moduł derivative_windows.py
----------------------------
Zawiera klasy okien do analizy pochodnych oraz drugich pochodnych.
"""

import numpy as np
from PyQt6 import QtWidgets, QtCore
import pyqtgraph as pg
from scipy.signal import savgol_filter
from utils import compute_zero_crossings  # import funkcji wykrywającej miejsca zerowe


class DerivativeWindow(QtWidgets.QDialog):
    """
    Okno do wyświetlania wykresu pierwszych pochodnych oraz wyszukiwania miejsc zerowych.
    """

    def __init__(self, x, deriv_y1, deriv_y2, parent=None):
        """
        Inicjalizacja okna pochodnych.

        Parameters:
            x (ndarray): Wartości osi x.
            deriv_y1 (ndarray): Obliczona pochodna dla utlenienia.
            deriv_y2 (ndarray): Obliczona pochodna dla redukcji.
        """
        super().__init__(parent)
        self.setWindowTitle("Pochodne utlenienia i redukcji")
        self.resize(800, 600)
        self.x = x
        self.orig_deriv_y1 = deriv_y1
        self.orig_deriv_y2 = deriv_y2
        self.current_curve1 = None
        self.current_curve2 = None
        self.intersectionPlot = None
        self.intersections = []  # przechowujemy miejsca zerowe
        self.init_ui()

    def init_ui(self):
        """Buduje interfejs okna, tworzy kontrolki i obszar wykresu."""
        main_layout = QtWidgets.QVBoxLayout(self)

        # Kontrolki wygładzania
        controls_layout = QtWidgets.QHBoxLayout()
        self.smoothingCheckBox = QtWidgets.QCheckBox("Wygładzanie (Savitzky-Golay)")
        self.smoothingCheckBox.setChecked(True)
        self.smoothingCheckBox.stateChanged.connect(self.update_plot)
        controls_layout.addWidget(self.smoothingCheckBox)
        controls_layout.addWidget(QtWidgets.QLabel("Okno:"))
        self.windowSpinBox = QtWidgets.QSpinBox()
        self.windowSpinBox.setRange(3, 101)
        self.windowSpinBox.setSingleStep(2)
        self.windowSpinBox.setValue(15)
        self.windowSpinBox.valueChanged.connect(self.update_plot)
        controls_layout.addWidget(self.windowSpinBox)
        controls_layout.addWidget(QtWidgets.QLabel("Stopień:"))
        self.polySpinBox = QtWidgets.QSpinBox()
        self.polySpinBox.setRange(1, 5)
        self.polySpinBox.setValue(3)
        self.polySpinBox.valueChanged.connect(self.update_plot)
        controls_layout.addWidget(self.polySpinBox)
        main_layout.addLayout(controls_layout)

        # Kontrolki zakresu zerowania
        intersection_layout = QtWidgets.QHBoxLayout()
        intersection_layout.addWidget(QtWidgets.QLabel("Zakres miejsc zerowych od:"))
        self.intMinSpin = QtWidgets.QDoubleSpinBox()
        self.intMinSpin.setRange(-1e9, 1e9)
        self.intMinSpin.setDecimals(3)
        self.intMinSpin.setValue(np.min(self.x))
        intersection_layout.addWidget(self.intMinSpin)
        intersection_layout.addWidget(QtWidgets.QLabel("do:"))
        self.intMaxSpin = QtWidgets.QDoubleSpinBox()
        self.intMaxSpin.setRange(-1e9, 1e9)
        self.intMaxSpin.setDecimals(3)
        self.intMaxSpin.setValue(np.max(self.x))
        intersection_layout.addWidget(self.intMaxSpin)
        self.findIntButton = QtWidgets.QPushButton("Znajdź miejsca zerowe")
        self.findIntButton.clicked.connect(self.find_intersections)
        intersection_layout.addWidget(self.findIntButton)
        main_layout.addLayout(intersection_layout)

        self.cursorLabel = QtWidgets.QLabel("x = ?, y = ?")
        main_layout.addWidget(self.cursorLabel)
        self.plot_widget = pg.PlotWidget(title="Wykres pochodnych")
        self.plot_widget.addLegend()
        main_layout.addWidget(self.plot_widget)
        self.plot_widget.scene().sigMouseMoved.connect(self.mouseMoved)
        self.update_plot()

    def update_plot(self):
        """Aktualizuje wykres pochodnych, stosując opcjonalne wygładzanie."""
        if self.smoothingCheckBox.isChecked():
            window_length = self.windowSpinBox.value()
            polyorder = self.polySpinBox.value()
            if window_length % 2 == 0:
                window_length += 1
            if window_length > len(self.orig_deriv_y1):
                window_length = len(self.orig_deriv_y1) if len(self.orig_deriv_y1) % 2 == 1 else len(
                    self.orig_deriv_y1) - 1
            try:
                smooth_y1 = savgol_filter(self.orig_deriv_y1, window_length, polyorder)
                smooth_y2 = savgol_filter(self.orig_deriv_y2, window_length, polyorder)
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Błąd", f"Nie udało się wygładzić danych: {str(e)}")
                smooth_y1 = self.orig_deriv_y1
                smooth_y2 = self.orig_deriv_y2
        else:
            smooth_y1 = self.orig_deriv_y1
            smooth_y2 = self.orig_deriv_y2

        self.current_curve1 = smooth_y1
        self.current_curve2 = smooth_y2
        self.plot_widget.clear()
        self.plot_widget.addLegend()
        self.plot_widget.plot(self.x, smooth_y1, pen=pg.mkPen(color='b', width=2), name='Pochodna utleniania')
        self.plot_widget.plot(self.x, smooth_y2, pen=pg.mkPen(color='r', width=2), name='Pochodna redukcji')
        if self.intersectionPlot is not None:
            self.plot_widget.removeItem(self.intersectionPlot)
            self.intersectionPlot = None

    def find_intersections(self):
        """
        Wyszukuje w zadanym zakresie miejsca zerowe obu krzywych pochodnych
        i zapisuje je analogicznie do zapisywania punktów przecięcia.
        """
        range_min = self.intMinSpin.value()
        range_max = self.intMaxSpin.value()
        # miejsca zerowe pochodnej utleniania
        zeros1 = compute_zero_crossings(self.x, self.current_curve1, range_min, range_max)
        # miejsca zerowe pochodnej redukcji
        zeros2 = compute_zero_crossings(self.x, self.current_curve2, range_min, range_max)

        intersections = zeros1 + zeros2
        self.intersections = intersections
        if intersections:
            xs = [pt[0] for pt in intersections]
            ys = [pt[1] for pt in intersections]
            self.intersectionPlot = pg.ScatterPlotItem(xs, ys, symbol='o', size=8, brush='y')
            self.plot_widget.addItem(self.intersectionPlot)
            lines = []
            for x, y in zeros1:
                lines.append(f"[Utlenianie] x = {x:.3f}, y = {y:.3f}")
            for x, y in zeros2:
                lines.append(f"[Redukcja]  x = {x:.3f}, y = {y:.3f}")
            msg = "Znalezione miejsca zerowe:\n" + "\n".join(lines)
            QtWidgets.QMessageBox.information(self, "Miejsca zerowe", msg)
        else:
            QtWidgets.QMessageBox.information(self, "Miejsca zerowe", "Brak miejsc zerowych w zadanym zakresie.")

    def mouseMoved(self, evt):
        """Aktualizuje etykietę z pozycją kursora podczas poruszania myszką."""
        pos = evt[0] if isinstance(evt, (list, tuple)) else evt
        if self.plot_widget.sceneBoundingRect().contains(pos):
            mouse_point = self.plot_widget.getViewBox().mapSceneToView(pos)
            self.cursorLabel.setText(f"x = {mouse_point.x():.3f}, y = {mouse_point.y():.3f}")


class SecondDerivativeWindow(QtWidgets.QDialog):
    """
    Okno do wyświetlania wykresu drugich pochodnych oraz wyszukiwania miejsc zerowych.
    """

    def __init__(self, x, second_deriv_y1, second_deriv_y2, parent=None):
        """
        Inicjalizacja okna drugich pochodnych.

        Parameters:
            x (ndarray): Wartości osi x.
            second_deriv_y1 (ndarray): Druga pochodna dla utlenienia.
            second_deriv_y2 (ndarray): Druga pochodna dla redukcji.
        """
        super().__init__(parent)
        self.setWindowTitle("Druga pochodna utlenienia i redukcji")
        self.resize(800, 600)
        self.x = x
        self.orig_second_deriv_y1 = second_deriv_y1
        self.orig_second_deriv_y2 = second_deriv_y2
        self.current_curve1 = None
        self.current_curve2 = None
        self.intersectionPlot = None
        self.intersections = []  # przechowujemy miejsca zerowe
        self.init_ui()

    def init_ui(self):
        """Buduje interfejs okna drugich pochodnych."""
        main_layout = QtWidgets.QVBoxLayout(self)

        # Kontrolki wygładzania
        controls_layout = QtWidgets.QHBoxLayout()
        self.smoothingCheckBox = QtWidgets.QCheckBox("Wygładzanie (Savitzky-Golay)")
        self.smoothingCheckBox.setChecked(True)
        self.smoothingCheckBox.stateChanged.connect(self.update_plot)
        controls_layout.addWidget(self.smoothingCheckBox)
        controls_layout.addWidget(QtWidgets.QLabel("Okno:"))
        self.windowSpinBox = QtWidgets.QSpinBox()
        self.windowSpinBox.setRange(3, 101)
        self.windowSpinBox.setSingleStep(2)
        self.windowSpinBox.setValue(15)
        self.windowSpinBox.valueChanged.connect(self.update_plot)
        controls_layout.addWidget(self.windowSpinBox)
        controls_layout.addWidget(QtWidgets.QLabel("Stopień:"))
        self.polySpinBox = QtWidgets.QSpinBox()
        self.polySpinBox.setRange(1, 5)
        self.polySpinBox.setValue(3)
        self.polySpinBox.valueChanged.connect(self.update_plot)
        controls_layout.addWidget(self.polySpinBox)
        main_layout.addLayout(controls_layout)

        # Kontrolki zakresu zerowania
        intersection_layout = QtWidgets.QHBoxLayout()
        intersection_layout.addWidget(QtWidgets.QLabel("Zakres miejsc zerowych od:"))
        self.intMinSpin = QtWidgets.QDoubleSpinBox()
        self.intMinSpin.setRange(-1e9, 1e9)
        self.intMinSpin.setDecimals(3)
        self.intMinSpin.setValue(np.min(self.x))
        intersection_layout.addWidget(self.intMinSpin)
        intersection_layout.addWidget(QtWidgets.QLabel("do:"))
        self.intMaxSpin = QtWidgets.QDoubleSpinBox()
        self.intMaxSpin.setRange(-1e9, 1e9)
        self.intMaxSpin.setDecimals(3)
        self.intMaxSpin.setValue(np.max(self.x))
        intersection_layout.addWidget(self.intMaxSpin)
        self.findIntButton = QtWidgets.QPushButton("Znajdź miejsca zerowe")
        self.findIntButton.clicked.connect(self.find_intersections)
        intersection_layout.addWidget(self.findIntButton)
        main_layout.addLayout(intersection_layout)

        self.cursorLabel = QtWidgets.QLabel("x = ?, y = ?")
        main_layout.addWidget(self.cursorLabel)
        self.plot_widget = pg.PlotWidget(title="Wykres drugiej pochodnej")
        self.plot_widget.addLegend()
        main_layout.addWidget(self.plot_widget)
        self.plot_widget.scene().sigMouseMoved.connect(self.mouseMoved)
        self.update_plot()

    def update_plot(self):
        """Aktualizuje wykres drugiej pochodnej z uwzględnieniem wygładzania."""
        if self.smoothingCheckBox.isChecked():
            window_length = self.windowSpinBox.value()
            polyorder = self.polySpinBox.value()
            if window_length % 2 == 0:
                window_length += 1
            if window_length > len(self.orig_second_deriv_y1):
                window_length = len(self.orig_second_deriv_y1) if len(self.orig_second_deriv_y1) % 2 == 1 else len(
                    self.orig_second_deriv_y1) - 1
            try:
                smooth_y1 = savgol_filter(self.orig_second_deriv_y1, window_length, polyorder)
                smooth_y2 = savgol_filter(self.orig_second_deriv_y2, window_length, polyorder)
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Błąd", f"Nie udało się wygładzić danych: {str(e)}")
                smooth_y1 = self.orig_second_deriv_y1
                smooth_y2 = self.orig_second_deriv_y2
        else:
            smooth_y1 = self.orig_second_deriv_y1
            smooth_y2 = self.orig_second_deriv_y2

        self.current_curve1 = smooth_y1
        self.current_curve2 = smooth_y2
        self.plot_widget.clear()
        self.plot_widget.addLegend()
        self.plot_widget.plot(self.x, smooth_y1, pen=pg.mkPen(color='b', width=2), name='Druga pochodna utleniania')
        self.plot_widget.plot(self.x, smooth_y2, pen=pg.mkPen(color='r', width=2), name='Druga pochodna redukcji')
        if self.intersectionPlot is not None:
            self.plot_widget.removeItem(self.intersectionPlot)
            self.intersectionPlot = None

    def find_intersections(self):
        """
        Wyszukuje w zadanym zakresie miejsca zerowe obu krzywych drugich pochodnych
        i zapisuje je analogicznie do zapisywania punktów przecięcia.
        """
        range_min = self.intMinSpin.value()
        range_max = self.intMaxSpin.value()
        # miejsca zerowe drugiej pochodnej utleniania
        zeros1 = compute_zero_crossings(self.x, self.current_curve1, range_min, range_max)
        # miejsca zerowe drugiej pochodnej redukcji
        zeros2 = compute_zero_crossings(self.x, self.current_curve2, range_min, range_max)

        intersections = zeros1 + zeros2
        self.intersections = intersections
        if intersections:
            xs = [pt[0] for pt in intersections]
            ys = [pt[1] for pt in intersections]
            self.intersectionPlot = pg.ScatterPlotItem(xs, ys, symbol='o', size=8, brush='y')
            self.plot_widget.addItem(self.intersectionPlot)
            lines = []
            for x, y in zeros1:
                lines.append(f"[Utlenianie] x = {x:.3f}, y = {y:.3f}")
            for x, y in zeros2:
                lines.append(f"[Redukcja]  x = {x:.3f}, y = {y:.3f}")
            msg = "Znalezione miejsca zerowe:\n" + "\n".join(lines)
            QtWidgets.QMessageBox.information(self, "Miejsca zerowe", msg)
        else:
            QtWidgets.QMessageBox.information(self, "Miejsca zerowe", "Brak miejsc zerowych w zadanym zakresie.")

    def mouseMoved(self, evt):
        """Aktualizuje etykietę z pozycją kursora podczas poruszania myszką."""
        pos = evt[0] if isinstance(evt, (list, tuple)) else evt
        if self.plot_widget.sceneBoundingRect().contains(pos):
            mouse_point = self.plot_widget.getViewBox().mapSceneToView(pos)
            self.cursorLabel.setText(f"x = {mouse_point.x():.3f}, y = {mouse_point.y():.3f}")
