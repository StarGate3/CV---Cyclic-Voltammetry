"""
Moduł main_window.py
--------------------
Zawiera główną klasę okna aplikacji (MainWindow), w której umieszczona jest logika interfejsu,
import danych, obliczeń oraz eksportu wyników do Excela.
"""

import sys
import numpy as np
import pandas as pd
from PyQt6 import QtWidgets, QtGui, QtCore
import pyqtgraph as pg
from scipy.signal import savgol_filter
import xlsxwriter

from dialogs import AxisSettingsDialog, BaselineSettingsDialog
from derivative_windows import DerivativeWindow, SecondDerivativeWindow
from utils import compute_intersections


class MainWindow(QtWidgets.QMainWindow):
    """
    Główne okno aplikacji do analizy woltamogramu cyklicznego.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CVision: Analiza woltamogramu cyklicznego")
        self.E_half_line = None
        self.plot_widget = pg.PlotWidget(title="Woltamogram")
        self.plot_widget.addLegend()
        self.init_ui()
        self.apply_theme("Ciemny")
        self.is_updating_baseline = False
        self.baseline_mode = None
        self.num_clicks = 0
        self.axis_settings = {
            'x_label': 'E [mV]',
            'y_label': 'I [μA]',
            'x_min': 0,
            'x_max': 10,
            'y_min': 0,
            'y_max': 10,
            'font': QtGui.QFont("Arial", 12)
        }
        self.update_axis_settings()
        self.baseline_settings = {
            'oxidation': {'x1': 0, 'y1': 0, 'x2': 10, 'y2': 0},
            'reduction': {'x1': 0, 'y1': 0, 'x2': 10, 'y2': 0}
        }
        self.baseline_region_oxidation = None
        self.baseline_region_reduction = None
        self.baseline_line_oxidation = None
        self.baseline_line_reduction = None
        self.peak_text_oxidation = None
        self.peak_text_reduction = None
        self.ip_a_line = None
        self.ip_c_line = None
        self.peak_curve_oxidation = None
        self.peak_curve_reduction = None
        self.x = None
        self.y1 = None
        self.y2 = None
        self.measurement_type = 0
        self.smoothingCheckBox = QtWidgets.QCheckBox("Wygładzanie (Savitzky-Golay)")
        self.windowSpinBox = QtWidgets.QSpinBox()
        self.windowSpinBox.setRange(3, 101)
        self.windowSpinBox.setSingleStep(2)
        self.windowSpinBox.setValue(15)
        self.polySpinBox = QtWidgets.QSpinBox()
        self.polySpinBox.setRange(1, 5)
        self.polySpinBox.setValue(3)
        self.raw_y1 = None
        self.raw_y2 = None
        self.smoothingCheckBox.stateChanged.connect(self.update_plot_from_raw_data)
        self.windowSpinBox.valueChanged.connect(self.update_plot_from_raw_data)
        self.polySpinBox.valueChanged.connect(self.update_plot_from_raw_data)
        self.setup_layout()
        self.resultsTable = QtWidgets.QTableWidget()
        self.resultsTable.setColumnCount(5)
        self.resultsTable.setHorizontalHeaderLabels(["Typ", "x_peak", "y_peak", "Baseline", "H/D"])
        self.centralLayout.addWidget(self.resultsTable)
        self.setStatusBar(QtWidgets.QStatusBar())
        self.proxy = pg.SignalProxy(self.plot_widget.scene().sigMouseMoved, rateLimit=60, slot=self.mouseMoved)
        self.plot_widget.scene().sigMouseClicked.connect(self.on_mouse_click)

    def setup_layout(self):
        """Buduje layout głównego okna aplikacji."""
        top_row1 = QtWidgets.QHBoxLayout()
        top_row2 = QtWidgets.QHBoxLayout()
        top_layout = QtWidgets.QVBoxLayout()
        top_layout.addLayout(top_row1)
        top_layout.addLayout(top_row2)
        self.measurement_type_combo = QtWidgets.QComboBox()
        self.measurement_type_combo.addItems([
            "Utlenianie",
            "Redukacja"
        ])
        for i in range(self.measurement_type_combo.count()):
            self.measurement_type_combo.setItemData(
                i,
                QtCore.Qt.AlignmentFlag.AlignCenter,
                QtCore.Qt.ItemDataRole.TextAlignmentRole
            )
        top_row1.addWidget(self.measurement_type_combo)
        btn_select_file = QtWidgets.QPushButton("Wybierz plik z danymi")
        btn_select_file.clicked.connect(self.open_file)
        top_row1.addWidget(btn_select_file)
        btn_baseline_settings = QtWidgets.QPushButton("Edytuj linię bazową (numerycznie)")
        btn_baseline_settings.clicked.connect(self.edit_baseline_settings)
        top_row1.addWidget(btn_baseline_settings)
        btn_clear = QtWidgets.QPushButton("Wyczyść wykres")
        btn_clear.clicked.connect(self.clear_plot)
        top_row1.addWidget(btn_clear)
        btn_axis_settings = QtWidgets.QPushButton("Edytuj ustawienia osi")
        btn_axis_settings.clicked.connect(self.edit_axis_settings)
        top_row1.addWidget(btn_axis_settings)
        btn_export = QtWidgets.QPushButton("Eksport do Excela")
        btn_export.clicked.connect(self.export_to_excel)
        top_row1.addWidget(btn_export)
        btn_help = QtWidgets.QPushButton("Help")
        btn_help.clicked.connect(self.show_help)
        top_row1.addWidget(btn_help)
        btn_about = QtWidgets.QPushButton("About")
        btn_about.clicked.connect(self.show_about)
        top_row1.addWidget(btn_about)
        btn_pick_ox = QtWidgets.QPushButton("Zakres utlenienia (2x klik)")
        btn_pick_ox.clicked.connect(self.pick_baseline_oxidation)
        top_row2.addWidget(btn_pick_ox)
        btn_pick_red = QtWidgets.QPushButton("Zakres redukcji (2x klik)")
        btn_pick_red.clicked.connect(self.pick_baseline_reduction)
        top_row2.addWidget(btn_pick_red)
        btn_compute_peak = QtWidgets.QPushButton("Oblicz parametry piku")
        btn_compute_peak.clicked.connect(self.compute_peak_parameters)
        top_row2.addWidget(btn_compute_peak)
        btn_derivative = QtWidgets.QPushButton("Oblicz pochodną")
        btn_derivative.clicked.connect(self.compute_derivative)
        top_row2.addWidget(btn_derivative)
        btn_second_derivative = QtWidgets.QPushButton("Oblicz drugą pochodną")
        btn_second_derivative.clicked.connect(self.compute_second_derivative)
        top_row2.addWidget(btn_second_derivative)
        self.combo_theme = QtWidgets.QComboBox()
        self.combo_theme.addItems(["Ciemny", "Jasny"])
        for i in range(self.combo_theme.count()):
            self.combo_theme.setItemData(
                i,
                QtCore.Qt.AlignmentFlag.AlignCenter,
                QtCore.Qt.ItemDataRole.TextAlignmentRole
            )
        self.combo_theme.currentTextChanged.connect(self.apply_theme)
        top_row2.addWidget(self.combo_theme)
        top_row2.addWidget(self.smoothingCheckBox)
        top_row2.addWidget(QtWidgets.QLabel("Okno:"))
        top_row2.addWidget(self.windowSpinBox)
        top_row2.addWidget(QtWidgets.QLabel("Stopień:"))
        top_row2.addWidget(self.polySpinBox)
        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        self.centralLayout = QtWidgets.QVBoxLayout(central_widget)
        self.centralLayout.addLayout(top_layout)
        self.centralLayout.addWidget(self.plot_widget)

    def mouseMoved(self, evt):
        """Wyświetla bieżące współrzędne kursora w pasku stanu."""
        pos = evt[0]
        if self.plot_widget.sceneBoundingRect().contains(pos):
            mouse_point = self.plot_widget.getViewBox().mapSceneToView(pos)
            self.statusBar().showMessage(f"x = {mouse_point.x():.3f}, y = {mouse_point.y():.3f}")

    def init_ui(self):
        """Dodatkowa inicjalizacja interfejsu (aktualnie pusta)."""
        pass

    def apply_theme(self, theme):
        """Zmienia motyw aplikacji na ciemny lub jasny."""
        if theme == "Ciemny":
            self.setStyleSheet("QWidget { background-color: #2e2e2e; color: white; }")
            self.plot_widget.setBackground('k')
            self.plot_widget.setStyleSheet("border: 1px solid white;")
        else:
            self.setStyleSheet("")
            self.plot_widget.setBackground('w')
            self.plot_widget.setStyleSheet("border: 1px solid black;")

    def open_file(self):
        """Otwiera okno wyboru pliku i importuje dane z wybranego pliku."""
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Wybierz plik z danymi", "",
                                                             "Pliki tekstowe (*.txt);;Wszystkie pliki (*)")
        if file_name:
            try:
                data = np.loadtxt(file_name)
                self.measurement_type = self.measurement_type_combo.currentIndex()
                if self.measurement_type == 0:
                    self.x = data[:, 0]
                    self.raw_y1 = data[:, 1]
                    self.raw_y2 = data[:, 2]
                else:
                    self.x = data[:, 0]
                    self.raw_y1 = data[:, 2]
                    self.raw_y2 = data[:, 1]
                if np.any(np.diff(self.x) < 0):
                    idx_sort = np.argsort(self.x)
                    self.x = self.x[idx_sort]
                    self.raw_y1 = self.raw_y1[idx_sort]
                    self.raw_y2 = self.raw_y2[idx_sort]
                self.update_plot_from_raw_data()
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Błąd", f"Nie udało się zaimportować danych z pliku.\n{str(e)}")

    def update_plot_from_raw_data(self):
        """Aktualizuje wykres główny na podstawie danych surowych i opcjonalnie stosuje wygładzanie."""
        if self.x is None or self.raw_y1 is None or self.raw_y2 is None:
            return
        if self.smoothingCheckBox.isChecked():
            window_length = self.windowSpinBox.value()
            polyorder = self.polySpinBox.value()
            if window_length % 2 == 0:
                window_length += 1
            if window_length > len(self.raw_y1):
                window_length = len(self.raw_y1) if len(self.raw_y1) % 2 == 1 else len(self.raw_y1) - 1
            self.y1 = savgol_filter(self.raw_y1, window_length, polyorder)
            self.y2 = savgol_filter(self.raw_y2, window_length, polyorder)
        else:
            self.y1 = self.raw_y1.copy()
            self.y2 = self.raw_y2.copy()
        self.plot_widget.clear()
        self.plot_widget.addLegend()
        self.plot_widget.plot(self.x, self.y1, pen=pg.mkPen(color='b', width=2), name='Utlenianie')
        self.plot_widget.plot(self.x, self.y2, pen=pg.mkPen(color='r', width=2), name='Redukcja')
        new_x_min = np.min(self.x)
        new_x_max = np.max(self.x)
        new_y_min = min(np.min(self.y1), np.min(self.y2))
        new_y_max = max(np.max(self.y1), np.max(self.y2))
        self.axis_settings['x_min'] = new_x_min
        self.axis_settings['x_max'] = new_x_max
        self.axis_settings['y_min'] = new_y_min
        self.axis_settings['y_max'] = new_y_max
        self.update_axis_settings()
        mid_x = (new_x_min + new_x_max) / 2
        self.baseline_settings['oxidation'] = {'x1': new_x_min, 'y1': new_y_min, 'x2': mid_x, 'y2': new_y_min}
        self.baseline_settings['reduction'] = {'x1': mid_x, 'y1': new_y_min, 'x2': new_x_max, 'y2': new_y_min}
        self.update_baseline_lines()

    def clear_plot(self):
        """Czyści wykres oraz resetuje wszystkie dane i elementy graficzne."""
        self.plot_widget.clear()
        self.plot_widget.addLegend()
        self.update_axis_settings()
        for item in [self.baseline_region_oxidation, self.baseline_region_reduction,
                     self.baseline_line_oxidation, self.baseline_line_reduction,
                     self.peak_text_oxidation, self.peak_text_reduction,
                     self.ip_a_line, self.ip_c_line, self.peak_curve_oxidation, self.peak_curve_reduction]:
            if item is not None:
                self.plot_widget.removeItem(item)
        if self.E_half_line is not None:
            self.plot_widget.removeItem(self.E_half_line)
            self.E_half_line = None
        self.baseline_region_oxidation = None
        self.baseline_region_reduction = None
        self.baseline_line_oxidation = None
        self.baseline_line_reduction = None
        self.peak_text_oxidation = None
        self.peak_text_reduction = None
        self.ip_a_line = None
        self.ip_c_line = None
        self.peak_curve_oxidation = None
        self.peak_curve_reduction = None
        self.resultsTable.setRowCount(0)

        self.x = None
        self.raw_y1 = None
        self.raw_y2 = None
        self.y1 = None
        self.y2 = None
        if hasattr(self, "deriv_y1"):
            self.deriv_y1 = None
        if hasattr(self, "deriv_y2"):
            self.deriv_y2 = None
        if hasattr(self, "second_deriv_y1"):
            self.second_deriv_y1 = None
        if hasattr(self, "second_deriv_y2"):
            self.second_deriv_y2 = None
        self.measurement_type = 0

    def edit_axis_settings(self):
        """Otwiera dialog edycji ustawień osi."""
        dialog = AxisSettingsDialog(self.axis_settings, self)
        dialog.applied.connect(self.on_axis_settings_applied)
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            self.axis_settings = dialog.get_settings()
            self.update_axis_settings()

    def on_axis_settings_applied(self, settings):
        """Aktualizuje ustawienia osi po zatwierdzeniu zmian w dialogu."""
        self.axis_settings = settings
        self.update_axis_settings()

    def edit_baseline_settings(self):
        """Otwiera dialog edycji ustawień linii bazowej."""
        dialog = BaselineSettingsDialog(self.baseline_settings, self)
        dialog.baseline_applied.connect(self.on_baseline_settings_applied)
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            self.baseline_settings = dialog.get_settings()
            self.update_baseline_lines()

    def on_baseline_settings_applied(self, settings):
        """Aktualizuje linię bazową na podstawie ustawień z dialogu."""
        for key in ['oxidation', 'reduction']:
            x1_new, y1_old, x2_new, y2_old = (
                settings[key]['x1'],
                settings[key]['y1'],
                settings[key]['x2'],
                settings[key]['y2'],
            )
            x1_old = self.baseline_settings[key]['x1']
            x2_old = self.baseline_settings[key]['x2']
            if x1_new != x1_old or x2_new != x2_old:
                m = (y2_old - y1_old) / (x2_old - x1_old) if x2_old != x1_old else 0
                y1_new = y2_old - m * (x2_old - x1_new) if x1_new != x1_old else y1_old
                y2_new = y1_old + m * (x2_new - x1_old) if x2_new != x2_old else y2_old
                settings[key]['y1'] = y1_new
                settings[key]['y2'] = y2_new
        self.baseline_settings = settings
        self.update_baseline_lines()

    def pick_baseline_oxidation(self):
        """Aktywuje tryb wyboru zakresu dla utlenienia poprzez dwukrotne kliknięcie."""
        self.baseline_mode = "oxidation"
        self.num_clicks = 0
        QtWidgets.QMessageBox.information(self, "Zakres utlenienia",
                                          "Kliknij dwa razy w obszar wykresu, aby wybrać punkty (x1,y1) oraz (x2,y2) dla utlenienia.")

    def pick_baseline_reduction(self):
        """Aktywuje tryb wyboru zakresu dla redukcji poprzez dwukrotne kliknięcie."""
        self.baseline_mode = "reduction"
        self.num_clicks = 0
        QtWidgets.QMessageBox.information(self, "Zakres redukcji",
                                          "Kliknij dwa razy w obszar wykresu, aby wybrać punkty (x1,y1) oraz (x2,y2) dla redukcji.")

    def on_mouse_click(self, event):
        """Obsługuje kliknięcia myszą w celu wyboru punktów dla linii bazowej."""
        if self.baseline_mode is None:
            return
        pos = event.scenePos()
        mouse_point = self.plot_widget.getViewBox().mapSceneToView(pos)
        x_click = mouse_point.x()
        if self.measurement_type == 0:
            data_oxidation = self.y1
            data_reduction = self.y2
        else:
            data_oxidation = self.y1
            data_reduction = self.y2
        if np.any(np.diff(self.x) < 0):
            sorted_indices = np.argsort(self.x)
            self.x = self.x[sorted_indices]
            data_oxidation = data_oxidation[sorted_indices]
            data_reduction = data_reduction[sorted_indices]
        y_curve = np.interp(x_click, self.x, data_oxidation if self.baseline_mode == "oxidation" else data_reduction)
        if self.num_clicks == 0:
            if self.baseline_mode == "oxidation":
                self.baseline_settings['oxidation']['x1'] = x_click
                self.baseline_settings['oxidation']['y1'] = y_curve
            else:
                self.baseline_settings['reduction']['x1'] = x_click
                self.baseline_settings['reduction']['y1'] = y_curve
            self.num_clicks = 1
        else:
            if self.baseline_mode == "oxidation":
                self.baseline_settings['oxidation']['x2'] = x_click
                self.baseline_settings['oxidation']['y2'] = y_curve
            else:
                self.baseline_settings['reduction']['x2'] = x_click
                self.baseline_settings['reduction']['y2'] = y_curve
            self.num_clicks = 0
            self.baseline_mode = None
            self.update_baseline_lines()

    def update_axis_settings(self):
        """Aktualizuje etykiety oraz zakresy osi wykresu."""
        x_label = self.axis_settings.get('x_label', 'Oś X')
        y_label = self.axis_settings.get('y_label', 'Prąd')
        font = self.axis_settings.get('font', QtGui.QFont("Arial", 12))
        self.plot_widget.setLabel('bottom', text=x_label, **{'font': font})
        self.plot_widget.setLabel('left', text=y_label, **{'font': font})
        x_min = self.axis_settings.get('x_min', 0)
        x_max = self.axis_settings.get('x_max', 10)
        y_min = self.axis_settings.get('y_min', 0)
        y_max = self.axis_settings.get('y_max', 10)
        self.plot_widget.setXRange(x_min, x_max)
        self.plot_widget.setYRange(y_min, y_max)

    def update_baseline_lines(self):
        """Rysuje na wykresie linie bazowe oraz regiony interaktywne dla utlenienia i redukcji."""
        self.is_updating_baseline = True
        if self.baseline_region_oxidation:
            self.plot_widget.removeItem(self.baseline_region_oxidation)
        if self.baseline_region_reduction:
            self.plot_widget.removeItem(self.baseline_region_reduction)
        ox = self.baseline_settings['oxidation']
        x1_ox, y1_ox, x2_ox, y2_ox = ox['x1'], ox['y1'], ox['x2'], ox['y2']
        self.baseline_region_oxidation = pg.LinearRegionItem(
            values=[min(x1_ox, x2_ox), max(x1_ox, x2_ox)],
            brush=(0, 0, 255, 50),
            movable=True
        )
        self.baseline_region_oxidation.sigRegionChanged.connect(self.on_oxidation_region_changed)
        self.plot_widget.addItem(self.baseline_region_oxidation)
        if self.baseline_line_oxidation:
            self.plot_widget.removeItem(self.baseline_line_oxidation)
        baseline_x_ox = [x1_ox, x2_ox]
        baseline_y_ox = [y1_ox, y2_ox]
        self.baseline_line_oxidation = self.plot_widget.plot(
            baseline_x_ox, baseline_y_ox,
            pen=pg.mkPen(color='b', width=2, style=QtCore.Qt.PenStyle.DashLine),
            name="Baseline Utlenienia"
        )
        red = self.baseline_settings['reduction']
        x1_red, y1_red, x2_red, y2_red = red['x1'], red['y1'], red['x2'], red['y2']
        if self.baseline_region_reduction:
            self.plot_widget.removeItem(self.baseline_region_reduction)
        self.baseline_region_reduction = pg.LinearRegionItem(
            values=[min(x1_red, x2_red), max(x1_red, x2_red)],
            brush=(255, 0, 0, 50),
            movable=True
        )
        self.baseline_region_reduction.sigRegionChanged.connect(self.on_reduction_region_changed)
        self.plot_widget.addItem(self.baseline_region_reduction)
        if self.baseline_line_reduction:
            self.plot_widget.removeItem(self.baseline_line_reduction)
        baseline_x_red = [x1_red, x2_red]
        baseline_y_red = [y1_red, y2_red]
        self.baseline_line_reduction = self.plot_widget.plot(
            baseline_x_red, baseline_y_red,
            pen=pg.mkPen(color='r', width=2, style=QtCore.Qt.PenStyle.DashLine),
            name="Baseline Redukcji"
        )
        self.is_updating_baseline = False

    def on_oxidation_region_changed(self):
        """Obsługuje zmianę regionu interaktywnego dla utlenienia."""
        if self.is_updating_baseline:
            return
        region = self.baseline_region_oxidation.getRegion()
        x_min, x_max = region
        old_y1 = self.baseline_settings['oxidation']['y1']
        old_y2 = self.baseline_settings['oxidation']['y2']
        self.baseline_settings['oxidation'] = {'x1': x_min, 'y1': old_y1, 'x2': x_max, 'y2': old_y2}
        self.update_baseline_lines()

    def on_reduction_region_changed(self):
        """Obsługuje zmianę regionu interaktywnego dla redukcji."""
        if self.is_updating_baseline:
            return
        region = self.baseline_region_reduction.getRegion()
        x_min, x_max = region
        old_y1 = self.baseline_settings['reduction']['y1']
        old_y2 = self.baseline_settings['reduction']['y2']
        self.baseline_settings['reduction'] = {'x1': x_min, 'y1': old_y1, 'x2': x_max, 'y2': old_y2}
        self.update_baseline_lines()

    def compute_peak_parameters(self):
        """Oblicza parametry piku na podstawie danych i aktualnych ustawień linii bazowych."""
        if self.x is None:
            QtWidgets.QMessageBox.warning(self, "Brak danych", "Najpierw zaimportuj dane.")
            return
        for item in [self.peak_text_oxidation, self.peak_text_reduction, self.ip_a_line, self.ip_c_line,
                     self.peak_curve_oxidation, self.peak_curve_reduction]:
            if item is not None:
                self.plot_widget.removeItem(item)
        self.peak_text_oxidation = None
        self.peak_text_reduction = None
        self.ip_a_line = None
        self.ip_c_line = None
        self.peak_curve_oxidation = None
        self.peak_curve_reduction = None
        results = ""
        ox = self.baseline_settings['oxidation']
        ox_x1, ox_y1, ox_x2, ox_y2 = ox['x1'], ox['y1'], ox['x2'], ox['y2']
        region_min = min(ox_x1, ox_x2)
        region_max = max(ox_x1, ox_x2)
        mask = (self.x >= region_min) & (self.x <= region_max)
        if np.any(mask):
            x_region = self.x[mask]
            y_region = self.y1[mask]
            idx_peak = np.argmax(y_region)
            x_peak = x_region[idx_peak]
            y_peak = y_region[idx_peak]
            baseline_val = ox_y1 + (ox_y2 - ox_y1) * (x_peak - ox_x1) / (ox_x2 - ox_x1) if ox_x2 != ox_x1 else ox_y1
            height = y_peak - baseline_val
            text = (f"Utlenienie:\n"
                    f"x_peak = {x_peak:.3f}\n"
                    f"y_peak = {y_peak:.3f}\n"
                    f"baseline = {baseline_val:.3f}\n"
                    f"height = {height:.3f}")
            self.peak_text_oxidation = pg.TextItem(text=text, color='b', anchor=(0.5, -1.0))
            self.peak_text_oxidation.setPos(x_peak, y_peak)
            self.plot_widget.addItem(self.peak_text_oxidation)
            results += f"Utlenienie: x_peak={x_peak:.3f}, y_peak={y_peak:.3f}, baseline={baseline_val:.3f}, height={height:.3f}\n"
            self.ip_a_line = self.plot_widget.plot([x_peak, x_peak], [baseline_val, y_peak],
                                                   pen=pg.mkPen(color='b', width=2, style=QtCore.Qt.PenStyle.DashLine),
                                                   name="Ip,a")
            baseline_curve = ox_y1 + (ox_y2 - ox_y1) * (x_region - ox_x1) / (ox_x2 - ox_x1)
            peak_height_curve = self.y1[mask] - baseline_curve
            self.peak_curve_oxidation = self.plot_widget.plot(x_region, peak_height_curve,
                                                              pen=pg.mkPen(color='c', width=2),
                                                              name="Peak Height Ox")
            self.insert_result_row("Utlenienie", x_peak, y_peak, baseline_val, height)
        else:
            results += "Utlenienie: brak danych w zadanym zakresie.\n\n"
        red = self.baseline_settings['reduction']
        red_x1, red_y1, red_x2, red_y2 = red['x1'], red['y1'], red['x2'], red['y2']
        region_min = min(red_x1, red_x2)
        region_max = max(red_x1, red_x2)
        mask = (self.x >= region_min) & (self.x <= region_max)
        if np.any(mask):
            x_region = self.x[mask]
            y_region = self.y2[mask]
            idx_peak = np.argmin(y_region)
            x_peak = x_region[idx_peak]
            y_peak = y_region[idx_peak]
            baseline_val = red_y1 + (red_y2 - red_y1) * (x_peak - red_x1) / (red_x2 - red_x1) if red_x2 != red_x1 else red_y1
            depth = baseline_val - y_peak
            text = (f"Redukcja:\n"
                    f"x_peak = {x_peak:.3f}\n"
                    f"y_peak = {y_peak:.3f}\n"
                    f"baseline = {baseline_val:.3f}\n"
                    f"depth = {depth:.3f}")
            self.peak_text_reduction = pg.TextItem(text=text, color='r', anchor=(0.5, -1.0))
            self.peak_text_reduction.setPos(x_peak, y_peak)
            self.plot_widget.addItem(self.peak_text_reduction)
            results += f"Redukcja: x_peak={x_peak:.3f}, y_peak={y_peak:.3f}, baseline={baseline_val:.3f}, depth={depth:.3f}\n"
            self.ip_c_line = self.plot_widget.plot([x_peak, x_peak], [y_peak, baseline_val],
                                                   pen=pg.mkPen(color='r', width=2, style=QtCore.Qt.PenStyle.DashLine),
                                                   name="Ip,c")
            baseline_curve = red_y1 + (red_y2 - red_y1) * (x_region - red_x1) / (red_x2 - red_x1)
            peak_height_curve = self.y2[mask] - baseline_curve
            self.peak_curve_reduction = self.plot_widget.plot(x_region, peak_height_curve,
                                                              pen=pg.mkPen(color='m', width=2),
                                                              name="Peak Height Red")
            self.insert_result_row("Redukcja", x_peak, y_peak, baseline_val, depth)
        else:
            results += "Redukcja: brak danych w zadanym zakresie.\n"
        if self.peak_text_oxidation is not None and self.peak_text_reduction is not None:
            E_half = (self.peak_text_oxidation.pos().x() + self.peak_text_reduction.pos().x()) / 2.0
            self.insert_result_row("E1/2", E_half, "", "", "")
            if self.E_half_line is not None:
                self.plot_widget.removeItem(self.E_half_line)
            self.E_half_line = pg.InfiniteLine(pos=E_half, angle=90,
                                               pen=pg.mkPen(color='g', width=2, style=QtCore.Qt.PenStyle.DashLine))
            self.plot_widget.addItem(self.E_half_line)
            results += f"E1/2: {E_half:.3f}\n"
        QtWidgets.QMessageBox.information(self, "Parametry piku", results)

    def insert_result_row(self, peak_type, x_peak, y_peak, baseline, h_or_d):
        """
        Wstawia nowy wiersz do tabeli wyników.

        Parameters:
            peak_type (str): Typ pomiaru (np. 'Utlenienie', 'Redukcja', 'E1/2').
            x_peak (float): Pozycja x piku.
            y_peak (float): Wartość y piku.
            baseline (float): Wartość linii bazowej.
            h_or_d (float): Wysokość lub głębokość piku.
        """
        row_position = self.resultsTable.rowCount()
        self.resultsTable.insertRow(row_position)
        self.resultsTable.setItem(row_position, 0, QtWidgets.QTableWidgetItem(str(peak_type)))
        self.resultsTable.setItem(row_position, 1, QtWidgets.QTableWidgetItem(f"{x_peak:.3f}" if x_peak != "" else ""))
        self.resultsTable.setItem(row_position, 2, QtWidgets.QTableWidgetItem(f"{y_peak:.3f}" if y_peak != "" else ""))
        self.resultsTable.setItem(row_position, 3, QtWidgets.QTableWidgetItem(f"{baseline:.3f}" if baseline != "" else ""))
        self.resultsTable.setItem(row_position, 4, QtWidgets.QTableWidgetItem(f"{h_or_d:.3f}" if h_or_d != "" else ""))

    def compute_derivative(self):
        """Oblicza pierwsze pochodne i otwiera okno analizy pochodnych."""
        if self.x is None or self.y1 is None or self.y2 is None:
            QtWidgets.QMessageBox.warning(self, "Brak danych", "Najpierw zaimportuj dane.")
            return
        # 1) Obliczamy pierwsze pochodne
        self.deriv_y1 = np.gradient(self.y1, self.x)
        self.deriv_y2 = np.gradient(self.y2, self.x)
        # 2) Otwieramy okno, by je wizualnie zbadać i odczytać miejsca zerowe
        derivative_window = DerivativeWindow(self.x, self.deriv_y1, self.deriv_y2, self)
        derivative_window.exec()
        # 3) Pobieramy znalezione miejsca zerowe (lista (x0, 0.0))
        zeros = derivative_window.intersections
        # 4) Wstawiamy je pojedynczo do tabeli wyników
        #    Jeśli chcesz tylko pierwsze zerowanie, weź zeros[0]
        if zeros:
            # przykład: weźmy pierwsze miejsce zerowe z utleniania, jeśli masz je oznaczone
            x0, y0 = zeros[0]
            # dopasuj nazwę typu; możesz użyć np. "Zero crossing first"
            self.insert_result_row("Zero crossing", x0, y0, "", "")
        # 5) (opcjonalnie) możesz też zapisać wszystkie w self.deriv_intersections
        self.deriv_intersections = zeros

        

    def compute_second_derivative(self):
        """Oblicza drugie pochodne i otwiera okno analizy drugich pochodnych."""
        if self.x is None or self.y1 is None or self.y2 is None:
            QtWidgets.QMessageBox.warning(self, "Brak danych", "Najpierw zaimportuj dane.")
            return
        # 1) Obliczamy drugą pochodną
        first_deriv_y1 = np.gradient(self.y1, self.x)
        first_deriv_y2 = np.gradient(self.y2, self.x)
        self.second_deriv_y1 = np.gradient(first_deriv_y1, self.x)
        self.second_deriv_y2 = np.gradient(first_deriv_y2, self.x)
        # 2) Pokaż okno analizy i zbierz miejsca zerowe
        second_derivative_window = SecondDerivativeWindow(self.x, self.second_deriv_y1, self.second_deriv_y2, self)
        second_derivative_window.exec()
        zeros2 = second_derivative_window.intersections
        # 3) Wstaw je do tabeli wyników
        if zeros2:
            for x0, y0 in zeros2:
                # Przykładowa etykieta w tabeli: "Zero crossing 2nd"
                self.insert_result_row("Zero crossing 2nd", x0, y0, "", "")
        # 4) Zapisz na przyszłość
        self.second_deriv_intersections = zeros2


    def export_to_excel(self):
        """Eksportuje dane, parametry i wykres do pliku Excel."""
        if self.x is None:
            QtWidgets.QMessageBox.warning(self, "Brak danych", "Brak danych do eksportu.")
            return
        filename, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Zapisz do Excela", "", "Excel Files (*.xlsx)")
        if not filename:
            return

        df = pd.DataFrame({
            "x": self.x,
            "y_ox": self.raw_y1 if self.raw_y1 is not None else np.nan,
            "y_red": self.raw_y2 if self.raw_y2 is not None else np.nan,
        })

        if self.smoothingCheckBox.isChecked():
            df["smoothed_y_ox"] = self.y1
            df["smoothed_y_red"] = self.y2

        if hasattr(self, "deriv_y1") and self.deriv_y1 is not None:
            df["deriv_ox"] = self.deriv_y1
        if hasattr(self, "deriv_y2") and self.deriv_y2 is not None:
            df["deriv_red"] = self.deriv_y2
        if hasattr(self, "second_deriv_y1") and self.second_deriv_y1 is not None:
            df["second_deriv_ox"] = self.second_deriv_y1
        if hasattr(self, "second_deriv_y2") and self.second_deriv_y2 is not None:
            df["second_deriv_red"] = self.second_deriv_y2

        table_rows = self.resultsTable.rowCount()
        table_data = []
        for row in range(table_rows):
            row_data = {}
            for col in range(self.resultsTable.columnCount()):
                header = self.resultsTable.horizontalHeaderItem(col).text()
                item = self.resultsTable.item(row, col)
                row_data[header] = item.text() if item is not None else ""
            table_data.append(row_data)
        df_params = pd.DataFrame(table_data)

        try:
            writer = pd.ExcelWriter(filename, engine='xlsxwriter')
            df.to_excel(writer, sheet_name="Dane", index=False)
            df_params.to_excel(writer, sheet_name="Parametry", index=False)
            if hasattr(self, "deriv_intersections") and self.deriv_intersections:
                df_deriv = pd.DataFrame(self.deriv_intersections, columns=["x", "y"])
                df_deriv.to_excel(writer, sheet_name="Przecięcia Pochodnej", index=False)
            if hasattr(self, "second_deriv_intersections") and self.second_deriv_intersections:
                df_second_deriv = pd.DataFrame(self.second_deriv_intersections, columns=["x", "y"])
                df_second_deriv.to_excel(writer, sheet_name="Przecięcia Drugiej Pochodnej", index=False)

            row_E_half = df_params[df_params["Typ"] == "E1/2"]
            if not row_E_half.empty:
                e_half_str = row_E_half["x_peak"].values[0]
                try:
                    self.E_half = float(e_half_str)
                except:
                    self.E_half = 0.0
            else:
                self.E_half = 0.0

            workbook = writer.book
            worksheet = writer.sheets["Dane"]

            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': '=Dane!$B$1',
                'categories': f"=Dane!$A$2:$A${len(self.x) + 1}",
                'values': f"=Dane!$B$2:$B${len(self.x) + 1}",
                'line': {'color': 'red'},
            })
            chart.add_series({
                'name': '=Dane!$C$1',
                'categories': f"=Dane!$A$2:$A${len(self.x) + 1}",
                'values': f"=Dane!$C$2:$C${len(self.x) + 1}",
                'line': {'color': 'blue'},
            })

            y_min = df[["y_ox", "y_red"]].min().min()
            y_max = df[["y_ox", "y_red"]].max().max()
            worksheet.write(len(self.x) + 1, 3, self.E_half)
            worksheet.write(len(self.x) + 1, 4, y_min)
            worksheet.write(len(self.x) + 2, 3, self.E_half)
            worksheet.write(len(self.x) + 2, 4, y_max)
            chart.add_series({
                'name': 'E1/2',
                'categories': f"=Dane!$D${len(self.x) + 2}:$D${len(self.x) + 3}",
                'values': f"=Dane!$E${len(self.x) + 2}:$E${len(self.x) + 3}",
                'line': {'color': 'green', 'dash_type': 'dash'},
            })

            if self.measurement_type == 0:
                chart.set_x_axis({
                    'name': 'E [mV]',
                    'name_font': {'name': 'Verdana', 'bold': True, 'size': 14},
                    'num_font': {'name': 'Calibri', 'size': 10},
                    'crossing': 'min',
                })
            else:
                chart.set_x_axis({
                    'name': 'E [mV]',
                    'name_font': {'name': 'Verdana', 'bold': True, 'size': 14},
                    'num_font': {'name': 'Calibri', 'size': 10},
                    'crossing': 'max',
                })
            chart.set_y_axis({
                'name': 'I [μA]',
                'name_font': {'name': 'Verdana', 'bold': True, 'size': 14},
                'num_font': {'name': 'Calibri', 'size': 10},
            })
            chart.set_size({'width': 600, 'height': 600})
            worksheet.insert_chart('G2', chart)

            writer.close()
            QtWidgets.QMessageBox.information(self, "Sukces", f"Dane oraz wykres zostały zapisane do pliku {filename}")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Błąd", f"Wystąpił błąd podczas zapisu do pliku:\n{str(e)}")

    def show_help(self):
        help_text = """
        <html>
        <body style="font-family:Arial; font-size:10pt;">
            <p><b>1. Wybór typu pomiaru</b><br/>
            Z rozwijanego menu wybierz “Utlenianie” lub “Redukcja”.</p>

            <p><b>2. Wczytanie danych</b><br/>
            Kliknij przycisk „Wybierz plik z danymi” i załaduj plik tekstowy (*.txt)
            zawierający trzy kolumny: E [mV], I_utlenianie [μA], I_redukcja [μA].</p>

            <p><b>3. Wygładzenie</b><br/>
            •! <i>W tym miejscu ustawienie wygładzania jest opcjonalne i zależy od jakości danych.</i><br/>
            • Zaznacz „Wygładzanie (Savitzky-Golay)”.<br/>
            <i>Uwaga:</i> niezalecane jest zwiększanie okna powyżej 15.</p>

            <p><b>4. Wybór linii bazowej</b><br/>
            <b>Utlenianie</b>: Kliknij „Zakres utlenienia (2× klik)” i wskaż dwa punkty.
            <b>Redukcja</b>: Kliknij „Zakres redukcji (2× klik)” i wskaż dwa punkty.
            
            Po wybraniu liniowego fragmentu należy ręcznie dostosować współrzędne punktów linii bazowej w taki sposób, 
            aby obejmowała ona cały pik.</p>

            <p><b>5. Obliczenie parametrów piku</b><br/>
            Kliknij „Oblicz parametry piku”. Program wyznaczy x_peak, y_peak, linię bazową, wysokość/głębokość piku i E₁/₂,
            a wyniki wyświetli na wykresie i w tabeli.</p>

            <p><b>6. Druga pochodna (procesy nieodwracalne)</b><br/>
            • Kliknij „Oblicz drugą pochodną”.<br/>
            • Opcja wygładzania jest opcjonalna, ale zalecana.<br/>
            • Podaj zakres poszukiwania miejsc zerowych.<br/>
            • Kliknij „Znajdź miejsca zerowe” – punkty zostaną pokazane na wykresie i zapisane w tabeli.</p>

            <p><b>7. Eksport do Excela</b><br/>
            Kliknij „Eksport do Excela”, wybierz nazwę pliku.
            Zapisane zostaną: surowe dane, dane wygładzone, pochodne, miejsca zerowe, wyniki piku i wykres.</p>

            <hr/>

            <p><b>Opcjonalne ustawienia</b><br/>
            • Tryb jasny/ciemny – przełącznik w górnym pasku.<br/>
            • Ręczna edycja osi – przycisk „Edytuj ustawienia osi”.</p>
        </body>
        </html>
        """
        QtWidgets.QMessageBox.information(self, "Help – instrukcja", help_text)

    def show_about(self):
        about_text = """
        <html>
        <body>
            <!-- sposób 1: użycie atrybutu align -->
            <h4 align="center">CVision</h4>

            <!-- albo sposób 2: CSS inline -->
            <!-- <h4 style="text-align:center;">CVision</h4> -->

            <h4>Analiza woltamogramu cyklicznego</h4>
            <p>Wersja: 2.0.0</p>
            <p>Autor: <b>StarGate3</b><br/>
            GitHub: <a href='https://github.com/StarGate3'>github.com/StarGate3</a>
            </p>
        </body>
        </html>
        """
        QtWidgets.QMessageBox.about(self, "About", about_text)
