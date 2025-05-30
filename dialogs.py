"""
Moduł dialogs.py
-----------------
Zawiera klasy dialogowe do ustawień osi oraz linii bazowej.
"""

from PyQt6 import QtWidgets, QtGui, QtCore


class AxisSettingsDialog(QtWidgets.QDialog):
    """
    Dialog umożliwiający ustawienia parametrów osi wykresu.
    """
    applied = QtCore.pyqtSignal(dict)

    def __init__(self, current_settings, parent=None):
        """
        Inicjalizacja dialogu z aktualnymi ustawieniami.
        """
        super().__init__(parent)
        self.setWindowTitle("Ustawienia osi")
        self.current_settings = current_settings
        self.init_ui()

    def init_ui(self):
        """Tworzy interfejs dialogu z polami do edycji ustawień osi."""
        layout = QtWidgets.QFormLayout(self)
        self.x_label_edit = QtWidgets.QLineEdit(self.current_settings.get('x_label', 'Oś X'))
        layout.addRow("Etykieta osi X:", self.x_label_edit)
        self.y_label_edit = QtWidgets.QLineEdit(self.current_settings.get('y_label', 'Wartości'))
        layout.addRow("Etykieta osi Y:", self.y_label_edit)

        # Ustawienia zakresu osi X
        self.x_min_spin = QtWidgets.QDoubleSpinBox()
        self.x_min_spin.setRange(-1e9, 1e9)
        self.x_min_spin.setDecimals(3)
        self.x_min_spin.setValue(self.current_settings.get('x_min', 0))
        self.x_max_spin = QtWidgets.QDoubleSpinBox()
        self.x_max_spin.setRange(-1e9, 1e9)
        self.x_max_spin.setDecimals(3)
        self.x_max_spin.setValue(self.current_settings.get('x_max', 10))
        x_range_layout = QtWidgets.QHBoxLayout()
        x_range_layout.addWidget(QtWidgets.QLabel("Min:"))
        x_range_layout.addWidget(self.x_min_spin)
        x_range_layout.addWidget(QtWidgets.QLabel("Max:"))
        x_range_layout.addWidget(self.x_max_spin)
        layout.addRow("Zakres osi X:", x_range_layout)

        # Ustawienia zakresu osi Y
        self.y_min_spin = QtWidgets.QDoubleSpinBox()
        self.y_min_spin.setRange(-1e9, 1e9)
        self.y_min_spin.setDecimals(3)
        self.y_min_spin.setValue(self.current_settings.get('y_min', 0))
        self.y_max_spin = QtWidgets.QDoubleSpinBox()
        self.y_max_spin.setRange(-1e9, 1e9)
        self.y_max_spin.setDecimals(3)
        self.y_max_spin.setValue(self.current_settings.get('y_max', 10))
        y_range_layout = QtWidgets.QHBoxLayout()
        y_range_layout.addWidget(QtWidgets.QLabel("Min:"))
        y_range_layout.addWidget(self.y_min_spin)
        y_range_layout.addWidget(QtWidgets.QLabel("Max:"))
        y_range_layout.addWidget(self.y_max_spin)
        layout.addRow("Zakres osi Y:", y_range_layout)

        # Wybór czcionki
        self.font = self.current_settings.get('font', QtGui.QFont("Arial", 12))
        self.font_button = QtWidgets.QPushButton("Wybierz czcionkę")
        self.font_button.clicked.connect(self.choose_font)
        self.font_label = QtWidgets.QLabel(self.font.toString())
        font_layout = QtWidgets.QHBoxLayout()
        font_layout.addWidget(self.font_button)
        font_layout.addWidget(self.font_label)
        layout.addRow("Czcionka:", font_layout)

        # Przyciski OK, Anuluj, Apply
        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Ok |
            QtWidgets.QDialogButtonBox.StandardButton.Cancel |
            QtWidgets.QDialogButtonBox.StandardButton.Apply,
            parent=self
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.button_box.button(QtWidgets.QDialogButtonBox.StandardButton.Apply).clicked.connect(self.on_apply)
        layout.addRow(self.button_box)

    def choose_font(self):
        """Otwiera okno wyboru czcionki i aktualizuje etykietę czcionki."""
        font, ok = QtWidgets.QFontDialog.getFont(self.font, self, "Wybierz czcionkę")
        if ok:
            self.font = font
            self.font_label.setText(self.font.toString())

    def on_apply(self):
        """Emituje sygnał z nowymi ustawieniami bez zamykania dialogu."""
        self.applied.emit(self.get_settings())

    def get_settings(self):
        """
        Zwraca słownik z ustawieniami osi.

        Returns:
            dict: Ustawienia osi.
        """
        return {
            'x_label': self.x_label_edit.text(),
            'y_label': self.y_label_edit.text(),
            'x_min': self.x_min_spin.value(),
            'x_max': self.x_max_spin.value(),
            'y_min': self.y_min_spin.value(),
            'y_max': self.y_max_spin.value(),
            'font': self.font
        }


class BaselineSettingsDialog(QtWidgets.QDialog):
    """
    Dialog umożliwiający numeryczną edycję ustawień linii bazowej dla utlenienia i redukcji.
    """
    baseline_applied = QtCore.pyqtSignal(dict)

    def __init__(self, current_settings, parent=None):
        """
        Inicjalizacja dialogu z aktualnymi ustawieniami linii bazowej.

        Parameters:
            current_settings (dict): Aktualne ustawienia bazowe.
        """
        super().__init__(parent)
        self.setWindowTitle("Ustawienia linii bazowej (numerycznie)")
        self.current_settings = current_settings
        # Zachowujemy oryginalne ustawienia do obliczeń
        self.initial_settings = {
            'oxidation': self.current_settings['oxidation'].copy(),
            'reduction': self.current_settings['reduction'].copy()
        }
        self.init_ui()

    def init_ui(self):
        """Tworzy interfejs dialogu z polami do ustawiania linii bazowej."""
        layout = QtWidgets.QFormLayout(self)
        # Utlenienie
        ox_x1 = QtWidgets.QDoubleSpinBox()
        ox_x1.setRange(-1e9, 1e9)
        ox_x1.setDecimals(3)
        ox_x1.setValue(self.current_settings['oxidation'].get('x1', 0))
        ox_y1 = QtWidgets.QDoubleSpinBox()
        ox_y1.setRange(-1e9, 1e9)
        ox_y1.setDecimals(3)
        ox_y1.setValue(self.current_settings['oxidation'].get('y1', 0))
        ox_x2 = QtWidgets.QDoubleSpinBox()
        ox_x2.setRange(-1e9, 1e9)
        ox_x2.setDecimals(3)
        ox_x2.setValue(self.current_settings['oxidation'].get('x2', 10))
        ox_y2 = QtWidgets.QDoubleSpinBox()
        ox_y2.setRange(-1e9, 1e9)
        ox_y2.setDecimals(3)
        ox_y2.setValue(self.current_settings['oxidation'].get('y2', 0))
        self.ox_label = QtWidgets.QLabel(self.get_oxidation_label_text())
        ox_layout = QtWidgets.QHBoxLayout()
        ox_layout.addWidget(QtWidgets.QLabel("x1:"))
        ox_layout.addWidget(ox_x1)
        ox_layout.addWidget(QtWidgets.QLabel("y1:"))
        ox_layout.addWidget(ox_y1)
        ox_layout.addWidget(QtWidgets.QLabel("x2:"))
        ox_layout.addWidget(ox_x2)
        ox_layout.addWidget(QtWidgets.QLabel("y2:"))
        ox_layout.addWidget(ox_y2)
        layout.addRow("Utlenienie:", ox_layout)
        self.ox_preview_label = QtWidgets.QLabel("Podgląd wartości: Punkty nie są jeszcze zainicjalizowane")
        layout.addRow(self.ox_preview_label)
        self.ox_x1 = ox_x1;
        self.ox_y1 = ox_y1;
        self.ox_x2 = ox_x2;
        self.ox_y2 = ox_y2

        # Redukcja
        red_x1 = QtWidgets.QDoubleSpinBox()
        red_x1.setRange(-1e9, 1e9)
        red_x1.setDecimals(3)
        red_x1.setValue(self.current_settings['reduction'].get('x1', 0))
        red_y1 = QtWidgets.QDoubleSpinBox()
        red_y1.setRange(-1e9, 1e9)
        red_y1.setDecimals(3)
        red_y1.setValue(self.current_settings['reduction'].get('y1', 0))
        red_x2 = QtWidgets.QDoubleSpinBox()
        red_x2.setRange(-1e9, 1e9)
        red_x2.setDecimals(3)
        red_x2.setValue(self.current_settings['reduction'].get('x2', 10))
        red_y2 = QtWidgets.QDoubleSpinBox()
        red_y2.setRange(-1e9, 1e9)
        red_y2.setDecimals(3)
        red_y2.setValue(self.current_settings['reduction'].get('y2', 0))
        self.red_label = QtWidgets.QLabel(self.get_reduction_label_text())
        red_layout = QtWidgets.QHBoxLayout()
        red_layout.addWidget(QtWidgets.QLabel("x1:"))
        red_layout.addWidget(red_x1)
        red_layout.addWidget(QtWidgets.QLabel("y1:"))
        red_layout.addWidget(red_y1)
        red_layout.addWidget(QtWidgets.QLabel("x2:"))
        red_layout.addWidget(red_x2)
        red_layout.addWidget(QtWidgets.QLabel("y2:"))
        red_layout.addWidget(red_y2)
        layout.addRow("Redukcja:", red_layout)
        layout.addRow("Podgląd wartości:", self.red_label)
        self.red_preview_label = QtWidgets.QLabel("Podgląd wartości: Punkty nie są jeszcze zainicjalizowane")
        layout.addRow(self.red_preview_label)
        self.red_x1 = red_x1;
        self.red_y1 = red_y1;
        self.red_x2 = red_x2;
        self.red_y2 = red_y2

        # Połączenia sygnałów do aktualizacji wartości i podglądu
        self.ox_x1.valueChanged.connect(self.update_y_values)
        self.ox_x1.valueChanged.connect(self.update_preview_labels)
        self.ox_x2.valueChanged.connect(self.update_y_values)
        self.ox_x2.valueChanged.connect(self.update_preview_labels)
        self.red_x1.valueChanged.connect(self.update_y_values)
        self.red_x1.valueChanged.connect(self.update_preview_labels)
        self.red_x2.valueChanged.connect(self.update_y_values)
        self.red_x2.valueChanged.connect(self.update_preview_labels)

        self.button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Ok |
            QtWidgets.QDialogButtonBox.StandardButton.Apply |
            QtWidgets.QDialogButtonBox.StandardButton.Cancel,
            parent=self
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.button_box.button(QtWidgets.QDialogButtonBox.StandardButton.Apply).clicked.connect(self.on_apply)
        layout.addRow(self.button_box)

    def update_y_values(self):
        """
        Aktualizuje wartości y dla obu linii bazowych na podstawie zmienionych wartości x
        przy zachowaniu oryginalnego nachylenia.
        """
        sender = self.sender()
        for key, x1_spin, y1_spin, x2_spin, y2_spin in [
            ("oxidation", self.ox_x1, self.ox_y1, self.ox_x2, self.ox_y2),
            ("reduction", self.red_x1, self.red_y1, self.red_x2, self.red_y2),
        ]:
            new_x1 = x1_spin.value()
            new_x2 = x2_spin.value()
            init = self.initial_settings[key]
            slope = (init['y2'] - init['y1']) / (init['x2'] - init['x1']) if init['x2'] != init['x1'] else 0
            if sender in (x1_spin, x2_spin):
                if sender == x1_spin:
                    y1_spin.setValue(init['y1'] + slope * (new_x1 - init['x1']))
                elif sender == x2_spin:
                    y2_spin.setValue(init['y2'] + slope * (new_x2 - init['x2']))
            else:
                y1_spin.setValue(init['y1'] + slope * (new_x1 - init['x1']))
                y2_spin.setValue(init['y2'] + slope * (new_x2 - init['x1']))
        self.ox_label.setText(self.get_oxidation_label_text())
        self.red_label.setText(self.get_reduction_label_text())

    def on_apply(self):
        """
        Zastosowanie zmian - przeliczenie ustawień i emisja sygnału.
        """
        for key, x1_spin, y1_spin, x2_spin, y2_spin in [
            ("oxidation", self.ox_x1, self.ox_y1, self.ox_x2, self.ox_y2),
            ("reduction", self.red_x1, self.red_y1, self.red_x2, self.red_y2),
        ]:
            init = self.initial_settings[key]
            new_x1 = x1_spin.value()
            new_x2 = x2_spin.value()
            slope = (init['y2'] - init['y1']) / (init['x2'] - init['x1']) if init['x2'] != init['x1'] else 0
            if new_x1 != init['x1']:
                y1_spin.setValue(init['y1'] + slope * (new_x1 - init['x1']))
            if new_x2 != init['x2']:
                y2_spin.setValue(init['y2'] + slope * (new_x2 - init['x2']))
        new_settings = self.get_settings()
        self.baseline_applied.emit(new_settings)
        self.initial_settings = {
            'oxidation': new_settings['oxidation'].copy(),
            'reduction': new_settings['reduction'].copy()
        }
        self.update_preview_labels()

    def get_oxidation_label_text(self):
        """Zwraca sformatowany tekst prezentujący ustawienia linii bazowej dla utlenienia."""
        if not hasattr(self, 'ox_x1'):
            return "Punkty nie są jeszcze zainicjalizowane"
        return f"x1: {self.ox_x1.value():.3f}, y1: {self.ox_y1.value():.3f} | x2: {self.ox_x2.value():.3f}, y2: {self.ox_y2.value():.3f}"

    def get_reduction_label_text(self):
        """Zwraca sformatowany tekst prezentujący ustawienia linii bazowej dla redukcji."""
        if not hasattr(self, 'red_x1'):
            return "Punkty nie są jeszcze zainicjalizowane"
        return f"x1: {self.red_x1.value():.3f}, y1: {self.red_y1.value():.3f} | x2: {self.red_x2.value():.3f}, y2: {self.red_y2.value():.3f}"

    def accept(self):
        """Przy zatwierdzaniu zmian stosuje ustawienia oraz zamyka dialog."""
        self.on_apply()
        super().accept()

    def get_settings(self):
        """
        Zwraca aktualne ustawienia linii bazowej.

        Returns:
            dict: Ustawienia dla utlenienia i redukcji.
        """
        return {
            'oxidation': {
                'x1': self.ox_x1.value(),
                'y1': self.ox_y1.value(),
                'x2': self.ox_x2.value(),
                'y2': self.ox_y2.value(),
            },
            'reduction': {
                'x1': self.red_x1.value(),
                'y1': self.red_y1.value(),
                'x2': self.red_x2.value(),
                'y2': self.red_y2.value(),
            }
        }

    def update_preview_labels(self):
        """Aktualizuje etykiety podglądu dla obu ustawień linii bazowej."""
        settings = self.get_settings()
        self.ox_preview_label.setText(
            f"x1: {settings['oxidation']['x1']:.3f}, y1: {settings['oxidation']['y1']:.3f} | x2: {settings['oxidation']['x2']:.3f}, y2: {settings['oxidation']['y2']:.3f}"
        )
        self.red_preview_label.setText(
            f"x1: {settings['reduction']['x1']:.3f}, y1: {settings['reduction']['y1']:.3f} | x2: {settings['reduction']['x2']:.3f}, y2: {settings['reduction']['y2']:.3f}"
        )
