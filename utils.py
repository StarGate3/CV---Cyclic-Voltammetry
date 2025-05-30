"""
Moduł utils.py
-----------------
Zawiera funkcje pomocnicze, np. obliczanie punktów przecięcia krzywych
oraz wykrywanie miejsc zerowych pojedynczej krzywej.
"""

import numpy as np

def compute_intersections(x: np.ndarray,
                          curve1: np.ndarray,
                          curve2: np.ndarray,
                          range_min: float,
                          range_max: float) -> list[tuple[float, float]]:
    """
    Oblicza punkty przecięcia dwóch krzywych (curve1 oraz curve2)
    na przedziale [range_min, range_max] metodą wykrywania zmiany znaku różnicy.

    Parameters:
        x (ndarray): Wartości osi x.
        curve1 (ndarray): Wartości pierwszej krzywej.
        curve2 (ndarray): Wartości drugiej krzywej.
        range_min (float): Dolna granica przedziału.
        range_max (float): Górna granica przedziału.

    Returns:
        list: Lista krotek (x, y) oznaczających punkty przecięcia.
    """
    mask = (x >= range_min) & (x <= range_max)
    if not np.any(mask):
        return []
    x_range = x[mask]
    y1_range = curve1[mask]
    y2_range = curve2[mask]
    d = y1_range - y2_range
    intersections: list[tuple[float, float]] = []
    for i in range(len(d) - 1):
        # Dokładne trafienie na zero
        if d[i] == 0:
            intersections.append((x_range[i], y1_range[i]))
        # Zmiana znaku między kolejnymi punktami
        elif d[i] * d[i + 1] < 0:
            r = d[i] / (d[i] - d[i + 1])
            x_int = x_range[i] + r * (x_range[i + 1] - x_range[i])
            y_int = y1_range[i] + r * (y1_range[i + 1] - y1_range[i])
            intersections.append((x_int, y_int))
    return intersections


def compute_zero_crossings(x: np.ndarray,
                           curve: np.ndarray,
                           range_min: float,
                           range_max: float) -> list[tuple[float, float]]:
    """
    Oblicza miejsca zerowe krzywej curve na przedziale [range_min, range_max]
    przez detekcję zmiany znaku i interpolację liniową.

    Parameters:
        x (ndarray): Wartości osi x.
        curve (ndarray): Wartości krzywej.
        range_min (float): Dolna granica przedziału.
        range_max (float): Górna granica przedziału.

    Returns:
        list: Lista krotek (x_zero, 0.0) oznaczających przybliżone miejsca zerowe.
    """
    mask = (x >= range_min) & (x <= range_max)
    if not np.any(mask):
        return []
    x_range = x[mask]
    y = curve[mask]
    zeros: list[tuple[float, float]] = []
    for i in range(len(y) - 1):
        # Trafienie dokładnie na zero
        if y[i] == 0:
            zeros.append((x_range[i], 0.0))
        # Zmiana znaku między sąsiednimi próbkami
        elif y[i] * y[i + 1] < 0:
            r = y[i] / (y[i] - y[i + 1])
            x0 = x_range[i] + r * (x_range[i + 1] - x_range[i])
            zeros.append((x0, 0.0))
    return zeros
