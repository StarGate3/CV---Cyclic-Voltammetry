import numpy as np
import pytest

from utils import compute_zero_crossings, compute_intersections


def test_compute_zero_crossings_single_crossing():
    x = np.array([0.0, 1.0])
    y = np.array([-1.0, 1.0])
    zeros = compute_zero_crossings(x, y, 0.0, 1.0)
    assert len(zeros) == 1
    x0, y0 = zeros[0]
    assert pytest.approx(0.5) == x0
    assert pytest.approx(0.0) == y0


def test_compute_intersections_simple_lines():
    x = np.array([0.0, 1.0])
    line1 = np.array([0.0, 1.0])
    line2 = np.array([1.0, 0.0])
    intersections = compute_intersections(x, line1, line2, 0.0, 1.0)
    assert len(intersections) == 1
    xi, yi = intersections[0]
    assert pytest.approx(0.5) == xi
    assert pytest.approx(0.5) == yi
