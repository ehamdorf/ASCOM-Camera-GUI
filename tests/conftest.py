"""
Pytest environment configuration and shared fixtures.
"""

import os
import sys
from pathlib import Path

import pytest

# Prevent popups during testing.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Make the project root importable.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PyQt5 import QtWidgets

# Avoid Astropy trying to download IERS.
try:
    from astropy.utils import iers

    iers.conf.auto_download = False
except Exception:
    pass


@pytest.fixture(scope="session")
def qapp():
    """
    QApplication instance for Qt tests.
    """
    app = QtWidgets.QApplication.instance()
    if app is None:
        app = QtWidgets.QApplication(["nirvana_hs_tests"])
    yield app
