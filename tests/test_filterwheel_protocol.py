"""
Tests for CFW-10 filter-wheel packet construction and serial protocol.
"""

import threading
import time
from unittest.mock import patch

import pytest

import NIRvana_HS_PyQt5_GUI as gui


def make_cfw():
    """
    Create a CFW10 instance without initialising a real serial port.
    """
    with patch("serial.Serial.__init__", return_value=None):
        fw = gui.CFW10(port=gui.CFW10_PORT, baudrate=gui.CFW10_BAUDRATE)

    fw.pkt = [0xA5, 0x03, 0x00, 0x00, 0x00, 0x00]
    fw.fname = gui.CFW10_FILTER_NAMES_SHORT
    fw._lock = threading.RLock()
    return fw


def test_filter_name_tables_have_ten_entries():
    """
    Both short and long filter tables should have 10 positions (CFW10 has 10 positions).
    """
    assert len(gui.CFW10_FILTER_NAMES_SHORT) == 10
    assert len(gui.CFW10_FILTER_NAMES_LONG) == 10


def test_cmd_builds_checksum():
    """
    The command packet should include a simple checksum.
    """
    fw = make_cfw()

    fw._cmd(17, 5)

    assert fw.pkt[0] == 0xA5
    assert fw.pkt[1] == 0x03
    assert fw.pkt[2] == 17
    assert fw.pkt[3] == 5
    assert fw.pkt[4] == 0
    assert fw.pkt[5] == sum(fw.pkt[:5])


def test_cmd_high_byte():
    """
    Values above 255 should be split into low/high bytes.
    """
    fw = make_cfw()

    fw._cmd(17, 300)

    assert fw.pkt[3] == 300 % 256
    assert fw.pkt[4] == 300 // 256
    assert fw.pkt[5] == sum(fw.pkt[:5])


def test_choose_filter_clamps_invalid_positions():
    """
    Invalid filter positions should be clamped to position 1.
    """
    fw = make_cfw()

    fw._choose_filter(0)
    assert fw.pkt[2] == 17
    assert fw.pkt[3] == 1

    fw._choose_filter(11)
    assert fw.pkt[3] == 1

    fw._choose_filter(10)
    assert fw.pkt[3] == 10


def test_wait_motor_off_timeout():
    """
    _wait_motor_off should return -1 if the motor does not stop within the timeout.
    """
    fw = make_cfw()
    
    # Mock _get_status to always return a "motor running" state (e.g., 0x10)
    with patch.object(fw, "_get_status", return_value=0x10):
        # Temporarily reduce timeout for the test
        original_timeout = gui.CFW10_MOTOR_OFF_TIMEOUT_S
        gui.CFW10_MOTOR_OFF_TIMEOUT_S = 0.1
        
        try:
            result = fw._wait_motor_off()
            assert result == -1
        finally:
            gui.CFW10_MOTOR_OFF_TIMEOUT_S = original_timeout


def test_wait_motor_off_comm_error():
    """
    _wait_motor_off should return -2 on communication error.
    """
    fw = make_cfw()
    
    # 0xFE indicates a comm error in _get_status
    with patch.object(fw, "_get_status", return_value=0xFE):
        result = fw._wait_motor_off()
        assert result == -2
