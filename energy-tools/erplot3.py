#! /usr/bin/env python3
# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4; -*-
#
# Copyright (C) 2014 Canonical Ltd.
# Author: Shih-Yuan Lee (FourDollars) <sylee@canonical.com>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import unittest
from sysinfo import SysInfo
from logging import debug, warning

__all__ = [
        "ErPLot3",
        "ErPLot3_2014",
        "ErPLot3_2016"]
    
class ErPLot3:
    """ErP Lot 3 calculator"""
    def __init__(self, sysinfo):
        self.sysinfo = sysinfo

    def calculate(self):
        print("\nErP Lot 3 from 1 July 2014:\n")
        early = ErPLot3_2014(self.sysinfo)
        if early.check_special_case():
            if early.computer_type == 3:
                print("\033[1;91mWARNING\033[0m: If discrete graphics card(s) providing total frame buffer bandwidths above 225 GB/s, use the requirement from 1 January 2016 instead.")
            else:
                print("\033[1;91mWARNING\033[0m: If discrete graphics card(s) providing total frame buffer bandwidths above 320 GB/s and\n         a PSU with a rated output power of at least 1000W, use the requirement from 1 January 2016 instead.")

        if self._verify_s3_s4(early):
            self._calculate(early)
        print("\nErP Lot 3 from 1 January 2016:\n")
        late = ErPLot3_2016(self.sysinfo)
        if self._verify_s3_s4(late):
            self._calculate(late)

    def _calculate(self, inst):
        if self.sysinfo.computer_type == 3:
            categories = ('A', 'B', 'C')
        else:
            categories = ('A', 'B', 'C', 'D')
        candidates = []
        for category in categories:
            ret = inst.category(category)
            if ret >= 0:
                candidates.append((category, ret))
        for cat, meet in candidates:
            if meet:
                print("  Category %s:" % cat)
            else:
                print("  Category %s if a discrete graphics card (dGfx) meeting the G3 (with FB Data Width > 128-bit), G4, G5, G6 or G7 classification:" % cat)
            TEC_BASE = inst.get_TEC_BASE(cat)
            TEC_MEMORY = inst.get_TEC_MEMORY(cat)
            TEC_STORAGE = inst.get_TEC_STORAGE()
            TEC_TV_TUNER = inst.get_TEC_TV_TUNER()
            TEC_AUDIO = inst.get_TEC_AUDIO()
            debug("TEC_BASE = %s, TEC_MEMORY = %s, TEC_STORAGE = %s, TEC_TV_TUNER = %s, TEC_AUDIO = %s" %
                (TEC_BASE, TEC_MEMORY, TEC_STORAGE, TEC_TV_TUNER, TEC_AUDIO))
            if inst.discrete_graphics_cards == 0:
                TEC_GRAPHICS = 0
                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_STORAGE + TEC_TV_TUNER + TEC_AUDIO + TEC_GRAPHICS
                debug("TEC_GRAPHICS = %s" % TEC_GRAPHICS)
                self._verifying(inst, E_TEC_MAX)
            elif inst.discrete_graphics_cards == 1:
                for gpu in ('G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7'):
                    TEC_GRAPHICS = inst.get_TEC_GRAPHICS(gpu)
                    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_STORAGE + TEC_TV_TUNER + TEC_AUDIO + TEC_GRAPHICS
                    debug("TEC_GRAPHICS = %s" % TEC_GRAPHICS)
                    self._verifying(inst, E_TEC_MAX, gpu)
            else:
                print("    No console output because of more than one discrete graphics card.")

    def _verify_s3_s4(self, inst):
        if self.sysinfo.computer_type != 3:
            if inst.sleep > 5:
                print("      Fail because P_SLEEP (%s) > 5.0" % inst.sleep)
                return False
            elif inst.sleep_wol > 5.7:
                print("      Fail because P_SLEEP_WOL (%s) > 5.7" % inst.sleep_wol)
                return False
        else:
            if inst.sleep > 3:
                print("      Fail because P_SLEEP (%s) > 3.0" % inst.sleep)
                return False
            elif inst.sleep_wol > 3.7:
                print("      Fail because P_SLEEP_WOL (%s) > 3.7" % inst.sleep_wol)
                return False

        if inst.off > 1.0:
            print("      Fail because P_OFF (%s) > 1.0" % inst.off)
            return False
        elif inst.off_wol > 1.7:
            print("      Fail because P_OFF (%s) > 1.7" % inst.off_wol)
            return False

        return True

    def _verifying(self, inst, E_TEC_MAX, gpu=None):
        msg = ''
        if gpu:
            if gpu == 'G1':
                msg = "G1 (FB_BW <= 16)"
            elif gpu == 'G2':
                msg = "G2 (16 < FB_BW <= 32)"
            elif gpu == 'G3':
                msg = "G3 (32 < FB_BW <= 64)"
            elif gpu == 'G4':
                msg = "G4 (64 < FB_BW <= 96)"
            elif gpu == 'G5':
                msg = "G5 (96 < FB_BW <= 128)"
            elif gpu == 'G6':
                msg = "G6 (FB_BW > 128 (with FB Data Width < 192-bit))"
            elif gpu == 'G7':
                msg = "G7 (FB_BW > 128 (with FB Data Width >= 192-bit))"
        E_TEC = inst.get_E_TEC()
        E_TEC_WOL = inst.get_E_TEC_WOL()
        if E_TEC > E_TEC_MAX:
            if msg:
                print("      %s (E_TEC) > %s (E_TEC_MAX) for %s, FAIL" % (E_TEC, E_TEC_MAX, msg))
            else:
                print("      %s (E_TEC) > %s (E_TEC_MAX), FAIL" % (E_TEC, E_TEC_MAX))
            return
        elif E_TEC_WOL > E_TEC_MAX:
            if msg:
                print("      %s (E_TEC_WOL) > %s (E_TEC_MAX) for %s, FAIL" % (E_TEC_WOL, E_TEC_MAX, msg))
            else:
                print("      %s (E_TEC_WOL) > %s (E_TEC_MAX), FAIL" % (E_TEC_WOL, E_TEC_MAX))
            return
        else:
            if gpu:
                print("      %s (E_TEC) <= %s (E_TEC_MAX), and %s (E_TEC_WOL) <= %s (E_TEC_MAX) for %s, PASS" %
                        (E_TEC, E_TEC_MAX, E_TEC_WOL, E_TEC_MAX, msg))
            else:
                print("      %s (E_TEC) <= %s (E_TEC_MAX), and %s (E_TEC_WOL) <= %s (E_TEC_MAX), PASS" %
                        (E_TEC, E_TEC_MAX, E_TEC_WOL, E_TEC_MAX))


class ErPLot3_2014:
    """ErP Lot 3 calculator from 1 July 2014"""
    def __init__(self, sysinfo):
        debug("=== ErP Lot 3 from 1 July 2014 ===")
        self.computer_type = sysinfo.computer_type
        self.cpu_core = sysinfo.cpu_core
        self.cpu_clock = sysinfo.cpu_clock
        self.memory_size = sysinfo.mem_size
        self.disk_number = sysinfo.disk_num
        self.discrete_graphics_cards = sysinfo.discrete_gpu_num
        self.discrete_audio = sysinfo.audio
        self.tv_tuner = sysinfo.tvtuner
        self.off = sysinfo.off
        self.off_wol = sysinfo.off_wol
        self.sleep = sysinfo.sleep
        self.sleep_wol = sysinfo.sleep_wol
        self.idle = sysinfo.short_idle

    def check_special_case(self):
        if self.computer_type == 1 or self.computer_type == 2:
            if self.cpu_core >= 6 and self.memory_size >= 16:
                return True
        elif self.computer_type == 3:
            if self.cpu_core >= 4 and self.memory_size >= 16:
                return True
        else:
            raise Exception("Should not be here.")
        return False
    def category(self, category):
        if self.computer_type == 1 or self.computer_type == 2:
            if category == 'D':
                if self.cpu_core >= 4:
                    if self.memory_size >= 4:
                        return 1
                    elif self.discrete_graphics_cards >= 1:
                        return 0
                    else:
                        return -1
                else:
                    return -1
            elif category == 'C':
                if self.cpu_core >= 3:
                    if self.memory_size >= 2 or self.discrete_graphics_cards >= 1:
                        return 1
                    else:
                        return -1
                else:
                    return -1
            elif category == 'B':
                if self.cpu_core >= 2 and self.memory_size >= 2:
                    return 1
                else:
                    return -1
            elif category == 'A':
                return 1
            else:
                raise Exception('Should not be here.')
        elif self.computer_type == 3:
            if category == 'C':
                if self.cpu_core >= 2 and self.memory_size >= 2 and self.discrete_graphics_cards >= 1:
                    return 0
                else:
                    return -1
            elif category == 'B':
                if self.discrete_graphics_cards >= 1:
                    return 1
                else:
                    return -1
            elif category == 'A':
                return 1
            else:
                raise Exception('Should not be here.')
        else:
            raise Exception('Should not be here.')

    def get_E_TEC(self):
        if self.computer_type == 3:
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

        E_TEC = ((T_OFF * self.off) + (T_SLEEP * self.sleep) + (T_IDLE * self.idle)) * 8760 / 1000
        return E_TEC

    def get_T_values(self):
        if self.computer_type == 3:
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)
        return (T_OFF, T_SLEEP, T_IDLE)

    def get_E_TEC_WOL(self):
        if self.computer_type == 3:
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

        E_TEC = ((T_OFF * self.off_wol) + (T_SLEEP * self.sleep_wol) + (T_IDLE * self.idle)) * 8760 / 1000
        return E_TEC

    def get_TEC_BASE(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'A':
                return 36
            elif category == 'B':
                return 48
            elif category == 'C':
                return 80.5
        # Desktop
        else:
            if category == 'A':
                return 133
            elif category == 'B':
                return 158
            elif category == 'C':
                return 188
            elif category == 'D':
                return 211
        raise Exception("Should not be here.")

    def get_TEC_GRAPHICS(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'G1':
                return 12
            elif category == 'G2':
                return 20
            elif category == 'G3':
                return 26
            elif category == 'G4':
                return 37
            elif category == 'G5':
                return 49
            elif category == 'G6':
                return 61
            elif category == 'G7':
                return 113
        # Desktop
        else:
            if category == 'G1':
                return 34
            elif category == 'G2':
                return 54
            elif category == 'G3':
                return 69
            elif category == 'G4':
                return 100
            elif category == 'G5':
                return 133
            elif category == 'G6':
                return 166
            elif category == 'G7':
                return 225
        raise Exception("Should not be here.")

    def additional_TEC_GRAPHICS(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'G1':
                return 7
            elif category == 'G2':
                return 12
            elif category == 'G3':
                return 15
            elif category == 'G4':
                return 22
            elif category == 'G5':
                return 29
            elif category == 'G6':
                return 36
            elif category == 'G7':
                return 66
        # Desktop
        else:
            if category == 'G1':
                return 20
            elif category == 'G2':
                return 32
            elif category == 'G3':
                return 41
            elif category == 'G4':
                return 59
            elif category == 'G5':
                return 78
            elif category == 'G6':
                return 98
            elif category == 'G7':
                return 133
        raise Exception("Should not be here.")

    def get_TEC_TV_TUNER(self):
        if self.computer_type == 3:
            if self.tv_tuner:
                return 2.1
            else:
                return 0
        else:
            if self.tv_tuner:
                return 15
            else:
                return 0

    def get_TEC_AUDIO(self):
        if self.computer_type == 3:
            return 0
        else:
            if self.discrete_audio:
                return 15
            else:
                return 0

    def get_TEC_MEMORY(self, category):
        # Notebook
        if self.computer_type == 3:
            if self.memory_size > 4:
                return 0.4 * (self.memory_size - 4)
            else:
                return 0
        # Desktop
        else:
            if category == 'D':
                return 1.0 * (self.memory_size - 4)
            else:
                if self.memory_size > 2:
                    return 1.0 * (self.memory_size - 2)
                else:
                    return 0

    def get_TEC_STORAGE(self):
        if self.disk_number == 0:
            return 0
        if self.computer_type == 3:
            return 3 * (self.disk_number - 1)
        else:
            return 25 * (self.disk_number - 1)

class ErPLot3_2016(ErPLot3_2014):
    """ErP Lot 3 calculator from 1 January 2016"""
    def __init__(self, sysinfo):
        ErPLot3_2014.__init__(self, sysinfo)
        debug("=== ErP Lot 3 from 1 January 2016 ===")

    def get_TEC_BASE(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'A':
                return 27
            elif category == 'B':
                return 36
            elif category == 'C':
                return 60.5
        # Desktop
        else:
            if category == 'A':
                return 94
            elif category == 'B':
                return 112
            elif category == 'C':
                return 134
            elif category == 'D':
                return 150
        raise Exception("Should not be here.")

    def get_TEC_GRAPHICS(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'G1':
                return 7
            elif category == 'G2':
                return 11
            elif category == 'G3':
                return 13
            elif category == 'G4':
                return 20
            elif category == 'G5':
                return 27
            elif category == 'G6':
                return 33
            elif category == 'G7':
                return 61
        # Desktop
        else:
            if category == 'G1':
                return 18
            elif category == 'G2':
                return 30
            elif category == 'G3':
                return 38
            elif category == 'G4':
                return 54
            elif category == 'G5':
                return 72
            elif category == 'G6':
                return 90
            elif category == 'G7':
                return 122
        raise Exception("Should not be here.")

    def additional_TEC_GRAPHICS(self, category):
        # Notebook
        if self.computer_type == 3:
            if category == 'G1':
                return 4
            elif category == 'G2':
                return 6
            elif category == 'G3':
                return 8
            elif category == 'G4':
                return 12
            elif category == 'G5':
                return 16
            elif category == 'G6':
                return 20
            elif category == 'G7':
                return 36
        # Desktop
        else:
            if category == 'G1':
                return 11
            elif category == 'G2':
                return 17
            elif category == 'G3':
                return 22
            elif category == 'G4':
                return 32
            elif category == 'G5':
                return 42
            elif category == 'G6':
                return 53
            elif category == 'G7':
                return 72
        raise Exception("Should not be here.")


class TestErPLot3(unittest.TestCase):
    def setUp(self):
        self.sysinfo = SysInfo({
            'Product Type': 1,
            'Computer Type': 3,
            'CPU Clock': 2.0,
            'CPU Cores': 2,
            'Discrete Audio': False,
            'Discrete Graphics': False,
            'Discrete Graphics Cards': 0,
            'Switchable Graphics': False,
            'Disk Number': 1,
            'Display Diagonal': 14,
            'Display Height': 768,
            'Display Width': 1366,
            'Enhanced Display': False,
            'Gigabit Ethernet': 1,
            'Memory Size': 8,
            'TV Tuner': False,
            'Off Mode': 1.0,
            'Off Mode with WOL': 1.0,
            'Sleep Mode': 1.7,
            'Sleep Mode with WOL': 1.7,
            'Long Idle Mode': 8.0,
            'Short Idle Mode': 10.0})

    def tearDown(self):
        self.sysinfo = None

    def test_desktop_category(self):
        self.sysinfo.computer_type = 1
        self.sysinfo.cpu_core = 4
        self.sysinfo.mem_size = 4
        self.sysinfo.discrete_gpu_num = 1

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 1)

        inst.memory_size = 2
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 0)

        inst.discrete_graphics_cards = 0
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), -1)

        inst.memory_size = 1
        inst.discrete_graphics_cards = 1
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 0)

        inst.cpu_core = 2
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), -1)
        self.assertEqual(inst.category('D'), -1)

        inst.memory_size = 2
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), -1)
        self.assertEqual(inst.category('D'), -1)

    def test_notebook_category(self):
        self.sysinfo.computer_type = 3
        self.sysinfo.cpu_core = 2
        self.sysinfo.mem_size = 2
        self.sysinfo.discrete_gpu_num = 1

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 0)
        self.assertRaises(Exception, inst.category, 'D')

        inst.memory_size = 1
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), -1)

        inst.discrete_graphics_cards = 0
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), -1)

    def test_TEC_BASE(self):
        self.sysinfo.computer_type = 1

        inst = ErPLot3_2014(self.sysinfo)

        self.assertEqual(inst.get_TEC_BASE('A'), 133)
        self.assertEqual(inst.get_TEC_BASE('B'), 158)
        self.assertEqual(inst.get_TEC_BASE('C'), 188)
        self.assertEqual(inst.get_TEC_BASE('D'), 211)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_BASE('A'), 36)
        self.assertEqual(inst.get_TEC_BASE('B'), 48)
        self.assertEqual(inst.get_TEC_BASE('C'), 80.5)

    def test_TEC_TV_TUNER(self):
        self.sysinfo.computer_type = 1
        self.sysinfo.tvtuner = True

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.get_TEC_TV_TUNER(), 15)
        inst.tv_tuner = False
        self.assertEqual(inst.get_TEC_TV_TUNER(), 0)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_TV_TUNER(), 0)
        inst.tv_tuner = True
        self.assertEqual(inst.get_TEC_TV_TUNER(), 2.1)

    def test_TEC_AUDIO(self):
        self.sysinfo.computer_type = 1
        self.sysinfo.audio = True

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.get_TEC_AUDIO(), 15)
        inst.discrete_audio = False
        self.assertEqual(inst.get_TEC_AUDIO(), 0)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_AUDIO(), 0)
        inst.discrete_audio = True
        self.assertEqual(inst.get_TEC_AUDIO(), 0)

    def test_TEC_MEMORY(self):
        self.sysinfo.computer_type = 1
        self.sysinfo.mem_size = 8

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.get_TEC_MEMORY('A'), 6)
        self.assertEqual(inst.get_TEC_MEMORY('B'), 6)
        self.assertEqual(inst.get_TEC_MEMORY('C'), 6)
        self.assertEqual(inst.get_TEC_MEMORY('D'), 4)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_MEMORY('A'), 1.6)
        self.assertEqual(inst.get_TEC_MEMORY('B'), 1.6)
        self.assertEqual(inst.get_TEC_MEMORY('C'), 1.6)
        self.assertEqual(inst.get_TEC_MEMORY('D'), 1.6)

    def test_TEC_STORAGE(self):
        self.sysinfo.computer_type = 1
        self.sysinfo.disk_num = 1

        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.get_TEC_STORAGE(), 0)
        inst.disk_number = 2
        self.assertEqual(inst.get_TEC_STORAGE(), 25)

        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_STORAGE(), 3)
        inst.disk_number = 1
        self.assertEqual(inst.get_TEC_STORAGE(), 0)

if __name__ == '__main__':
    unittest.main()
