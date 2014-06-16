#! /usr/bin/env python
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
from sysinfo import SysInfo as SysInfo

class ErPLot3_2014:
    """ErP Lot 3 calculator from 1 July 2014"""
    def __init__(self, sysinfo):
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

    def get_E_TEC_wol(self):
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

    def get_MEMORY(self, category):
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
    def test_desktop_category(self):
        sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                cpu_core=4,
                mem_size=4,
                discrete_gpu_num=1)
        inst = ErPLot3_2014(sysinfo)
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
        sysinfo = SysInfo(
                auto=True,
                computer_type=3,
                cpu_core=2,
                mem_size=2,
                discrete_gpu_num=1)
        inst = ErPLot3_2014(sysinfo)
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
        sysinfo = SysInfo(auto=True, computer_type=1)
        inst = ErPLot3_2014(sysinfo)
        self.assertEqual(inst.get_TEC_BASE('A'), 133)
        self.assertEqual(inst.get_TEC_BASE('B'), 158)
        self.assertEqual(inst.get_TEC_BASE('C'), 188)
        self.assertEqual(inst.get_TEC_BASE('D'), 211)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_BASE('A'), 36)
        self.assertEqual(inst.get_TEC_BASE('B'), 48)
        self.assertEqual(inst.get_TEC_BASE('C'), 80.5)

    def test_TEC_TV_TUNER(self):
        sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                tvtuner=True)
        inst = ErPLot3_2014(sysinfo)
        self.assertEqual(inst.get_TEC_TV_TUNER(), 15)
        inst.tv_tuner = False
        self.assertEqual(inst.get_TEC_TV_TUNER(), 0)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_TV_TUNER(), 0)
        inst.tv_tuner = True
        self.assertEqual(inst.get_TEC_TV_TUNER(), 2.1)

    def test_TEC_AUDIO(self):
        sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                audio=True)
        inst = ErPLot3_2014(sysinfo)
        self.assertEqual(inst.get_TEC_AUDIO(), 15)
        inst.discrete_audio = False
        self.assertEqual(inst.get_TEC_AUDIO(), 0)
        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_AUDIO(), 0)
        inst.discrete_audio = True
        self.assertEqual(inst.get_TEC_AUDIO(), 0)

    def test_TEC_MEMORY(self):
        sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                mem_size=8)
        inst = ErPLot3_2014(sysinfo)
        self.assertEqual(inst.get_MEMORY('A'), 6)
        self.assertEqual(inst.get_MEMORY('B'), 6)
        self.assertEqual(inst.get_MEMORY('C'), 6)
        self.assertEqual(inst.get_MEMORY('D'), 4)
        inst.computer_type = 3
        self.assertEqual(inst.get_MEMORY('A'), 1.6)
        self.assertEqual(inst.get_MEMORY('B'), 1.6)
        self.assertEqual(inst.get_MEMORY('C'), 1.6)
        self.assertEqual(inst.get_MEMORY('D'), 1.6)

    def test_TEC_STORAGE(self):
        sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                disk_num=1)
        inst = ErPLot3_2014(sysinfo)
        self.assertEqual(inst.get_TEC_STORAGE(), 0)
        inst.disk_number = 2
        self.assertEqual(inst.get_TEC_STORAGE(), 25)

        inst.computer_type = 3
        self.assertEqual(inst.get_TEC_STORAGE(), 3)
        inst.disk_number = 1
        self.assertEqual(inst.get_TEC_STORAGE(), 0)

if __name__ == '__main__':
    unittest.main()
