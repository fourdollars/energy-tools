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
            if self.audio:
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
        if self.computer_type == 3:
            return 3 * (self.disk_number - 1)
        else:
            return 25 * (self.disk_number - 1)

class ErPLot3_2016(ErPLot3_2014):
    """ErP Lot 3 calculator from 1 January 2016"""
    pass

class TestErPLot3(unittest.TestCase):
    def test_desktop_category(self):
        self.sysinfo = SysInfo(
                auto=True,
                computer_type=1,
                cpu_core=4,
                mem_size=4,
                discrete_gpu_num=1)
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 1)

        self.sysinfo.mem_size = 2
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 0)

        self.sysinfo.discrete_gpu_num=0
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), -1)

        self.sysinfo.discrete_gpu_num=1
        self.sysinfo.mem_size = 1
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), 1)
        self.assertEqual(inst.category('D'), 0)

        self.sysinfo.cpu_core=2
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), -1)
        self.assertEqual(inst.category('D'), -1)

        self.sysinfo.mem_size = 2
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), -1)
        self.assertEqual(inst.category('D'), -1)

    def test_notebook_category(self):
        self.sysinfo = SysInfo(
                auto=True,
                computer_type=3,
                cpu_core=2,
                mem_size=2,
                discrete_gpu_num=1)
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), 0)
        self.assertRaises(Exception, inst.category, 'D')

        self.sysinfo.mem_size = 1
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), 1)
        self.assertEqual(inst.category('C'), -1)

        self.sysinfo.discrete_gpu_num=0
        inst = ErPLot3_2014(self.sysinfo)
        self.assertEqual(inst.category('A'), 1)
        self.assertEqual(inst.category('B'), -1)
        self.assertEqual(inst.category('C'), -1)

    def tearDown(self):
        #print(">>> Core: %d, Memory: %d, GPU: %d" % (self.sysinfo.cpu_core, self.sysinfo.mem_size, self.sysinfo.discrete_gpu_num))
        pass

if __name__ == '__main__':
    unittest.main()
