# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
#
# Copyright (C) 2018 Canonical Ltd.
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

from logging import debug, warning
from math import tanh

class EnergyStar70:
    """Energy Star 7.0 calculator"""
    def __init__(self, sysinfo):
        self.sysinfo = sysinfo
        debug("=== Energy Star 7.0 ===")

    # Requirements for Desktop, Integrated Desktop, and Notebook Computers
    def equation_one(self):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, Thin Client and Notebook Computers"""
        (P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE) = self.sysinfo.get_power_consumptions()
        if self.sysinfo.product_type == 4 or self.sysinfo.computer_type == 3:
            (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.25, 0.35, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

        E_TEC = ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_LONG_IDLE * T_LONG_IDLE) + (P_SHORT_IDLE * T_SHORT_IDLE)) * 8760 / 1000

        debug("T_OFF = %s, T_SLEEP = %s, T_LONG_IDLE = %s, T_SHORT_IDLE = %s" % (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE))
        debug("P_OFF = %s, P_SLEEP = %s, P_LONG_IDLE = %s, P_SHORT_IDLE = %s" % (P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE))

        return E_TEC

    def equation_two(self, gpu_category, FB_BW=0):
        """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop, and Notebook Computers"""
        (core, clock, memory, disk) = self.sysinfo.get_basic_info()

        P = core * clock
        debug("P = %s" % (P))

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            if P <= 3:
                TEC_BASE = 69.0
            elif self.sysinfo.discrete:
                if P <= 9:
                    TEC_BASE = 115.0
                else:
                    TEC_BASE = 135.0
            else:
                if P <= 6:
                    TEC_BASE = 112.0
                elif P <= 7:
                    TEC_BASE = 120.0
                else:
                    TEC_BASE = 135.0
        else:
            if P <= 2:
                TEC_BASE = 6.5
            elif P <= 8:
                TEC_BASE = 8.0
            else:
                TEC_BASE = 14.0

        if self.sysinfo.computer_type != 3:
            TEC_MEMORY = 0.8 * memory
        else:
            TEC_MEMORY = 2.4 + 0.294 * memory

        if self.sysinfo.switchable:
            if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
                TEC_SWITCHABLE = 0.5 * 36
            else:
                TEC_SWITCHABLE = 0
            TEC_GRAPHICS = 0
        else:
            TEC_SWITCHABLE = 0
            if self.sysinfo.discrete:
                if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
                    if gpu_category == 'G1':
                        TEC_GRAPHICS = 36
                    elif gpu_category == 'G2':
                        TEC_GRAPHICS = 51
                    elif gpu_category == 'G3':
                        TEC_GRAPHICS = 64
                    elif gpu_category == 'G4':
                        TEC_GRAPHICS = 83
                    elif gpu_category == 'G5':
                        TEC_GRAPHICS = 105 
                    elif gpu_category == 'G6':
                        TEC_GRAPHICS = 115 
                    elif gpu_category == 'G7':
                        TEC_GRAPHICS = 130 
                else:
                    TEC_GRAPHICS = 29.3 * tanh(0.0038 * FB_BW - 0.137) + 13.4
            else:
                TEC_GRAPHICS = 0

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * self.sysinfo.get_1glan_num()
        else:
            TEC_EEE = 0

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            TEC_STORAGE = 26 * (disk - 1)
        else:
            TEC_STORAGE = 2.6 * (disk - 1)

        if self.sysinfo.computer_type != 1:
            (EP, r, A) = self.equation_three()

        if self.sysinfo.computer_type == 2 or self.sysinfo.product_type == 4:
            TEC_INT_DISPLAY = 8.76 * 0.35 * (1 + EP) * (4 * r + 0.05 * A)
        elif self.sysinfo.computer_type == 3:
            TEC_INT_DISPLAY = 8.76 * 0.30 * (1 + EP) * (0.43 * r + 0.0263 * A)
        else:
            TEC_INT_DISPLAY = 0
        debug("TEC_BASE = %s, TEC_MEMORY = %s, TEC_GRAPHICS = %s, TEC_SWITCHABLE = %s, TEC_EEE = %s, TEC_STORAGE = %s, TEC_INT_DISPLAY = %s" % (TEC_BASE, TEC_MEMORY, TEC_GRAPHICS, TEC_SWITCHABLE, TEC_EEE, TEC_STORAGE, TEC_INT_DISPLAY))

        return TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE + TEC_INT_DISPLAY + TEC_SWITCHABLE + TEC_EEE

    def equation_three(self):
        """Equation 3: Calculation of Allowance for Enhanced-performance Integrated Displays"""
        (diagonal, enhanced_performance_display) = self.sysinfo.get_display()
        if enhanced_performance_display:
            if diagonal >= 27.0:
                EP = 0.75
            else:
                EP = 0.3
        else:
            EP = 0
        (width, height) = self.sysinfo.get_resolution()
        r = 1.0 * width * height / 1000000
        A =  1.0 * self.sysinfo.get_screen_area()
        debug("EP = %s, r = %s, A = %s" % (EP, r, A))
        return (EP, r, A)

    def equation_four(self):
        """Equation 4: P_TEC Calculation for Workstations""" 
        (P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE) = self.sysinfo.get_power_consumptions()
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.35, 0.10, 0.15, 0.40)
        P_TEC = P_OFF * T_OFF + P_SLEEP * T_SLEEP + P_LONG_IDLE * T_LONG_IDLE + P_SHORT_IDLE * T_SHORT_IDLE
        return P_TEC

    def equation_five(self):
        """Equation 5: P_TEC_MAX Calculation for Workstations"""
        (T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.10, 0.15, 0.40)
        P_EEE = 0.2 * self.sysinfo.get_1glan_num()
        P_MAX = self.sysinfo.max_power
        N_HDD = self.sysinfo.disk_num
        P_TEC_MAX = 0.28 * (P_MAX + N_HDD * 5) + 8.76 * P_EEE * (T_SLEEP + T_LONG_IDLE + T_SHORT_IDLE)
        return P_TEC_MAX

    def equation_six(self, discrete, wol):
        """Equation 6: Calculation of E_TEC_MAX for Thin Clients"""
        TEC_BASE = 31

        if discrete:
            TEC_GRAPHICS = 36
        else:
            TEC_GRAPHICS = 0

        if wol:
            TEC_WOL = 2
        else:
            TEC_WOL = 0

        if self.sysinfo.integrated_display:
            (EP, r, A) = self.equation_three()
            TEC_INT_DISPLAY = 8.76 * 0.35 * (1 + EP) * (4 * r + 0.05 * A)
        else:
            TEC_INT_DISPLAY = 0
        debug("TEC_INT_DISPLAY = %s" % (TEC_INT_DISPLAY))

        TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * self.sysinfo.get_1glan_num()

        E_TEC_MAX = TEC_BASE + TEC_GRAPHICS + TEC_WOL + TEC_INT_DISPLAY + TEC_EEE

        return E_TEC_MAX
