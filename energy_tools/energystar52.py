# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
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

from logging import debug, warning

class EnergyStar52:
    """Energy Star 5.2 calculator"""
    def __init__(self, sysinfo):
        self.sysinfo = sysinfo
        debug("=== Energy Star 5.2 ===")

    def qualify_desktop_category(self, category, discrete=False, over_frame_buffer_width_128=False):
        (core, clock, memory, disk) = self.sysinfo.get_basic_info()

        if category == 'D':
            if core >= 4:
                if memory >= 4:
                    return True
                elif discrete and over_frame_buffer_width_128:
                    return True
        elif category == 'C':
            if core > 2:
                if memory >= 2:
                    return True
                elif discrete:
                    return True
        elif category == 'B':
            if core == 2 and memory >= 2:
                return True
        elif category == 'A':
            return True
        return False

    def qualify_netbook_category(self, category, discrete=False, over_frame_buffer_width_128=False):
        (core, clock, memory, disk) = self.sysinfo.get_basic_info()

        if category == 'C':
            if core >= 2 and memory >= 2:
                if discrete and over_frame_buffer_width_128:
                    return True
        elif category == 'B':
            if discrete:
                return True
        elif category == 'A':
            return True
        return False

    def equation_one(self):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, and Notebook Computers"""
        (P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE) = self.sysinfo.get_power_consumptions()
        P_IDLE = P_SHORT_IDLE

        if self.sysinfo.computer_type == 3:
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

        E_TEC = ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE)) * 8760 / 1000

        debug("T_OFF = %s, T_SLEEP = %s, T_IDLE = %s" % (T_OFF, T_SLEEP, T_IDLE))
        debug("P_OFF = %s, P_SLEEP = %s, P_IDLE = %s" % (P_OFF, P_SLEEP, P_IDLE))

        return E_TEC

    def equation_two(self, over_frame_buffer_width_128=False, over_frame_buffer_width_64=False):
        """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop, and Notebook Computers"""

        (core, clock, memory, disk) = self.sysinfo.get_basic_info()

        result = []

        ## Maximum TEC Allowances for Desktop and Integrated Desktop Computers
        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            if disk > 1:
                TEC_STORAGE = 25.0 * (disk - 1)
            else:
                TEC_STORAGE = 0.0

            if self.qualify_desktop_category('A'):
                TEC_BASE = 148.0

                if memory > 2:
                    TEC_MEMORY = 1.0 * (memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_frame_buffer_width_128:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_desktop_category('B'):
                TEC_BASE = 175.0

                if memory > 2:
                    TEC_MEMORY = 1.0 * (memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_frame_buffer_width_128:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_desktop_category('C', self.sysinfo.discrete):
                TEC_BASE = 209.0

                if memory > 2:
                    TEC_MEMORY = 1.0 * (memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_frame_buffer_width_128:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('C', E_TEC_MAX))

            if self.qualify_desktop_category('D', self.sysinfo.discrete, over_frame_buffer_width_128):
                TEC_BASE = 234.0

                if memory > 4:
                    TEC_MEMORY = 1.0 * (memory - 4)
                else:
                    TEC_MEMORY = 0.0

                if over_frame_buffer_width_128:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('D', E_TEC_MAX))

        ## Maximum TEC Allowances for Notebook Computers
        else:
            if memory > 4:
                TEC_MEMORY = 0.4 * (memory - 4)
            else:
                TEC_MEMORY = 0.0

            if disk > 1:
                TEC_STORAGE = 3.0 * (disk - 1)
            else:
                TEC_STORAGE = 0.0

            if self.qualify_netbook_category('A'):
                TEC_BASE = 40.0
                TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_netbook_category('B', self.sysinfo.discrete):
                TEC_BASE = 53.0

                if over_frame_buffer_width_64:
                    TEC_GRAPHICS = 3.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_netbook_category('C', self.sysinfo.discrete, over_frame_buffer_width_128):
                TEC_BASE = 88.5
                TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('C', E_TEC_MAX))

        debug("TEC_BASE = %s, TEC_MEMORY = %s, TEC_STORAGE = %s, TEC_GRAPHICS = %s" % (TEC_BASE, TEC_MEMORY, TEC_STORAGE, TEC_GRAPHICS))
        debug("E_TEC_MAX = %s" % (E_TEC_MAX))

        return result

    def equation_three(self):
        """Equation 3: P_TEC Calculation for Workstations""" 
        (P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE) = self.sysinfo.get_power_consumptions()
        P_IDLE = P_SHORT_IDLE

        (T_OFF, T_SLEEP, T_IDLE) = (0.35, 0.10, 0.55)
        P_TEC = (P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE) 

        return P_TEC

    def equation_four(self):
        """Equation 4: P_TEC_MAX Calculation for Workstations"""
        P_MAX = self.sysinfo.max_power
        N_HDD = self.sysinfo.disk_num

        P_TEC_MAX = 0.28 * (P_MAX + (N_HDD * 5))

        return P_TEC_MAX

    def equation_five(self, wol):
        """Equation 5: Calculation of P_OFF_MAX for Small-scale Servers"""
        (core, clock, memory, disk) = self.sysinfo.get_basic_info()

        P_OFF_BASE = 2.0
        if wol:
            P_OFF_WOL = 0.7
        else:
            P_OFF_WOL = 0
        P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

        if (core > 1 or self.sysinfo.more_discrete) and memory >= 1:
            category = 'B'
            P_IDLE_MAX = 65.0
        else:
            category = 'A'
            P_IDLE_MAX = 50.0

        return (category, P_OFF_MAX, P_IDLE_MAX)

    def equation_six(self, wol):
        """Equation 6: Calculation of P_OFF_MAX for Thin Clients"""
        P_OFF_BASE = 2.0
        if wol:
            P_OFF_WOL = 0.7
        else:
            P_OFF_WOL = 0
        P_OFF_MAX = P_OFF_BASE + P_OFF_WOL
        return P_OFF_MAX

    def equation_seven(self, wol):
        """Equation 7: Calculation of P_SLEEP_MAX for Thin Clients"""
        P_SLEEP_BASE = 2.0
        if wol:
            P_SLEEP_WOL = 0.7
        else:
            P_SLEEP_WOL = 0
        P_SLEEP_MAX = P_SLEEP_BASE + P_SLEEP_WOL
        return P_SLEEP_MAX
