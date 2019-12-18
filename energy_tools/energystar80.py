# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
#
# Copyright (C) 2019 Canonical Ltd.
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


from logging import debug, error
from math import tanh


class EnergyStar80(object):
    """Energy Star 8.0 calculator"""
    def __init__(self, sysinfo):
        self.sysinfo = sysinfo
        debug("=== Energy Star 8.0 ===")

    # Requirements for Desktop, Integrated Desktop, and Notebook Computers
    def equation_one(self):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop,
                       Thin Client and Notebook Computers"""
        (p_off, p_sleep, p_long_idle, p_short_idle) = \
            self.sysinfo.get_power_consumptions()
        if self.sysinfo.product_type == 4:
            (t_off, t_sleep, t_long_idle, t_short_idle) = \
                (0.45, 0.05, 0.15, 0.35)
        elif self.sysinfo.computer_type == 3:
            (t_off, t_sleep, t_long_idle, t_short_idle) = \
                (0.25, 0.35, 0.1, 0.3)
        else:
            (t_off, t_sleep, t_long_idle, t_short_idle) = \
                (0.15, 0.45, 0.1, 0.3)

        e_tec = ((p_off * t_off) + (p_sleep * t_sleep) +
                 (p_long_idle * t_long_idle) +
                 (p_short_idle * t_short_idle)) * 8760 / 1000

        debug("T_OFF = %s, T_SLEEP = %s, T_LONG_IDLE = %s, T_SHORT_IDLE = %s" %
              (t_off, t_sleep, t_long_idle, t_short_idle))
        debug("P_OFF = %s, P_SLEEP = %s, P_LONG_IDLE = %s, P_SHORT_IDLE = %s" %
              (p_off, p_sleep, p_long_idle, p_short_idle))

        return e_tec

    def equation_two(self, fb_bw, mobile_workstation=False):
        """Equation 2: E_TEC_MAX Calculation for
                       Desktop, Integrated Desktop, and Notebook Computers"""
        (core, clock, memory, disk) = self.sysinfo.get_basic_info()
        storage_info = self.sysinfo.profile

        pscore = core * clock
        debug("P = %s" % (pscore))

        if self.sysinfo.computer_type == 1:
            if self.sysinfo.discrete:
                if pscore <= 8:
                    tec_base = 35.0
                else:
                    tec_base = 45.0
            else:
                if pscore <= 8:
                    tec_base = 26.0
                else:
                    tec_base = 46.0
        elif self.sysinfo.computer_type == 2:
            if pscore <= 8:
                tec_base = 9.0
            else:
                tec_base = 27.0
        elif self.sysinfo.computer_type == 3:
            if pscore <= 2:
                tec_base = 6.5
            elif pscore < 8:
                tec_base = 8.0
            else:
                tec_base = 14.0
        else:
            error("It should not reach here.")

        if self.sysinfo.computer_type != 3:
            tec_memory = 1.7 + 0.24 * memory
        else:
            tec_memory = 2.4 + 0.294 * memory

        if self.sysinfo.switchable:
            if self.sysinfo.computer_type != 3:
                tec_switchable = 14.4
            else:
                tec_switchable = 0
            tec_graphics = 0
        elif self.sysinfo.discrete:
            tec_switchable = 0
            if self.sysinfo.computer_type != 3:
                tec_graphics = 50.4 * tanh(0.0038 * fb_bw - 0.137) + 23
            else:
                tec_graphics = 29.3 * tanh(0.0038 * fb_bw - 0.137) + 13.4
        else:
            tec_switchable = 0
            tec_graphics = 0

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            if self.sysinfo.get_10glan_num():
                tec_glan10 = 18.0
            else:
                tec_glan10 = 0
            if self.sysinfo.get_1to10glan_num():
                tec_glan1to10 = 4.0
            else:
                tec_glan1to10 = 0
        else:
            tec_glan1to10 = 0
            tec_glan10 = 0

        if disk > 1:
            if self.sysinfo.computer_type == 3:
                tec_storage = storage_info["3.5 inch HDD"] * 0.0 \
                    + storage_info["2.5 inch HDD"] * 2.6 \
                    + storage_info["Hybrid HDD/SSD"] * 2.6 \
                    + storage_info["SSD"] * 2.6
            else:
                tec_storage = storage_info["3.5 inch HDD"] * 16.5 \
                    + storage_info["2.5 inch HDD"] * 2.1 \
                    + storage_info["Hybrid HDD/SSD"] * 0.8 \
                    + storage_info["SSD"] * 0.4
        else:
            tec_storage = 0

        if self.sysinfo.computer_type != 1:
            (e_p, resolution, area) = self.equation_three()

        if self.sysinfo.computer_type == 2:
            if area < 190:
                tec_int_display = (3.43*resolution + 0.148*area + 1.30)*(1+e_p)
            elif area < 210:
                tec_int_display = (3.43*resolution + 0.018*area + 26.1)*(1+e_p)
            elif area < 315:
                tec_int_display = (3.43*resolution + 0.078*area + 13.2)*(1+e_p)
            else:  # area >= 315
                tec_int_display = (3.43*resolution + 0.156*area - 11.3)*(1+e_p)
        elif self.sysinfo.computer_type == 3:
            tec_int_display = 8.76*0.30*(1+e_p)*(0.43*resolution+0.0263*area)
        else:
            tec_int_display = 0

        if self.sysinfo.computer_type == 3 and mobile_workstation:
            tec_mobile_workstation = 4.0
        else:
            tec_mobile_workstation = 0

        msg = "TEC_BASE=%s, TEC_MEMORY=%s, TEC_GRAPHICS=%s, " + \
            "TEC_SWITCHABLE=%s, TEC_10GLAN=%s, TEC_1TO10GLAN=%s, " + \
            "TEC_STORAGE=%s, TEC_INT_DISPLAY=%s, TEC_MOBILEWORKSTATION=%s"

        debug(msg % (tec_base, tec_memory, tec_graphics, tec_switchable,
                     tec_glan10, tec_glan1to10, tec_storage, tec_int_display,
                     tec_mobile_workstation))

        return tec_base + tec_memory + tec_graphics + tec_storage + \
            tec_int_display + tec_switchable + tec_glan10 + tec_glan1to10 + \
            tec_mobile_workstation

    def equation_three(self):
        """Equation 3: Calculation of Allowance for
                       Enhanced-performance Integrated Displays"""
        (diagonal, enhanced_performance_display) = self.sysinfo.get_display()
        if enhanced_performance_display:
            if diagonal >= 27.0:
                e_p = 0.75
            else:
                e_p = 0.3
        else:
            e_p = 0
        (width, height) = self.sysinfo.get_resolution()
        resolution = 1.0 * width * height / 1000000
        area = 1.0 * self.sysinfo.get_screen_area()
        debug("EP = %s, r = %s, A = %s" % (e_p, resolution, area))
        return (e_p, resolution, area)

    def equation_four(self):
        """Equation 4: P_TEC Calculation for Workstations"""
        (p_off, p_sleep, p_long_idle, p_short_idle) = \
            self.sysinfo.get_power_consumptions()
        (t_off, t_sleep, t_long_idle, t_short_idle) = (0.10, 0.35, 0.20, 0.35)
        p_tec = p_off * t_off \
            + p_sleep * t_sleep \
            + p_long_idle * t_long_idle \
            + p_short_idle * t_short_idle
        return p_tec

    def equation_five(self):
        """Equation 5: P_TEC_MAX Calculation for Workstations"""
        p_max = self.sysinfo.max_power
        n_hdd = self.sysinfo.disk_num
        p_tec_max = 0.28 * (p_max + n_hdd * 5)
        return p_tec_max

    def equation_six(self, discrete, wol):
        """Equation 6: Calculation of E_TEC_MAX for Thin Clients"""
        tec_base = 31

        if discrete:
            tec_graphics = 36
        else:
            tec_graphics = 0

        if wol:
            tec_wol = 2
        else:
            tec_wol = 0

        if self.sysinfo.integrated_display:
            (e_p, resolution, area) = self.equation_three()
            if area < 190:
                tec_int_display = (3.43*resolution + 0.148*area + 1.30)*(1+e_p)
            elif area < 210:
                tec_int_display = (3.43*resolution + 0.018*area + 26.1)*(1+e_p)
            elif area < 315:
                tec_int_display = (3.43*resolution + 0.078*area + 13.2)*(1+e_p)
            else:  # area >= 315
                tec_int_display = (3.43*resolution + 0.156*area - 11.3)*(1+e_p)
        else:
            tec_int_display = 0
        debug("TEC_INT_DISPLAY = %s" % (tec_int_display))

        e_tec_max = tec_base + tec_graphics + tec_wol + tec_int_display

        return e_tec_max
