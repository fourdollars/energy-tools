# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
#
# Copyright (C) 2020 Canonical Ltd.
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
from .sysinfo import SysInfo
from logging import debug, warning
from .common import result_filter

__all__ = ["ErPLot26"]


class ErPLot26:
    """ErP Lot 26 (Networked Standby) calculator"""
    def __init__(self, sysinfo):
        self.sysinfo = sysinfo

    def calculate(self):
        print("\nErP Lot 26 Tier 3 (1 Jan 2019):\n")
        self._verify_s3_s5()

    def _verify_s3_s5(self):
        passed = True

        if self.sysinfo.sleep_wol > 2:
            print("  Failed. P_SLEEP_WOL (%s) > 2.0"
                  % self.sysinfo.sleep_wol)
            passed = False

        if self.sysinfo.off_wol > 0.5:
            print("  Failed. P_OFF_WOL (%s) > 0.5" % self.sysinfo.off_wol)
            passed = False

        if passed:
            print("  Pass. P_SLEEP_WOL (%s) <= 2.0 and P_OFF_WOL (%s) <= 0.5"
                  % (self.sysinfo.sleep_wol, self.sysinfo.off_wol))
