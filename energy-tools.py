#!/usr/bin/env python
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

import logging
from logging import debug, warning
import argparse
from excel_output import *
from energystar52 import EnergyStar52
from energystar60 import EnergyStar60
from sysinfo import SysInfo
from erplot3 import ErPLot3

def energystar_calculate(sysinfo):
    if sysinfo.product_type == 1:

        # Energy Star 5.2
        print("Energy Star 5.2:")
        estar52 = EnergyStar52(sysinfo)
        E_TEC = estar52.equation_one()

        over_128 = estar52.equation_two(True, True)
        between_64_and_128 = estar52.equation_two(False, True)
        under_64 = estar52.equation_two(False, False)
        debug(over_128)
        debug(between_64_and_128)
        debug(under_64)
        different=False

        for i,j,k in zip(over_128, between_64_and_128, under_64):
            (cat1, max1) = i
            (cat2, max2) = j
            (cat3, max3) = k
            if cat1 != cat2 or max1 != max2 or cat2 != cat3 or max2 != max3:
                different=True
        else:
            if different is True:
                if sysinfo.computer_type == 3:
                    print("\n  If GPU Frame Buffer Width <= 64 bits,")
                    for i in under_64:
                        (category, E_TEC_MAX) = i
                        if E_TEC <= E_TEC_MAX:
                            result = 'PASS'
                            operator = '<='
                        else:
                            result = 'FAIL'
                            operator = '>'
                        print("    Category %s: %s (E_TEC) %s %s (E_TEC_MAX), %s" % (category, E_TEC, operator, E_TEC_MAX, result))
                    print("\n  If 64 bits < GPU Frame Buffer Width <= 128 bits,")
                    for i in between_64_and_128:
                        (category, E_TEC_MAX) = i
                        if E_TEC <= E_TEC_MAX:
                            result = 'PASS'
                            operator = '<='
                        else:
                            result = 'FAIL'
                            operator = '>'
                        print("    Category %s: %s (E_TEC) %s %s (E_TEC_MAX), %s" % (category, E_TEC, operator, E_TEC_MAX, result))
                else:
                    print("\n  If GPU Frame Buffer Width <= 128 bits,")
                    for i in between_64_and_128:
                        (category, E_TEC_MAX) = i
                        if E_TEC <= E_TEC_MAX:
                            result = 'PASS'
                            operator = '<='
                        else:
                            result = 'FAIL'
                            operator = '>'
                        print("    Category %s: %s (E_TEC) %s %s (E_TEC_MAX), %s" % (category, E_TEC, operator, E_TEC_MAX, result))
                print("\n  If GPU Frame Buffer Width > 128 bits,")
                for i in over_128:
                    (category, E_TEC_MAX) = i
                    if E_TEC <= E_TEC_MAX:
                        result = 'PASS'
                        operator = '<='
                    else:
                        result = 'FAIL'
                        operator = '>'
                    print("    Category %s: %s (E_TEC) %s %s (E_TEC_MAX), %s" % (category, E_TEC, operator, E_TEC_MAX, result))
            else:
                for i in under_64:
                    (category, E_TEC_MAX) = i
                    if E_TEC <= E_TEC_MAX:
                        result = 'PASS'
                        operator = '<='
                    else:
                        result = 'FAIL'
                        operator = '>'
                    print("\n  Category %s: %s (E_TEC) %s %s (E_TEC_MAX), %s" % (category, E_TEC, operator, E_TEC_MAX, result))

        # Energy Star 6.0
        print("\nEnergy Star 6.0:\n")
        estar60 = EnergyStar60(sysinfo)
        E_TEC = estar60.equation_one()

        lower = 1.015
        if sysinfo.computer_type == 2:
            higher = 1.04
        else:
            higher = 1.03

        for AllowancePSU in (1, lower, higher):
            if sysinfo.discrete:
                if AllowancePSU == 1:
                    print("  If power supplies do not meet the requirements of Power Supply Efficiency Allowance,")
                elif AllowancePSU == lower:
                    print("  If power supplies meet lower efficiency requirements,")
                elif AllowancePSU == higher:
                    print("  If power supplies meet higher efficiency requirements,")
                for gpu in ('G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7'):
                    E_TEC_MAX = estar60.equation_two(gpu) * AllowancePSU
                    if E_TEC <= E_TEC_MAX:
                        result = 'PASS'
                        operator = '<='
                    else:
                        result = 'FAIL'
                        operator = '>'
                    if gpu == 'G1':
                        gpu = "G1 (FB_BW <= 16)"
                    elif gpu == 'G2':
                        gpu = "G2 (16 < FB_BW <= 32)"
                    elif gpu == 'G3':
                        gpu = "G3 (32 < FB_BW <= 64)"
                    elif gpu == 'G4':
                        gpu = "G4 (64 < FB_BW <= 96)"
                    elif gpu == 'G5':
                        gpu = "G5 (96 < FB_BW <= 128)"
                    elif gpu == 'G6':
                        gpu = "G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)"
                    elif gpu == 'G7':
                        gpu = "G7 (FB_BW > 128; Frame Buffer Data Width >= 192 bits)"
                    print("    %s (E_TEC) %s %s (E_TEC_MAX) for %s, %s" % (E_TEC, operator, E_TEC_MAX, gpu, result))
            else:
                if AllowancePSU == 1:
                    print("  If power supplies do not meet the requirements of Power Supply Efficiency Allowance,")
                elif AllowancePSU == lower:
                    print("  If power supplies meet lower efficiency requirements,")
                elif AllowancePSU == higher:
                    print("  If power supplies meet higher efficiency requirements,")
                E_TEC_MAX = estar60.equation_two('G1') * AllowancePSU
                if E_TEC <= E_TEC_MAX:
                    result = 'PASS'
                    operator = '<='
                else:
                    result = 'FAIL'
                    operator = '>'
                print("    %s (E_TEC) %s %s (E_TEC_MAX), %s" % (E_TEC, operator, E_TEC_MAX, result))
    elif sysinfo.product_type == 2:
        # Energy Star 5.2
        print("Energy Star 5.2:")
        estar52 = EnergyStar52(sysinfo)
        P_TEC = estar52.equation_three()
        P_TEC_MAX = estar52.equation_four()
        if P_TEC <= P_TEC_MAX:
            result = 'PASS'
            operator = '<='
        else:
            result = 'FAIL'
            operator = '>'
        print("  %s (P_TEC) %s %s (P_TEC_MAX), %s" % (P_TEC, operator, P_TEC_MAX, result))

        # Energy Star 6.0
        print("Energy Star 6.0:")
        estar60 = EnergyStar60(sysinfo)
        P_TEC = estar60.equation_four()
        P_TEC_MAX = estar60.equation_five()
        if P_TEC <= P_TEC_MAX:
            result = 'PASS'
            operator = '<='
        else:
            result = 'FAIL'
            operator = '>'
        print("  %s (P_TEC) %s %s (P_TEC_MAX), %s" % (P_TEC, operator, P_TEC_MAX, result))
    elif sysinfo.product_type == 3:
        # Energy Star 5.2
        print("Energy Star 5.2:")
        estar52 = EnergyStar52(sysinfo)
        for wol in (True, False):
            (category, P_OFF_MAX, P_IDLE_MAX) = estar52.equation_five(wol)
            P_OFF = sysinfo.off
            P_IDLE = sysinfo.short_idle

            if P_OFF <= P_OFF_MAX and P_IDLE <= P_IDLE_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'

            if P_OFF <= P_OFF_MAX:
                op1 = '<='
            else:
                op1 = '>'

            if P_IDLE <= P_IDLE_MAX:
                op2 = '<='
            else:
                op2 = '>'
            if wol:
                print("  If Wake-On-LAN (WOL) is enabled by default upon shipment.")
            else:
                print("  If Wake-On-LAN (WOL) is disabled by default upon shipment.")
            print("    Category %s: %s (P_OFF) %s %s (P_OFF_MAX), %s (P_IDLE) %s %s (P_IDLE_MAX), %s" % (category, P_OFF, op1, P_OFF_MAX, P_IDLE, op2, P_IDLE_MAX, result))

        # Energy Star 6.0
        print("Energy Star 6.0:")
        estar60 = EnergyStar60(sysinfo)
        for wol in (True, False):
            P_OFF = sysinfo.off
            P_OFF_MAX = estar60.equation_six(wol)
            P_IDLE = sysinfo.short_idle
            P_IDLE_MAX = estar60.equation_seven()

            if P_OFF <= P_OFF_MAX and P_IDLE <= P_IDLE_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'

            if P_OFF <= P_OFF_MAX:
                op1 = '<='
            else:
                op1 = '>'

            if P_IDLE <= P_IDLE_MAX:
                op2 = '<='
            else:
                op2 = '>'
            if wol:
                print("  If Wake-On-LAN (WOL) is enabled by default upon shipment.")
            else:
                print("  If Wake-On-LAN (WOL) is disabled by default upon shipment.")
            print("    %s (P_OFF) %s %s (P_OFF_MAX), %s (P_IDLE) %s %s (P_IDLE_MAX), %s" % (P_OFF, op1, P_OFF_MAX, P_IDLE, op2, P_IDLE_MAX, result))

    elif sysinfo.product_type == 4:
        # Energy Star 5.2
        print("Energy Star 5.2:")
        estar52 = EnergyStar52(sysinfo)
        for wol in (True, False):
            if wol:
                print("  If Wake-On-LAN (WOL) is enabled by default upon shipment.")
            else:
                print("  If Wake-On-LAN (WOL) is disabled by default upon shipment.")

            P_OFF = sysinfo.off
            P_OFF_MAX = estar52.equation_six(wol)

            P_SLEEP = sysinfo.sleep
            P_SLEEP_MAX = estar52.equation_seven(wol)

            P_IDLE = sysinfo.short_idle
            if sysinfo.media_codec:
                P_IDLE_MAX = 15.0
                category = 'B'
            else:
                P_IDLE_MAX = 12.0
                category = 'A'

            print("    Category %s:" % (category))

            if P_OFF <= P_OFF_MAX:
                op1 = '<='
            else:
                op1 = '>'
            print("      %s (P_OFF) %s %s (P_OFF_MAX)" % (P_OFF, op1, P_OFF_MAX))

            if P_SLEEP <= P_SLEEP_MAX:
                op2 = '<='
            else:
                op2 = '>'
            print("      %s (P_SLEEP) %s %s (P_SLEEP_MAX)" % (P_SLEEP, op2, P_SLEEP_MAX))

            if P_IDLE <= P_IDLE_MAX:
                op3 = '<='
            else:
                op3 = '>'
            print("      %s (P_IDLE) %s %s (P_IDLE_MAX)" % (P_IDLE, op3, P_IDLE_MAX))


            if P_OFF <= P_OFF_MAX and P_SLEEP <= P_SLEEP_MAX and P_IDLE <= P_IDLE_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("        %s" % (result))
        # Energy Star 6.0
        print("Energy Star 6.0:")
        estar60 = EnergyStar60(sysinfo)
        E_TEC = estar60.equation_one()
        for discrete in (True, False):
            for wol in (True, False):
                E_TEC_MAX = estar60.equation_eight(discrete, wol)
                if discrete:
                    msg1 = "it has Discrete Graphics enabled"
                else:
                    msg1 = "it doesn't have Discrete Graphics enabled"
                if wol:
                    msg2 = "Wake-On-LAN (WOL) is enabled"
                else:
                    msg2 = "Wake-On-LAN (WOL) is disabled"
                print("  If %s and %s by default upon shipment," % (msg1, msg2))
                if E_TEC <= E_TEC_MAX:
                    operator = '<='
                    result = 'PASS'
                else:
                    operator = '>'
                    result = 'FAIL'
                print("    %s (E_TEC) %s %s (E_TEC_MAX), %s" % (E_TEC, operator, E_TEC_MAX, result))
    else:
        raise Exception('This is a bug when you see this.')

def main():
    version = "v0.0"
    print("Energy Tools %s for Energy Star 5.2/6.0 and ErP Lot 3\n" % (version)+ '=' * 80)
    if args.test == 1:
        # Test case from Energy Star 5.2/6.0 for Notebooks
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=3,
                cpu_core=2, cpu_clock=2.0,
                mem_size=8, disk_num=1,
                width=1366, height=768, eee=1,
                diagonal=14, ep=False,
                discrete=False, switchable=True,
                off=1.0, sleep=1.7, long_idle=8.0, short_idle=10.0)
    elif args.test == 2:
        # Test case for Workstations
        sysinfo = SysInfo(
                auto=True,
                product_type=2, disk_num=2, eee=0,
                off=2, sleep=4, long_idle=50, short_idle=80, max_power=180)
    elif args.test == 3:
        # Test case for Small-scale Servers 
        sysinfo = SysInfo(
                auto=True,
                product_type=3, mem_size=4,
                cpu_core=1, more_discrete=False,
                eee=1, disk_num=1,
                off=2.7, short_idle=65.0)
    elif args.test == 4:
        # Test case for Thin Clients
        sysinfo = SysInfo(
                auto=True,
                product_type=4,
                integrated_display=True, width=1366, height=768, diagonal=14, ep=True,
                off=2.7, sleep=2.7, long_idle=15.0, short_idle=15.0, media_codec=True)
    elif args.test == 5:
        # Test case from OEM/ODM only for Energy Star 5.2
        # Category B: 19.16688 (E_TEC) <= 60.8 (E_TEC_MAX), PASS
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=3,
                cpu_core=2, cpu_clock=1.8,
                mem_size=16, disk_num=1,
                width=1366, height=768, eee=1,
                diagonal=14, ep=False,
                discrete=True, switchable=False,
                off=0.27, sleep=0.61, long_idle=6.55, short_idle=6.55)
    elif args.test == 6:
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=2,
                cpu_core=2, cpu_clock=2.4,
                mem_size=4, disk_num=1,
                width=1680, height=1050, eee=1,
                diagonal=27, ep=True,
                discrete=False, switchable=True,
                off=12, sleep=23, long_idle=34, short_idle=45)
    elif args.test == 7:
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=3,
                cpu_core=2, cpu_clock=2.0,
                mem_size=8, disk_num=1,
                width=1366, height=768, eee=1,
                diagonal=14, ep=False,
                discrete=True, switchable=False,
                off=1.0, sleep=1.7, long_idle=8.0, short_idle=10.0)
    elif args.test == 8:
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=2,
                cpu_core=4, cpu_clock=2.4,
                mem_size=4, disk_num=1,
                width=1680, height=1050, eee=1,
                diagonal=27, ep=True,
                discrete=False, switchable=True,
                off=12, sleep=23, long_idle=34, short_idle=45)
    elif args.test == 9:
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=1,
                cpu_core=4, cpu_clock=2.4,
                mem_size=4, disk_num=1, eee=1,
                discrete=False, switchable=True,
                off=12, sleep=23, long_idle=34, short_idle=45)
    elif args.test == 10:
        sysinfo = SysInfo(
                auto=True,
                product_type=1, computer_type=2,
                cpu_core=3, cpu_clock=2.26,
                mem_size=4, disk_num=1, eee=1, ep=True,
                discrete=True, discrete_gpu_num=1, diagonal=12.1,
                off=0.486, off_wol=0.66, sleep=0.74, sleep_wol=0.74, long_idle=10.6, short_idle=15)
    else:
        sysinfo = SysInfo()

    energystar_calculate(sysinfo)
    erplot3_calculate(sysinfo)
    generate_excel(sysinfo, version, args.output)

def erplot3_calculate(sysinfo):
    if sysinfo.product_type != 1:
        return
    erplot3 = ErPLot3(sysinfo)
    erplot3.calculate()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--debug",  help="print debug messages", action="store_true")
    parser.add_argument("-t", "--test",  help="use test case", type=int)
    parser.add_argument("-o", "--output",  help="output Excel file", type=str)
    args = parser.parse_args()
    if args.debug:
        logging.basicConfig(format='<%(levelname)s> %(message)s', level=logging.DEBUG)
    main()
