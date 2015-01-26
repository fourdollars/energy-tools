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

import logging, os, json
from logging import debug, warning, error
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
    version = "v1.3.3"
    print("Energy Tools %s for Energy Star 5.2/6.0 and ErP Lot 3\n" % (version)+ '=' * 80)
    if args.test == 1:
        # Test case for Notebooks of Energy Star 5.2 & 6.0
        # E_TEC: 33.03 kWh/year, E_TEC_MAX: 41.6 kWh/year, PASS for 5.2
        # E_TEC: 40.7 kWh/year, E_TEC_MAX: 39.0 kWh/year, FAIL for 6.0
        sysinfo = SysInfo({
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
    elif args.test == 2:
        # Test case for Workstations of Energy Star 5.2
        # P_TEC: 45.1 W, P_MAX: 53.2 W, PASS for 5.2
        sysinfo = SysInfo({
            'Product Type': 2,
            'Disk Number': 2,
            'Gigabit Ethernet': 0,
            'Off Mode': 2.0,
            'Sleep Mode': 4.0,
            'Long Idle Mode': 50.0,
            'Short Idle Mode': 80.0,
            'Maximum Power': 180.0})
    elif args.test == 3:
        # Test case for Small-scale Servers
        sysinfo = SysInfo({
            'Product Type': 3,
            'Memory Size': 4,
            'CPU Clock': 2.0, # TODO: remove this
            'CPU Cores': 1,
            'More Discrete Graphics': False,
            'Gigabit Ethernet': 1,
            'Disk Number': 1,
            'Off Mode': 2.7,
            'Short Idle Mode': 65.0})
    elif args.test == 4:
        # Test case for Thin Clients
        sysinfo = SysInfo({
            'Product Type': 4,
            'Integrated Display': True,
            'Display Width': 1366,
            'Display Height': 768,
            'Display Diagonal': 14,
            'Enhanced Display': True,
            'Discrete Graphics': False,
            'Off Mode': 2.7,
            'Sleep Mode': 2.7,
            'Long Idle Mode': 15.0,
            'Short Idle Mode': 15.0,
            'Gigabit Ethernet': 1,
            'Media Codec': True})
    elif args.test == 5:
        # Test case from OEM/ODM only for Energy Star 5.2
        # Category B: 19.16688 (E_TEC) <= 60.8 (E_TEC_MAX), PASS
        sysinfo = SysInfo({
            'Product Type': 1,
            'Computer Type': 3,
            'CPU Clock': 1.8,
            'CPU Cores': 2,
            'Discrete Audio': False,
            'Discrete Graphics': True,
            'Discrete Graphics Cards': 1,
            'Switchable Graphics': False,
            'Disk Number': 1,
            'Display Diagonal': 14,
            'Display Height': 768,
            'Display Width': 1366,
            'Enhanced Display': False,
            'Gigabit Ethernet': 1,
            'Memory Size': 16,
            'TV Tuner': False,
            'Off Mode': 0.27,
            'Off Mode with WOL': 0.27,
            'Sleep Mode': 0.61,
            'Sleep Mode with WOL': 0.61,
            'Long Idle Mode': 6.55,
            'Short Idle Mode': 6.55})
    elif args.profile:
        if os.path.exists(args.profile):
            with open(args.profile, "r") as data:
                tmp = data.read().replace('\n', '')
                profile = json.loads(tmp)
                sysinfo = SysInfo(profile)
        else:
            error('Can not open %s to read.' % args.profile)
            return
    else:
        sysinfo = SysInfo()

    energystar_calculate(sysinfo)
    erplot3_calculate(sysinfo)
    generate_excel(sysinfo, version, args.output)

    if args.export:
        sysinfo.save(args.export)

def erplot3_calculate(sysinfo):
    if sysinfo.product_type != 1:
        return
    erplot3 = ErPLot3(sysinfo)
    erplot3.calculate()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", "--debug",   help="print debug messages",  action="store_true")
    parser.add_argument("-t", "--test",    help="use test case",         type=int)
    parser.add_argument("-o", "--output",  help="output Excel file",     type=str)
    parser.add_argument("-p", "--profile", help="import system profile", type=str)
    parser.add_argument("-x", "--export",  help="export system profile", type=str)
    args = parser.parse_args()
    if args.debug:
        logging.basicConfig(format='<%(levelname)s> %(message)s', level=logging.DEBUG)
    else:
        logging.basicConfig(format='<%(levelname)s> %(message)s')
    main()
