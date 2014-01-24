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

from sysinfo import SystemInformation

def question_str(prompt, length, validator):
    while True:
        s = raw_input(prompt + "\n>> ")
        if len(s) == length and set(s).issubset(validator):
            print('')
            return s
        print("The valid input '" + validator + "'.")

def question_num(prompt):
    while True:
        s = raw_input(prompt + "\n>> ")
        try:
            num = float(s)
            print('')
            return num
        except ValueError:
            print "Oops!  That was no valid number.  Try again..."

import inspect
def lineno():
    """Returns the current line number in our program."""
    return " (DEBUG: line %s)" % str(inspect.currentframe().f_back.f_lineno) 

class EnergyStar52:
    """Energy Star 5.2 calculator"""
    def __init__(self, sysinfo):
        self.core = sysinfo.get_cpu_info()
        self.disk = sysinfo.get_disk_info()
        self.memory = sysinfo.get_memory_info()

    def qualify_desktop_category(self, category, gpu_discrete, gpu_width):
        if category == 'D':
            if self.core >= 4:
                if self.memory >= 4:
                    return True
                elif gpu_width == 'y':
                    return True
        elif category == 'C':
            if self.core >= 2:
                if self.memory >= 2:
                    return True
                elif gpu_discrete == 'y':
                    return True
        elif category == 'B':
            if self.core == 2 and self.memory >= 2:
                return True
        elif category == 'A':
            return True
        return False

    def qualify_netbook_category(self, category, gpu_discrete, gpu_width):
        if category =='C':
            if self.core >= 2 and self.memory >= 2:
                if gpu_discrete == 'y' and gpu_width == 'y':
                    return True
        elif category =='B':
            if gpu_discrete == 'y':
                return True
        elif category =='A':
            return True
        return False

    # Requirements for Desktop, Integrated Desktop, and Notebook Computers
    def equation_one(self, computer, P_OFF, P_SLEEP, P_IDLE):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, and Notebook Computers"""
        if computer == '3':
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

        E_TEC = ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE)) * 8760 / 1000

        return E_TEC

    def equation_two(self, computer, gpu_type, gpu_width):
        """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop"""

        result = []

        ## Maximum TEC Allowances for Desktop and Integrated Desktop Computers
        if computer == '1' or computer == '2':
            if self.disk > 1:
                TEC_STORAGE = 25.0 * (self.disk - 1)
            else:
                TEC_STORAGE = 0.0

            if self.qualify_desktop_category('A', gpu_type, gpu_width):
                TEC_BASE = 148.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if gpu_width == 'y':
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_desktop_category('B', gpu_type, gpu_width):
                TEC_BASE = 175.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if gpu_width == 'y':
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_desktop_category('C', gpu_type, gpu_width):
                TEC_BASE = 209.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if gpu_width == 'y':
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('C', E_TEC_MAX))

            if self.qualify_desktop_category('D', gpu_type, gpu_width):
                TEC_BASE = 234.0

                if self.memory > 4:
                    TEC_MEMORY = 1.0 * (self.memory - 4)
                else:
                    TEC_MEMORY = 0.0

                if gpu_width == 'y':
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('D', E_TEC_MAX))

        ## Maximum TEC Allowances for Notebook Computers
        else:
            if self.memory > 4:
                TEC_MEMORY = 0.4 * (self.memory - 4)
            else:
                TEC_MEMORY = 0.0

            if self.disk > 1:
                TEC_STORAGE = 3.0 * (self.disk - 1)
            else:
                TEC_STORAGE = 0.0

            if self.qualify_netbook_category('A', gpu_type, gpu_width):
                TEC_BASE = 40.0
                TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_netbook_category('B', gpu_type, gpu_width):
                TEC_BASE = 53.0

                if gpu_width == 'y':
                    TEC_GRAPHICS = 3.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_netbook_category('C', gpu_type, gpu_width):
                TEC_BASE = 88.5
                TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('C', E_TEC_MAX))

        return result

    # Requirements for Workstations
    def equation_three():
        """Equation 3: P_TEC Calculation for Workstations""" 
        P_TEC = (P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE) 

    def equation_four():
        """Equation 4: P_TEC_MAX Calculation for Workstations"""
        P_TEC_MAX <= 0.28 * (P_MAX + (N_HDD * 5))

    # Requirements for Small-scale Servers
    def equation_five():
        """Equation 5: Calculation of P_OFF_MAX for Small-scale Servers"""
        P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

    # Requirements for Thin Clients
    def equation_six():
        """Equation 6: Calculation of P_OFF_MAX for Thin Clients"""
        P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

    def equation_seven():
        """Equation 7: Calculation of P_SLEEP_MAX for Thin Clients"""
        P_SLEEP_MAX = P_SLEEP_BASE + P_SLEEP_WOL

def main():
    sysinfo = SystemInformation()

    product_type = question_str("""Which product type would you like to verify?
 [1] Desktop, Integrated Desktop, and Notebook Computers
 [2] Workstations (Not implemented yet)
 [3] Small-scale Servers (Not implemented yet)
 [4] Thin Clients (Not implemented yet)""", 1, "1234")

    if product_type == '1':
        computer_type = question_str("""Which type of computer do you use?
 [1] Desktop
 [2] Integrated Desktop
 [3] Notebook""", 1, "123")

        if computer_type == '3':
            gpu_bit = '64'
        else:
            gpu_bit = '128'

        discrete = question_str("Is there a discrete GPU? [y/n]", 1, "yn")

        P_OFF = question_num("What is the power consumption in Off Mode?")
        P_SLEEP = question_num("What is the power consumption in Sleep Mode?")
        P_IDLE = question_num("What is the power consumption in Idle Mode?")

        print("Energy Star 5.2:");
        estar52 = EnergyStar52(sysinfo)
        E_TEC = estar52.equation_one(computer_type, P_OFF, P_SLEEP, P_IDLE)

        over_gpu_width = estar52.equation_two(computer_type, discrete, 'y')
        print("\n  If GPU Frame Buffer Width > %s," % (gpu_bit))
        for i in over_gpu_width:
            (category, E_TEC_MAX) = i
            if E_TEC <= E_TEC_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("    Category %s, E_TEC = %s, E_TEC_MAX = %s, %s" % (category, E_TEC, E_TEC_MAX, result))

        under_gpu_width = estar52.equation_two(computer_type, discrete, 'n')
        print("\n  If GPU Frame Buffer Width <= %s," % (gpu_bit))
        for i in under_gpu_width:
            (category, E_TEC_MAX) = i
            if E_TEC <= E_TEC_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("    Category %s, E_TEC = %s, E_TEC_MAX = %s, %s" % (category, E_TEC, E_TEC_MAX, result))

    elif product_type == '2':
        raise Exception('Not implemented yet.')
    elif product_type == '3':
        raise Exception('Not implemented yet.')
    elif product_type == '4':
        raise Exception('Not implemented yet.')
    else:
        raise Exception('This is a bug when you see this.')

if __name__ == '__main__':
    main()
