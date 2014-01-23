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

GPUFrameBufferWidth = "Is GPU Frame Buffer Width bigger than %s-bit? [y/n]"
DiscreteGPU = "Is there a Discrete GPU? [y/n]"

def question(prompt, length, validator):
    while True:
        s = raw_input(prompt + ' ')
        if len(s) == length and set(s).issubset(validator):
            print('')
            return s
        print("The valid input '" + validator + "'.")

import inspect
def lineno():
    """Returns the current line number in our program."""
    return " (DEBUG: line %s)" % str(inspect.currentframe().f_back.f_lineno) 

# Requirements for Desktop, Integrated Desktop, and Notebook Computers
def equation_one(si):
    """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, and Notebook Computers"""
    E_TEC = (8760 / 1000) * ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE))

def equation_two(si):
    """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop
    disk: the number of internal storage elements
    gpu:  GPU Frame Buffer Width
    """

    computer = question("""[1] Desktop
[2] Integrated Desktop
[3] Notebook
Which type of computer do you use? [1-3] """, 1, "123")

    core = si.get_cpu_info()
    memory = si.get_memory_info()
    category = ''
    gpu_width = ''
    gpu_type = ''

    ## Determine the category
    # Desktop & Integrated Desktop
    if computer == '1' or computer == '2':
        if core >= 4:
            if memory < 4:
                gpu_width = question(GPUFrameBufferWidth % (128) + lineno(), 1, "yn")
                if gpu_width == 'y':
                    category = 'D'
            else:
                category = 'D'

        if not category and core >= 2:
            if memory >= 2:
                category = 'C'
            else:
                gpu_type = question(DiscreteGPU + lineno(), 1, "yn")
                if gpu_type == 'y':
                    category = 'C'

        if not category and core == 2:
            if memory >= 2:
                category = 'B'

        if not category:
            category = 'A'
    # Notebook
    else:
        if core >= 2 and memory >= 2:
            gpu_width = question(GPUFrameBufferWidth % (64) + lineno(), 1, "yn")
            if gpu_width == 'y':
                category = 'C'

        if not category:
            gpu_type = question(DiscreteGPU + lineno(), 1, "yn")
            if gpu_type == 'y':
                category = 'B'

        if not category:
            category = 'A'

    print("Category %s" % (category))

    ## Maximum TEC Allowances for Desktop and Integrated Desktop Computers
    if computer == '1' or computer == '2':
        if category == 'A':
            TEC_BASE = 148.0

            if memory > 2:
                TEC_MEMORY = 1.0 * (memory - 2)
            else:
                TEC_MEMORY = 0.0

            if not gpu_width:
                gpu_width = question(GPUFrameBufferWidth % (128) + lineno(), 1, "yn")
            if gpu_width == 'y':
                TEC_GRAPHICS = 50.0
            else:
                TEC_GRAPHICS = 35.0
        elif category == 'B':
            TEC_BASE = 175.0

            if memory > 2:
                TEC_MEMORY = 1.0 * (memory - 2)
            else:
                TEC_MEMORY = 0.0

            if not gpu_width:
                gpu_width = question(GPUFrameBufferWidth % (128) + lineno(), 1, "yn")
            if gpu_width == 'y':
                TEC_GRAPHICS = 50.0
            else:
                TEC_GRAPHICS = 35.0
        elif category == 'C':
            TEC_BASE = 209.0

            if memory > 2:
                TEC_MEMORY = 1.0 * (memory - 2)
            else:
                TEC_MEMORY = 0.0

            if not gpu_width:
                gpu_width = question(GPUFrameBufferWidth % (128) + lineno(), 1, "yn")
            if gpu_width == 'y':
                TEC_GRAPHICS = 50.0
            else:
                TEC_GRAPHICS = 0.0
        elif category == 'D':
            TEC_BASE = 234.0

            if memory > 4:
                TEC_MEMORY = 1.0 * (memory - 4)
            else:
                TEC_MEMORY = 0.0

            if not gpu_width:
                gpu_width = question(GPUFrameBufferWidth % (128) + lineno(), 1, "yn")
            if gpu_width == 'y':
                TEC_GRAPHICS = 50.0
            else:
                TEC_GRAPHICS = 0.0
        else:
            raise Exception('This is a bug when you see this.')

        disk = si.get_disk_info()
        if disk > 1:
            TEC_STORAGE = 25.0 * (disk - 1)
        else:
            TEC_STORAGE = 0.0
    ## Maximum TEC Allowances for Notebook Computers
    else:
        if category == 'A':
            TEC_BASE = 40.0
            TEC_GRAPHICS = 0.0
        elif category == 'B':
            TEC_BASE = 53.0
            if not gpu_width:
                gpu_width = question(GPUFrameBufferWidth % (64) + lineno(), 1, "yn")
            if gpu_width == 'y':
                TEC_GRAPHICS = 3.0
            else:
                TEC_GRAPHICS = 0.0
        elif category == 'C':
            TEC_BASE = 88.5
            TEC_GRAPHICS = 0.0
        else:
            raise Exception('This is a bug when you see this.')

        if memory > 4:
            TEC_MEMORY = 0.4 * (memory - 4)
        else:
            TEC_MEMORY = 0.0

        disk = si.get_disk_info()
        if disk > 1:
            TEC_STORAGE = 3.0 * (disk - 1)
        else:
            TEC_STORAGE = 0.0
    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
    return E_TEC_MAX

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
    si = SystemInformation()

    product_type = question("""[1] Desktop, Integrated Desktop, and Notebook Computers
[2] Workstations
[3] Small-scale Servers
[4] Thin Clients
Which product type would you like to verify? [1-4] """, 1, "1234")

    if product_type == '1':
        E_TEC_MAX = equation_two(si)
        print("E_TEC_MAX = %s" % (E_TEC_MAX,))
        E_TEC = equation_one(si)
        if E_TEC <= E_TEC_MAX:
            print("%s (E_TEC) <= %s (E_TEC_MAX) Passed." % (E_TEC, E_TEC_MAX))
        else:
            print("%s (E_TEC) > %s (E_TEC_MAX) Failed." % (E_TEC, E_TEC_MAX))
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
