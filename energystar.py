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

import subprocess

def debug(message):
    DEBUG=False
    if DEBUG:
        print(message)

def question_str(prompt, length, validator):
    while True:
        s = raw_input(prompt + "\n>> ")
        if len(s) == length and set(s).issubset(validator):
            print('')
            return s
        print("The valid input '" + validator + "'.")

def question_bool(prompt):
    while True:
        s = raw_input(prompt + " [y/n]\n>> ")
        if len(s) == 1 and set(s).issubset("YyNn01"):
            print('')
            if s == 'Y' or s == 'y' or s == '1':
                return True
            else:
                return False

def question_int(prompt, maximum):
    while True:
        s = raw_input(prompt + "\n>> ")
        if set(s).issubset("0123456789"):
            try:
                num = int(s)
                if num <= maximum:
                    print('')
                    return num
            except ValueError:
                print("Please input a positive integer less than or equal to %s." % (maximum))
        print("Please input a positive integer less than or equal to %s." % (maximum))

def question_num(prompt):
    while True:
        s = raw_input(prompt + "\n>> ")
        try:
            num = float(s)
            print('')
            return num
        except ValueError:
            print "Oops!  That was no valid number.  Try again..."


class SysInfo:
    def __init__(self,
            auto=False,
            cpu_core=0,
            cpu_clock=0.0,
            mem_size=0.0,
            disk_num=0,
            width=0,
            height=0,
            diagonal=0,
            ep=False,
            w=0,
            h=0,
            product_type=0,
            computer_type=0,
            off=0,
            sleep=0,
            long_idle=0,
            short_idle=0,
            eee=0,
            discrete=False,
            switchable=False,
            power_supply=''):
        self.auto = auto
        self.cpu_core = cpu_core
        self.cpu_clock = cpu_clock
        self.mem_size = mem_size
        self.disk_num = disk_num
        self.w = w
        self.h = h
        self.width = width
        self.height = height
        self.diagonal = diagonal
        self.ep = ep
        self.product_type = product_type
        self.computer_type = computer_type
        self.eee = eee
        self.off = off
        self.sleep = sleep
        self.long_idle = long_idle
        self.short_idle = short_idle
        self.discrete = discrete
        self.switchable = switchable
        self.power_supply = power_supply

    def get_cpu_core(self):
        if self.cpu_core:
            return self.cpu_core

        try:
            subprocess.check_output('cat /proc/cpuinfo | grep cores', shell=True)
        except subprocess.CalledProcessError:
            self.cpu_core = 1
        else:
            self.cpu_core = int(subprocess.check_output('cat /proc/cpuinfo | grep "cpu cores" | sort -ru | head -n 1 | cut -d: -f2 | xargs', shell=True).strip())

        return self.cpu_core

    def get_cpu_clock(self):
        if self.cpu_clock:
            return self.cpu_clock

        self.cpu_clock = float(subprocess.check_output("cat /proc/cpuinfo | grep 'model name' | sort -u | cut -d: -f2 | cut -d@ -f2 | xargs | sed 's/GHz//'", shell=True).strip())

        return self.cpu_clock

    def get_mem_size(self):
        if self.mem_size:
            return self.mem_size

        for size in subprocess.check_output("sudo dmidecode -t 17 | grep Size | awk '{print $2}'", shell=True).split('\n'):
            if size:
                self.mem_size = self.mem_size + int(size)
        self.mem_size = self.mem_size / 1024

        return self.mem_size

    def get_disk_num(self):
        if self.disk_num:
            return self.disk_num

        self.disk_num = len(subprocess.check_output('ls /sys/block | grep sd', shell=True).strip().split('\n'))

        return self.disk_num

    def set_display(self, width, height, diagonal, ep):
        self.width = width
        self.height = height
        self.diagonal = diagonal
        self.ep = ep

    def get_display(self):
        return (self.width, self.height, self.diagonal, self.ep)

    def get_resolution(self):
        if self.w == 0 or self.h == 0:
            (width, height) = subprocess.check_output("xrandr --current | grep current | sed 's/.*current \\([0-9]*\\) x \\([0-9]*\\).*/\\1 \\2/'", shell=True).strip().split(' ')
            self.w = int(width)
            self.h = int(height)
        return (self.w, self.h)

class EnergyStar52:
    """Energy Star 5.2 calculator"""
    def __init__(self, sysinfo):
        self.core = sysinfo.get_cpu_core()
        self.disk = sysinfo.get_disk_num()
        self.memory = sysinfo.get_mem_size()
        self.sysinfo = sysinfo

    def qualify_desktop_category(self, category, gpu_discrete, gpu_width):
        if category == 'D':
            if self.core >= 4:
                if self.memory >= 4:
                    return True
                elif gpu_width:
                    return True
        elif category == 'C':
            if self.core >= 2:
                if self.memory >= 2:
                    return True
                elif gpu_discrete:
                    return True
        elif category == 'B':
            if self.core == 2 and self.memory >= 2:
                return True
        elif category == 'A':
            return True
        return False

    def qualify_netbook_category(self, category, gpu_discrete, over_gpu_width):
        if category =='C':
            if self.core >= 2 and self.memory >= 2:
                if gpu_discrete and over_gpu_width:
                    return True
        elif category =='B':
            if gpu_discrete:
                return True
        elif category =='A':
            return True
        return False

    # Requirements for Desktop, Integrated Desktop, and Notebook Computers
    def equation_one(self, P_OFF, P_SLEEP, P_IDLE):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, and Notebook Computers"""
        if self.sysinfo.computer_type == 3:
            (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

        E_TEC = ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_IDLE * T_IDLE)) * 8760 / 1000

        return E_TEC

    def equation_two(self, over_gpu_width):
        """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop, and Notebook Computers"""

        result = []

        ## Maximum TEC Allowances for Desktop and Integrated Desktop Computers
        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            if self.disk > 1:
                TEC_STORAGE = 25.0 * (self.disk - 1)
            else:
                TEC_STORAGE = 0.0

            if self.qualify_desktop_category('A', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 148.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_gpu_width:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_desktop_category('B', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 175.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_gpu_width == 'y':
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 35.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_desktop_category('C', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 209.0

                if self.memory > 2:
                    TEC_MEMORY = 1.0 * (self.memory - 2)
                else:
                    TEC_MEMORY = 0.0

                if over_gpu_width:
                    TEC_GRAPHICS = 50.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('C', E_TEC_MAX))

            if self.qualify_desktop_category('D', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 234.0

                if self.memory > 4:
                    TEC_MEMORY = 1.0 * (self.memory - 4)
                else:
                    TEC_MEMORY = 0.0

                if over_gpu_width:
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

            if self.qualify_netbook_category('A', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 40.0
                TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('A', E_TEC_MAX))

            if self.qualify_netbook_category('B', self.sysinfo.discrete, over_gpu_width):
                TEC_BASE = 53.0

                if over_gpu_width:
                    TEC_GRAPHICS = 3.0
                else:
                    TEC_GRAPHICS = 0.0

                E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
                result.append(('B', E_TEC_MAX))

            if self.qualify_netbook_category('C', self.sysinfo.discrete, over_gpu_width):
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

class EnergyStar60:
    """Energy Star 6.0 calculator"""
    def __init__(self, sysinfo):
        self.clock = sysinfo.get_cpu_clock()
        self.core = sysinfo.get_cpu_core()
        self.disk = sysinfo.get_disk_num()
        self.memory = sysinfo.get_mem_size()
        self.sysinfo = sysinfo

    # Requirements for Desktop, Integrated Desktop, and Notebook Computers
    def equation_one(self, P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE):
        """Equation 1: TEC Calculation (E_TEC) for Desktop, Integrated Desktop, Thin Client and Notebook Computers"""
        if self.sysinfo.computer_type == 3:
            (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.25, 0.35, 0.1, 0.3)
        else:
            (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

        E_TEC = ((P_OFF * T_OFF) + (P_SLEEP * T_SLEEP) + (P_LONG_IDLE * T_LONG_IDLE) + (P_SHORT_IDLE * T_SHORT_IDLE)) * 8760 / 1000

        return E_TEC

    def equation_two(self, gpu_category):
        """Equation 2: E_TEC_MAX Calculation for Desktop, Integrated Desktop, and Notebook Computers"""

        P = self.core * self.clock

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            if P <= 3:
                TEC_BASE = 69.0
            elif self.sysinfo.switchable or self.sysinfo.discrete:
                if P <= 6:
                    TEC_BASE = 112.0
                elif P <= 7:
                    TEC_BASE = 120.0
                else:
                    TEC_BASE = 135.0
            else:
                if P <= 9:
                    TEC_BASE = 115.0
                else:
                    TEC_BASE = 135.0
        else:
            if P <= 2:
                TEC_BASE = 14.0
            elif self.sysinfo.switchable or self.sysinfo.discrete:
                if P <= 5.2:
                    TEC_BASE = 22.0
                elif P <= 8:
                    TEC_BASE = 24.0
                else:
                    TEC_BASE = 28.0
            else:
                if P <= 9:
                    TEC_BASE = 16.0
                else:
                    TEC_BASE = 18.0
        debug("TEC_BASE = %s" % (TEC_BASE))

        TEC_MEMORY = 0.8 * self.memory
        debug("TEC_MEMORY = %s" % (TEC_MEMORY))

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
                if gpu_category == 'G1':
                    TEC_GRAPHICS = 14
                elif gpu_category == 'G2':
                    TEC_GRAPHICS = 20
                elif gpu_category == 'G3':
                    TEC_GRAPHICS = 26
                elif gpu_category == 'G4':
                    TEC_GRAPHICS = 32
                elif gpu_category == 'G5':
                    TEC_GRAPHICS = 42
                elif gpu_category == 'G6':
                    TEC_GRAPHICS = 48
                elif gpu_category == 'G7':
                    TEC_GRAPHICS = 60
        else:
            TEC_GRAPHICS = 0
        debug("TEC_GRAPHICS = %s" % (TEC_GRAPHICS))

        if self.sysinfo.switchable:
            if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
                TEC_SWITCHABLE = 0.5 * TEC_GRAPHICS
            else:
                TEC_SWITCHABLE = 0
        else:
            TEC_SWITCHABLE = 0
        debug("TEC_SWITCHABLE = %s" % (TEC_SWITCHABLE))

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * self.sysinfo.eee
        else:
            TEC_EEE = 8.76 * 0.2 * (0.10 + 0.30) * self.sysinfo.eee
        debug("TEC_EEE = %s" % (TEC_EEE))

        if self.sysinfo.computer_type == 1 or self.sysinfo.computer_type == 2:
            TEC_STORAGE = 26 * (self.disk - 1)
        else:
            TEC_STORAGE = 2.6 * (self.disk - 1)
        debug("TEC_STORAGE = %s" % (TEC_STORAGE))

        if self.sysinfo.computer_type != 1:
            (width, height, diagonal, enhanced_performance_display) = self.sysinfo.get_display()
            if enhanced_performance_display:
                if diagonal:
                    EP = 0.75
                else:
                    EP = 0.3
            else:
                EP = 0
            (w, h) = self.sysinfo.get_resolution()
            r = 1.0 * w * h / 1000000
            A = width * height
            debug("EP = %s, r = %s, A = %s" % (EP, r, A))

        if self.sysinfo.computer_type == 2:
            TEC_INT_DISPLAY = 8.76 * 0.35 * (1 + EP) * (4 * r + 0.05 * A)
        elif self.sysinfo.computer_type == 3:
            TEC_INT_DISPLAY = 8.76 * 0.30 * (1 + EP) * (2 * r + 0.02 * A)
        else:
            TEC_INT_DISPLAY = 0
        debug("TEC_INT_DISPLAY = %s" % (TEC_INT_DISPLAY))

        return TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE + TEC_INT_DISPLAY + TEC_SWITCHABLE + TEC_EEE

def qualifying(sysinfo):
    if not sysinfo.auto:
        product_type = question_int("""Which product type would you like to verify?
 [1] Desktop, Integrated Desktop, and Notebook Computers
 [2] Workstations (Not implemented yet)
 [3] Small-scale Servers (Not implemented yet)
 [4] Thin Clients (Not implemented yet)""", 4)
        sysinfo.product_type = product_type
    else:
        product_type = sysinfo.product_type

    if product_type == 1:
        if not sysinfo.auto:
            computer_type = question_int("""Which type of computer do you use?
 [1] Desktop
 [2] Integrated Desktop
 [3] Notebook""", 3)
            sysinfo.computer_type = computer_type
        else:
            computer_type = sysinfo.computer_type

        # GPU Information
        if not sysinfo.auto:
            if question_bool("Is there a discrete GPU?"):
                sysinfo.discrete = True
            else:
                sysinfo.discrete = False
            if question_bool("Is it a switchable GPU?"):
                sysinfo.switchable = True
            else:
                sysinfo.switchable = False

        # Power Consumption
        if not sysinfo.auto:
            P_OFF = question_num("What is the power consumption in Off Mode?")
            sysinfo.off = P_OFF
        else:
            P_OFF = sysinfo.off
        if not sysinfo.auto:
            P_SLEEP = question_num("What is the power consumption in Sleep Mode?")
            sysinfo.sleep = P_SLEEP
        else:
            P_SLEEP = sysinfo.sleep
        if not sysinfo.auto:
            P_LONG_IDLE = question_num("What is the power consumption in Long Idle Mode?")
            sysinfo.long_idle = P_LONG_IDLE
        else:
            P_LONG_IDLE = sysinfo.long_idle
        if not sysinfo.auto:
            P_SHORT_IDLE = question_num("What is the power consumption in Short Idle Mode?")
            sysinfo.short_idle = P_SHORT_IDLE
        else:
            P_SHORT_IDLE = sysinfo.short_idle

        # Energy Star 5.2
        print("Energy Star 5.2:");
        estar52 = EnergyStar52(sysinfo)
        E_TEC = estar52.equation_one(P_OFF, P_SLEEP, P_SHORT_IDLE)

        if computer_type == 3:
            gpu_bit = '64'
        else:
            gpu_bit = '128'

        over_gpu_width = estar52.equation_two(True)
        print("\n  If GPU Frame Buffer Width > %s bits," % (gpu_bit))
        for i in over_gpu_width:
            (category, E_TEC_MAX) = i
            if E_TEC <= E_TEC_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("    Category %s, E_TEC = %s, E_TEC_MAX = %s, %s" % (category, E_TEC, E_TEC_MAX, result))

        under_gpu_width = estar52.equation_two(False)
        print("\n  If GPU Frame Buffer Width <= %s bits," % (gpu_bit))
        for i in under_gpu_width:
            (category, E_TEC_MAX) = i
            if E_TEC <= E_TEC_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("    Category %s, E_TEC = %s, E_TEC_MAX = %s, %s" % (category, E_TEC, E_TEC_MAX, result))

        # Power Supply
        if not sysinfo.auto:
            power_supply = question_str("\nDoes it use external power supply or internal power supply? [e/i]", 1, "ei")
            sysinfo.power_supply = power_supply
        else:
            power_supply = sysinfo.power_supply

        # Screen size
        if computer_type != '1':
            if not sysinfo.auto:
                width = question_num("What is the physical width of the display in inches?")
                height = question_num("What is the physical height of the display in inches?")
                diagonal = question_bool("Is the physical diagonal of the display bigger than or equal to 27 inches?")
                ep = question_bool("Is there an Enhanced Perforcemance Display?")
                sysinfo.set_display(width, height, diagonal, ep)

        if not sysinfo.auto:
            sysinfo.eee = question_num("How many Gigabit Ethernet ports?")

        # Energy Star 6.0
        print("\nEnergy Star 6.0:\n");
        estar60 = EnergyStar60(sysinfo)
        E_TEC = estar60.equation_one(P_OFF, P_SLEEP, P_LONG_IDLE, P_SHORT_IDLE)

        if sysinfo.discrete:
            for gpu in ('G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7'):
                E_TEC_MAX = estar60.equation_two(gpu)
                if E_TEC <= E_TEC_MAX:
                    result = 'PASS'
                else:
                    result = 'FAIL'
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
                print("  E_TEC = %s, E_TEC_MAX = %s for %s, %s" % (E_TEC, E_TEC_MAX, gpu, result))
        else:
            E_TEC_MAX = estar60.equation_two('G1')
            if E_TEC <= E_TEC_MAX:
                result = 'PASS'
            else:
                result = 'FAIL'
            print("  E_TEC = %s, E_TEC_MAX %s, %s" % (E_TEC, E_TEC_MAX, result))

    elif product_type == '2':
        raise Exception('Not implemented yet.')
    elif product_type == '3':
        raise Exception('Not implemented yet.')
    elif product_type == '4':
        raise Exception('Not implemented yet.')
    else:
        raise Exception('This is a bug when you see this.')

def main():
    sysinfo = SysInfo()
#    sysinfo = SysInfo(
#            auto=True,
#            product_type=1, computer_type=3,
#            cpu_core=2, cpu_clock=2.0,
#            mem_size=8, disk_num=1,
#            w=1366, h=768, eee=1,
#            width=12, height=6.95, diagonal=False,
#            discrete=False, switchable=False,
#            off=1.0, sleep=1.7, long_idle=8.0, short_idle=10.0)
    qualifying(sysinfo)

if __name__ == '__main__':
    main()
