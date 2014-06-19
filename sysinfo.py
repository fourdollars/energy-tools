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

from logging import debug, warning

import math
import os
import subprocess

def question_str(prompt, length, validator):
    while True:
        s = raw_input(prompt + "\n>> ")
        if len(s) == length and set(s).issubset(validator):
            print('-'*80)
            return s
        print("The valid input '" + validator + "'.")

def question_bool(prompt):
    while True:
        s = raw_input(prompt + " [y/n]\n>> ")
        if len(s) == 1 and set(s).issubset("YyNn01"):
            print('-'*80)
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
                    print('-'*80)
                    return num
            except ValueError:
                print("Please input a positive integer less than or equal to %s." % (maximum))
        print("Please input a positive integer less than or equal to %s." % (maximum))

def question_num(prompt):
    while True:
        s = raw_input(prompt + "\n>> ")
        try:
            num = float(s)
            print('-'*80)
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
            diagonal=0,
            ep=False,
            width=0,
            height=0,
            product_type=0,
            computer_type=0,
            off=0,
            off_wol=0,
            sleep=0,
            sleep_wol=0,
            long_idle=0,
            short_idle=0,
            eee=-1,
            tvtuner=False,
            audio=False,
            discrete=False,
            discrete_gpu_num=0,
            switchable=False,
            max_power=0,
            more_discrete=False,
            media_codec=False,
            integrated_display=False):
        self.auto = auto
        self.cpu_core = cpu_core
        self.cpu_clock = cpu_clock
        self.mem_size = mem_size
        self.disk_num = disk_num
        self.width = width
        self.height = height
        self.diagonal = diagonal
        self.ep = ep
        self.product_type = product_type
        self.computer_type = computer_type
        self.eee = eee
        self.tvtuner = tvtuner
        self.audio = audio
        self.off = off
        self.off_wol = off_wol
        self.sleep = sleep
        self.sleep_wol = sleep_wol
        self.long_idle = long_idle
        self.short_idle = short_idle
        self.discrete = discrete
        self.discrete_gpu_num = discrete_gpu_num
        self.switchable = switchable
        self.max_power = max_power
        self.more_discrete = more_discrete
        self.media_codec = media_codec
        self.integrated_display = integrated_display

        if auto:
            return

        if eee != -1:
            self.eee = eee
        else:
            self.eee = 0
            for eth in os.listdir("/sys/class/net/"):
                if eth.startswith('eth') and subprocess.call('sudo ethtool %s | grep 1000 >/dev/null 2>&1' % eth, shell=True) == 0:
                    self.eee = self.eee + 1

        # Product type
        self.product_type = question_int("""Which product type would you like to verify?
[1] Desktop, Integrated Desktop, and Notebook Computers
[2] Workstations
[3] Small-scale Servers
[4] Thin Clients""", 4)

        if self.product_type == 1:
            # Computer type
            self.computer_type = question_int("""Which type of computer do you use?
[1] Desktop
[2] Integrated Desktop
[3] Notebook""", 3)
            self.tvtuner = question_bool("Is there a television tuner?")

            if self.computer_type != 3:
                self.audio = question_bool("Is there a discrete audio?")

            # GPU Information
            self.discrete_gpu_num = question_num("How many discrete graphics cards?")
            if self.discrete_gpu_num > 0:
                self.switchable = False
                self.discrete = True
            elif question_bool("Does it have switchable graphics and automated switching enabled by default?"):
                self.switchable = True
                # Those with switchable graphics may not apply the Discrete Graphics allowance.
                self.discrete = False
            else:
                self.switchable = False
                self.discrete = False

            # Screen size
            if self.computer_type != 1:
                self.diagonal = question_num("What is the display diagonal in inches?")
                self.ep = question_bool("Is there an Enhanced-perforcemance Integrated Display?")

            # Power Consumption
            self.off = question_num("What is the power consumption in Off Mode?")
            self.off_wol = question_num("What is the power consumption in Off Mode with Wake-on-LAN enabled?")
            self.sleep = question_num("What is the power consumption in Sleep Mode?")
            self.sleep_wol = question_num("What is the power consumption in Sleep Mode with Wake-on-LAN enabled?")
            self.long_idle = question_num("What is the power consumption in Long Idle Mode?")
            self.short_idle = question_num("What is the power consumption in Short Idle Mode?")
        elif self.product_type == 2:
            self.off = question_num("What is the power consumption in Off Mode?")
            self.sleep = question_num("What is the power consumption in Sleep Mode?")
            self.long_idle = question_num("What is the power consumption in Long Idle Mode?")
            self.short_idle = question_num("What is the power consumption in Short Idle Mode?")
            self.max_power = question_num("What is the maximum power consumption?")
        elif self.product_type == 3:
            self.off = question_num("What is the power consumption in Off Mode?")
            self.short_idle = question_num("What is the power consumption in Short Idle Mode?")
            if self.get_cpu_core() < 2:
                self.more_discrete = question_bool("Does it have more than one discrete graphics device?")
        elif self.product_type == 4:
            self.off = question_num("What is the power consumption in Off Mode?")
            self.sleep = question_num("What is the power consumption in Sleep Mode?\n(You can input the power consumption in Long Idle Mode, if it lacks a discrete System Sleep Mode)")
            self.long_idle = question_num("What is the power consumption in Long Idle Mode?")
            self.short_idle = question_num("What is the power consumption in Short Idle Mode?")
            self.media_codec = question_bool("Does it support local multimedia encode/decode?")
            self.discrete = question_bool("Does it have discrete graphics?")
            self.integrated_display = question_bool("Does it have integrated display?")
            if self.integrated_display:
                self.diagonal = question_num("What is the display diagonal in inches?")
                self.ep = question_bool("Is it an Enhanced-perforcemance Integrated Display?")

    def get_cpu_core(self):
        if self.cpu_core:
            return self.cpu_core

        try:
            subprocess.check_output('cat /proc/cpuinfo | grep cores', shell=True)
        except subprocess.CalledProcessError:
            self.cpu_core = 1
        else:
            self.cpu_core = int(subprocess.check_output('cat /proc/cpuinfo | grep "cpu cores" | sort -ru | head -n 1 | cut -d: -f2 | xargs', shell=True).strip())

        debug("CPU core: %s" % (self.cpu_core))
        return self.cpu_core

    def get_cpu_clock(self):
        if self.cpu_clock:
            return self.cpu_clock

        self.cpu_clock = float(subprocess.check_output("cat /proc/cpuinfo | grep 'model name' | sort -u | cut -d: -f2 | cut -d@ -f2 | xargs | sed 's/GHz//'", shell=True).strip())

        debug("CPU clock: %s GHz" % (self.cpu_clock))
        return self.cpu_clock

    def get_mem_size(self):
        if self.mem_size:
            return self.mem_size

        for size in subprocess.check_output("sudo dmidecode -t 17 | grep 'Size:.*MB' | awk '{print $2}'", shell=True).split('\n'):
            if size:
                self.mem_size = self.mem_size + int(size)
        self.mem_size = self.mem_size / 1024

        debug("Memory size: %s GB" % (self.mem_size))
        return self.mem_size

    def get_disk_num(self):
        if self.disk_num:
            return self.disk_num

        self.disk_num = len(subprocess.check_output('ls /sys/block | grep sd', shell=True).strip().split('\n'))

        debug("Disk number: %s" % (self.disk_num))
        return self.disk_num

    def set_display(self, diagonal, ep):
        self.diagonal = diagonal 
        self.ep = ep

    def get_display(self):
        return (self.diagonal, self.ep)

    def get_resolution(self):
        if self.width == 0 or self.height == 0:
            (width, height) = subprocess.check_output("xrandr --current | grep current | sed 's/.*current \\([0-9]*\\) x \\([0-9]*\\).*/\\1 \\2/'", shell=True).strip().split(' ')
            self.width = int(width)
            self.height = int(height)
        debug("Resolution: %s x %s" % (self.width, self.height))
        return (self.width, self.height)

    def get_power_consumptions(self):
        return (self.off, self.sleep, self.long_idle, self.short_idle)

    def get_basic_info(self):
        return (self.get_cpu_core(), self.get_cpu_clock(), self.get_mem_size(), self.get_disk_num())
