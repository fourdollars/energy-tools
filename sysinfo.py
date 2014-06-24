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

import json
import math
import os
import re
import subprocess

class SysInfo:
    def question_str(self, prompt, length, validator, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = raw_input(prompt + "\n>> ")
            if len(s) == length and set(s).issubset(validator):
                print('-'*80)
                self.profile[name] = s
                return s
            print("The valid input '" + validator + "'.")

    def question_bool(self, prompt, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = raw_input(prompt + " [y/n]\n>> ")
            if len(s) == 1 and set(s).issubset("YyNn01"):
                print('-'*80)
                if s == 'Y' or s == 'y' or s == '1':
                    self.profile[name] = True
                    return True
                else:
                    self.profile[name] = False
                    return False

    def question_int(self, prompt, maximum, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = raw_input(prompt + "\n>> ")
            if set(s).issubset("0123456789"):
                try:
                    num = int(s)
                    if num <= maximum:
                        print('-'*80)
                        self.profile[name] = num
                        return num
                except ValueError:
                    print("Please input a positive integer less than or equal to %s." % (maximum))
            print("Please input a positive integer less than or equal to %s." % (maximum))

    def question_num(self, prompt, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = raw_input(prompt + "\n>> ")
            try:
                num = float(s)
                print('-'*80)
                self.profile[name] = num
                return num
            except ValueError:
                print "Oops!  That was no valid number.  Try again..."

    def __init__(self, profile=None):
        if profile:
            self.profile = profile
        else:
            self.profile = {}

        # Product type
        self.product_type = self.question_int("""Which product type would you like to verify?
[1] Desktop, Integrated Desktop, and Notebook Computers
[2] Workstations
[3] Small-scale Servers
[4] Thin Clients""", 4, "Product Type")

        if self.product_type == 1:
            # Computer type
            self.computer_type = self.question_int("""Which type of computer do you use?
[1] Desktop
[2] Integrated Desktop
[3] Notebook""", 3, "Computer Type")
            self.tvtuner = self.question_bool("Is there a television tuner?", "TV Tuner")

            if self.computer_type != 3:
                self.audio = self.question_bool("Is there a discrete audio?", "Discrete Audio")
            else:
                self.audio = False

            # GPU Information
            self.discrete_gpu_num = self.question_num("How many discrete graphics cards?", "Discrete Graphics Cards")
            if self.discrete_gpu_num > 0:
                self.switchable = False
                self.discrete = True
            elif self.question_bool("Does it have switchable graphics and automated switching enabled by default?", "Switchable Graphics"):
                self.switchable = True
                # Those with switchable graphics may not apply the Discrete Graphics allowance.
                self.discrete = False
            else:
                self.switchable = False
                self.discrete = False

            # Screen size
            if self.computer_type != 1:
                self.diagonal = self.question_num("What is the display diagonal in inches?", "Display Diagonal")
                self.ep = self.question_bool("Is there an Enhanced-perforcemance Integrated Display?", "Enhanced Display")

            # Power Consumption
            self.off = self.question_num("What is the power consumption in Off Mode?", "Off Mode")
            self.off_wol = self.question_num("What is the power consumption in Off Mode with Wake-on-LAN enabled?", "Off Mode with WOL")
            self.sleep = self.question_num("What is the power consumption in Sleep Mode?", "Sleep Mode")
            self.sleep_wol = self.question_num("What is the power consumption in Sleep Mode with Wake-on-LAN enabled?", "Sleep Mode with WOL")
            self.long_idle = self.question_num("What is the power consumption in Long Idle Mode?", "Long Idle Mode")
            self.short_idle = self.question_num("What is the power consumption in Short Idle Mode?", "Short Idle Mode")
        elif self.product_type == 2:
            self.off = self.question_num("What is the power consumption in Off Mode?", "Off Mode")
            self.sleep = self.question_num("What is the power consumption in Sleep Mode?", "Sleep Mode")
            self.long_idle = self.question_num("What is the power consumption in Long Idle Mode?", "Long Idle Mode")
            self.short_idle = self.question_num("What is the power consumption in Short Idle Mode?", "Short Idle Mode")
            self.max_power = self.question_num("What is the maximum power consumption?", "Maximum Power")
        elif self.product_type == 3:
            self.off = self.question_num("What is the power consumption in Off Mode?", "Off Mode")
            self.short_idle = self.question_num("What is the power consumption in Short Idle Mode?", "Short Idle Mode")
            if self.get_cpu_core() < 2:
                self.more_discrete = self.question_bool("Does it have more than one discrete graphics device?", "More Discrete Graphics")
        elif self.product_type == 4:
            self.off = self.question_num("What is the power consumption in Off Mode?", "Off Mode")
            self.sleep = self.question_num("What is the power consumption in Sleep Mode?\n(You can input the power consumption in Long Idle Mode, if it lacks a discrete System Sleep Mode)", "Sleep Mode")
            self.long_idle = self.question_num("What is the power consumption in Long Idle Mode?", "Long Idle Mode")
            self.short_idle = self.question_num("What is the power consumption in Short Idle Mode?", "Short Idle Mode")
            self.media_codec = self.question_bool("Does it support local multimedia encode/decode?", "Media Codec")
            self.discrete = self.question_bool("Does it have discrete graphics?", "Discrete Graphics")
            self.integrated_display = self.question_bool("Does it have integrated display?", "Integrated Display")
            if self.integrated_display:
                self.diagonal = self.question_num("What is the display diagonal in inches?", "Display Diagonal")
                self.ep = self.question_bool("Is it an Enhanced-perforcemance Integrated Display?", "Enhanced Display")

        # Ethernet
        if self.profile:
            self.eee = self.profile["Gigabit Ethernet"]
        else:
            self.eee = 0
            for eth in os.listdir("/sys/class/net/"):
                if eth.startswith('eth') and subprocess.call('sudo ethtool %s | grep 1000 >/dev/null 2>&1' % eth, shell=True) == 0:
                    self.eee = self.eee + 1

        if self.profile:
            if "Disk Number" in self.profile:
                self.disk_num = self.profile["Disk Number"]
            if "CPU Cores" in self.profile:
                self.cpu_core = self.profile["CPU Cores"]
            if "CPU Clock" in self.profile:
                self.cpu_clock = self.profile["CPU Clock"]
            if "Memory Size" in self.profile:
                self.mem_size = self.profile["Memory Size"]

    def _get_cpu_vendor(self):
        vendor=subprocess.check_output("cat /proc/cpuinfo | grep 'vendor_id' | grep -ioE '(intel|amd)'", shell=True).strip()
        if re.match("intel", vendor, re.IGNORECASE):
            return 'intel'
        if re.match("amd", vendor, re.IGNORECASE):
            return 'amd'
        return 'unknown'

    def get_cpu_core(self):
        if self.profile:
            self.cpu_core = self.profile["CPU Cores"]
            return self.cpu_core
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
        if self.profile:
            self.cpu_clock = self.profile["CPU Clock"]
            return self.cpu_clock
        if self.cpu_clock:
            return self.cpu_clock

        cpu = self._get_cpu_vendor()
        if cpu == 'intel':
            self.cpu_clock = float(subprocess.check_output("cat /proc/cpuinfo | grep 'model name' | sort -u | cut -d: -f2 | cut -d@ -f2 | xargs | sed 's/GHz//'", shell=True).strip())
        elif cpu == 'amd':
            self.cpu_clock = float(subprocess.check_output("sudo dmidecode -t processor | grep 'Max Speed' | cut -d: -f2 | xargs | sed 's/MHz//'", shell=True).strip()) / 1000
        else:
            raise Exception('Unknown CPU Vendor')

        debug("CPU clock: %s GHz" % (self.cpu_clock))
        return self.cpu_clock

    def get_mem_size(self):
        if self.profile:
            self.mem_size = self.profile["Memory Size"]
            return self.mem_size
        if self.mem_size:
            return self.mem_size

        for size in subprocess.check_output("sudo dmidecode -t 17 | grep 'Size:.*MB' | awk '{print $2}'", shell=True).split('\n'):
            if size:
                self.mem_size = self.mem_size + int(size)
        self.mem_size = self.mem_size / 1024

        debug("Memory size: %s GB" % (self.mem_size))
        return self.mem_size

    def get_disk_num(self):
        if self.profile:
            self.disk_num = self.profile["Disk Number"]
            return self.disk_num
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
        if self.profile:
            self.width = self.profile["Display Width"]
            self.height = self.profile["Display Height"]
            return (self.width, self.height)
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

    def save(self, filename):
        with open(filename, "w") as data:
            data.write(json.dumps(self.profile, sort_keys=True, indent=4, separators=(',', ': ')) + '\n')
