# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
#
# Copyright (C) 2014-2018 Canonical Ltd.
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

from logging import debug, warning, error
from pathlib import Path

import json
import math
import os
import re
import subprocess


class SysInfo:
    def edid_decode(self):
        monitor = None
        for edid in Path('/sys/devices').glob('**/edid'):
            try:
                with open(edid, 'rb') as f:
                    content = f.read()
                    if len(content) != 0:
                        monitor = edid
                        break
            except PermissionError as err:
                if 'SNAP_NAME' in os.environ and os.environ['SNAP_NAME'] == 'energy-tools':
                    error('Please execute `snap connect energy-tools:hardware-observe` to get the permissions.')
                raise err
        debug("EDID location is %s" % (monitor))

        if monitor is None:
            return None

        width = None
        height = None
        width_mm = None
        height_mm = None

        edid = subprocess.check_output("""edid-decode < %s | \
                                       grep 'Detailed mode:' -A 2""" % monitor,
                                       shell=True,
                                       encoding='utf8',
                                       stderr=open(os.devnull, 'w'))
        debug(edid)

        for line in edid.split('\n'):
            m = re.search(r'(\d+) mm x (\d+) mm', line)
            if m:
                width_mm = m.group(1)
                height_mm = m.group(2)
                continue

            m = re.search(r'(\d+)\s+\d+\s+\d+\s+\d+\s+hborder\s+\d+', line)
            if m:
                width = m.group(1)
                continue

            m = re.search(r'(\d+)\s+\d+\s+\d+\s+\d+\s+vborder\s+\d+', line)
            if m:
                height = m.group(1)
                continue

            if width_mm and height_mm and width and height:
                break
        debug('%s %s %s %s' % (width, height, width_mm, height_mm))

        self.width_mm = int(width_mm)
        self.height_mm = int(height_mm)
        self.width = int(width)
        self.height = int(height)

    def question_str(self, prompt, length, validator, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = input(prompt + "\n>> ")
            if len(s) == length and set(s).issubset(validator):
                print('-'*80)
                self.profile[name] = s
                return s
            print("The valid input '" + validator + "'.")

    def question_bool(self, prompt, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = input(prompt + " [y/n]\n>> ")
            if len(s) == 1 and set(s).issubset("YyNn01"):
                print('-'*80)
                if s == 'Y' or s == 'y' or s == '1':
                    self.profile[name] = True
                    return True
                else:
                    self.profile[name] = False
                    return False

    def question_int(self, prompt, maximum, name=None):
        if name and name in self.profile:
            return self.profile[name]
        while True:
            s = input(prompt + "\n>> ")
            if not set(s).issubset("0123456789"):
                print("Please input a positive integer <= %s." % (maximum))
            try:
                num = int(s)
                if num <= maximum:
                    print('-'*80)
                    if name:
                        self.profile[name] = num
                    return num
            except ValueError:
                print("Please input a positive integer <= %s." % (maximum))

    def question_num(self, prompt, name):
        if name in self.profile:
            return self.profile[name]
        while True:
            s = input(prompt + "\n>> ")
            try:
                num = float(s)
                print('-'*80)
                self.profile[name] = num
                return num
            except ValueError:
                print("Please input a valid number.")

    def get_diagonal(self):
        key = 'Display Diagonal'
        if key in self.profile:
            return self.profile[key]
        else:
            if self.width_mm is None or self.height_mm is None:
                self.edid_decode()
            diagonal_mm = math.sqrt(self.width_mm ** 2 + self.height_mm ** 2)
            self.profile[key] = diagonal_mm / 25.4
            return self.profile[key]

    def get_screen_area(self):
        key = 'Screen Area'
        if key in self.profile:
            return self.profile[key]
        else:
            if self.width_mm is None or self.height_mm is None:
                self.edid_decode()
            self.profile[key] = self.width_mm * self.height_mm / 25.4 / 25.4
            return self.profile[key]

    def __init__(self, profile=None, chassis=0, manual=False):
        self.ep = False
        self.diagonal = 0.0
        self.width = None
        self.height = None
        self.width_mm = None
        self.height_mm = None

        if not manual and not profile and os.path.exists('/sys/class/dmi/id/chassis_type'):
            try:
                with open('/sys/class/dmi/id/chassis_type') as chassis_type:
                    chassis = int(chassis_type.read().strip())
            except PermissionError as err:
                if 'SNAP_NAME' in os.environ and os.environ['SNAP_NAME'] == 'energy-tools':
                    error('Please execute `snap connect energy-tools:hardware-observe` to get the permissions.')
                raise err

        if profile:
            self.profile = profile
        else:
            self.profile = {}

        # Assume there is a X Window System
        if 'DISPLAY' not in os.environ:
            os.environ['DISPLAY'] = ':0'

        product_type = ""

        if chassis == 3 or chassis == 16:
            self.profile['Product Type'] = 1
            self.profile['Computer Type'] = 1
            product_type = "Desktop"
        elif chassis == 13:
            self.profile['Product Type'] = 1
            self.profile['Computer Type'] = 2
            product_type = "Integrated Desktop"
        elif chassis == 10:
            self.profile['Product Type'] = 1
            self.profile['Computer Type'] = 3
            product_type = "Notebook"

        if product_type:
            warning("According to /sys/class/dmi/id/chassis_type = " +
                    str(chassis) + ", this should be a " + product_type +
                    " Computer. If it is wrong, please use '-m' option to skip this detection.")

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
            self.tvtuner = self.question_bool("Is there a television tuner?",
                                              "TV Tuner")

            if self.computer_type != 3:
                self.audio = self.question_bool("Is there a discrete audio?",
                                                "Discrete Audio")
            else:
                self.audio = False

            # GPU Information
            if self.question_bool("""Does it have switchable graphics and automated switching enabled by default?
Switchable Graphics: Functionality that allows Discrete Graphics to be disabled
when not required in favor of Integrated Graphics.""", "Switchable Graphics"):
                self.switchable = True
                # Switchable Graphics can not apply with Discrete Graphics
                self.discrete = False
                self.discrete_gpu_num = 0
                self.fb_bw = 0
                self.profile["FB_BW"] = 0
            else:
                self.discrete_gpu_num = self.question_int(
                    """Discrete Graphics (dGfx):
    A graphics processor (GPU) which must contain a local memory controller
    interface and local graphics-specific memory.
How many discrete graphics cards?""", 10,
                    "Discrete Graphics Cards")
                if self.discrete_gpu_num > 0:
                    self.switchable = False
                    self.discrete = True
                    self.fb_bw = self.question_num("""How many is the display frame buffer bandwidth in gigabytes per second (GB/s) (abbr FB_BW)?
    This is a manufacturer declared parameter and should be calculated as follows
    (Data Rate [Mhz] * Frame Buffer Data Width [bits]) / ( 8 * 1000 ) """,
                                                   "Frame Buffer Bandwidth")
                    self.profile["FB_BW"] = self.fb_bw
                else:
                    self.switchable = False
                    self.discrete = False
                    self.fb_bw = 0
                    self.profile["FB_BW"] = 0

            # Screen size
            if self.computer_type != 1:
                self.diagonal = self.get_diagonal()
                self.screen_area = self.get_screen_area()
                (width, height) = self.get_resolution()
                if width * height >= 2300000:
                    self.ep = self.question_bool("""Is it an Enhanced-perforcemance Integrated Display?
  i. Contrast ratio of at least 60:1 measured at a horizontal viewing angle of at least 85° from the
     perpendicular on a flat screen and at least 83° from the perpendicular on a curved screen,
     with or without a screen cover glass;
 ii. A native resolution greater than or equal to 2.3 megapixels (MP); and
iii. Color Gamut greater than or equal to 32.9% of CIE LUV.""",
                                                 "Enhanced Display")
                else:
                    self.ep = False

            # Check the storage type
            if self.product_type == 1:
                storage_tuple = ("Unknown / System Disk",
                                 "3.5 inch HDD",
                                 "2.5 inch HDD",
                                 "Hybrid HDD/SSD",
                                 "SSD")
                disk_num = self.get_disk_num()

                for disk_type in storage_tuple:
                    if disk_type in self.profile:
                        disk_num = disk_num - self.profile[disk_type]
                    else:
                        self.profile[disk_type] = 0

                if disk_num > 1:
                    for disk in subprocess.check_output(
                            'ls /sys/block | grep -e sd -e nvme -e emmc',
                            shell=True, encoding='utf8').strip().split('\n'):
                        disk_type = self.question_int("""Which storage type for /sys/block/%s?
[0] Unknown / System Disk (by checking the output of `lsblk`)
[1] 3.5" HDD
[2] 2.5" HDD
[3] Hybrid HDD/SSD
[4] SSD (including M.2 port solutions)""" % disk, 4)
                        self.profile[storage_tuple[disk_type]] = \
                            self.profile[storage_tuple[disk_type]] + 1
                else:
                    self.profile["Unknown / System Disk"] = 1

            # Power Consumption
            self.off = self.question_num(
                "What is the power consumption in Off Mode?", "Off Mode")
            if self._check_wol():
                self.profile["Wake-on-LAN"] = True
                self.off_wol = self.question_num(
                    "What is the power consumption in Off Mode with Wake-on-LAN enabled?",
                    "Off Mode with WOL")
            else:
                self.profile["Wake-on-LAN"] = False
                self.off_wol = self.off
            self.sleep = self.question_num(
                "What is the power consumption in Sleep Mode?", "Sleep Mode")
            if self._check_wol():
                self.profile["Wake-on-LAN"] = True
                self.sleep_wol = self.question_num(
                    "What is the power consumption in Sleep Mode with Wake-on-LAN enabled?",
                    "Sleep Mode with WOL")
            else:
                self.profile["Wake-on-LAN"] = False
                self.sleep_wol = self.sleep
            self.long_idle = self.question_num(
                "What is the power consumption in Long Idle Mode?",
                "Long Idle Mode")
            self.short_idle = self.question_num(
                "What is the power consumption in Short Idle Mode?",
                "Short Idle Mode")
        elif self.product_type == 2:
            self.off = self.question_num(
                "What is the power consumption in Off Mode?",
                "Off Mode")
            self.sleep = self.question_num(
                "What is the power consumption in Sleep Mode?",
                "Sleep Mode")
            self.long_idle = self.question_num(
                "What is the power consumption in Long Idle Mode?",
                "Long Idle Mode")
            self.short_idle = self.question_num(
                "What is the power consumption in Short Idle Mode?",
                "Short Idle Mode")
            self.max_power = self.question_num(
                "What is the maximum power consumption?",
                "Maximum Power")
            self.get_disk_num()
        elif self.product_type == 3:
            self.off = self.question_num(
                "What is the power consumption in Off Mode?",
                "Off Mode")
            self.short_idle = self.question_num(
                "What is the power consumption in Short Idle Mode?",
                "Short Idle Mode")
            if self.get_cpu_core() < 2:
                self.more_discrete = self.question_bool(
                    "Does it have more than one discrete graphics device?",
                    "More Discrete Graphics")
            else:
                self.more_discrete = False
        elif self.product_type == 4:
            self.off = self.question_num(
                "What is the power consumption in Off Mode?", "Off Mode")
            self.sleep = self.question_num(
                """What is the power consumption in Sleep Mode?
(You can input the power consumption in Long Idle Mode, if it lacks a discrete System Sleep Mode)""",
                "Sleep Mode")
            self.long_idle = self.question_num(
                "What is the power consumption in Long Idle Mode?",
                "Long Idle Mode")
            self.short_idle = self.question_num(
                "What is the power consumption in Short Idle Mode?",
                "Short Idle Mode")
            self.media_codec = self.question_bool(
                "Does it support local multimedia encode/decode?",
                "Media Codec")
            self.discrete = self.question_bool(
                "Does it have discrete graphics?",
                "Discrete Graphics")
            self.integrated_display = self.question_bool(
                "Does it have integrated display?",
                "Integrated Display")
            if self.integrated_display:
                self.diagonal = self.get_diagonal()
                self.screen_area = self.get_screen_area()
                (width, height) = self.get_resolution()
                if width * height >= 2300000:
                    self.ep = self.question_bool("""Is it an Enhanced-perforcemance Integrated Display?
  i. Contrast ratio of at least 60:1 measured at a horizontal viewing angle of at least 85° from the
     perpendicular on a flat screen and at least 83° from the perpendicular on a curved screen,
     with or without a screen cover glass;
 ii. A native resolution greater than or equal to 2.3 megapixels (MP); and
iii. Color Gamut greater than or equal to 32.9% of CIE LUV.""", "Enhanced Display")
                else:
                    self.ep = False

        # Ethernet
        if "Gigabit Ethernet" in self.profile:
            self.one_glan = self.profile["Gigabit Ethernet"]
        else:
            self._check_ethernet_num()

        if "10 Gigabit Ethernet" in self.profile:
            self.ten_glan = self.profile["10 Gigabit Ethernet"]
        else:
            self._check_ethernet_num()

        if "Disk Number" in self.profile:
            self.disk_num = self.profile["Disk Number"]
        if "CPU Cores" in self.profile:
            self.cpu_core = self.profile["CPU Cores"]
        if "CPU Clock" in self.profile:
            self.cpu_clock = self.profile["CPU Clock"]
        if "Memory Size" in self.profile:
            self.mem_size = self.profile["Memory Size"]

    def _check_wol(self):
        if "Wake-on-LAN" in self.profile:
            return self.profile["Wake-on-LAN"]
        for dev in os.listdir("/sys/class/net/"):
            if dev.startswith('eth') or dev.startswith('en'):
                wakeup = os.path.join("/sys/class/net/", dev,
                                      "/device/power/wakeup")
                if os.path.exists(wakeup):
                    with open(wakeup, 'r') as f:
                        value = f.read()
                        if 'enabled' in value:
                            self.profile["Wake-on-LAN"] = True
                            return True
        self.profile["Wake-on-LAN"] = False
        return False

    def _check_ethernet_num(self):
        self.one_glan = 0
        self.ten_glan = 0
        for dev in os.listdir("/sys/class/net/"):
            if dev.startswith('eth') or dev.startswith('en'):
                eee_enabled = False
                try:
                    output = subprocess.check_output(
                        "ethtool --show-eee " + dev,
                        shell=True, encoding='utf8')
                except subprocess.CalledProcessError:
                    warning("`ethtool --show-eee " + dev + "` failed. Please check it.")
                    continue
                for line in output.split('\t'):
                    if eee_enabled:
                        if "10000baseT/Full" in line:
                            self.ten_glan = self.ten_glan + 1
                            break
                        elif "1000baseT/Full" in line:
                            self.one_glan = self.one_glan + 1
                            break
                    elif "EEE status: enabled" in line:
                        eee_enabled = True

    def _get_cpu_vendor(self):
        vendor = subprocess.check_output(
            "cat /proc/cpuinfo | grep 'vendor_id' | grep -ioE '(intel|amd)'",
            shell=True, encoding='utf8').strip()
        if re.match("intel", vendor, re.IGNORECASE):
            return 'intel'
        if re.match("amd", vendor, re.IGNORECASE):
            return 'amd'
        return 'unknown'

    def get_cpu_core(self):
        if "CPU Cores" in self.profile:
            self.cpu_core = self.profile["CPU Cores"]
            return self.cpu_core

        try:
            subprocess.check_output('cat /proc/cpuinfo | grep cores',
                                    shell=True, encoding='utf8')
        except subprocess.CalledProcessError:
            warning("Can not check the core number by /proc/cpuinfo. Assume the core number is 1")
            self.cpu_core = 1
        else:
            self.cpu_core = self._int_cmd(
                "grep 'cpu cores' /proc/cpuinfo | sort -u | awk '{print $4}'")

        debug("CPU core: %s" % (self.cpu_core))
        self.profile["CPU Cores"] = self.cpu_core
        return self.cpu_core

    def _float_cmd(self, command):
        return float(subprocess.check_output(
            command, shell=True, encoding='utf8').strip())

    def _int_cmd(self, command):
        return int(subprocess.check_output(
            command, shell=True, encoding='utf8').strip())

    def get_cpu_clock(self):
        if "CPU Clock" in self.profile:
            self.cpu_clock = self.profile["CPU Clock"]
            return self.cpu_clock

        if self._get_cpu_vendor() == 'intel':
            self.cpu_clock = self._float_cmd(
                "grep -oP '[0-9.]+GHz' /proc/cpuinfo | uniq | sed 's/GHz//'")
        else:
            self.cpu_clock = self.question_num("What is CPU frequency (GHz)?",
                                               "CPU Clock")

        debug("CPU clock: %s GHz" % (self.cpu_clock))
        self.profile["CPU Clock"] = self.cpu_clock

        return self.cpu_clock

    def get_mem_size(self):
        if "Memory Size" in self.profile:
            self.mem_size = self.profile["Memory Size"]
            return self.mem_size

        total_online = 0
        block_size = 0
        for online in Path('/sys/devices/system/memory').glob('*/online'):
            with open(online, 'r') as f:
                if f.read().strip() == '1':
                    total_online = total_online + 1

        with open('/sys/devices/system/memory/block_size_bytes') as f:
            block_size = int(f.read().strip(), 16)

        self.mem_size = block_size * total_online / 1024 / 1024 / 1024

        debug("Memory size: %s GB" % (self.mem_size))
        self.profile["Memory Size"] = self.mem_size

        return self.mem_size

    def get_disk_num(self):
        if "Disk Number" in self.profile:
            self.disk_num = self.profile["Disk Number"]
            return self.disk_num

        self.disk_num = len(subprocess.check_output(
            'ls /sys/block | grep -e sd -e nvme -e emmc',
            shell=True, encoding='utf8').strip().split('\n'))

        debug("Disk number: %s" % (self.disk_num))
        self.profile["Disk Number"] = self.disk_num
        return self.disk_num

    def get_1glan_num(self):
        if "Gigabit Ethernet" in self.profile:
            self.one_glan = self.profile["Gigabit Ethernet"]
            return self.one_glan

        self._check_ethernet_num()

        debug("Gigabit Ethernet: %s" % (self.one_glan))
        self.profile["Gigabit Ethernet"] = self.one_glan
        return self.one_glan

    def get_10glan_num(self):
        if "10 Gigabit Ethernet" in self.profile:
            self.ten_glan = self.profile["10 Gigabit Ethernet"]
            return self.ten_glan

        self._check_ethernet_num()

        debug("10 Gigabit Ethernet: %s" % (self.ten_glan))
        self.profile["10 Gigabit Ethernet"] = self.ten_glan
        return self.ten_glan

    def set_display(self, diagonal, ep):
        self.diagonal = diagonal
        self.ep = ep

    def get_display(self):
        return (self.diagonal, self.ep)

    def get_dmi_info(self, info):
        base = '/sys/devices/virtual/dmi/id/'
        if os.path.exists(base + info):
            with open(base + info, "r") as data:
                return data.read().strip().replace(' ', '_')

    def get_bios_version(self):
        if "BIOS version" not in self.profile:
            self.profile["BIOS version"] = self.get_dmi_info('bios_version')
        return self.profile["BIOS version"]

    def get_product_name(self):
        if "Product name" not in self.profile:
            self.profile["Product name"] = self.get_dmi_info('product_name')
        return self.profile["Product name"]

    def get_resolution(self):
        if "Display Width" in self.profile \
                and "Display Height" in self.profile:
            self.width = self.profile["Display Width"]
            self.height = self.profile["Display Height"]
            return (self.width, self.height)

        if self.width is None or self.height is None:
            self.edid_decode()
        self.profile["Display Width"] = self.width
        self.profile["Display Height"] = self.height
        return (self.width, self.height)

    def get_power_consumptions(self):
        return (self.off, self.sleep, self.long_idle, self.short_idle)

    def get_basic_info(self):
        return (self.get_cpu_core(), self.get_cpu_clock(),
                self.get_mem_size(), self.get_disk_num())

    def report(self, filename):
        product_types = ('Desktop, Integrated Desktop, and Notebook Computers',
                         'Workstations', 'Small-scale Servers', 'Thin Clients')
        computer_types = ('Desktop', 'Integrated Desktop', 'Notebook')
        units = {"CPU Clock": "GHz",
                 "Display Diagonal": "inches",
                 "Display Height": "pixels",
                 "Display Width": "pixels",
                 "Memory Size": "GB",
                 "Screen Area": "inches^2"}
        with open(filename, "w") as data:
            data.write('Devicie configuration details:')
            data.write('\n\tProduct Type: ' +
                       product_types[self.profile['Product Type'] - 1])
            data.write('\n\tComputer Type: ' +
                       computer_types[self.profile['Computer Type'] - 1])

            for key in sorted(self.profile.keys()):
                if key not in ('Off Mode', 'Sleep Mode', 'Short Idle Mode',
                               'Long Idle Mode', 'Product Type',
                               'Computer Type', 'BIOS version'):
                    if key in units:
                        if isinstance(self.profile[key], float):
                            data.write('\n\t' + key + ': ' +
                                       str(round(self.profile[key], 2)) +
                                       ' ' + units[key])
                        else:
                            data.write('\n\t' + key + ': ' +
                                       str(self.profile[key]) + ' '
                                       + units[key])
                    else:
                        data.write('\n\t' + key + ': ' +
                                   str(self.profile[key]))
            if os.path.isfile('/var/lib/ubuntu_dist_channel'):
                with open('/var/lib/ubuntu_dist_channel', 'r') as f:
                    for line in f:
                        if line.startswith('#'):
                            continue
                        data.write('\nManifest version: ' + line.strip())

            data.write('\nBIOS version: ' + self.get_bios_version())

            for k in ('Short Idle Mode',
                      'Long Idle Mode',
                      'Sleep Mode',
                      'Off Mode'):
                data.write('\n' + k + ': ' + str(self.profile[k]) + ' W')
            data.write('\n')

    def save(self, filename):
        try:
            with open(filename, "w") as data:
                data.write(json.dumps(self.profile, sort_keys=True, indent=4,
                                      separators=(',', ': ')) + '\n')
        except PermissionError as err:
            if 'SNAP_NAME' in os.environ and os.environ['SNAP_NAME'] == 'energy-tools':
                error('Please execute `snap connect energy-tools:home` to get the permissions.')
            raise err
