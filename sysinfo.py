#!/usr/bin/python
"""
Collect essential system information indicated in the Dell ENV0379 document,
and put the result into a JSON file (sysinfo.json).

TODO: Simplify the string process of some outputs.  Because this program
      rewrites a shell script, some string processes are too complicated.
"""
import subprocess
import logging
import os
import re
import json
import argparse

class SystemInformation:
    def __init__(self):
        self.logger = logging.getLogger('SystemInformation')
        self.sysinfo = {}

    def get_misc_info(self):
        self.sysinfo['misc_config_name'] = subprocess.check_output('cat /sys/class/dmi/id/product_name', shell=True).strip()
        self.sysinfo['misc_test_date'] = subprocess.check_output('date +%Y/%m/%d', shell=True).strip()

        self.logger.debug('misc_config_name = %s' % self.sysinfo['misc_config_name'])
        self.logger.debug('misc_test_date = %s' % self.sysinfo['misc_test_date'])

    def get_os_info(self):
        self.sysinfo['os_id'] = subprocess.check_output('lsb_release -i | cut -d: -f2 | tr -d \'\t\'', shell=True).strip()
        self.sysinfo['os_codename'] = subprocess.check_output("lsb_release -c | cut -d: -f2 | tr -d '\t'", shell=True).strip()

        self.logger.debug('os_id = %s' % self.sysinfo['os_id'])
        self.logger.debug('os_codename = %s' % self.sysinfo['os_codename'])

    def get_bios_info(self):
        self.sysinfo['bios_version'] = subprocess.check_output('cat /sys/class/dmi/id/bios_version', shell=True).strip()

        self.logger.debug('bios_version = %s' % self.sysinfo['bios_version'])

    def get_cpu_info(self):
        """
        Requests:
          * CPU model
          * CPU core number
          * CPU clock speed (GHz)

        http://www.richweb.com/cpu_info
        """
        self.sysinfo['cpu_model'] = subprocess.check_output('cat /proc/cpuinfo | grep "model name" | sort -u | cut -d: -f2 | cut -d@ -f1 | xargs', shell=True).strip()
        self.sysinfo['cpu_clock'] = subprocess.check_output('cat /proc/cpuinfo | grep "model name" | sort -u | cut -d: -f2 | cut -d@ -f2 | xargs', shell=True).strip()

        try:
            subprocess.check_output('cat /proc/cpuinfo | grep cores', shell=True)
        except subprocess.CalledProcessError:
            self.sysinfo['cpu_core'] = 1
        else:
            self.sysinfo['cpu_core'] = subprocess.check_output('cat /proc/cpuinfo | grep "cpu cores" | sort -ru | head -n 1 | cut -d: -f2 | xargs', shell=True).strip()

        self.logger.debug('cpu_model = %s' % self.sysinfo['cpu_model'])
        self.logger.debug('cpu_clock = %s' % self.sysinfo['cpu_clock'])
        self.logger.debug('cpu_core = %s' % self.sysinfo['cpu_core'])

        return int(self.sysinfo['cpu_core'])

    def get_memory_info(self):
        # Requests:
        #   * Total Memory (GB)
        #   * DIMM installed
        size = 0
        for i in subprocess.check_output("sudo dmidecode -t 17 | grep Size | awk '{print $2}'", shell=True).split('\n'):
            if i:
                size = size + int(i)
        self.sysinfo['memory_size'] = size / 1024

        self.logger.debug('memory_size = %s' % self.sysinfo['memory_size'])

        return size

    def get_graphic_info(self):
        """
        Requests:
          * Switchable Mode or Discrete Mode
          * UMA or DIS
          * ? GPU model name
        """
        # there might be multi-graphics, so we collect all the graphics
        # information into a temp file, and handle each of them seperately.
        tmpfile = subprocess.check_output('tempfile', shell=True).strip()
        subprocess.call('lspci -nn | grep VGA > %s' % tmpfile, shell=True)

        self.sysinfo['graphic_num'] = 0
        with open(tmpfile) as f:
            vgainfo = f.readline()
            key = 'graphic_%d' % self.sysinfo['graphic_num']
            self.sysinfo[key] = ' '.join(vgainfo.split(' ')[2:]).strip()
            self.sysinfo['graphic_num'] += 1

        os.remove(tmpfile)

        self.logger.debug('graphic_num = %s' % self.sysinfo['graphic_num'])
        for i in xrange(self.sysinfo['graphic_num']):
            key = 'graphic_%d' % i
            self.logger.debug('%s = %s' % (key, self.sysinfo[key]))

    def get_disk_info(self):
        """
        Requests:
          * HDD Quan. (#)
          * HDD Model name
          * HDD Speed
          * HDD Capacity
        https://wiki.ubuntu.com/UnitsPolicy
        http://www.cyberciti.biz/faq/find-hard-disk-hardware-specs-on-linux/
        """
        self.sysinfo['disk_list'] = subprocess.check_output('ls /sys/block | grep sd', shell=True).strip().split('\n')

        for i in xrange(len(self.sysinfo['disk_list'])):
            self.sysinfo['disk%s_model' % i] = subprocess.check_output('cat /sys/block/%s/device/model' % self.sysinfo['disk_list'][i], shell=True).strip()
            # Here we use SI standard (base-10 units) rather than IEC standard (base-2 units).
            # For more details, please refer to Ubuntu UnitsPolicy.
            self.sysinfo['disk%s_size' % i] = '%dGB' % (int(subprocess.check_output('cat /sys/block/%s/size' % self.sysinfo['disk_list'][i], shell=True)) * 512 / 1000000000)

        for i in xrange(len(self.sysinfo['disk_list'])):
            self.logger.debug('%s = %s' % ('disk%s_model' % i, self.sysinfo['disk%s_model' % i]))
            self.logger.debug('%s = %s' % ('disk%s_size' % i, self.sysinfo['disk%s_size' % i]))

        return len(self.sysinfo['disk_list'])

    def get_screen_info(self):
        # Requests:
        #   * Panel vendor & model name
        #   * Panel Size (in.)
        self.sysinfo['screen_resolution'] = subprocess.check_output("xrandr -d :0.0 2> /dev/null | grep '*' | xargs | cut -d' ' -f1", shell=True).strip()

        self.logger.debug('screen_resolution = %s' % self.sysinfo['screen_resolution'])

    def output_json(self, output_file):
        self.pcjson = json.dumps(self.sysinfo)
        with open('{}.json'.format(output_file), 'w') as f:
            f.write(self.pcjson)
        f.close()

class MyArgumentParser(object):
    """
    Command-line argument parser
    """
    def __init__(self):
        """
        Create parser object
        """
        description = 'Collect essential system information'
        parser = argparse.ArgumentParser(description=description)
        parser.add_argument('-o', '--output-file', default='sysinfo',
                            help=('file name used for output and log files'))
        self.parser = parser

    def parse(self):
        """
        Parse command-line arguments
        """
        args, extra_args = self.parser.parse_known_args()
        return args, extra_args

def main():
    args, extra_args = MyArgumentParser().parse()

    logging.basicConfig(filename='{}.log'.format(args.output_file), filemode='w', level=logging.DEBUG)

    si = SystemInformation()
    si.get_misc_info()
    si.get_os_info()
    si.get_bios_info()
    si.get_cpu_info()
    si.get_memory_info()
    si.get_graphic_info()
    si.get_disk_info()
    si.get_screen_info()
    si.output_json(args.output_file)

if __name__ == '__main__':
    main()

# Default requirements from ODM
#
# Basic information:
#   Project Name
#   Data phase
#   Testing date
#   OS version
#   BIOS version
#   Switchable Mode or Discrete Mode
#   UMA or DIS
#
# System configuration:
#   CPU model
#   CPU core number
#   CPU clock speed (GHz)
#   Total Memory (GB)
#   # DIMM installed
#   HDD Quan. (#)
#   HDD Model name
#   HDD Speed
#   HDD Capacity
#   Panel vendor & model name
#   Panel Size (in.)
#   Screen resolution
#   GPU model name
#   Adapter Mfg
#   Adapter model name
#   Adapter size watt
#
# Test data:
#   Short Idle max brightness (Watt)
#   Short Idle 90 nits (Watt)
#   Long Idle (Watt)
#   Power of S3 (watt)
#   Power of S5 (watt)
