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

from distutils.core import setup
import glob, os

setup(name='energy-tools',
      version='1.5.7',
      description='Energy Tools for Energy Star and ErP Lot 3',
      long_description='This program is designed to collect the system profile and calculate\nthe results of Energy Star (5.2 & 6.0 & 7.0) and ErP Lot 3 (Jul. 2014 & Jan. 2016).',
      platforms=['Linux'],
      license='GPLv3',
      author='Shih-Yuan Lee (FourDollars)',
      author_email='sylee@canonical.com',
      url='https://code.launchpad.net/~oem-solutions-engineers/somerville/energy-tools',
      packages=['energy-tools'],
     )
