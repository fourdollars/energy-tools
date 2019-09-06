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


def result_filter(result, value, maximum):
    if maximum >= value:
        delta = (maximum - value) * 100 / maximum
        if delta < 5.0:
            return "marginally %s (%s%% to fail)" % (result, round(delta, 2))
        else:
            return result
    else:
        delta = (value - maximum) * 100 / maximum
        return "%s (%s%% to pass)" % (result, round(delta, 2))
