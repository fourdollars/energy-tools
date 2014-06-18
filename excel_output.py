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

__all__ = [
    "generate_excel",
    "generate_excel_for_computers",
    "generate_excel_for_workstations",
    "generate_excel_for_small_scale_servers",
    "generate_excel_for_thin_clients"]

from logging import debug, warning

def generate_excel(sysinfo, version, output):
    if not output:
        return

    excel = ExcelMaker(version, output)

    if sysinfo.product_type == 1:
        generate_excel_for_computers(excel, sysinfo)
    elif sysinfo.product_type == 2:
        generate_excel_for_workstations(excel, sysinfo)
    elif sysinfo.product_type == 3:
        generate_excel_for_small_scale_servers(excel, sysinfo)
    elif sysinfo.product_type == 4:
        generate_excel_for_thin_clients(excel, sysinfo)

class ExcelMaker:
    def __init__(self, version, output):
        try:
            from xlsxwriter import Workbook
        except:
            warning("You need to install Python xlsxwriter module or you can not output Excel format file.")
            return
        self.book = Workbook(output)
        self.book.set_properties({'comments':"Created by Energy Tools %s from Canonical Ltd." % (version)})
        self.sheet = self.book.add_worksheet()
        self.adjust_column_width()
        self.setup_theme()
        self.row = 1
        self.column = 'A'
        self.pos = {}

    def save(self):
        self.book.close()

    def adjust_column_width(self):
        sheet = self.sheet
        sheet.set_column('A:A', 38)
        sheet.set_column('B:B', 37)
        sheet.set_column('C:C', 1)
        sheet.set_column('D:D', 13)
        sheet.set_column('E:E', 6)
        sheet.set_column('F:F', 15)
        sheet.set_column('G:G', 6)
        sheet.set_column('H:H', 6)
        sheet.set_column('I:I', 6)
        sheet.set_column('J:J', 6)

    def setup_theme(self):
        book = self.book
        theme = {}
        theme["header"] = book.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'fg_color': '#CFE2F3'})
        theme["left"] = book.add_format({'align': 'left'})
        theme["right"] = book.add_format({'align': 'right'})
        theme["center"] = book.add_format({'align': 'center'})
        theme["field"] = book.add_format({
            'border': 1,
            'fg_color': '#F3F3F3'})
        theme["field1"] = book.add_format({
            'left': 1,
            'right': 1,
            'fg_color': '#F3F3F3'})
        theme["fieldC"] = book.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#F3F3F3'})
        theme["field2"] = book.add_format({
            'left': 1,
            'right': 1,
            'bottom': 1,
            'fg_color': '#F3F3F3'})
        theme["unsure"] = book.add_format({'border': 1, 'fg_color': '#D9EAD3'})
        theme["value"] = book.add_format({'border': 1})
        theme["value1"] = book.add_format({
            'left': 1,
            'right': 1})
        theme["value1"].set_num_format('0.00')
        theme["value2"] = book.add_format({
            'left': 1,
            'right': 1,
            'bottom': 1})
        theme["value2"].set_num_format('0.00')
        theme["value3"] = book.add_format({
            'left': 1,
            'right': 1})
        theme["value3"].set_num_format('0%')
        theme["value4"] = book.add_format({
            'left': 1,
            'right': 1,
            'bottom': 1})
        theme["value4"].set_num_format('0%')
        theme["float1"] = book.add_format({'border': 1})
        theme["float1"].set_num_format('0.0')
        theme["float2"] = book.add_format({'border': 1})
        theme["float2"].set_num_format('0.00')
        theme["float3"] = book.add_format({
            'left': 1,
            'right': 1})
        theme["float3"].set_num_format('0.000')
        theme["result"] = book.add_format({
            'border': 1,
            'fg_color': '#F4CCCC'})
        theme["result_value"] = book.add_format({
            'border': 1,
            'fg_color': '#FFF2CC'})
        theme["result_value"].set_num_format('0.00')
        self.theme = theme

    def ncell(self, width, height, field, formula=None, value=None, validator=None, twin=False):
        if type(value) is list:
            validator = value
            value = None
        if value is None:
            if formula is None:
                value = field
            else:
                value = formula
                formula = None
        debug("Field: %s, Formula: %s, Value: %s, Validator: %s" % (field, formula, value, validator))

        if field in palette:
            (i, j) = palette[field]
            theme1 = self.theme[i]
            if j:
                theme2 = self.theme[j]
            else:
                theme2 = self.theme['value']
        else:
            theme1 = self.theme['field']
            theme2 = self.theme['value']
        debug("Theme 1: %s, Theme 2: %s" % (theme1, theme2))

        end_column = chr(ord(self.column) + width - 1)
        next_column = chr(ord(self.column) + 1)
        end_row = self.row + height - 1
        start_cell = "%s%s" % (self.column, self.row)
        next_cell = "%s%s" % (next_column, self.row)
        end_cell = "%s%s" % (end_column, end_row)
        if start_cell != end_cell:
            debug("Position: %s%s" % (start_cell, end_cell))
        else:
            debug("Position: %s" % start_cell)

        if width > 1 or height > 1:
            self.sheet.merge_range("%s:%s" % (start_cell, end_cell), value, theme1)
            if formula:
                self.sheet.write("%s" % start_cell, formula, theme1, value)
        else:
            if twin:
                self.sheet.write("%s" % start_cell, field, theme1)
                self.sheet.write("%s" % next_cell, value, theme2)
                if validator:
                    self.sheet.data_validation("%s" % next_cell, {
                        'validate': 'list',
                        'source': validator})
            else:
                if formula:
                    self.sheet.write("%s" % start_cell, formula, theme1, value)
                else:
                    self.sheet.write("%s" % start_cell, value, theme1)
        if twin:
            self.pos[field] = next_cell
        else:
            self.pos[field] = start_cell
        self.row = self.row + height

    def cell(self, field, formula=None, value=None, validator=None):
        self.ncell(1, 1, field, formula, value, validator)

    def tcell(self, field, formula=None, value=None, validator=None):
        self.ncell(1, 1, field, formula, value, validator, True)

    def separator(self):
        self.row = self.row + 1

    def jump(self, column, row):
        self.column = column
        self.row = row

palette = {
        # Header
        'General': ('header', None),
        'Graphics': ('header', None),
        'Power Consumption': ('header', None),
        'Display': ('header', None),
        'Energy Star 5.2': ('header', None),
        'Energy Star 6.0': ('header', None),

        # Inner header
        'Category': ('fieldC', None),
        'A': ('fieldC', None),
        'B': ('fieldC', None),
        'C': ('fieldC', None),
        'D': ('fieldC', None),

        # Cell without ceil and floor by percentage
        'T_OFF': ('field1', 'value3'),
        'T_SLEEP': ('field1', 'value3'),
        'T_LONG_IDLE': ('field1', 'value3'),

        # Cell with different background color by percentage
        'T_IDLE': ('field2', 'value4'),
        'T_SHORT_IDLE': ('field2', 'value4'),

        # Cell without ceil and floor and float 1
        'P_OFF': ('field1', 'value1'),
        'P_SLEEP': ('field1', 'value1'),
        'P_LONG_IDLE': ('field1', 'value1'),
        'TEC_BASE': ('value1', 'value1'),
        'TEC_MEMORY': ('value1', 'value1'),
        'TEC_GRAPHICS': ('value1', 'value1'),
        'TEC_STORAGE': ('value2', 'value2'),

        # Cell with different background color and float 1
        'P_IDLE': ('field2', 'value2'),
        'P_SHORT_IDLE': ('field2', 'value2'),

        # Important result
        'E_TEC': ('result', 'result_value'),
        'E_TEC_MAX': ('result_value', None),

        # Prompt
        'None': ('center', None),
        'field1': ('field1', None),
        'result': ('result', None),

        # Unsure
        "IEEE 802.3az compliant Gigabit Ethernet": ('field', 'unsure'),
        "GPU Frame Buffer Width": ('field', 'unsure'),
        "Graphics Category": ('field', 'unsure'),
        "Enhanced-performance Integrated Display": ('field', 'unsure'),

        # Float 1
        "Physical Diagonal (inch)": ('field', 'float1'),

        # Float 2
        "Off mode (W)": ('field', 'float2'),
        "Off mode (W) with WOL": ('field', 'float2'),
        "Sleep mode (W)": ('field', 'float2'),
        "Sleep mode (W) with WOL": ('field', 'float2'),
        "Long idle mode (W)": ('field', 'float2'),
        "Short idle mode (W)": ('field', 'float2'),

        # Float 3
        'CPU clock (GHz)': ('field', 'float3'),
        }

def generate_excel_for_computers(excel, sysinfo):

    excel.ncell(2, 1, "General")
    excel.tcell("Product Type", "Desktop, Integrated Desktop, and Notebook")

    if sysinfo.computer_type == 1:
        msg = "Desktop"
    elif sysinfo.computer_type == 2:
        msg = "Integrated Desktop"
    else:
        msg = "Notebook"

    excel.tcell("Computer Type", msg)
    excel.tcell("CPU cores", sysinfo.cpu_core)

    excel.tcell("CPU clock (GHz)", sysinfo.cpu_clock)
    excel.tcell("Memory size (GB)", sysinfo.mem_size)
    excel.tcell("Number of Hard Drives", sysinfo.disk_num)
    excel.tcell("Number of Discrete Graphics Cards", sysinfo.discrete_gpu_num)
    if sysinfo.tvtuner:
        msg = "Yes"
    else:
        msg = "No"
    excel.tcell("Discrete television tuner", msg, ["Yes", "No"])
    if sysinfo.audio:
        msg = "Yes"
    else:
        msg = "No"
    excel.tcell("Discrete audio card", msg, ["Yes", "No"])
    excel.tcell("IEEE 802.3az compliant Gigabit Ethernet", sysinfo.eee)

    excel.separator()

    excel.ncell(2, 1, "Graphics")
    if sysinfo.switchable:
        msg = "Switchable"
    elif sysinfo.discrete:
        msg = "Discrete"
    else:
        msg = "Integrated"
    excel.tcell("Graphics Type", msg, ['Integrated', 'Switchable', 'Discrete'])
    if sysinfo.computer_type == 3:
        msg = "<= 64-bit"
        validator = ['<= 64-bit', '> 64-bit and <= 128-bit', '> 128-bit']
    else:
        msg = "<= 128-bit"
        validator = ['<= 128-bit', '> 128-bit']
    excel.tcell("GPU Frame Buffer Width", msg, validator)
    excel.tcell("Graphics Category", "G1 (FB_BW <= 16)", [
        'G1 (FB_BW <= 16)',
        'G2 (16 < FB_BW <= 32)',
        'G3 (32 < FB_BW <= 64)',
        'G4 (64 < FB_BW <= 96)',
        'G5 (96 < FB_BW <= 128)',
        'G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)',
        'G7 (FB_BW > 128; Frame Buffer Data Width >= 192 bits)'])

    excel.separator()

    excel.ncell(2, 1, "Power Consumption")
    excel.tcell("Off mode (W)", sysinfo.off)
    excel.tcell("Off mode (W) with WOL", sysinfo.off_wol)
    excel.tcell("Sleep mode (W)", sysinfo.sleep)
    excel.tcell("Sleep mode (W) with WOL", sysinfo.sleep_wol)
    excel.tcell("Long idle mode (W)", sysinfo.long_idle)
    excel.tcell("Short idle mode (W)", sysinfo.short_idle)

    excel.separator()

    if sysinfo.computer_type != 1:
        excel.ncell(2, 1, "Display")
        if sysinfo.ep:
            msg = "Yes"
        else:
            msg = "No"
        excel.tcell("Enhanced-performance Integrated Display", msg, ["Yes", "No"])
        excel.tcell("Physical Diagonal (inch)", sysinfo.diagonal)
        excel.tcell("Screen Width (px)", sysinfo.width)
        excel.tcell("Screen Height (px)", sysinfo.height)

    excel.jump('D', 1)
    if sysinfo.computer_type == 3:
        width = 6
    else:
        width = 7
    excel.ncell(width, 1, "Energy Star 5.2")

    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

    excel.tcell("T_OFF", '=IF(EXACT(%s,"Notebook"),0.6,0.55' % excel.pos["Computer Type"], T_OFF)
    excel.tcell("T_SLEEP", '=IF(EXACT(%s,"Notebook"),0.1,0.05' % excel.pos["Computer Type"], T_SLEEP)
    excel.tcell("T_IDLE", '=IF(EXACT(%s,"Notebook"),0.3,0.4' % excel.pos["Computer Type"], T_IDLE)

    excel.tcell("P_OFF", '=%s' % excel.pos["Off mode (W)"], sysinfo.off)
    excel.tcell("P_SLEEP", '=%s' % excel.pos["Sleep mode (W)"], sysinfo.sleep)
    excel.tcell("P_IDLE", '=%s' % excel.pos["Short idle mode (W)"], sysinfo.short_idle)

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_IDLE * sysinfo.short_idle) * 8760 / 1000

    excel.tcell("E_TEC", "=(%s*%s+%s*%s+%s*%s)*8760/1000" % (
        excel.pos["T_OFF"], excel.pos["P_OFF"],
        excel.pos["T_SLEEP"], excel.pos["P_SLEEP"],
        excel.pos["T_IDLE"], excel.pos["P_IDLE"]), E_TEC)

    excel.jump('F', 2)
    excel.ncell(1, 2, "Category")
    excel.jump('G', 2)
    excel.ncell(1, 2, "A")
    excel.jump('H', 2)
    excel.ncell(1, 2, "B")
    excel.jump('I', 2)
    excel.ncell(1, 2, "C")

    if sysinfo.computer_type != 3:
        excel.jump('J', 2)
        excel.ncell(1, 2, "D")

    if sysinfo.computer_type == 3:
        excel.jump('H', 2)
        if sysinfo.discrete:
            msg = "B"
        else:
            msg = " "
        excel.ncell(1, 2, 'B', '=IF(EXACT(%s,"Discrete"), "B", "")' % excel.pos["Graphics Type"], msg)

        excel.jump('I', 2)
        excel.ncell(1, 2, 'C', '=IF(AND(EXACT(%s,"Discrete"), EXACT(%s, "> 128-bit"), %s>=2, %s>=2), "C", "")' % (
            excel.pos["Graphics Type"],
            excel.pos["GPU Frame Buffer Width"],
            excel.pos["CPU cores"],
            excel.pos["Memory size (GB)"]), " ")
    else:
        excel.jump('H', 2)
        if sysinfo.cpu_core == 2 and sysinfo.mem_size >=2:
            msg = "B"
        else:
            msg = ""
        excel.ncell(1, 2, 'B', '=IF(AND(%s=2,%s>=2), "B", "")' % (
            excel.pos["CPU cores"],
            excel.pos["Memory size (GB)"]), msg)

        excel.jump('I', 2)
        if sysinfo.cpu_core > 2 and (sysinfo.mem_size >= 2 or sysinfo.discrete):
            msg = "C"
        else:
            msg = " "
        excel.ncell(1, 2, 'C', '=IF(AND(%s>2,OR(%s>=2,EXACT(%s,"Discrete"))), "C", "")' % (
            excel.pos["CPU cores"],
            excel.pos["Memory size (GB)"],
            excel.pos["Graphics Type"]), msg)

        excel.jump('J', 2)
        if sysinfo.cpu_core >= 4 and sysinfo.mem_size >= 4:
            msg = "D"
        else:
            msg = " "
        excel.ncell(1, 2, 'D', '=IF(AND(%s>=4,OR(%s>=4,AND(EXACT(%s,"Discrete"),EXACT(%s,"> 128-bit")))), "D", "")' % (
            excel.pos["CPU cores"],
            excel.pos["Memory size (GB)"],
            excel.pos["Graphics Type"],
            excel.pos["GPU Frame Buffer Width"]), msg)

    excel.jump('F', 4)
    excel.cell("field1", "TEC_BASE")
    excel.cell("field1", "TEC_MEMORY")
    excel.cell("field1", "TEC_GRAPHICS")
    excel.cell("field1", "TEC_STORAGE")
    excel.cell("result", "E_TEC_MAX")

    # Category A
    excel.jump('G', 4)
    if sysinfo.computer_type == 3:
        # Notebook
        TEC_BASE = 40

        if sysinfo.mem_size > 4:
            TEC_MEMORY = 0.4 * (sysinfo.mem_size - 4)
        else:
            TEC_MEMORY = 0

        TEC_GRAPHICS = 0

        if sysinfo.disk_num > 1:
            TEC_STORAGE = 3.0 * (sysinfo.disk_num - 1)
        else:
            TEC_STORAGE = 0

        excel.cell("TEC_BASE", TEC_BASE)
        excel.cell("TEC_MEMORY", "=IF(%s>4, 0.4*(%s-4), 0)" % (excel.pos["Memory size (GB)"], excel.pos["Memory size (GB)"]), TEC_MEMORY)
        excel.cell("TEC_GRAPHICS", TEC_GRAPHICS)
        excel.cell("TEC_STORAGE", "=IF(%s>1, 3*(%s-1), 0)" % (excel.pos["Number of Hard Drives"], excel.pos["Number of Hard Drives"]), TEC_STORAGE)
    else:
        # Desktop
        TEC_BASE = 148

        if sysinfo.mem_size > 2:
            TEC_MEMORY = 1.0 * (sysinfo.mem_size - 2)
        else:
            TEC_MEMORY = 0

        TEC_GRAPHICS = 35

        if sysinfo.disk_num > 1:
            TEC_STORAGE = 25 * (sysinfo.disk_num - 1)
        else:
            TEC_STORAGE = 0
 
        excel.cell("TEC_BASE", TEC_BASE)
        excel.cell("TEC_MEMORY", "=IF(%s>2, 1.0*(%s-2), 0)" % (excel.pos["Memory size (GB)"], excel.pos["Memory size (GB)"]), TEC_MEMORY)
        excel.cell("TEC_GRAPHICS", '=IF(EXACT(%s,"> 128-bit"), 50, 35)' % excel.pos["GPU Frame Buffer Width"], TEC_GRAPHICS)
        excel.cell("TEC_STORAGE", "=IF(%s>1, 25*(%s-1), 0)" % (excel.pos["Number of Hard Drives"], excel.pos["Number of Hard Drives"]), TEC_STORAGE)

    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
    excel.cell("E_TEC_MAX", "=%s+%s+%s+%s" % (
        excel.pos["TEC_BASE"],
        excel.pos["TEC_MEMORY"],
        excel.pos["TEC_GRAPHICS"],
        excel.pos["TEC_STORAGE"]), E_TEC_MAX)
    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    excel.cell('None', '=IF(%s<=%s, "PASS", "FAIL")' % (
        excel.pos["E_TEC"],
        excel.pos["E_TEC_MAX"]), RESULT)

    # Category B
    excel.jump('H', 4)
    if sysinfo.computer_type == 3:
        # Notebook
        if sysinfo.discrete:
            TEC_BASE = 53

            if sysinfo.mem_size > 4:
                TEC_MEMORY = 0.4 * (sysinfo.mem_size - 4)
            else:
                TEC_MEMORY = 0

            TEC_GRAPHICS = 0

            if sysinfo.disk_num > 1:
                TEC_STORAGE = 3.0 * (sysinfo.disk_num - 1)
            else:
                TEC_STORAGE = 0

            E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE

            if E_TEC <= E_TEC_MAX:
                RESULT = "PASS"
            else:
                RESULT = "FAIL"
            print('Hello Kitty')
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%s, "B"), 53, "")' % excel.pos["B"], TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%s, "B"), %s, "")' % (excel.pos["B"], excel.pos["TEC_MEMORY"]), TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%s, "B"), IF(EXACT(%s, "<= 64-bit"), 0, 3), "")' % (excel.pos["B"], excel.pos["GPU Frame Buffer Width"]), TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%s, "B"), %s, "")' % (excel.pos["B"], excel.pos["TEC_STORAGE"]), TEC_STORAGE)
    else:
        # Desktop
        if sysinfo.cpu_core == 2 and sysinfo.mem_size >= 2:
            TEC_BASE = 175

            if sysinfo.mem_size > 2:
                TEC_MEMORY = 1.0 * (sysinfo.mem_size - 2)
            else:
                TEC_MEMORY = 0

            TEC_GRAPHICS = 35

            if sysinfo.disk_num > 1:
                TEC_STORAGE = 25 * (sysinfo.disk_num - 1)
            else:
                TEC_STORAGE = 0

            E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE

            if E_TEC <= E_TEC_MAX:
                RESULT = "PASS"
            else:
                RESULT = "FAIL"
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%s, "B"), 175, "")' % excel.pos["B"], TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%s, "B"), %s, "")' % (excel.pos["B"], excel.pos["TEC_MEMORY"]), TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%s, "B"), %s, "")' % (excel.pos["B"], excel.pos["TEC_GRAPHICS"]), TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%s, "B"), %s, "")' % (excel.pos["B"], excel.pos["TEC_STORAGE"]), TEC_STORAGE)

    excel.cell('E_TEC_MAX', '=IF(EXACT(%s, "B"), %s+%s+%s+%s, "")' % (
        excel.pos["B"],
        excel.pos["TEC_BASE"],
        excel.pos["TEC_MEMORY"],
        excel.pos["TEC_GRAPHICS"],
        excel.pos["TEC_STORAGE"]), E_TEC_MAX)
    excel.cell('None', '=IF(EXACT(%s, "B"), IF(%s<=%s, "PASS", "FAIL"), "")' % (
        excel.pos["B"],
        excel.pos["E_TEC"],
        excel.pos["E_TEC_MAX"]), RESULT)

    # Category C
    excel.jump('I', 4)
    if sysinfo.computer_type == 3:
        # Notebook
        TEC_BASE = ""
        TEC_MEMORY = ""
        TEC_GRAPHICS = ""
        TEC_STORAGE = ""
        E_TEC_MAX = ""
        RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%s, "C"), 88.5, "")' % excel.pos["C"], TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%s, "C"), %s, "")' % (excel.pos["C"], excel.pos["TEC_MEMORY"]), TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%s, "C"), 0, "")' % excel.pos["C"], TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%s, "C"), %s, "")' % (excel.pos["C"], excel.pos["TEC_STORAGE"]), TEC_STORAGE)
    else:
        # Desktop
        if sysinfo.cpu_core > 2 and (sysinfo.mem_size >= 2 or sysinfo.discrete):
            TEC_BASE = 209

            if sysinfo.mem_size > 2:
                TEC_MEMORY = 1.0 * (sysinfo.mem_size - 2)
            else:
                TEC_MEMORY = 0

            TEC_GRAPHICS = 0 

            if sysinfo.disk_num > 1:
                TEC_STORAGE = 25 * (sysinfo.disk_num - 1)
            else:
                TEC_STORAGE = 0

            E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE

            if E_TEC <= E_TEC_MAX:
                RESULT = "PASS"
            else:
                RESULT = "FAIL"
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%s, "C"), 209, "")' % excel.pos["C"], TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%s, "C"), G5, "")' % excel.pos["C"], TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%s, "C"), IF(EXACT(%s, "> 128-bit"), 50, 0), "")' % (excel.pos["C"], excel.pos["GPU Frame Buffer Width"]), TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%s, "C"), G7, "")' % excel.pos["C"], TEC_STORAGE)

    excel.cell('E_TEC_MAX', '=IF(EXACT(%s, "C"), %s+%s+%s+%s, "")' % (
        excel.pos["C"],
        excel.pos["TEC_BASE"],
        excel.pos["TEC_MEMORY"],
        excel.pos["TEC_GRAPHICS"],
        excel.pos["TEC_STORAGE"]), E_TEC_MAX)
    excel.cell('None', '=IF(EXACT(%s, "C"), IF(%s<=%s, "PASS", "FAIL"), "")' % (
        excel.pos["C"],
        excel.pos["E_TEC"],
        excel.pos["E_TEC_MAX"]), RESULT)

    excel.jump('J', 4)
    # Category D
    if sysinfo.computer_type != 3:
        # Desktop
        if sysinfo.cpu_core >= 4 and sysinfo.mem_size >= 4:
            TEC_BASE = 234

            if sysinfo.mem_size > 4:
                TEC_MEMORY = 1.0 * (sysinfo.mem_size - 4)
            else:
                TEC_MEMORY = 0

            TEC_GRAPHICS = 0 

            if sysinfo.disk_num > 1:
                TEC_STORAGE = 25 * (sysinfo.disk_num - 1)
            else:
                TEC_STORAGE = 0

            E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE

            if E_TEC <= E_TEC_MAX:
                RESULT = "PASS"
            else:
                RESULT = "FAIL"
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%s, "D"), 234, "")' % excel.pos["D"], TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%s, "D"), IF(%s>4, %s-4, 0), "")' % (
            excel.pos["D"],
            excel.pos["Memory size (GB)"],
            excel.pos["Memory size (GB)"]), TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%s, "D"), %s, "")' % (excel.pos["D"], excel.pos["TEC_GRAPHICS"]), TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%s, "D"), %s, "")' % (excel.pos["D"], excel.pos["TEC_STORAGE"]), TEC_STORAGE)
        excel.cell('E_TEC_MAX', '=IF(EXACT(%s, "D"), %s+%s+%s+%s, "")' % (
            excel.pos["D"],
            excel.pos["TEC_BASE"],
            excel.pos["TEC_MEMORY"],
            excel.pos["TEC_GRAPHICS"],
            excel.pos["TEC_STORAGE"]), E_TEC_MAX)
        excel.cell('None', '=IF(EXACT(%s, "D"), IF(%s<=%s, "PASS", "FAIL"), "")' % (
            excel.pos["D"],
            excel.pos["E_TEC"],
            excel.pos["E_TEC_MAX"]), RESULT)

    excel.jump('D', 10)
    excel.ncell(4, 1, "Energy Star 6.0")

    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.25, 0.35, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

    excel.tcell("T_OFF", '=IF(EXACT(%s,"Notebook"),0.25,0.45' % excel.pos["Computer Type"], T_OFF)
    excel.tcell("T_SLEEP", '=IF(EXACT(%s,"Notebook"),0.35,0.05' % excel.pos["Computer Type"], T_SLEEP)
    excel.tcell("T_LONG_IDLE", '=IF(EXACT(%s,"Notebook"),0.1,0.15' % excel.pos["Computer Type"], T_LONG_IDLE)
    excel.tcell("T_SHORT_IDLE", '=IF(EXACT(%s,"Notebook"),0.3,0.35' % excel.pos["Computer Type"], T_SHORT_IDLE)

    excel.tcell("P_OFF", '=%s' % excel.pos["Off mode (W)"], sysinfo.off)
    excel.tcell("P_SLEEP", '=%s' % excel.pos["Sleep mode (W)"], sysinfo.sleep)
    excel.tcell("P_LONG_IDLE", '=%s' % excel.pos["Long idle mode (W)"], sysinfo.long_idle)
    excel.tcell("P_SHORT_IDLE", '=%s' % excel.pos["Short idle mode (W)"], sysinfo.short_idle)

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_LONG_IDLE * sysinfo.long_idle + T_SHORT_IDLE * sysinfo.short_idle) * 8760 / 1000

    excel.tcell("E_TEC", "=(%s*%s+%s*%s+%s*%s+%s*%s)*8760/1000" % (
        excel.pos["T_OFF"], excel.pos["P_OFF"],
        excel.pos["T_SLEEP"], excel.pos["P_SLEEP"],
        excel.pos["T_LONG_IDLE"], excel.pos["P_LONG_IDLE"],
        excel.pos["T_SHORT_IDLE"], excel.pos["P_SHORT_IDLE"]), E_TEC)

    # XXX TODO
    excel.save()
    return

    excel.jump('D', 21)
    excel.ncell(4, 1, "Power Supply Efficiency Allowance requirements:", "None", ['None', 'Lower', 'Higher'])

    excel.jump('F', 11)
    excel.float3("ALLOWANCE_PSU", '=IF(OR(EXACT(%s, "Notebook"), EXACT(%s, "Desktop")), IF(EXACT(%s, "Higher"), 0.03, IF(EXACT(%s, "Lower"), 0.015, 0)), IF(EXACT(%s, "Higher"), 0.04, IF(EXACT(%s, "Lower"), 0.015, 0)))' % (
        excel.pos["Computer Type"],
        excel.pos["Computer Type"],
        excel.pos["Power Supply Efficiency Allowance requirements:"],
        excel.pos["Power Supply Efficiency Allowance requirements:"],
        excel.pos["Power Supply Efficiency Allowance requirements:"],
        excel.pos["Power Supply Efficiency Allowance requirements:"]), 0)

    P = sysinfo.cpu_core * sysinfo.cpu_clock
    if sysinfo.computer_type == 3:
        if sysinfo.discrete:
            if P > 9:
                TEC_BASE = 18
            elif P > 2:
                TEC_BASE = 16
            else:
                TEC_BASE = 14
        else:
            if P > 8:
                TEC_BASE = 28
            elif P > 5.2:
                TEC_BASE = 24
            elif P > 2:
                TEC_BASE = 22
            else:
                TEC_BASE = 14
    else:
        if sysinfo.discrete:
            if P > 9:
                TEC_BASE = 135
            elif P > 3:
                TEC_BASE = 115 
            else:
                TEC_BASE = 69
        else:
            if P > 7:
                TEC_BASE = 135
            elif P > 6:
                TEC_BASE = 120
            elif P > 3:
                TEC_BASE = 112
            else:
                TEC_BASE = 69
    sheet.write("F12", "TEC_BASE", field1)
    if sysinfo.computer_type == 3:
        sheet.write("G12", '=IF(EXACT(B11,"Discrete"), IF(I11>9, 18, IF(AND(I11<=9, I11>2), 16, 14)), IF(I11>8, 28, IF(AND(I11<=8, I11>5.2), 24, IF(AND(I11<=5.2, I11>2), 22, 14))))', value1, TEC_BASE)
    else:
        sheet.write("G12", '=IF(EXACT(B11,"Discrete"), IF(I11>9, 135, IF(AND(I11<=9, I11>3), 115, 69)), IF(I11>7, 135, IF(AND(I11<=7, I11>6), 120, IF(AND(I11<=6, I11>3), 112, 69))))', value1, TEC_BASE)

    TEC_MEMORY = 0.8 * sysinfo.mem_size
    sheet.write("F13", "TEC_MEMORY", field1)
    sheet.write("G13", '=B6*0.8', value1, TEC_MEMORY)

    if sysinfo.discrete:
        if sysinfo.computer_type == 3:
            TEC_GRAPHICS = 14
        else:
            TEC_GRAPHICS = 36
    else:
        TEC_GRAPHICS = 0
    sheet.write("F14", "TEC_GRAPHICS", field1)
    if sysinfo.computer_type == 3:
        sheet.write("G14", '=IF(EXACT(B11, "Discrete"), IF(EXACT(B13, "G1 (FB_BW <= 16)"), 14, IF(EXACT(B13, "G2 (16 < FB_BW <= 32)"), 20, IF(EXACT(B13, "G3 (32 < FB_BW <= 64)"), 26, IF(EXACT(B13, "G4 (64 < FB_BW <= 96)"), 32, IF(EXACT(B13, "G5 (96 < FB_BW <= 128)"), 42, IF(EXACT(B13, "G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)"), 48, 60)))))), 0)', value1, TEC_GRAPHICS)
    else:
        sheet.write("G14", '=IF(EXACT(B11, "Discrete"), IF(EXACT(B13, "G1 (FB_BW <= 16)"), 36, IF(EXACT(B13, "G2 (16 < FB_BW <= 32)"), 51, IF(EXACT(B13, "G3 (32 < FB_BW <= 64)"), 64, IF(EXACT(B13, "G4 (64 < FB_BW <= 96)"), 83, IF(EXACT(B13, "G5 (96 < FB_BW <= 128)"), 105, IF(EXACT(B13, "G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)"), 115, 130)))))), 0)', value1, TEC_GRAPHICS)
    
    if sysinfo.disk_num > 1:
        if sysinfo.computer_type == 3:
            TEC_STORAGE = 2.6 * (sysinfo.disk_num - 1)
        else:
            TEC_STORAGE = 26 * (sysinfo.disk_num - 1)
    else:
        TEC_STORAGE = 0
    sheet.write("F15", "TEC_STORAGE", field1)
    if sysinfo.computer_type == 3:
        sheet.write("G15", '=IF(B7>1,2.6*(B7-1),0)', value1, TEC_STORAGE)
    else:
        sheet.write("G15", '=IF(B7>1,26*(B7-1),0)', value1, TEC_STORAGE)

    sheet.write("H11", "P:", right)
    sheet.write("I11", '=B4*B5', left, P)

    EP_LABEL = "EP:"
    if sysinfo.ep:
        if sysinfo.diagonal >= 27:
            EP = 0.75
        else:
            EP = 0.3
        sheet.write("B22", "Yes", value0)
    else:
        if sysinfo.computer_type == 1:
            EP = ""
            EP_LABEL = ""
        else:
            EP = 0
    sheet.write("H12", '=IF(EXACT(B3, "Desktop"), "", "EP:")', right, EP_LABEL)
    sheet.write("I12", '=IF(EXACT(B3, "Desktop"), "", IF(EXACT(B22, "Yes"), IF(B17 >= 27, 0.75, 0.3), 0))', left, EP)

    r_label = "r:"
    if sysinfo.computer_type != 1:
        r = 1.0 * sysinfo.width * sysinfo.height / 1000000
    else:
        if sysinfo.computer_type == 1:
            r = ""
            r_label = ""
        else:
            r = 0
    sheet.write("H13", '=IF(EXACT(B3, "Desktop"), "", "r:"', right, r_label)
    sheet.write("I13", '=IF(EXACT(B3, "Desktop"), "", B18 * B19 / 1000000)', left, r)

    A_LABEL = "A:"
    if sysinfo.computer_type != 1:
        A =  1.0 * sysinfo.diagonal * sysinfo.diagonal * sysinfo.width * sysinfo.height / (sysinfo.width ** 2 + sysinfo.height ** 2)
    else:
        if sysinfo.computer_type == 1:
            A = ""
            A_LABEL = ""
        else:
            A = 0
    sheet.write("H14", '=IF(EXACT(B3, "Desktop"), "", "A:"', right, A_LABEL)
    sheet.merge_range("I14:J14", A, left)
    sheet.write("I14", '=IF(EXACT(B3, "Desktop"), "", B17 * B17 * B18 * B19 / (B18 * B18 + B19 * B19))', left, A)

    sheet.write("F16", "TEC_INT_DISPLAY", field1)
    if sysinfo.computer_type == 3:
        TEC_INT_DISPLAY = 8.76 * 0.3 * (1+EP) * (2*r + 0.02*A)
    elif sysinfo.computer_type == 2:
        TEC_INT_DISPLAY = 8.76 * 0.35 * (1+EP) * (4*r + 0.05*A)
    else:
        TEC_INT_DISPLAY = 0
    sheet.write("G16", '=IF(EXACT(B3, "Notebook"), 8.76 * 0.3 * (1+I12) * (2*I13 + 0.02*I14), IF(EXACT(B3, "Integrated Desktop"), 8.76 * 0.35 * (1+I12) * (4*I13 + 0.05*I14), 0))', value1, TEC_INT_DISPLAY)

    sheet.write("F17", "TEC_SWITCHABLE", field1)
    if sysinfo.computer_type == 3:
        TEC_SWITCHABLE = 0
    elif sysinfo.switchable:
        TEC_SWITCHABLE = 0.5 * 36
    else:
        TEC_SWITCHABLE = 0
    sheet.write("G17", '=IF(EXACT(B3, "Notebook"), 0, IF(EXACT(B11, "Switchable"), 0.5 * 36, 0))', value1, TEC_SWITCHABLE)

    sheet.write("F18", "TEC_EEE", field1)
    if sysinfo.computer_type == 3:
        TEC_EEE = 8.76 * 0.2 * (0.1 + 0.3) * sysinfo.eee
    else:
        TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * sysinfo.eee
    sheet.write("G18", '=IF(EXACT(B3, "Notebook"), 8.76 * 0.2 * (0.1 + 0.3) * B8, 8.76 * 0.2 * (0.15 + 0.35) * B8)', value1, TEC_EEE)

    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE + TEC_INT_DISPLAY + TEC_SWITCHABLE + TEC_EEE
    sheet.write("F19", "E_TEC_MAX", result)
    sheet.write("G19", "=(1+G11)*(G12+G13+G14+G15+G16+G17+G18)", result_value, E_TEC_MAX)

    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    sheet.write("G20", '=IF(E19<=G19, "PASS", "FAIL")', center, RESULT)

def generate_excel_for_workstations(book, sysinfo, version):

    sheet = book.add_worksheet()

    sheet.set_column('A:A', 38)
    sheet.set_column('B:B', 12)

    center = book.add_format({'align': 'center'})

    header = book.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'fg_color': '#CFE2F3'})

    field = book.add_format({
        'border': 1,
        'fg_color': '#F3F3F3'})
    field1 = book.add_format({
        'left': 1,
        'right': 1,
        'fg_color': '#F3F3F3'})

    float2 = book.add_format({'border': 1})
    float2.set_num_format('0.00')

    result = book.add_format({
        'border': 1,
        'fg_color': '#F4CCCC'})
    result_value = book.add_format({
        'border': 1,
        'fg_color': '#FFF2CC'})
    result_value.set_num_format('0.00')

    value = book.add_format({'border': 1})
    value1 = book.add_format({
        'left': 1,
        'right': 1})
    value1.set_num_format('0.00')

    sheet.merge_range("A1:B1", "General", header)
    sheet.write('A2', "Product Type", field)
    sheet.write('B2', "Workstations", value)
    sheet.write("A3", "Number of Hard Drives", field)
    sheet.write("B3", sysinfo.disk_num, value)
    sheet.write("A4", "IEEE 802.3az compliant Gigabit Ethernet", field)
    sheet.write("B4", sysinfo.eee, value)

    sheet.merge_range("A6:B6", "Power Consumption", header)
    sheet.write("A7", "Off mode (W)", field)
    sheet.write("B7", sysinfo.off, float2)
    sheet.write("A8", "Sleep mode (W)", field)
    sheet.write("B8", sysinfo.sleep, float2)
    sheet.write("A9", "Long idle mode (W)", field)
    sheet.write("B9", sysinfo.long_idle, float2)
    sheet.write("A10", "Short idle mode (W)", field)
    sheet.write("B10", sysinfo.short_idle, float2)
    sheet.write("A11", "Max Power (W)", field)
    sheet.write("B11", sysinfo.max_power, float2)

    sheet.merge_range("A13:B13", "Energy Star 5.2", header)

    (T_OFF, T_SLEEP, T_IDLE) = (0.35, 0.1, 0.55)

    sheet.write('A14', 'T_OFF', field1)
    sheet.write('B14', T_OFF, value1)

    sheet.write('A15', 'T_SLEEP', field1)
    sheet.write('B15', T_SLEEP, value1)

    sheet.write('A16', 'T_IDLE', field1)
    sheet.write('B16', T_IDLE, value1)

    P_TEC = T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_IDLE * sysinfo.short_idle

    sheet.write('A17', 'P_TEC', result)
    sheet.write('B17', '=(B14*B7+B15*B8+B16*B10)', result_value, P_TEC)

    P_MAX = 0.28 * (sysinfo.max_power + sysinfo.disk_num * 5)
    sheet.write('A18', 'P_MAX', result)
    sheet.write('B18', '=0.28*(B11+B3*5)', result_value, P_MAX)

    if P_TEC <= P_MAX:
        RESULT = 'PASS'
    else:
        RESULT = 'FAIL'
    sheet.write('B19', '=IF(B17 <= B18, "PASS", "FAIL")', center, RESULT)

    sheet.merge_range("A20:B20", "Energy Star 6.0", header)

    (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.35, 0.1, 0.15, 0.4)

    sheet.write('A21', 'T_OFF', field1)
    sheet.write('B21', T_OFF, value1)

    sheet.write('A22', 'T_SLEEP', field1)
    sheet.write('B22', T_SLEEP, value1)

    sheet.write('A23', 'T_LONG_IDLE', field1)
    sheet.write('B23', T_LONG_IDLE, value1)

    sheet.write('A24', 'T_SHORT_IDLE', field1)
    sheet.write('B24', T_SHORT_IDLE, value1)

    P_TEC = T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_LONG_IDLE * sysinfo.long_idle + T_SHORT_IDLE * sysinfo.short_idle

    sheet.write('A25', 'P_TEC', result)
    sheet.write('B25', '=(B21*B7+B22*B8+B23*B9+B24*B10)', result_value, P_TEC)

    P_MAX = 0.28 * (sysinfo.max_power + sysinfo.disk_num * 5) + 8.76 * sysinfo.eee * (sysinfo.sleep + sysinfo.long_idle + sysinfo.short_idle)
    sheet.write('A26', 'P_MAX', result)
    sheet.write('B26', '=0.28*(B11+B3*5) + 8.76*B4*(B8 + B9 + B10)', result_value, P_MAX)

    if P_TEC <= P_MAX:
        RESULT = 'PASS'
    else:
        RESULT = 'FAIL'
    sheet.write('B27', '=IF(B25 <= B26, "PASS", "FAIL")', center, RESULT)

def generate_excel_for_small_scale_servers(book, sysinfo, version):

    sheet = book.add_worksheet()

    sheet.set_column('A:A', 45)
    sheet.set_column('B:B', 15)

    header = book.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'fg_color': '#CFE2F3'})

    field = book.add_format({
        'border': 1,
        'fg_color': '#F3F3F3'})
    field1 = book.add_format({
        'left': 1,
        'right': 1,
        'fg_color': '#F3F3F3'})

    value0 = book.add_format({'border': 1, 'fg_color': '#D9EAD3'})
    value = book.add_format({'border': 1})
    value1 = book.add_format({
        'left': 1,
        'right': 1})
    value1.set_num_format('0.00')
    center = book.add_format({'align': 'center'})

    float2 = book.add_format({'border': 1})
    float2.set_num_format('0.00')

    result = book.add_format({
        'border': 1,
        'fg_color': '#F4CCCC'})
    result_value = book.add_format({
        'border': 1,
        'fg_color': '#FFF2CC'})
    result_value.set_num_format('0.00')

    sheet.merge_range("A1:B1", "General", header)

    sheet.write('A2', "Product Type", field)
    sheet.write('B2', "Small-scale Servers", value)

    sheet.write("A3", "Wake-On-LAN (WOL) by default upon shipment", field)
    sheet.write("B3", 'Enabled', value0)
    sheet.data_validation('B3', {
        'validate': 'list',
        'source': [
            'Enabled',
            'Disabled']})

    sheet.write("A4", "More than one physical core?", field)
    if sysinfo.cpu_core > 1:
        sheet.write("B4", 'Yes', value)
    else:
        sheet.write("B4", 'No', value)
    sheet.data_validation('B4', {
        'validate': 'list',
        'source': [
            'Yes',
            'No']})

    sheet.write("A5", "More than one discrete processor?", field)
    if sysinfo.more_discrete:
        sheet.write("B5", 'Yes', value)
    else:
        sheet.write("B5", 'No', value)
    sheet.data_validation('B5', {
        'validate': 'list',
        'source': [
            'Yes',
            'No']})

    sheet.write("A6", "More than or equal to one gigabyte of system memory?", field)
    if sysinfo.mem_size >= 1:
        sheet.write("B6", 'Yes', value)
    else:
        sheet.write("B6", 'No', value)
    sheet.data_validation('B6', {
        'validate': 'list',
        'source': [
            'Yes',
            'No']})

    sheet.write("A7", "IEEE 802.3az compliant Gigabit Ethernet", field)
    sheet.write("B7", sysinfo.eee, value)

    sheet.write("A8", "Number of Hard Drives", field)
    sheet.write("B8", sysinfo.disk_num, value)

    sheet.merge_range("A10:B10", "Power Consumption", header)
    sheet.write("A11", "Off mode (W)", field)
    sheet.write("B11", sysinfo.off, float2)
    sheet.write("A12", "Short idle mode (W)", field)
    sheet.write("B12", sysinfo.short_idle, float2)

    sheet.merge_range("A14:B14", "Energy Star 5.2", header)

    P_OFF_BASE = 2
    P_OFF_WOL = 0.7
    P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

    sheet.write('A15', 'P_OFF_BASE', field1)
    sheet.write('B15', P_OFF_BASE, value1)

    sheet.write('A16', 'P_OFF_WOL', field1)
    sheet.write('B16', '=IF(EXACT(B3, "Enabled"), 0.7, 0)', value1, P_OFF_WOL)

    sheet.write('A17', 'P_OFF_MAX', result)
    sheet.write('B17', '=B15+B16', result_value, P_OFF_MAX)

    if (sysinfo.cpu_core > 1 or sysinfo.more_discrete) and sysinfo.mem_size >= 1:
        P_IDLE_MAX = 65
        category = 'A'
    else:
        P_IDLE_MAX = 50
        category = 'B'

    sheet.write('A18', "Product Category", field)
    sheet.write('B18', '=IF(AND(OR(EXACT(B4, "Yes"), EXACT(B5, "Yes")), EXACT(B6, "Yes")), "B", "A")', value, category)

    sheet.write('A19', 'P_IDLE_MAX', result)
    sheet.write('B19', '=IF(EXACT(B18, "B"), 65, 50)', result_value, P_IDLE_MAX)

    if sysinfo.off <= P_OFF_MAX and sysinfo.short_idle <= P_IDLE_MAX:
        RESULT = 'PASS'
    else:
        RESULT = 'FAIL'
    sheet.write('B20', '=IF(AND(B11 <= B18, B12 <= B19), "PASS", "FAIL")', center, RESULT)

    sheet.merge_range("A21:B21", "Energy Star 6.0", header)

    P_OFF_BASE = 1
    P_OFF_WOL = 0.4
    P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

    sheet.write('A22', 'P_OFF_BASE', field1)
    sheet.write('B22', P_OFF_BASE, value1)

    sheet.write('A23', 'P_OFF_WOL', field1)
    sheet.write('B23', '=IF(EXACT(B3, "Enabled"), 0.4, 0)', value1, P_OFF_WOL)

    sheet.write('A24', 'P_OFF_MAX', result)
    sheet.write('B24', '=B22+B23', result_value, P_OFF_MAX)

    P_IDLE_BASE = 24
    sheet.write('A25', 'P_IDLE_BASE', field1)
    sheet.write('B25', P_IDLE_BASE, value1)

    P_IDLE_HDD = 8
    sheet.write('A26', 'P_IDLE_HDD', field1)
    sheet.write('B26', P_IDLE_HDD, value1)

    P_EEE = 0.2 * sysinfo.eee
    sheet.write('A27', 'P_EEE', field1)
    sheet.write('B27', '=0.2*B7', value1, P_EEE)

    P_IDLE_MAX = P_IDLE_BASE + (sysinfo.disk_num - 1) * P_IDLE_HDD + P_EEE
    sheet.write('A28', 'P_IDLE_MAX', result)
    sheet.write('B28', '=B25 + (B8 - 1) * B26 + B27', result_value, P_IDLE_MAX)

    if sysinfo.off <= P_OFF_MAX and sysinfo.short_idle <= P_IDLE_MAX:
        RESULT = 'PASS'
    else:
        RESULT = 'FAIL'
    sheet.write('B29', '=IF(AND(B11 <= B24, B12 <= B28), "PASS", "FAIL")', center, RESULT)

def generate_excel_for_thin_clients(book, sysinfo, version):

    sheet = book.add_worksheet()

    sheet.set_column('A:A', 45)
    sheet.set_column('B:B', 10)
    sheet.set_column('C:C', 1)
    sheet.set_column('D:D', 15)
    sheet.set_column('E:E', 8)
    sheet.set_column('F:F', 5)

    header = book.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'fg_color': '#CFE2F3'})

    field = book.add_format({
        'border': 1,
        'fg_color': '#F3F3F3'})
    field1 = book.add_format({
        'left': 1,
        'right': 1,
        'fg_color': '#F3F3F3'})

    value0 = book.add_format({'border': 1, 'fg_color': '#D9EAD3'})
    value = book.add_format({'border': 1})
    value1 = book.add_format({
        'left': 1,
        'right': 1})
    value1.set_num_format('0.00')
    center = book.add_format({'align': 'center'})
    left = book.add_format({'align': 'left'})
    right = book.add_format({'align': 'right'})

    float2 = book.add_format({'border': 1})
    float2.set_num_format('0.00')

    result = book.add_format({
        'border': 1,
        'fg_color': '#F4CCCC'})
    result_value = book.add_format({
        'border': 1,
        'fg_color': '#FFF2CC'})
    result_value.set_num_format('0.00')

    sheet.merge_range("A1:B1", "General", header)

    sheet.write('A2', "Product Type", field)
    sheet.write('B2', "Thin Clients", value)

    sheet.write("A3", "Wake-On-LAN (WOL) by default upon shipment", field)
    sheet.write("B3", 'Enabled', value0)
    sheet.data_validation('B3', {
        'validate': 'list',
        'source': [
            'Enabled',
            'Disabled']})

    sheet.write("A4", "Local multimedia encode/decode support", field)
    if sysinfo.media_codec:
        sheet.write("B4", 'Yes', value)
    else:
        sheet.write("B4", 'No', value)
    sheet.data_validation('B4', {
        'validate': 'list',
        'source': [
            'Yes',
            'No']})

    sheet.write("A5", "Discrete Graphics", field)
    if sysinfo.discrete:
        sheet.write("B5", 'Yes', value)
    else:
        sheet.write("B5", 'No', value)
    sheet.data_validation('B5', {
        'validate': 'list',
        'source': [
            'Yes',
            'No']})

    sheet.write("A6", "IEEE 802.3az compliant Gigabit Ethernet", field)
    sheet.write("B6", sysinfo.eee, value)

    sheet.merge_range("A8:B8", "Power Consumption", header)
    sheet.write("A9", "Off mode (W)", field)
    sheet.write("B9", sysinfo.off, float2)
    sheet.write("A10", "Sleep mode (W)", field)
    sheet.write("B10", sysinfo.sleep, float2)
    sheet.write("A11", "Long idle mode (W)", field)
    sheet.write("B11", sysinfo.long_idle, float2)
    sheet.write("A12", "Short idle mode (W)", field)
    sheet.write("B12", sysinfo.short_idle, float2)


    if sysinfo.integrated_display:
        sheet.merge_range("A14:B14", "Display", header)

        sheet.write("A15", "Enhanced-performance Integrated Display", field)
        sheet.write("B15", "No", value0)
        sheet.data_validation('B15', {
            'validate': 'list',
            'source': [
                'Yes',
                'No']})

        sheet.write("A16", "Physical Diagonal (inch)", field)
        sheet.write("B16", sysinfo.diagonal, value)

        sheet.write("A17", "Screen Width (px)", field)
        sheet.write("B17", sysinfo.width, value)

        sheet.write("A18", "Screen Height (px)", field)
        sheet.write("B18", sysinfo.height, value)

    sheet.merge_range("D1:E1", "Energy Star 5.2", header)

    sheet.write("D2", "Product Category", field)
    sheet.write("E2", '=IF(EXACT(B4, "Yes"), "B", "A")', value, "B")

    P_OFF_BASE = 2
    P_OFF_WOL = 0.7
    P_OFF_MAX = P_OFF_BASE + P_OFF_WOL

    sheet.write("D3", "P_OFF_BASE", field1)
    sheet.write("E3", P_OFF_BASE, value1)

    sheet.write("D4", "P_WOL_BASE", field1)
    sheet.write("E4", '=IF(EXACT(B3, "Enabled"), 0.7, 0)', value1, P_OFF_WOL)

    sheet.write("D5", "P_OFF_MAX", result)
    sheet.write("E5", '=E3+E4', result_value, P_OFF_MAX)

    P_SLEEP_BASE = 2
    P_SLEEP_WOL = 0.7
    P_SLEEP_MAX = P_SLEEP_BASE + P_SLEEP_WOL

    sheet.write("D6", "P_SLEEP_BASE", field1)
    sheet.write("E6", P_SLEEP_BASE, value1)

    sheet.write("D7", "P_WOL_BASE", field1)
    sheet.write("E7", '=IF(EXACT(B3, "Enabled"), 0.7, 0)', value1, P_SLEEP_WOL)

    sheet.write("D8", "P_SLEEP_MAX", result)
    sheet.write("E8", '=B19+B20', result_value, P_SLEEP_MAX)

    P_IDLE_MAX = 15
    sheet.write("D9", "P_IDLE_MAX", result)
    sheet.write("E9", '=IF(EXACT(B4, "Yes"), 15, 12)', result_value, P_IDLE_MAX)

    if sysinfo.off <= P_OFF_MAX and sysinfo.sleep <= P_SLEEP_MAX and sysinfo.short_idle <= P_IDLE_MAX:
        RESULT = 'PASS'
    else:
        RESULT = 'FAIL'
    sheet.write('E10', '=IF(AND(B9 <= E5, B10 <= E8, B12 <= E9), "PASS", "FAIL")', center, RESULT)

    sheet.merge_range("D11:E11", "Energy Star 6.0", header)

    (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

    sheet.write("D12", "T_OFF", field1)
    sheet.write("E12", T_OFF, value1)

    sheet.write("D13", "T_SLEEP", field1)
    sheet.write("E13", T_SLEEP, value1)

    sheet.write("D14", "T_LONG_IDLE", field1)
    sheet.write("E14", T_LONG_IDLE, value1)

    sheet.write("D15", "T_SHORT_IDLE", field1)
    sheet.write("E15", T_SHORT_IDLE, value1)

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_LONG_IDLE * sysinfo.long_idle + T_SHORT_IDLE * sysinfo.short_idle) * 8760 / 1000
    sheet.write("D16", "E_TEC", result)
    sheet.write("E16", "=(B9*E12+B10*E13+B11*E14+B12*E15)*8760/1000", result_value, E_TEC)

    TEC_BASE = 60
    sheet.write("D17", "TEC_BASE", field1)
    sheet.write("E17", TEC_BASE, value1)

    TEC_GRAPHICS = 36
    sheet.write("D18", "TEC_GRAPHICS", field1)
    sheet.write("E18", '=IF(EXACT(B5, "Yes"), 36, 0)', value1, TEC_GRAPHICS)

    TEC_WOL = 2
    sheet.write("D19", "TEC_WOL", field1)
    sheet.write("E19", '=IF(EXACT(B3, "Enabled"), 2, 0)', value1, TEC_WOL)

    EP_LABEL = "EP:"
    if sysinfo.integrated_display:
        if sysinfo.ep:
            sheet.write("B15", "Yes", value0)
            if sysinfo.diagonal >= 27:
                EP = 0.75
            else:
                EP = 0.3
        else:
            EP = 0
    else:
        EP = ""
        EP_LABEL = ""
    sheet.write("F12", '=IF(EXACT(A14, "Display"), "EP:", "")', right, EP_LABEL)
    sheet.write("G12", '=IF(EXACT(A14, "Display"), IF(EXACT(B15, "Yes"), IF(B16 >= 27, 0.75, 0.3), 0), "")', left, EP)

    r_label = "r:"
    if sysinfo.integrated_display:
        r = 1.0 * sysinfo.width * sysinfo.height / 1000000
    else:
        r = ""
        r_label = ""
    sheet.write("F13", '=IF(EXACT(A14, "Display"), "r:", ""', right, r_label)
    sheet.write("G13", '=IF(EXACT(A14, "Display"), B17 * B18 / 1000000, "")', left, r)

    A_LABEL = "A:"
    if sysinfo.integrated_display:
        A =  1.0 * sysinfo.diagonal * sysinfo.diagonal * sysinfo.width * sysinfo.height / (sysinfo.width ** 2 + sysinfo.height ** 2)
    else:
        A = ""
        A_LABEL = ""
    sheet.write("F14", '=IF(EXACT(A14, "Display"), "A:", "")', right, A_LABEL)
    sheet.merge_range("G14:H14", A, left)
    sheet.write("G14", '=IF(EXACT(A14, "Display"), B16 * B16 * B17 * B18 / (B17 * B17 + B18 * B18), "")', left, A)

    if sysinfo.integrated_display:
        TEC_INT_DISPLAY = 8.76 * 0.35 * (1+EP) * (4*r + 0.05*A)
    else:
        TEC_INT_DISPLAY = 0
    sheet.write("D20", "TEC_INT_DISPLAY", field1)
    sheet.write("E20", '=IF(EXACT(A14, "Display"), 8.76 * 0.35 * (1+G12) * (4*G13 + 0.05*G14), 0)', value1, TEC_INT_DISPLAY)

    TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * sysinfo.eee
    sheet.write("D21", "TEC_EEE", field1)
    sheet.write("E21", '=8.76 * 0.2 * (0.15 + 0.35) * B6', value1, TEC_EEE)

    E_TEC_MAX = TEC_BASE + TEC_GRAPHICS + TEC_WOL + TEC_INT_DISPLAY + TEC_EEE
    sheet.write("D22", "E_TEC_MAX", result)
    sheet.write("E22", '=E17 + E18 + E19 + E20 + E21', result_value, E_TEC_MAX)

    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    sheet.write("E23", '=IF(E16<=E22, "PASS", "FAIL")', center, RESULT)