# -*- coding: utf-8; indent-tabs-mode: nil; tab-width: 4; c-basic-offset: 4;-*-
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
from .erplot3 import *

G1 = 'G1 (FB_BW <= 16)'
G2 = 'G2 (16 < FB_BW <= 32)'
G3 = 'G3 (32 < FB_BW <= 64)'
G4 = 'G4 (64 < FB_BW <= 96)'
G5 = 'G5 (96 < FB_BW <= 128)'
G6 = 'G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)'
G7 = 'G7 (FB_BW > 128; Frame Buffer Data Width >= 192 bits)'
GN_LIST = [G1, G2, G3, G4, G5, G6, G7]

def generate_excel(sysinfo, version, output):
    if not output:
        return
    elif not output.endswith('.xlsx'):
        warning('Please use xlsx as the suffix of Excel file.') 
        output = output + '.xlsx'

    if sysinfo.product_type == 1:
        excel = ExcelMaker(version, output)
    else:
        try:
            from xlsxwriter import Workbook
        except:
            warning("You need to install Python xlsxwriter module or you can not output Excel format file.")
            return

        book = Workbook(output)
        book.set_properties({'comments':"Energy Tools %s" % (version)})

    if sysinfo.product_type == 1:
        generate_excel_for_computers(excel, sysinfo)
    elif sysinfo.product_type == 2:
        generate_excel_for_workstations(book, sysinfo, version)
    elif sysinfo.product_type == 3:
        generate_excel_for_small_scale_servers(book, sysinfo, version)
    elif sysinfo.product_type == 4:
        generate_excel_for_thin_clients(book, sysinfo, version)

    if sysinfo.product_type == 1:
        excel.save()
    else:
        book.close()

def formula_strip(formula):
    return ' '.join(formula.split())

class ExcelMaker:
    def __init__(self, version, output):
        try:
            from xlsxwriter import Workbook
        except:
            warning("You need to install Python xlsxwriter module or you can not output Excel format file.")
            return
        self.book = Workbook(output)
        self.book.set_properties({'comments':"Energy Tools %s" % (version)})
        self.sheet = self.book.add_worksheet()
        self.adjust_column_width()
        self.setup_theme()
        self.row = 1
        self.column = 'A'
        self.pos = {
                'g1': G1,
                'g2': G2,
                'g3': G3,
                'g4': G4,
                'g5': G5,
                'g6': G6,
                'g7': G7}

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
        sheet.set_column('K:K', 6)
        sheet.set_column('L:L', 6)
        sheet.set_column('M:M', 6)
        sheet.set_column('N:N', 6)

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
        theme["warning"] = book.add_format({'color': '#FF0000'})
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

    def ncell(self, width, height, label, formula=None, value=None, validator=None, abbr=None, twin=False):
        if type(value) is list:
            validator = value
            value = None
        if value is None:
            if formula is None:
                value = label
            else:
                value = formula
                formula = None
        debug("Field: %s, Formula: %s, Value: %s, Validator: %s" % (label, formula, value, validator))

        if label in palette:
            (i, j) = palette[label]
            debug("Theme: (%s, %s)" % (i, j))
            if not twin and width == 1 and height == 1 and j:
                theme1 = self.theme[j]
            else:
                theme1 = self.theme[i]
            if j:
                theme2 = self.theme[j]
            else:
                theme2 = self.theme['value']
        else:
            theme1 = self.theme['field']
            theme2 = self.theme['value']
            debug("Theme: (%s, %s)" % ('field', 'value'))

        end_column = chr(ord(self.column) + width - 1)
        if twin:
            next_column = chr(ord(self.column) + width)
        else:
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
        if twin:
            self.sheet.write("%s" % start_cell, label, theme1)
            if formula:
                formula = formula % self.pos
                self.sheet.write("%s" % next_cell, formula, theme2, value)
            else:
                self.sheet.write("%s" % next_cell, value, theme2)
            if validator:
                self.sheet.data_validation("%s" % next_cell, {
                    'validate': 'list',
                    'source': validator})
        else:
            if formula:
                formula = formula % self.pos
                self.sheet.write("%s" % start_cell, formula, theme1, value)
            else:
                self.sheet.write("%s" % start_cell, value, theme1)
        if twin:
            if abbr:
                self.pos[abbr] = next_cell
            else:
                self.pos[label] = next_cell
        else:
            if abbr:
                self.pos[abbr] = start_cell 
            else:
                self.pos[label] = start_cell
        self.row = self.row + height

    def cell(self, label, formula=None, value=None, validator=None, abbr=None):
        self.ncell(1, 1, label, formula, value, validator, abbr)

    def tcell(self, label, formula=None, value=None, validator=None, abbr=None):
        self.ncell(1, 1, label, formula, value, validator, abbr, True)

    def up(self, step=1):
        self.row = self.row - step
        return self

    def down(self, step=1):
        self.row = self.row + step
        return self

    def left(self, step=1):
        self.column = chr(ord(self.column) - step)
        return self

    def right(self, step=1):
        self.column = chr(ord(self.column) + step)
        return self

    def position(self):
        return (self.column, self.row)

    def jump(self, column, row):
        self.column = column
        self.row = row
        return self

    def shift(self, column, row, x, y):
        self.column = chr(ord(column) + x)
        self.row = row + y
        return self

palette = {
        # Header
        'General': ('header', None),
        'Graphics': ('header', None),
        'Additional Discrete Graphics': ('header', None),
        'Power Consumption': ('header', None),
        'Display': ('header', None),
        'Energy Star 5.2': ('header', None),
        'Energy Star 6.0': ('header', None),
        'ErP Lot 3': ('header', None),
        'from 1 July 2014': ('header', None),
        'from 1 January 2016': ('header', None),

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
        'P_OFF_WOL': ('field1', 'value1'),
        'P_SLEEP': ('field1', 'value1'),
        'P_SLEEP_WOL': ('field1', 'value1'),
        'P_LONG_IDLE': ('field1', 'value1'),
        'ALLOWANCE_PSU': ('field1', 'float3'),
        'TEC_BASE': ('field1', 'value1'),
        'TEC_MEMORY': ('field1', 'value1'),
        'TEC_GRAPHICS': ('field1', 'value1'),
        'TEC_STORAGE': ('field1', 'value1'),
        'TEC_INT_DISPLAY': ('field1', 'value1'),
        'TEC_SWITCHABLE': ('field1', 'value1'),
        'TEC_EEE': ('field1', 'value1'),
        'TEC_TV_TUNER': ('field1', 'value1'),
        'TEC_AUDIO': ('field1', 'value1'),

        # Cell with different background color and float 1
        'P_IDLE': ('field2', 'value2'),
        'P_SHORT_IDLE': ('field2', 'value2'),

        # Important result
        'E_TEC': ('result', 'result_value'),
        'E_TEC_WOL': ('result', 'result_value'),
        'E_TEC_MAX': ('result', 'result_value'),

        # Prompt
        'warning': ('warning', None),
        'center': ('center', None),
        'field1': ('field1', None),
        'result': ('result', None),
        'P:': ('right', 'left'),
        'EP': ('left', None),
        'r': ('left', None),
        'A:': ('left', None),
        'left': ('left', None),
        'right': ('right', None),

        # Unsure
        "IEEE 802.3az compliant Gigabit Ethernet": ('field', 'unsure'),
        "GPU Frame Buffer Width": ('field', 'unsure'),
        "Graphics Category": ('field', 'unsure'),
        "Enhanced-performance Integrated Display": ('field', 'unsure'),
        "Power Supply Efficiency Allowance requirements:": ('field', 'unsure'),

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
    # General Information
    excel.ncell(2, 1, "General")
    excel.tcell("Product Type", "Desktop, Integrated Desktop, and Notebook")

    if sysinfo.computer_type == 1:
        msg = "Desktop"
    elif sysinfo.computer_type == 2:
        msg = "Integrated Desktop"
    else:
        msg = "Notebook"

    excel.tcell("Computer Type", msg, ["Desktop", "Integrated Desktop", "Notebook"], abbr='computer')
    excel.tcell("CPU cores", sysinfo.cpu_core, abbr='cpu_core')

    excel.tcell("CPU clock (GHz)", sysinfo.cpu_clock, abbr='cpu_clock')
    excel.tcell("Memory size (GB)", sysinfo.mem_size, abbr='memory')
    excel.tcell("Number of Hard Drives", sysinfo.disk_num, abbr='disk_number')
    excel.tcell("Number of Discrete Graphics Cards", sysinfo.discrete_gpu_num, abbr='gpu_number')
    if sysinfo.tvtuner:
        msg = "Yes"
    else:
        msg = "No"
    excel.tcell("Discrete television tuner", msg, ["Yes", "No"], abbr='tvtuner')
    if sysinfo.audio:
        msg = "Yes"
    else:
        msg = "No"
    excel.tcell("Discrete audio card", msg, ["Yes", "No"], abbr='audio')
    excel.tcell("IEEE 802.3az compliant Gigabit Ethernet", sysinfo.eee, abbr='eee')

    excel.down()

    excel.ncell(2, 1, "Graphics")
    if sysinfo.switchable:
        msg = "Switchable"
    elif sysinfo.discrete:
        msg = "Discrete"
    else:
        msg = "Integrated"
    excel.tcell("Graphics Type", msg, ['Integrated', 'Switchable', 'Discrete'], abbr='gpu_type')
    if sysinfo.computer_type == 3:
        msg = "<= 64-bit"
        validator = ['<= 64-bit', '> 64-bit and <= 128-bit', '> 128-bit']
    else:
        msg = "<= 128-bit"
        validator = ['<= 128-bit', '> 128-bit']
    excel.tcell("GPU Frame Buffer Width", msg, validator, abbr='gpu_width')
    excel.tcell("Graphics Category", G1, GN_LIST, abbr='gpu_category')

    excel.down()

    excel.ncell(2, 1, "Power Consumption")
    excel.tcell("Off mode (W)", sysinfo.off, abbr='off')
    excel.tcell("Off mode (W) with WOL", sysinfo.off_wol, abbr='off_wol')
    excel.tcell("Sleep mode (W)", sysinfo.sleep, abbr='sleep')
    excel.tcell("Sleep mode (W) with WOL", sysinfo.sleep_wol, abbr='sleep_wol')
    excel.tcell("Long idle mode (W)", sysinfo.long_idle, abbr='long_idle')
    excel.tcell("Short idle mode (W)", sysinfo.short_idle, abbr='short_idle')

    excel.down()

    excel.ncell(2, 1, "Display")
    if sysinfo.ep:
        msg = "Yes"
    else:
        msg = "No"
    excel.tcell("Enhanced-performance Integrated Display", msg, ["Yes", "No"], abbr='ep_display')
    excel.tcell("Physical Diagonal (inch)", sysinfo.diagonal, abbr='diagonal')
    excel.tcell("Screen Width (px)", sysinfo.width, abbr='width')
    excel.tcell("Screen Height (px)", sysinfo.height, abbr='height')

    if sysinfo.discrete_gpu_num > 1:
        excel.down()
        excel.ncell(2, 1, "Additional Discrete Graphics")
        for i in range(int(sysinfo.discrete_gpu_num) - 1):
            palette["dGfx #%d" % (i+2)] = ('field', 'unsure')
            excel.tcell("dGfx #%d" % (i+2), G1, GN_LIST, abbr='dGfx%d' % (i+2))

    # Energy Star 5.2
    excel.jump('D', 1)
    if sysinfo.computer_type == 3:
        width = 6
    else:
        width = 7
    excel.ncell(width, 1, "Energy Star 5.2")

    # E_TEC
    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

    excel.tcell("T_OFF", '=IF(EXACT(%(computer)s,"Notebook"),0.6,0.55', T_OFF, abbr='t_off')
    excel.tcell("T_SLEEP", '=IF(EXACT(%(computer)s,"Notebook"),0.1,0.05', T_SLEEP, abbr='t_sleep')
    excel.tcell("T_IDLE", '=IF(EXACT(%(computer)s,"Notebook"),0.3,0.4', T_IDLE, abbr='t_idle')

    excel.tcell("P_OFF", '=%(off)s', sysinfo.off, abbr='p_off')
    excel.tcell("P_SLEEP", '=%(sleep)s', sysinfo.sleep, abbr='p_sleep')
    excel.tcell("P_IDLE", '=%(short_idle)s', sysinfo.short_idle, abbr='p_idle')

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_IDLE * sysinfo.short_idle) * 8760 / 1000

    excel.tcell("E_TEC", "=(%(t_off)s*%(p_off)s+%(t_sleep)s*%(p_sleep)s+%(t_idle)s*%(p_idle)s)*8760/1000", E_TEC)

    # Prompt
    excel.jump('F', 2)
    excel.ncell(1, 2, "Category")
    excel.up(2).right()
    excel.ncell(1, 2, "A")
    excel.up(2).right()
    excel.ncell(1, 2, "B")
    excel.up(2).right()
    excel.ncell(1, 2, "C")

    if sysinfo.computer_type != 3:
        excel.jump('J', 2)
        excel.ncell(1, 2, "D")

    if sysinfo.computer_type == 3:
        excel.jump('H', 2)
        if sysinfo.discrete:
            msg = "B"
        else:
            msg = ""
        excel.ncell(1, 2, 'B', '=IF(EXACT(%(gpu_type)s,"Discrete"), "B", "")', msg)

        excel.jump('I', 2)
        excel.ncell(1, 2, 'C', '=IF(AND(EXACT(%(gpu_type)s,"Discrete"), EXACT(%(gpu_width)s, "> 128-bit"), %(cpu_core)s>=2, %(memory)s>=2), "C", "")', "")
    else:
        excel.jump('H', 2)
        if sysinfo.cpu_core == 2 and sysinfo.mem_size >=2:
            msg = "B"
        else:
            msg = ""
        excel.ncell(1, 2, 'B', '=IF(AND(%(cpu_core)s=2,%(memory)s>=2), "B", "")', msg)

        excel.jump('I', 2)
        if sysinfo.cpu_core > 2 and (sysinfo.mem_size >= 2 or sysinfo.discrete):
            msg = "C"
        else:
            msg = ""
        excel.ncell(1, 2, 'C', '=IF(AND(%(cpu_core)s>2,OR(%(memory)s>=2,EXACT(%(gpu_type)s,"Discrete"))), "C", "")', msg)

        excel.jump('J', 2)
        if sysinfo.cpu_core >= 4 and sysinfo.mem_size >= 4:
            msg = "D"
        else:
            msg = ""
        excel.ncell(1, 2, 'D', '=IF(AND(%(cpu_core)s>=4,OR(%(memory)s>=4,AND(EXACT(%(gpu_type)s,"Discrete"),EXACT(%(gpu_width)s,"> 128-bit")))), "D", "")', msg)

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
        excel.cell("TEC_MEMORY", "=IF(%(memory)s>4, 0.4*(%(memory)s-4), 0)", TEC_MEMORY)
        excel.cell("TEC_GRAPHICS", TEC_GRAPHICS)
        excel.cell("TEC_STORAGE", "=IF(%(disk_number)s>1, 3*(%(disk_number)s-1), 0)", TEC_STORAGE)
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
        excel.cell("TEC_MEMORY", "=IF(%(memory)s>2, 1.0*(%(memory)s-2), 0)", TEC_MEMORY)
        excel.cell("TEC_GRAPHICS", '=IF(EXACT(%(gpu_width)s,"> 128-bit"), 50, 35)', TEC_GRAPHICS)
        excel.cell("TEC_STORAGE", "=IF(%(disk_number)s>1, 25*(%(disk_number)s-1), 0)", TEC_STORAGE)

    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
    excel.cell("E_TEC_MAX", "=%(TEC_BASE)s+%(TEC_MEMORY)s+%(TEC_GRAPHICS)s+%(TEC_STORAGE)s", E_TEC_MAX)
    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    excel.cell('center', '=IF(%(E_TEC)s<=%(E_TEC_MAX)s, "PASS", "FAIL")', RESULT)

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
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        excel.cell('TEC_BASE', '=IF(EXACT(%(B)s, "B"), 53, "")', TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%(B)s, "B"), %(TEC_MEMORY)s, "")', TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%(B)s, "B"), IF(EXACT(%(gpu_width)s, "<= 64-bit"), 0, 3), "")', TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%(B)s, "B"), %(TEC_STORAGE)s, "")', TEC_STORAGE)
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
        excel.cell('TEC_BASE', '=IF(EXACT(%(B)s, "B"), 175, "")', TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%(B)s, "B"), %(TEC_MEMORY)s, "")', TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%(B)s, "B"), %(TEC_GRAPHICS)s, "")', TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%(B)s, "B"), %(TEC_STORAGE)s, "")', TEC_STORAGE)

    excel.cell('E_TEC_MAX', '=IF(EXACT(%(B)s, "B"), %(TEC_BASE)s+%(TEC_MEMORY)s+%(TEC_GRAPHICS)s+%(TEC_STORAGE)s, "")', E_TEC_MAX)
    excel.cell('center', '=IF(EXACT(%(B)s, "B"), IF(%(E_TEC)s<=%(E_TEC_MAX)s, "PASS", "FAIL"), "")', RESULT)

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
        excel.cell('TEC_BASE', '=IF(EXACT(%(C)s, "C"), 88.5, "")', TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%(C)s, "C"), %(TEC_MEMORY)s, "")', TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%(C)s, "C"), 0, "")', TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%(C)s, "C"), %(TEC_STORAGE)s, "")', TEC_STORAGE)
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
        excel.cell('TEC_BASE', '=IF(EXACT(%(C)s, "C"), 209, "")', TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%(C)s, "C"), G5, "")', TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%(C)s, "C"), IF(EXACT(%(gpu_width)s, "> 128-bit"), 50, 0), "")', TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%(C)s, "C"), G7, "")', TEC_STORAGE)

    excel.cell('E_TEC_MAX', '=IF(EXACT(%(C)s, "C"), %(TEC_BASE)s+%(TEC_MEMORY)s+%(TEC_GRAPHICS)s+%(TEC_STORAGE)s, "")', E_TEC_MAX)
    excel.cell('center', '=IF(EXACT(%(C)s, "C"), IF(%(E_TEC)s<=%(E_TEC_MAX)s, "PASS", "FAIL"), "")', RESULT)

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
        excel.cell('TEC_BASE', '=IF(EXACT(%(D)s, "D"), 234, "")', TEC_BASE)
        excel.cell('TEC_MEMORY', '=IF(EXACT(%(D)s, "D"), IF(%(memory)s>4, %(memory)s-4, 0), "")', TEC_MEMORY)
        excel.cell('TEC_GRAPHICS', '=IF(EXACT(%(D)s, "D"), %(TEC_GRAPHICS)s, "")', TEC_GRAPHICS)
        excel.cell('TEC_STORAGE', '=IF(EXACT(%(D)s, "D"), %(TEC_STORAGE)s, "")', TEC_STORAGE)
        excel.cell('E_TEC_MAX', '=IF(EXACT(%(D)s, "D"), %(TEC_BASE)s+%(TEC_MEMORY)s+%(TEC_GRAPHICS)s+%(TEC_STORAGE)s, "")', E_TEC_MAX)
        excel.cell('center', '=IF(EXACT(%(D)s, "D"), IF(%(E_TEC)s<=%(E_TEC_MAX)s, "PASS", "FAIL"), "")', RESULT)

    # Energy Star 6.0
    excel.jump('D', 10)
    excel.ncell(4, 1, "Energy Star 6.0")

    # E_TEC
    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.25, 0.35, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

    excel.tcell("T_OFF", '=IF(EXACT(%(computer)s,"Notebook"),0.25,0.45', T_OFF, abbr='t_off')
    excel.tcell("T_SLEEP", '=IF(EXACT(%(computer)s,"Notebook"),0.35,0.05', T_SLEEP, abbr='t_sleep')
    excel.tcell("T_LONG_IDLE", '=IF(EXACT(%(computer)s,"Notebook"),0.1,0.15', T_LONG_IDLE, abbr='t_long_idle')
    excel.tcell("T_SHORT_IDLE", '=IF(EXACT(%(computer)s,"Notebook"),0.3,0.35', T_SHORT_IDLE, abbr='t_short_idle')

    excel.tcell("P_OFF", '=%(off)s', sysinfo.off, abbr='p_off')
    excel.tcell("P_SLEEP", '=%(sleep)s', sysinfo.sleep, abbr='p_sleep')
    excel.tcell("P_LONG_IDLE", '=%(long_idle)s', sysinfo.long_idle, abbr='p_long_idle')
    excel.tcell("P_SHORT_IDLE", '=%(short_idle)s', sysinfo.short_idle, abbr='p_short_idle')

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_LONG_IDLE * sysinfo.long_idle + T_SHORT_IDLE * sysinfo.short_idle) * 8760 / 1000

    excel.tcell("E_TEC", "=(%(t_off)s*%(p_off)s+%(t_sleep)s*%(p_sleep)s+%(t_long_idle)s*%(p_long_idle)s+%(t_short_idle)s*%(p_short_idle)s)*8760/1000", E_TEC)

    # P, EP, r, A
    P = sysinfo.cpu_core * sysinfo.cpu_clock
    excel.jump('H', 11)
    excel.tcell('P:', '=%(cpu_core)s*%(cpu_clock)s', P)

    EP_LABEL = "EP:"
    if sysinfo.ep:
        if sysinfo.diagonal >= 27:
            EP = 0.75
        else:
            EP = 0.3
    else:
        if sysinfo.computer_type == 1:
            EP = ""
            EP_LABEL = ""
        else:
            EP = 0
    excel.jump('H', 12)
    excel.cell('right', '=IF(EXACT(%(computer)s, "Desktop"), "", "EP:")', EP_LABEL)
    excel.jump('I', 12)
    excel.cell('EP', '=IF(EXACT(%(computer)s, "Desktop"), "", IF(EXACT(%(ep_display)s, "Yes"), IF(%(diagonal)s >= 27, 0.75, 0.3), 0))', EP)

    r_label = "r:"
    if sysinfo.computer_type != 1:
        r = 1.0 * sysinfo.width * sysinfo.height / 1000000
    else:
        if sysinfo.computer_type == 1:
            r = ""
            r_label = ""
        else:
            r = 0
    excel.jump('H', 13)
    excel.cell('right', '=IF(EXACT(%(computer)s, "Desktop"), "", "r:"', r_label)
    excel.jump('I', 13)
    excel.cell('r', '=IF(EXACT(%(computer)s, "Desktop"), "", %(width)s * %(height)s / 1000000)', r)

    A_LABEL = "A:"
    if sysinfo.computer_type != 1:
        A =  1.0 * sysinfo.diagonal * sysinfo.diagonal * sysinfo.width * sysinfo.height / (sysinfo.width ** 2 + sysinfo.height ** 2)
    else:
        if sysinfo.computer_type == 1:
            A = ""
            A_LABEL = ""
        else:
            A = 0
    excel.jump('H', 14)
    excel.cell('right', '=IF(EXACT(%(computer)s, "Desktop"), "", "A:"', A_LABEL)
    excel.jump('I', 14)
    excel.ncell(2, 1, 'A:', '=IF(EXACT(%(computer)s, "Desktop"), "", %(diagonal)s*%(diagonal)s*%(width)s*%(height)s/(%(width)s*%(width)s+%(height)s*%(height)s)', A)

    excel.jump('F', 11)

    # TEC_BASE
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

    if sysinfo.computer_type == 3:
        excel.tcell("TEC_BASE", '=IF(EXACT(%(gpu_type)s,"Discrete"), IF(%(P:)s>9, 18, IF(AND(%(P:)s<=9, %(P:)s>2), 16, 14)), IF(%(P:)s>8, 28, IF(AND(%(P:)s<=8, %(P:)s>5.2), 24, IF(AND(%(P:)s<=5.2, %(P:)s>2), 22, 14))))', TEC_BASE)
    else:
        excel.tcell("TEC_BASE", '=IF(EXACT(%(gpu_type)s,"Discrete"), IF(%(P:)s>9, 135, IF(AND(%(P:)s<=9, %(P:)s>3), 115, 69)), IF(%(P:)s>7, 135, IF(AND(%(P:)s<=7, %(P:)s>6), 120, IF(AND(%(P:)s<=6, %(P:)s>3), 112, 69))))', TEC_BASE)

    # TEC_MEMORY
    TEC_MEMORY = 0.8 * sysinfo.mem_size
    excel.tcell("TEC_MEMORY", '=%(memory)s*0.8', TEC_MEMORY)

    # TEC_GRAPHICS
    if sysinfo.discrete:
        if sysinfo.computer_type == 3:
            TEC_GRAPHICS = 14
        else:
            TEC_GRAPHICS = 36
    else:
        TEC_GRAPHICS = 0
    if sysinfo.computer_type == 3:
        excel.tcell("TEC_GRAPHICS", '=IF(EXACT(%(gpu_type)s, "Discrete"), IF(EXACT(%(gpu_category)s, "%(g1)s"), 14, IF(EXACT(%(gpu_category)s, "%(g2)s"), 20, IF(EXACT(%(gpu_category)s, "%(g3)s"), 26, IF(EXACT(%(gpu_category)s, "%(g4)s"), 32, IF(EXACT(%(gpu_category)s, "%(g5)s"), 42, IF(EXACT(%(gpu_category)s, "%(g6)s"), 48, 60)))))), 0)', TEC_GRAPHICS)
    else:
        excel.tcell("TEC_GRAPHICS", '=IF(EXACT(%(gpu_type)s, "Discrete"), IF(EXACT(%(gpu_category)s, "%(g1)s"), 36, IF(EXACT(%(gpu_category)s, "%(g2)s"), 51, IF(EXACT(%(gpu_category)s, "%(g3)s"), 64, IF(EXACT(%(gpu_category)s, "%(g4)s"), 83, IF(EXACT(%(gpu_category)s, "%(g5)s"), 105, IF(EXACT(%(gpu_category)s, "%(g6)s"), 115, 130)))))), 0)', TEC_GRAPHICS)
    
    # TEC_STORAGE
    if sysinfo.disk_num > 1:
        if sysinfo.computer_type == 3:
            TEC_STORAGE = 2.6 * (sysinfo.disk_num - 1)
        else:
            TEC_STORAGE = 26 * (sysinfo.disk_num - 1)
    else:
        TEC_STORAGE = 0
    if sysinfo.computer_type == 3:
        excel.tcell("TEC_STORAGE", '=IF(%(disk_number)s>1,2.6*(%(disk_number)s-1),0)', TEC_STORAGE)
    else:
        excel.tcell("TEC_STORAGE", '=IF(%(disk_number)s>1,26*(%(disk_number)s-1),0)', TEC_STORAGE)

    # TEC_INT_DISPLAY
    if sysinfo.computer_type == 3:
        TEC_INT_DISPLAY = 8.76 * 0.3 * (1+EP) * (2*r + 0.02*A)
    elif sysinfo.computer_type == 2:
        TEC_INT_DISPLAY = 8.76 * 0.35 * (1+EP) * (4*r + 0.05*A)
    else:
        TEC_INT_DISPLAY = 0
    excel.tcell("TEC_INT_DISPLAY", '=IF(EXACT(%(computer)s, "Notebook"), 8.76 * 0.3 * (1+%(EP)s) * (2*%(r)s + 0.02*%(A:)s), IF(EXACT(%(computer)s, "Integrated Desktop"), 8.76 * 0.35 * (1+%(EP)s) * (4*%(r)s + 0.05*%(A:)s), 0))', TEC_INT_DISPLAY)

    # TEC_SWITCHABLE
    if sysinfo.computer_type == 3:
        TEC_SWITCHABLE = 0
    elif sysinfo.switchable:
        TEC_SWITCHABLE = 0.5 * 36
    else:
        TEC_SWITCHABLE = 0
    excel.tcell('TEC_SWITCHABLE', '=IF(EXACT(%(computer)s, "Notebook"), 0, IF(EXACT(%(gpu_type)s, "Switchable"), 0.5 * 36, 0))', TEC_SWITCHABLE)

    # TEC_EEE
    if sysinfo.computer_type == 3:
        TEC_EEE = 8.76 * 0.2 * (0.1 + 0.3) * sysinfo.eee
    else:
        TEC_EEE = 8.76 * 0.2 * (0.15 + 0.35) * sysinfo.eee
    excel.tcell('TEC_EEE','=IF(EXACT(%(computer)s, "Notebook"), 8.76 * 0.2 * (0.1 + 0.3) * %(eee)s, 8.76 * 0.2 * (0.15 + 0.35) * %(eee)s)', TEC_EEE)

    # ALLOWANCE_PSU
    (column, row) = excel.position()
    excel.jump('D', 21)
    excel.ncell(4, 1, "Power Supply Efficiency Allowance requirements:", "None", ['None', 'Lower', 'Higher'], twin=True, abbr='psu')

    excel.jump(column, row)
    excel.tcell("ALLOWANCE_PSU", '=IF(OR(EXACT(%(computer)s, "Notebook"), EXACT(%(computer)s, "Desktop")), IF(EXACT(%(psu)s, "Higher"), 0.03, IF(EXACT(%(psu)s, "Lower"), 0.015, 0)), IF(EXACT(%(psu)s, "Higher"), 0.04, IF(EXACT(%(psu)s, "Lower"), 0.015, 0)))', 0)

    # E_TEC_MAX
    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE + TEC_INT_DISPLAY + TEC_SWITCHABLE + TEC_EEE
    excel.tcell('E_TEC_MAX', "=(1+%(ALLOWANCE_PSU)s)*(%(TEC_BASE)s+%(TEC_MEMORY)s+%(TEC_GRAPHICS)s+%(TEC_STORAGE)s+%(TEC_INT_DISPLAY)s+%(TEC_SWITCHABLE)s+%(TEC_EEE)s)", E_TEC_MAX)

    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    excel.right()
    excel.cell('center', '=IF(%(E_TEC)s<=%(E_TEC_MAX)s, "PASS", "FAIL")', RESULT)

    # ErP Lot 3
    excel.jump('D', 23)
    (column, row) = excel.position()
    excel.ncell(3, 1, "ErP Lot 3")
    if sysinfo.computer_type == 3:
        width = 3
    else:
        width = 4
    excel.shift(column, row, 3, 0)
    excel.ncell(width, 1, "from 1 July 2014")
    excel.shift(column, row, 3 + width, 0)
    excel.ncell(width, 1, "from 1 January 2016")

    early = ErPLot3_2014(sysinfo)
    late = ErPLot3_2016(sysinfo)

    excel.shift(column, row, 0, 1)
    (T_OFF, T_SLEEP, T_IDLE) = early.get_T_values()
    E_TEC = early.get_E_TEC()
    E_TEC_WOL = early.get_E_TEC_WOL()

    excel.tcell("T_OFF", '=IF(EXACT(%(computer)s,"Notebook"),0.6,0.55', T_OFF, abbr='t_off')
    excel.tcell("T_SLEEP", '=IF(EXACT(%(computer)s,"Notebook"),0.1,0.05', T_SLEEP, abbr='t_sleep')
    excel.tcell("T_IDLE", '=IF(EXACT(%(computer)s,"Notebook"),0.3,0.4', T_IDLE, abbr='t_idle')

    excel.tcell("P_OFF", '=%(off)s', sysinfo.off, abbr='p_off')
    excel.tcell("P_OFF_WOL", '=%(off_wol)s', sysinfo.off_wol, abbr='p_off_wol')
    excel.tcell("P_SLEEP", '=%(sleep)s', sysinfo.sleep, abbr='p_sleep')
    excel.tcell("P_SLEEP_WOL", '=%(sleep_wol)s', sysinfo.sleep_wol, abbr='p_sleep_wol')
    excel.tcell("P_IDLE", '=%(short_idle)s', sysinfo.short_idle, abbr='p_idle')

    excel.tcell("E_TEC", "=(%(t_off)s*%(p_off)s+%(t_sleep)s*%(p_sleep)s+%(t_idle)s*%(p_idle)s)*8760/1000", E_TEC)

    excel.tcell("E_TEC_WOL", "=(%(t_off)s*%(p_off_wol)s+%(t_sleep)s*%(p_sleep_wol)s+%(t_idle)s*%(p_idle)s)*8760/1000", E_TEC_WOL)

    if sysinfo.computer_type == 3:
        if sysinfo.sleep > 3 or sysinfo.sleep_wol > 3.7:
            RESULT = 'FAIL'
        else:
            RESULT = 'PASS'
    else:
        if sysinfo.sleep > 5 or sysinfo.sleep_wol > 5.7:
            RESULT = 'FAIL'
        else:
            RESULT = 'PASS'

    excel.cell('center', '=IF(EXACT(%(computer)s,"Notebook"), IF(OR(%(sleep)s>3.0,%(sleep_wol)s>3.7), "FAIL", "PASS"), IF(OR(%(sleep)s>5.0,%(sleep_wol)s>5.7), "FAIL", "PASS")', RESULT)

    if sysinfo.off > 1 or sysinfo.off_wol > 1.7:
        RESULT = 'FAIL'
    else:
        RESULT = 'PASS'

    excel.up()
    excel.right()
    excel.cell('center', '=IF(OR(%(off)s>1.0,%(off_wol)s>1.7), "FAIL", "PASS")', RESULT)

    excel.shift(column, row, 2, 1)
    excel.ncell(1, 3, "Category")
    excel.cell("field1", "TEC_BASE")
    excel.cell("field1", "TEC_MEMORY")
    excel.cell("field1", "TEC_GRAPHICS")
    excel.cell("field1", "TEC_TV_TUNER")
    excel.cell("field1", "TEC_AUDIO")
    excel.cell("field1", "TEC_STORAGE")
    excel.cell("result", "E_TEC_MAX")

    for step in (0, width):
        if step == 0:
            erplot3 = early
        else:
            erplot3 = late
        gfx_A_pos = None
        tuner_A_pos = None
        audio_A_pos = None
        storage_A_pos = None
        for i in range(width):
            cat = chr(ord('A') + i)
            meet = early.category(cat)
            if meet > 0:
                msg = cat
            else:
                msg = ''
            excel.shift(column, row, 3 + step + i, 1)
            if sysinfo.computer_type != 3:
                if cat == 'A':
                    formula = None
                elif cat == 'B':
                    formula = '=IF(AND(%(cpu_core)s>=2, %(memory)s>=2), "B", "")'
                elif cat == 'C':
                    formula = '=IF(AND(%(cpu_core)s>=3, OR(%(memory)s>=2, %(gpu_number)s>=1)), "C", "")'
                elif cat == 'D':
                    formula = '=IF(\
                                    AND(\
                                        %(cpu_core)s>=4,\
                                        OR(\
                                            %(memory)s>=4,\
                                            AND(\
                                                %(gpu_number)s>=1,\
                                                OR(\
                                                    AND(\
                                                        EXACT(%(gpu_category)s, "%(g3)s"),\
                                                        EXACT(%(gpu_width)s, "> 128-bit")\
                                                    ),\
                                                    EXACT(%(gpu_category)s, "%(g4)s"),\
                                                    EXACT(%(gpu_category)s, "%(g5)s"),\
                                                    EXACT(%(gpu_category)s, "%(g6)s"),\
                                                    EXACT(%(gpu_category)s, "%(g7)s")\
                                                )\
                                            )\
                                        )\
                                    ),\
                                    "D",\
                                    ""\
                                )'
                    formula = formula_strip(formula)
                else:
                    raise Exception('Should not be here.')
            else:
                if cat == 'A':
                    formula = None
                elif cat == 'B':
                    formula = '=IF(%(gpu_number)s>=1, "B", "")'
                elif cat == 'C':
                    formula = '=IF(\
                                    AND(\
                                        %(cpu_core)s >= 2,\
                                        %(memory)s >= 2,\
                                        %(gpu_number)s >= 1,\
                                        OR(\
                                            AND(\
                                                EXACT(%(gpu_category)s, "%(g3)s"),\
                                                EXACT(%(gpu_width)s, "> 128-bit")\
                                            ),\
                                            EXACT(%(gpu_category)s, "%(g4)s"),\
                                            EXACT(%(gpu_category)s, "%(g5)s"),\
                                            EXACT(%(gpu_category)s, "%(g6)s"),\
                                            EXACT(%(gpu_category)s, "%(g7)s")\
                                        )\
                                    ),\
                                    "C",\
                                    ""\
                                )'
                    formula = formula_strip(formula)
                else:
                    raise Exception('Should not be here.')
            if formula:
                excel.ncell(1, 3, cat, formula, msg)
            else:
                excel.ncell(1, 3, msg)

            # TEC_BASE for ErP Lot 3
            TEC_BASE = erplot3.get_TEC_BASE(cat)
            if cat == 'A':
                formula = None
            else:
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), ' + str(TEC_BASE) + ', "")'
            if meet > 0:
                msg = TEC_BASE
            else:
                msg = ''
            if formula:
                excel.cell('TEC_BASE', formula, msg)
            else:
                excel.cell('TEC_BASE', msg)

            # TEC_MEMORY for ErP Lot 3
            TEC_MEMORY = erplot3.get_TEC_MEMORY(cat)
            if sysinfo.computer_type != 3:
                if cat == 'A':
                    formula = '=IF(%(memory)s > 2, 1.0 * (%(memory)s - 2), 0)'
                elif cat == 'D':
                    formula = '=IF(EXACT(%(D)s, "D"), IF(%(memory)s > 4, 1.0 * (%(memory)s - 4), 0), "")'
                else:
                    formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), IF(%(memory)s > 2, 1.0 * (%(memory)s - 2), 0), "")'
            else:
                if cat == 'A':
                    formula = '=IF(%(memory)s > 4, 0.4 * (%(memory)s - 4), 0)'
                else:
                    formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), IF(%(memory)s > 4, 0.4 * (%(memory)s - 4), 0), "")'
            if meet > 0:
                msg = TEC_MEMORY
            else:
                msg = ''
            excel.cell('TEC_MEMORY', formula, msg)

            # TEC_GRAPHICS for ErP Lot 3
            if sysinfo.discrete_gpu_num > 0:
                TEC_GRAPHICS = erplot3.get_TEC_GRAPHICS('G1')
                for i in range(sysinfo.discrete_gpu_num - 1):
                    TEC_GRAPHICS = TEC_GRAPHICS + erplot3.additional_TEC_GRAPHICS('G1')
            else:
                TEC_GRAPHICS = 0

            if cat == 'A':
                if sysinfo.computer_type != 3:
                    if step == 0:
                        # Desktop from July 2014
                        formula = '=IF(EXACT(%(gpu_type)s, "Discrete"),\
                            IF(EXACT(%(gpu_category)s, "%(g1)s"), 34,\
                            IF(EXACT(%(gpu_category)s, "%(g2)s"), 54,\
                            IF(EXACT(%(gpu_category)s, "%(g3)s"), 69,\
                            IF(EXACT(%(gpu_category)s, "%(g4)s"), 100,\
                            IF(EXACT(%(gpu_category)s, "%(g5)s"), 133,\
                            IF(EXACT(%(gpu_category)s, "%(g6)s"), 166,\
                            IF(EXACT(%(gpu_category)s, "%(g7)s"), 225, 0\
                            ))))))))'
                    else:
                        # Desktop from January 1 2016
                        formula = '=IF(EXACT(%(gpu_type)s, "Discrete"),\
                            IF(EXACT(%(gpu_category)s, "%(g1)s"), 18,\
                            IF(EXACT(%(gpu_category)s, "%(g2)s"), 30,\
                            IF(EXACT(%(gpu_category)s, "%(g3)s"), 38,\
                            IF(EXACT(%(gpu_category)s, "%(g4)s"), 54,\
                            IF(EXACT(%(gpu_category)s, "%(g5)s"), 72,\
                            IF(EXACT(%(gpu_category)s, "%(g6)s"), 90,\
                            IF(EXACT(%(gpu_category)s, "%(g7)s"), 122, 0\
                            ))))))))'
                    for i in range(int(sysinfo.discrete_gpu_num) - 1):
                        dGfx = "dGfx%d" % (i+2)
                        if step == 0:
                            # Desktop from July 2014
                            formula = formula + ' + IF(EXACT(%(' + dGfx + ')s, "%(g1)s"), 20, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g2)s"), 32, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g3)s"), 41, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g4)s"), 59, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g5)s"), 78, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g6)s"), 98, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g7)s"), 133, 0' \
                                + ')))))))'
                        else:
                            # Desktop from January 1 2016
                            formula = formula + ' + IF(EXACT(%(' + dGfx + ')s, "%(g1)s"), 11, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g2)s"), 17, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g3)s"), 22, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g4)s"), 32, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g5)s"), 42, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g6)s"), 53, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g7)s"), 72, 0' \
                                + ')))))))'
                else:
                    if step == 0:
                        # Notebook from July 2014
                        formula = '=IF(EXACT(%(gpu_type)s, "Discrete"),\
                            IF(EXACT(%(gpu_category)s, "%(g1)s"), 12,\
                            IF(EXACT(%(gpu_category)s, "%(g2)s"), 20,\
                            IF(EXACT(%(gpu_category)s, "%(g3)s"), 26,\
                            IF(EXACT(%(gpu_category)s, "%(g4)s"), 37,\
                            IF(EXACT(%(gpu_category)s, "%(g5)s"), 49,\
                            IF(EXACT(%(gpu_category)s, "%(g6)s"), 61,\
                            IF(EXACT(%(gpu_category)s, "%(g7)s"), 113, 0\
                            ))))))))'
                    else:
                        # Notebook from January 1 2016
                        formula = '=IF(EXACT(%(gpu_type)s, "Discrete"),\
                            IF(EXACT(%(gpu_category)s, "%(g1)s"), 7,\
                            IF(EXACT(%(gpu_category)s, "%(g2)s"), 11,\
                            IF(EXACT(%(gpu_category)s, "%(g3)s"), 13,\
                            IF(EXACT(%(gpu_category)s, "%(g4)s"), 20,\
                            IF(EXACT(%(gpu_category)s, "%(g5)s"), 27,\
                            IF(EXACT(%(gpu_category)s, "%(g6)s"), 33,\
                            IF(EXACT(%(gpu_category)s, "%(g7)s"), 61, 0\
                            ))))))))'
                    for i in range(int(sysinfo.discrete_gpu_num) - 1):
                        dGfx = "dGfx%d" % (i+2)
                        if step == 0:
                            # Notebook from July 2014
                            formula = formula + ' + IF(EXACT(%(' + dGfx + ')s, "%(g1)s"), 7, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g2)s"), 12, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g3)s"), 15, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g4)s"), 22, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g5)s"), 29, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g6)s"), 36, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g7)s"), 66, 0' \
                                + ')))))))'
                        else:
                            # Notebook from January 1 2016
                            formula = formula + ' + IF(EXACT(%(' + dGfx + ')s, "%(g1)s"), 4, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g2)s"), 6, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g3)s"), 8, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g4)s"), 12, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g5)s"), 16, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g6)s"), 20, ' \
                                + 'IF(EXACT(%(' + dGfx + ')s, "%(g7)s"), 36, 0' \
                                + ')))))))'
            else:
                if gfx_A_pos == None:
                    gfx_A_pos = excel.pos["TEC_GRAPHICS"]
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), ' + gfx_A_pos + ', "")'
            formula = formula_strip(formula)
            if meet > 0:
                msg = TEC_GRAPHICS
            else:
                msg = ''
            excel.cell('TEC_GRAPHICS', formula, msg)

            # TEC_TV_TUNER for ErP Lot 3
            TEC_TV_TUNER = erplot3.get_TEC_TV_TUNER()
            if cat == 'A':
                formula = '=IF(EXACT(%(tvtuner)s, "Yes"), IF(EXACT(%(computer)s, "Notebook"), 2.1, 15), 0)'
            else:
                if tuner_A_pos == None:
                    tuner_A_pos = excel.pos["TEC_TV_TUNER"]
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), ' + tuner_A_pos + ', "")'
            if meet > 0:
                msg = TEC_TV_TUNER
            else:
                msg = ''
            excel.cell('TEC_TV_TUNER', formula, msg)

            # TEC_AUDIO for ErP Lot 3
            TEC_AUDIO = erplot3.get_TEC_AUDIO()
            if cat == 'A':
                formula = '=IF(EXACT(%(audio)s, "Yes"), IF(EXACT(%(computer)s, "Notebook"), 0, 15), 0)'
            else:
                if audio_A_pos == None:
                    audio_A_pos = excel.pos["TEC_AUDIO"]
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), ' + audio_A_pos + ', "")'
            if meet > 0:
                msg = TEC_AUDIO
            else:
                msg = ''
            excel.cell('TEC_AUDIO', formula, msg)

            # TEC_STORAGE for ErP Lot 3
            TEC_STORAGE = erplot3.get_TEC_STORAGE()
            if cat == 'A':
                formula = '=IF(%(disk_number)s > 1, IF(EXACT(%(computer)s, "Notebook"), 3 * (%(disk_number)s - 1), 25 * (%(disk_number)s - 1)), 0)'
            else:
                if storage_A_pos == None:
                    storage_A_pos = excel.pos["TEC_STORAGE"]
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), ' + storage_A_pos + ', "")'
            if meet > 0:
                msg = TEC_STORAGE
            else:
                msg = ''
            excel.cell('TEC_STORAGE', formula, msg)

            # E_TEC_MAX for ErP Lot 3
            E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_STORAGE + TEC_TV_TUNER + TEC_AUDIO + TEC_GRAPHICS
            if cat == 'A':
                formula = '=%(TEC_BASE)s + %(TEC_MEMORY)s + %(TEC_STORAGE)s + %(TEC_TV_TUNER)s + %(TEC_AUDIO)s + %(TEC_GRAPHICS)s'
            else:
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), %(TEC_BASE)s + %(TEC_MEMORY)s + %(TEC_STORAGE)s + %(TEC_TV_TUNER)s + %(TEC_AUDIO)s + %(TEC_GRAPHICS)s, "")'
            if meet > 0:
                msg = E_TEC_MAX
            else:
                msg = ''
            excel.cell('E_TEC_MAX', formula, msg)

            # Compute the result of ErP Lot 3
            if E_TEC > E_TEC_MAX or E_TEC_WOL > E_TEC_MAX:
                result = 'FAIL'
            else:
                result = 'PASS'
            if cat == 'A':
                formula = '=IF(AND(%(E_TEC)s <= %(E_TEC_MAX)s, %(E_TEC_WOL)s <= %(E_TEC_MAX)s), "PASS", "FAIL")'
            else:
                formula = '=IF(EXACT(%(' + cat +')s, "' + cat + '"), IF(AND(%(E_TEC)s <= %(E_TEC_MAX)s, %(E_TEC_WOL)s <= %(E_TEC_MAX)s), "PASS", "FAIL"), "")'
            if meet > 0:
                msg = result
            else:
                msg = ''
            excel.cell('center', formula, msg)

    excel.jump(column, row).down(12)

    if early.check_special_case():
        notebook_msg = "If discrete graphics card(s) providing total frame buffer bandwidths above 225 GB/s, use the requirement from 1 January 2016 instead."
        desktop_msg = "If discrete graphics card(s) providing total frame buffer bandwidths above 320 GB/s and a PSU with a rated output power of at least 1000W, use the requirement from 1 January 2016 instead."
        if early.computer_type == 3:
            msg = notebook_msg
        else:
            msg = desktop_msg
        excel.cell('warning', '=IF(EXACT(%(computer)s, "Notebook"), IF(AND(%(cpu_core)s >= 4, %(memory)s >=16), "' + notebook_msg + '", ""), IF(AND(%(cpu_core)s >= 6, %(memory)s >=16), "' + desktop_msg + '", "") )', msg)

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

    P_MAX = 0.28 * (sysinfo.max_power + sysinfo.disk_num * 5) + 8.76 * 0.2 * sysinfo.eee * (T_SLEEP + T_LONG_IDLE + T_SHORT_IDLE)
    sheet.write('A26', 'P_MAX', result)
    sheet.write('B26', '=0.28*(B11+B3*5) + 8.76*0.2*B4*(B22 + B23 + B24)', result_value, P_MAX)

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
