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

def generate_excel(sysinfo, version, output):
    if not output:
        return

    try:
        from xlsxwriter import Workbook
    except:
        warning("You need to install Python xlsxwriter module or you can not output Excel format file.")
        return

    book = Workbook(output)
    book.set_properties({'comments':"Created by Energy Tools %s from Canonical Ltd." % (version)})

    if sysinfo.product_type == 1:
        generate_excel_for_computers(book, sysinfo, version)
    elif sysinfo.product_type == 2:
        generate_excel_for_workstations(book, sysinfo, version)
    elif sysinfo.product_type == 3:
        generate_excel_for_small_scale_servers(book, sysinfo, version)
    elif sysinfo.product_type == 4:
        generate_excel_for_thin_clients(book, sysinfo, version)

    book.close()

def generate_excel_for_computers(book, sysinfo, version):
    sheet = book.add_worksheet()

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

    header = book.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'fg_color': '#CFE2F3'})
    left = book.add_format({'align': 'left'})
    right = book.add_format({'align': 'right'})
    center = book.add_format({'align': 'center'})
    field = book.add_format({
        'border': 1,
        'fg_color': '#F3F3F3'})
    field1 = book.add_format({
        'left': 1,
        'right': 1,
        'fg_color': '#F3F3F3'})
    fieldC = book.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F3F3F3'})
    field2 = book.add_format({
        'left': 1,
        'right': 1,
        'bottom': 1,
        'fg_color': '#F3F3F3'})
    value0 = book.add_format({'border': 1, 'fg_color': '#D9EAD3'})
    value = book.add_format({'border': 1})
    value1 = book.add_format({
        'left': 1,
        'right': 1})
    value1.set_num_format('0.00')
    value2 = book.add_format({
        'left': 1,
        'right': 1,
        'bottom': 1})
    value2.set_num_format('0.00')
    value3 = book.add_format({
        'left': 1,
        'right': 1})
    value3.set_num_format('0%')
    value4 = book.add_format({
        'left': 1,
        'right': 1,
        'bottom': 1})
    value4.set_num_format('0%')
    float2 = book.add_format({'border': 1})
    float2.set_num_format('0.00')
    float3 = book.add_format({
        'left': 1,
        'right': 1})
    float3.set_num_format('0.000')
    result = book.add_format({
        'border': 1,
        'fg_color': '#F4CCCC'})
    result_value = book.add_format({
        'border': 1,
        'fg_color': '#FFF2CC'})
    result_value.set_num_format('0.00')

    sheet.merge_range("A1:B1", "General", header)

    sheet.write("A2", "Product Type", field)
    if sysinfo.product_type == 1:
        sheet.write("B2", "Desktop, Integrated Desktop, and Notebook", value)

    sheet.write("A3", "Computer Type", field)
    if sysinfo.computer_type == 1:
        sheet.write("B3", "Desktop", value)
    elif sysinfo.computer_type == 2:
        sheet.write("B3", "Integrated Desktop", value)
    else:
        sheet.write("B3", "Notebook", value)

    sheet.write("A4", "CPU cores", field)
    sheet.write("B4", sysinfo.cpu_core, value)

    sheet.write("A5", "CPU clock (GHz)", field)
    sheet.write("B5", sysinfo.cpu_clock, float2)

    sheet.write("A6", "Memory size (GB)", field)
    sheet.write("B6", sysinfo.mem_size, value)

    sheet.write("A7", "Number of Hard Drives", field)
    sheet.write("B7", sysinfo.disk_num, value)

    sheet.write("A8", "IEEE 802.3az compliant Gigabit Ethernet", field)
    sheet.write("B8", sysinfo.eee, value)

    sheet.merge_range("A10:B10", "Graphics", header)

    sheet.write("A11", "Graphics Type", field)
    if sysinfo.switchable:
        sheet.write("B11", "Switchable", value)
    elif sysinfo.discrete:
        sheet.write("B11", "Discrete", value)
    else:
        sheet.write("B11", "Integrated", value)
    sheet.data_validation('B11', {
        'validate': 'list',
        'source': [
            'Integrated',
            'Switchable',
            'Discrete']})

    sheet.write("A12", "GPU Frame Buffer Width", field)

    if sysinfo.computer_type == 3:
        sheet.write("B12", "<= 64-bit", value0)
        sheet.data_validation('B12', {
            'validate': 'list',
            'source': [
                '<= 64-bit',
                '> 64-bit and <= 128-bit',
                '> 128-bit']})
    else:
        sheet.write("B12", "<= 128-bit", value0)
        sheet.data_validation('B12', {
            'validate': 'list',
            'source': [
                '<= 128-bit',
                '> 128-bit']})

    sheet.write("A13", "Graphics Category", field)
    sheet.write("B13", "G1 (FB_BW <= 16)", value0)
    sheet.data_validation('B13', {
        'validate': 'list',
        'source': [
            'G1 (FB_BW <= 16)',
            'G2 (16 < FB_BW <= 32)',
            'G3 (32 < FB_BW <= 64)',
            'G4 (64 < FB_BW <= 96)',
            'G5 (96 < FB_BW <= 128)',
            'G6 (FB_BW > 128; Frame Buffer Data Width < 192 bits)',
            'G7 (FB_BW > 128; Frame Buffer Data Width >= 192 bits)']})

    sheet.merge_range("A15:B15", "Power Consumption", header)
    sheet.write("A16", "Off mode (W)", field)
    sheet.write("B16", sysinfo.off, float2)
    sheet.write("A17", "Sleep mode (W)", field)
    sheet.write("B17", sysinfo.sleep, float2)
    sheet.write("A18", "Long idle mode (W)", field)
    sheet.write("B18", sysinfo.long_idle, float2)
    sheet.write("A19", "Short idle mode (W)", field)
    sheet.write("B19", sysinfo.short_idle, float2)

    if sysinfo.computer_type != 1:
        sheet.merge_range("A21:B21", "Display", header)

        sheet.write("A22", "Enhanced-performance Integrated Display", field)
        sheet.write("B22", "No", value0)
        sheet.data_validation('B22', {
            'validate': 'list',
            'source': [
                'Yes',
                'No']})

        sheet.write("A23", "Physical Diagonal (inch)", field)
        sheet.write("B23", sysinfo.diagonal, value)

        sheet.write("A24", "Screen Width (px)", field)
        sheet.write("B24", sysinfo.width, value)

        sheet.write("A25", "Screen Height (px)", field)
        sheet.write("B25", sysinfo.height, value)

    sheet.merge_range("D21:G21", "Power Supply Efficiency Allowance requirements:", field)
    sheet.write("H21", "None", value0)
    sheet.data_validation('H21', {
        'validate': 'list',
        'source': [
            'None',
            'Lower',
            'Higher']})

    if sysinfo.computer_type == 3:
        sheet.merge_range("D1:I1", "Energy Star 5.2", header)
    else:
        sheet.merge_range("D1:J1", "Energy Star 5.2", header)

    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_IDLE) = (0.6, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_IDLE) = (0.55, 0.05, 0.4)

    sheet.write("D2", "T_OFF", field1)
    sheet.write("D3", "T_SLEEP", field1)
    sheet.write("D4", "T_IDLE", field2)
    sheet.write("E2", '=IF(EXACT(E,"Notebook"),0.6,0.55', value3, T_OFF)
    sheet.write("E3", '=IF(EXACT(E,"Notebook"),0.1,0.05', value3, T_SLEEP)
    sheet.write("E4", '=IF(EXACT(E,"Notebook"),0.3,0.4', value4, T_IDLE)

    sheet.write("D5", "P_OFF", field1)
    sheet.write("D6", "P_SLEEP", field1)
    sheet.write("D7", "P_IDLE", field2)
    sheet.write("E5", "=B16", value1, sysinfo.off)
    sheet.write("E6", "=B17", value1, sysinfo.sleep)
    sheet.write("E7", "=B19", value2, sysinfo.short_idle)

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_IDLE * sysinfo.short_idle) * 8760 / 1000
    sheet.write("D8", "E_TEC", result)
    sheet.write("E8", "=(E2*E5+E3*E6+E4*E7)*8760/1000", result_value, E_TEC)

    sheet.merge_range("F2:F3", "Category", fieldC)
    sheet.merge_range("G2:G3", "A", fieldC)
    sheet.merge_range("H2:H3", "B", fieldC)
    sheet.merge_range("I2:I3", "C", fieldC)
    if sysinfo.computer_type != 3:
        sheet.merge_range("J2:J3", "D", fieldC)

    if sysinfo.computer_type == 3:
        sheet.write("H2:H3", '=IF(EXACT(B11,"Discrete"), "B", "")', fieldC, "B")
        sheet.write("I2:I3", '=IF(AND(EXACT(B11,"Discrete"), EXACT(B12, "> 128-bit"), B4>=2, B6>=2), "C", "")', fieldC, "C")
    else:
        sheet.write("H2:H3", '=IF(AND(B4=2,B6>=2), "B", "")', fieldC, "B")
        sheet.write("I2:I3", '=IF(AND(B4>2,OR(B6>=2,EXACT(B11,"Discrete"))), "C", "")', fieldC, "C")
        sheet.write("J2:J3", '=IF(AND(B4>=4,OR(B6>=4,AND(EXACT(B11,"Discrete"),EXACT(B12,"> 128-bit")))), "D", "")', fieldC, "D")

    sheet.write("F4", "TEC_BASE", field1)
    sheet.write("F5", "TEC_MEMORY", field1)
    sheet.write("F6", "TEC_GRAPHICS", field1)
    sheet.write("F7", "TEC_STORAGE", field1)
    sheet.write("F8", "E_TEC_MAX", result)

    # Category A
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

        sheet.write("G4", TEC_BASE, value1)
        sheet.write("G5", "=IF(B6>4, 0.4*(B6-4), 0)", value1, TEC_MEMORY)
        sheet.write("G6", TEC_GRAPHICS, value1)
        sheet.write("G7", "=IF(B7>1, 3*(B7-1), 0)", value2, TEC_STORAGE)
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

        sheet.write("G4", TEC_BASE, value1)
        sheet.write("G5", "=IF(B6>2, 1.0*(B6-2), 0)", value1, TEC_MEMORY)
        sheet.write("G6", '=IF(EXACT(B12,"> 128-bit"), 50, 35)', value1, TEC_GRAPHICS)
        sheet.write("G7", "=IF(B7>1, 25*(B7-1), 0)", value2, TEC_STORAGE)

    E_TEC_MAX = TEC_BASE + TEC_MEMORY + TEC_GRAPHICS + TEC_STORAGE
    sheet.write("G8", "=G4+G5+G6+G7", result_value, E_TEC_MAX)

    if E_TEC <= E_TEC_MAX:
        RESULT = "PASS"
    else:
        RESULT = "FAIL"
    sheet.write("G9", '=IF(E8<=G8, "PASS", "FAIL")', center, RESULT)

    # Category B
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
        sheet.write("H4", '=IF(EXACT(H2,"B"), 53, "")', value1, TEC_BASE)
        sheet.write("H5", '=IF(EXACT(H2,"B"), G5, "")', value1, TEC_MEMORY)
        sheet.write("H6", '=IF(EXACT(H2,"B"), IF(EXACT(B12, "<= 64-bit"), 0, 3), "")', value1, TEC_GRAPHICS)
        sheet.write("H7", '=IF(EXACT(H2, "B"), G7, "")', value2, TEC_STORAGE)
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
        else:
            TEC_BASE = ""
            TEC_MEMORY = ""
            TEC_GRAPHICS = ""
            TEC_STORAGE = ""
            E_TEC_MAX = ""
            RESULT = ""
        sheet.write("H4", '=IF(EXACT(H2, "B"), 175, "")', value1, TEC_BASE)
        sheet.write("H5", '=IF(EXACT(H2,"B"), G5, "")', value1, TEC_MEMORY)
        sheet.write("H6", '=IF(EXACT(H2,"B"), IF(EXACT(B12, "<= 128-bit"), 35, 50), "")', value1, TEC_GRAPHICS)
        sheet.write("H7", '=IF(EXACT(H2, "B"), G7, "")', value2, TEC_STORAGE)

    sheet.write("H8", '=IF(EXACT(H2, "B"), H4+H5+H6+H7, "")', result_value, E_TEC_MAX)
    sheet.write("H9", '=IF(EXACT(H2, "B"),IF(E8<=H8, "PASS", "FAIL"), "")', center, RESULT)

    # Category C
    if sysinfo.computer_type == 3:
        # Notebook
        TEC_BASE = ""
        TEC_MEMORY = ""
        TEC_GRAPHICS = ""
        TEC_STORAGE = ""
        E_TEC_MAX = ""
        RESULT = ""
        sheet.write("I4", '=IF(EXACT(I2, "C"), 88.5, "")', value1, TEC_BASE)
        sheet.write("I5", '=IF(EXACT(I2, "C"), G5, "")', value1, TEC_MEMORY)
        sheet.write("I6", '=IF(EXACT(I2, "C"), 0, "")', value1, TEC_GRAPHICS)
        sheet.write("I7", '=IF(EXACT(I2, "C"), G7, "")', value2, TEC_STORAGE)
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
        sheet.write("I4", '=IF(EXACT(I2, "C"), 209, "")', value1, TEC_BASE)
        sheet.write("I5", '=IF(EXACT(I2, "C"), IF(B6>2, 1.0*(B6-2), 0), "")', value1, TEC_MEMORY)
        sheet.write("I6", '=IF(EXACT(I2, "C"), IF(EXACT(B12,"> 128-bit"), 50, 0), "")', value1, TEC_GRAPHICS)
        sheet.write("I7", '=IF(EXACT(I2, "C"), G7, "")', value2, TEC_STORAGE)

    sheet.write("I8", '=IF(EXACT(I2, "C"), I4+I5+I6+I7, "")', result_value, E_TEC_MAX)
    sheet.write("I9", '=IF(EXACT(I2, "C"), IF(E8<=I8, "PASS", "FAIL"), "")', center, RESULT)

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
        sheet.write("J4", '=IF(EXACT(J2, "D"), 234, "")', value1, TEC_BASE)
        sheet.write("J5", '=IF(EXACT(J2, "D"), IF(B6>4, B6-4, 0), "")', value1, TEC_MEMORY)
        sheet.write("J6", '=IF(EXACT(J2, "D"), IF(EXACT(B12,"> 128-bit"), 50, 0), "")', value1, TEC_GRAPHICS)
        sheet.write("J7", '=IF(EXACT(J2, "D"), G7, "")', value2, TEC_STORAGE)
        sheet.write("J8", '=IF(EXACT(J2, "D"), J4+J5+J6+J7, "")', result_value, E_TEC_MAX)
        sheet.write("J9", '=IF(EXACT(J2, "D"), IF(E8<=J8, "PASS", "FAIL"), "")', center, RESULT)

    sheet.merge_range("D10:G10", "Energy Star 6.0", header)

    if sysinfo.computer_type == 3:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.25, 0.35, 0.1, 0.3)
    else:
        (T_OFF, T_SLEEP, T_LONG_IDLE, T_SHORT_IDLE) = (0.45, 0.05, 0.15, 0.35)

    sheet.write("D11", "T_OFF", field1)
    sheet.write("E11", '=IF(EXACT(B3,"Notebook"),0.25,0.45', value3, T_OFF)
    sheet.write("D12", "T_SLEEP", field1)
    sheet.write("E12", '=IF(EXACT(B3,"Notebook"),0.35,0.05', value3, T_SLEEP)
    sheet.write("D13", "T_LONG_IDLE", field1)
    sheet.write("E13", '=IF(EXACT(B3,"Notebook"),0.1,0.15', value3, T_LONG_IDLE)
    sheet.write("D14", "T_SHORT_IDLE", field2)
    sheet.write("E14", '=IF(EXACT(B3,"Notebook"),0.3,0.35', value4, T_SHORT_IDLE)

    sheet.write("D15", "P_OFF", field1)
    sheet.write("E15", "=B16", value1, sysinfo.off)
    sheet.write("D16", "P_SLEEP", field1)
    sheet.write("E16", "=B17", value1, sysinfo.sleep)
    sheet.write("D17", "P_LONG_IDLE", field1)
    sheet.write("E17", "=B18", value1, sysinfo.long_idle)
    sheet.write("D18", "P_SHORT_IDLE", field1)
    sheet.write("E18", "=B19", value2, sysinfo.short_idle)

    E_TEC = (T_OFF * sysinfo.off + T_SLEEP * sysinfo.sleep + T_LONG_IDLE * sysinfo.long_idle + T_SHORT_IDLE * sysinfo.short_idle) * 8760 / 1000
    sheet.write("D19", "E_TEC", result)
    sheet.write("E19", "=(E11*E15+E12*E16+E13*E17+E14*E18)*8760/1000", result_value, E_TEC)

    sheet.write("F11", "ALLOWANCE_PSU", field1)
    sheet.write("G11", '=IF(OR(EXACT(B3, "Notebook"), EXACT(B3, "Desktop")), IF(EXACT(H21, "Higher"), 0.03, IF(EXACT(H21, "Lower"), 0.015, 0)), IF(EXACT(H21, "Higher"), 0.04, IF(EXACT(H21, "Lower"), 0.015, 0)))', float3)

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
