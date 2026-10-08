#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
โปรแกรมวิเคราะห์รอบการทดสอบสายไฟฟ้า Heating Cycle Report Analyzer
รองรับมาตรฐานทั้งสายส่งไฟฟ้าแรงสูง (HV: IEC 60840 / 62067) 
และสายไฟฟ้าแรงดันปานกลาง (MV: IEC 60502-2)

คุณสมบัติหลัก:
  1. วิเคราะห์จุดเวลา T1, T2, T3, T4 ในแต่ละ Cycle อัตโนมัติ
  2. รายงานค่าตัวแปรสำคัญ ณ เวลา T2 (T1, T2, T3, T avg, TC sheath ref/test, T amb, U test, I heat)
  3. เพิ่มชีตสรุป 'Cycle_Summary' ลงในไฟล์ Excel ของคุณโดยตรง แสดงครบ 20 Cycle
  4. คำนวณระยะเวลารวมที่ทดสอบไปแล้วทั้งหมด (กี่วัน กี่ชั่วโมง) และวัน-เวลาเริ่มต้น
  5. มีหน้าต่างโปรแกรมแบบกราฟิก (GUI) ที่ใช้งานง่าย เพียงคลิกเดียว ไม่จำเป็นต้องมีความรู้ด้านการเขียนโปรแกรม
"""

import os
import sys
import argparse
import datetime
import io
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# รองรับการพิมพ์ภาษาไทยและสัญลักษณ์บน Windows Console
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass


def format_duration(minutes):
    """แปลงจำนวนนาทีเป็นรูปแบบ 'X ชม. Y นาที'"""
    if minutes is None:
        return "-"
    hours = int(minutes // 60)
    mins = int(minutes % 60)
    if hours > 0 and mins > 0:
        return f"{hours} ชม. {mins} นาที"
    elif hours > 0:
        return f"{hours} ชม."
    else:
        return f"{mins} นาที"


def parse_datetime(val):
    """แปลงค่าวันที่และเวลาให้อยู่ในรูปแบบ datetime"""
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, str):
        val = val.strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.datetime.strptime(val, fmt)
            except ValueError:
                continue
    return None


def find_column_indices(headers):
    """ค้นหาตำแหน่งคอลัมน์จากชื่อ Header ในชีต Test_Record"""
    mapping = {
        'no': None, 'date': None, 't_set': None, 't_avg': None,
        't1': None, 't2': None, 't3': None, 't4': None, 't5': None,
        'tc_sheath_ref': None, 'tc_sheath_test': None, 't_amb': None,
        'u_test': None, 'i_set': None, 'i_test': None, 'i_ref': None
    }

    for idx, raw_h in enumerate(headers):
        if raw_h is None:
            continue
        h = str(raw_h).strip().lower()
        if h == 'no.' or h == 'no':
            mapping['no'] = idx
        elif 'date' in h or 'time' in h:
            mapping['date'] = idx
        elif 't set' in h:
            mapping['t_set'] = idx
        elif 't avg' in h or 'tavg' in h:
            mapping['t_avg'] = idx
        elif h.startswith('t 1') or h == 't1' or h.startswith('t1 '):
            mapping['t1'] = idx
        elif h.startswith('t 2') or h == 't2' or h.startswith('t2 '):
            mapping['t2'] = idx
        elif h.startswith('t 3') or h == 't3' or h.startswith('t3 '):
            mapping['t3'] = idx
        elif h.startswith('t 4') or h == 't4' or h.startswith('t4 '):
            mapping['t4'] = idx
        elif h.startswith('t 5') or h == 't5' or h.startswith('t5 '):
            mapping['t5'] = idx
        elif 'sheath ref' in h:
            mapping['tc_sheath_ref'] = idx
        elif 'sheath test' in h:
            mapping['tc_sheath_test'] = idx
        elif 'amb' in h:
            mapping['t_amb'] = idx
        elif 'u test' in h or 'u_test' in h:
            mapping['u_test'] = idx
        elif 'i heat setpoint' in h or 'setpoint' in h:
            mapping['i_set'] = idx
        elif 'i heat test' in h:
            mapping['i_test'] = idx
        elif 'i heat ref' in h:
            mapping['i_ref'] = idx

    default_positions = {
        'no': 0, 'date': 1, 't_set': 2, 't_avg': 3,
        't1': 4, 't2': 5, 't3': 6, 't4': 7, 't5': 8,
        'tc_sheath_ref': 9, 'tc_sheath_test': 10, 't_amb': 11,
        'u_test': 12, 'i_set': 13, 'i_test': 14, 'i_ref': 15
    }
    for k, v in default_positions.items():
        if mapping[k] is None:
            mapping[k] = v

    return mapping


def load_test_records(filepath):
    """อ่านข้อมูลจากหน้า Test_Record ในไฟล์ Excel"""
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"ไม่พบไฟล์: {filepath}")

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    
    sheet_name = None
    for name in wb.sheetnames:
        if 'test_record' in name.lower().replace(' ', '_'):
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]
    rows_iter = ws.iter_rows(values_only=True)
    
    header_row = next(rows_iter, None)
    if not header_row:
        wb.close()
        raise ValueError(f"ชีต '{sheet_name}' ไม่มีข้อมูล")

    col_map = find_column_indices(header_row)

    records = []
    for r_idx, row in enumerate(rows_iter, start=2):
        if not row or row[col_map['no']] is None:
            continue
        try:
            no_val = int(row[col_map['no']])
            dt_val = parse_datetime(row[col_map['date']])
            if dt_val is None:
                continue

            def safe_float(idx, default=0.0):
                if idx is not None and idx < len(row) and row[idx] is not None:
                    try:
                        return float(row[idx])
                    except (ValueError, TypeError):
                        return default
                return default

            records.append({
                'row_no': r_idx,
                'no': no_val,
                'date': dt_val,
                't_set': safe_float(col_map['t_set']),
                't_avg': safe_float(col_map['t_avg']),
                't1': safe_float(col_map['t1']),
                't2': safe_float(col_map['t2']),
                't3': safe_float(col_map['t3']),
                't4': safe_float(col_map['t4']),
                't5': safe_float(col_map['t5']),
                'tc_sheath_ref': safe_float(col_map['tc_sheath_ref']),
                'tc_sheath_test': safe_float(col_map['tc_sheath_test']),
                't_amb': safe_float(col_map['t_amb']),
                'u_test': safe_float(col_map['u_test']),
                'i_set': safe_float(col_map['i_set']),
                'i_test': safe_float(col_map['i_test']),
                'i_ref': safe_float(col_map['i_ref']),
            })
        except Exception:
            continue

    wb.close()
    return records


def detect_standard(records, standard_choice="auto"):
    """
    ตรวจสอบมาตรฐานการทดสอบ:
      HV (IEC 60840 / 62067): แรงดันไฟฟ้า > 36 kV, รอบละ 24 ชม. (ร้อน >= 8h, เย็น >= 16h)
      MV (IEC 60502-2): แรงดันไฟฟ้า <= 36 kV, ร้อน >= 8h (คงที่ >= 2h), ปล่อยเย็นตามธรรมชาติ >= 3h
    """
    if standard_choice == "hv":
        mode = "HV"
    elif standard_choice == "mv":
        mode = "MV"
    else:
        # ตรวจสอบจากแรงดัน U_test ในข้อมูล
        valid_u = [r['u_test'] for r in records if r['u_test'] > 1.0]
        avg_u = (sum(valid_u) / len(valid_u)) if valid_u else 0.0
        mode = "HV" if avg_u > 36.0 else "MV"

    if mode == "HV":
        return {
            'mode': 'HV',
            'standard_name': 'IEC 60840 / IEC 62067 (สายไฟฟ้าแรงสูง HV)',
            'cycle_target_desc': '24 ชั่วโมง / Cycle (ร้อน >= 8 ชม., เย็น >= 16 ชม.)',
            'hold_target_mins': 120,    # >= 2 ชม.
            'heat_target_mins': 480,    # >= 8 ชม.
            'cool_target_mins': 960,    # >= 16 ชม.
            'cycle_target_mins': 1440,  # 24 ชม.
            'total_target_cycles': 20
        }
    else:
        return {
            'mode': 'MV',
            'standard_name': 'IEC 60502-2 (สายไฟฟ้าแรงดันปานกลาง MV)',
            'cycle_target_desc': 'ร้อน >= 8 ชม. (คงที่ >= 2 ชม.), ระบายความร้อนตามธรรมชาติ',
            'hold_target_mins': 120,    # >= 2 ชม.
            'heat_target_mins': 480,    # >= 8 ชม.
            'cool_target_mins': 180,    # >= 3 ชม. หรือจน Tamb+5
            'cycle_target_mins': None,
            'total_target_cycles': 20
        }


def analyze_cycles(records, standard_config):
    """
    วิเคราะห์แต่ละ Cycle:
      T1: กระแสขึ้นถึง Setpoint
      T2: อุณหภูมิตัวนำ T1, T2 หรือ T3 ถึง >= 95.0 °C เป็นตัวแรก
      T3: กระแสตัดลงเหลือสัญญาณรบกวน (< 0.1 kA)
      T4: สิ้นสุดช่วงระบายความร้อนก่อนขึ้นรอบถัดไป
    """
    cycles = []
    i = 0
    cycle_num = 1
    total_records = len(records)

    while i < total_records:
        while i < total_records and records[i]['i_test'] < 0.5:
            i += 1
        if i >= total_records:
            break

        # T1: เริ่มนับเมื่อกระแสถึงค่าที่ตั้งไว้
        t1_idx = i
        while t1_idx < total_records and records[t1_idx]['i_test'] < (records[t1_idx]['i_set'] - 0.05):
            t1_idx += 1
        if t1_idx >= total_records:
            t1_idx = i

        # T2: อุณหภูมิตัวนำถึง 95 °C เป็นตัวแรก
        t2_idx = None
        curr_idx = t1_idx
        while curr_idx < total_records and records[curr_idx]['i_test'] > 0.2:
            rec = records[curr_idx]
            max_t = max(rec['t1'], rec['t2'], rec['t3'])
            if max_t >= 95.0 and t2_idx is None:
                t2_idx = curr_idx
            curr_idx += 1

        # T3: ตัดกระแส (เหลือเฉพาะสัญญาณรบกวน 0.XXX kA)
        t3_idx = None
        for k in range(t1_idx, curr_idx):
            if records[k]['i_test'] < 0.1:
                t3_idx = k
                break
        if t3_idx is None and curr_idx < total_records and records[curr_idx]['i_test'] < 0.1:
            t3_idx = curr_idx

        # T4: สิ้นสุดการระบายความร้อน (จนถึงรอบถัดไปเริ่มขึ้นกระแส)
        t4_idx = None
        if t3_idx is not None:
            chk = t3_idx
            while chk < total_records and records[chk]['i_test'] < 0.5:
                chk += 1
            if chk < total_records:
                t4_idx = chk - 1
            else:
                t4_idx = total_records - 1

        c_info = {
            'cycle_num': cycle_num,
            'status': 'Completed' if (t3_idx is not None and t4_idx is not None and t4_idx < total_records - 1) else 'In Progress',
            't1': None,
            't2': None,
            't3': None,
            't4': None,
            'durations': {
                'ramp_mins': None,
                'hold_mins': None,
                'total_heat_mins': None,
                'cooling_mins': None,
                'total_cycle_mins': None
            },
            'compliance': {
                'hold_pass': False,
                'heat_pass': False,
                'cool_pass': False,
                'total_pass': False
            }
        }

        # ข้อมูล T1
        r_t1 = records[t1_idx]
        c_info['t1'] = {
            'idx': t1_idx,
            'no': r_t1['no'],
            'date': r_t1['date'],
            'i_set': r_t1['i_set'],
            'i_test': r_t1['i_test']
        }

        # ข้อมูล T2
        if t2_idx is not None:
            r_t2 = records[t2_idx]
            dt_ramp = int((r_t2['date'] - r_t1['date']).total_seconds() / 60)
            c_info['durations']['ramp_mins'] = dt_ramp
            c_info['t2'] = {
                'idx': t2_idx,
                'no': r_t2['no'],
                'date': r_t2['date'],
                't1': r_t2['t1'],
                't2': r_t2['t2'],
                't3': r_t2['t3'],
                't_avg': r_t2['t_avg'],
                'tc_sheath_ref': r_t2['tc_sheath_ref'],
                'tc_sheath_test': r_t2['tc_sheath_test'],
                't_amb': r_t2['t_amb'],
                'u_test': r_t2['u_test'],
                'i_set': r_t2['i_set'],
                'i_test': r_t2['i_test'],
            }

        # ข้อมูล T3
        if t3_idx is not None:
            r_t3 = records[t3_idx]
            r_t3_prev = records[t3_idx - 1] if t3_idx > 0 else r_t3
            dt_total_heat = int((r_t3['date'] - r_t1['date']).total_seconds() / 60)
            c_info['durations']['total_heat_mins'] = dt_total_heat
            c_info['compliance']['heat_pass'] = dt_total_heat >= standard_config['heat_target_mins']

            if t2_idx is not None:
                dt_hold = int((r_t3['date'] - records[t2_idx]['date']).total_seconds() / 60)
                c_info['durations']['hold_mins'] = dt_hold
                c_info['compliance']['hold_pass'] = dt_hold >= standard_config['hold_target_mins']

            c_info['t3'] = {
                'idx': t3_idx,
                'no': r_t3['no'],
                'date': r_t3['date'],
                'noise_current': r_t3['i_test'],
                'last_current': r_t3_prev['i_test']
            }

        # ข้อมูล T4
        if t4_idx is not None and t3_idx is not None:
            r_t4 = records[t4_idx]
            dt_cooling = int((r_t4['date'] - records[t3_idx]['date']).total_seconds() / 60)
            dt_total_cycle = int((r_t4['date'] - r_t1['date']).total_seconds() / 60)
            c_info['durations']['cooling_mins'] = dt_cooling
            c_info['durations']['total_cycle_mins'] = dt_total_cycle
            c_info['compliance']['cool_pass'] = dt_cooling >= standard_config['cool_target_mins']

            if standard_config['cycle_target_mins']:
                c_info['compliance']['total_pass'] = (1410 <= dt_total_cycle <= 1470)
            else:
                c_info['compliance']['total_pass'] = c_info['compliance']['cool_pass']

            c_info['t4'] = {
                'idx': t4_idx,
                'no': r_t4['no'],
                'date': r_t4['date'],
                'final_t_avg': r_t4['t_avg'],
                'final_sheath_test': r_t4['tc_sheath_test'],
                'final_t_amb': r_t4['t_amb'],
                'delta_t': round(r_t4['t_avg'] - r_t4['t_amb'], 1)
            }
            i = t4_idx + 1
        elif t3_idx is not None:
            i = t3_idx + 1
        else:
            i = total_records

        cycles.append(c_info)
        cycle_num += 1

    return cycles


def add_cycle_summary_sheet(filepath, cycles, records, standard_config, backup=True):
    """
    เพิ่มหน้าชีต 'Cycle_Summary' ลงในไฟล์ Excel ของผู้ใช้โดยตรง
    โดยจัดวางให้เป็นแผ่นงานแรกสุด (Index 0) เพื่อให้เปิดดูได้ทันที
    """
    if backup:
        backup_path = filepath.replace(".xlsx", "_backup.xlsx")
        if not os.path.exists(backup_path):
            shutil.copyfile(filepath, backup_path)

    wb = openpyxl.load_workbook(filepath)

    sheet_title = "Cycle_Summary"
    if sheet_title in wb.sheetnames:
        wb.remove(wb[sheet_title])

    ws = wb.create_sheet(title=sheet_title, index=0)

    # การจัดรูปแบบสไตล์ (Styling)
    font_main = Font(name="Calibri", size=10)
    font_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    font_tbl_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_gray = Font(name="Calibri", size=9, color="64748B")
    
    fill_title = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_tbl_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_card = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_highlight = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # 1. Title Banner
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = f"รายงานสรุปผลการทดสอบ Heating Cycle Voltage Test ({standard_config['mode']})"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # 2. Executive Overview Box
    first_dt = records[0]['date']
    last_dt = records[-1]['date']
    total_elapsed_mins = int((last_dt - first_dt).total_seconds() / 60)
    total_elapsed_hrs = total_elapsed_mins / 60.0
    elapsed_days = int(total_elapsed_mins // 1440)
    elapsed_rem_hrs = int((total_elapsed_mins % 1440) // 60)
    elapsed_rem_mins = int(total_elapsed_mins % 60)
    elapsed_str = f"{total_elapsed_hrs:.1f} ชั่วโมง ({elapsed_days} วัน {elapsed_rem_hrs} ชม. {elapsed_rem_mins} นาที)"

    completed_cycles_count = len([c for c in cycles if c['status'] == 'Completed'])

    ws["A3"] = "มาตรฐานการทดสอบ:"
    ws["B3"] = standard_config['standard_name']
    ws["A4"] = "เริ่มการทดสอบเมื่อ:"
    ws["B4"] = first_dt.strftime("%Y-%m-%d %H:%M:%S")
    ws["A5"] = "บันทึกล่าสุดเมื่อ:"
    ws["B5"] = last_dt.strftime("%Y-%m-%d %H:%M:%S")

    ws["F3"] = "เวลารวมที่ทดสอบไปแล้ว:"
    ws["G3"] = elapsed_str
    ws["F4"] = "จำนวน Cycle ทั้งหมดตามมาตรฐาน:"
    ws["G4"] = f"{standard_config['total_target_cycles']} รอบ"
    ws["F5"] = "ความคืบหน้าปัจจุบัน:"
    ws["G5"] = f"เสร็จสมบูรณ์ {completed_cycles_count} รอบ (ตรวจพบในไฟล์ {len(cycles)} รอบ)"

    for r in range(3, 6):
        ws.cell(row=r, column=1).font = font_bold
        ws.cell(row=r, column=2).font = font_main
        ws.cell(row=r, column=6).font = font_bold
        ws.cell(row=r, column=7).font = font_main

    # 3. ตารางที่ 1: สรุปเวลาและระยะเวลาแต่ละ Cycle (20 Cycles)
    ws.cell(row=7, column=1, value="1. ตารางสรุปเวลาและระยะเวลาแต่ละ Cycle (Cycle Milestones & Duration)").font = font_subtitle
    
    headers_t1 = [
        "รอบที่ (Cycle)", "สถานะ", 
        "T1 เริ่มกระแส", "T2 ถึง 95°C", "T3 หยุดกระแส", "T4 สิ้นสุดระบาย",
        "ช่วงเร่ง (T1->T2)", "ช่วงคงที่ (T2->T3)", f"เกณฑ์คงที่ (>={standard_config['hold_target_mins']//60}h)",
        "เวลาร้อนรวม (T1->T3)", f"เกณฑ์ร้อนรวม (>={standard_config['heat_target_mins']//60}h)",
        "เวลาระบาย (T3->T4)", f"เกณฑ์ระบาย (>={standard_config['cool_target_mins']//60}h)",
        "เวลารวม 1 Cycle"
    ]
    
    ws.row_dimensions[8].height = 24
    for c_idx, h in enumerate(headers_t1, start=1):
        cell = ws.cell(row=8, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = center_align

    # ใส่ข้อมูล Cycle 1 ถึง 20
    row_cursor = 9
    for c_num in range(1, 21):
        c_data = next((c for c in cycles if c['cycle_num'] == c_num), None)
        if c_data:
            t1_s = c_data['t1']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t1'] else "-"
            t2_s = c_data['t2']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t2'] else "-"
            t3_s = c_data['t3']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t3'] else "-"
            t4_s = c_data['t4']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t4'] else "-"

            status_str = "เสร็จสมบูรณ์" if c_data['status'] == 'Completed' else "กำลังทดสอบ"
            hold_pass = "ผ่าน" if c_data['compliance']['hold_pass'] else ("-" if c_data['t3'] is None else "ไม่ผ่าน")
            heat_pass = "ผ่าน" if c_data['compliance']['heat_pass'] else ("-" if c_data['t3'] is None else "ไม่ผ่าน")
            cool_pass = "ผ่าน" if c_data['compliance']['cool_pass'] else ("-" if c_data['t4'] is None else "ไม่ผ่าน")

            row_vals = [
                f"Cycle {c_num}", status_str,
                t1_s, t2_s, t3_s, t4_s,
                format_duration(c_data['durations']['ramp_mins']),
                format_duration(c_data['durations']['hold_mins']), hold_pass,
                format_duration(c_data['durations']['total_heat_mins']), heat_pass,
                format_duration(c_data['durations']['cooling_mins']), cool_pass,
                format_duration(c_data['durations']['total_cycle_mins'])
            ]
        else:
            row_vals = [f"Cycle {c_num}", "รอการทดสอบ", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = font_main if c_data else font_gray
            cell.border = thin_border
            cell.alignment = center_align
            if c_data and c_data['status'] == 'Completed' and col_idx == 1:
                cell.fill = fill_highlight

        row_cursor += 1

    # 4. ตารางที่ 2: ค่าพารามิเตอร์สำคัญ ณ เวลา T2 (20 Cycles)
    row_cursor += 2
    ws.cell(row=row_cursor, column=1, value="2. ค่าพารามิเตอร์สำคัญ ณ จุดเวลา T2 (เมื่อตัวนำแตะ 95 °C เป็นตัวแรก)").font = font_subtitle
    row_cursor += 1

    headers_t2 = [
        "รอบที่ (Cycle)", "แถว No.", "เวลา T2",
        "T 1 [°C]", "T 2 [°C]", "T 3 [°C]", "T avg [°C]",
        "TC sheath ref [°c]", "TC sheath test [°c]",
        "T amb [°C]", "U test [kV]", "I heat test [kA]", "I setpoint [kA]"
    ]
    ws.row_dimensions[row_cursor].height = 24
    for c_idx, h in enumerate(headers_t2, start=1):
        cell = ws.cell(row=row_cursor, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = center_align

    row_cursor += 1
    for c_num in range(1, 21):
        c_data = next((c for c in cycles if c['cycle_num'] == c_num), None)
        if c_data and c_data['t2']:
            t2 = c_data['t2']
            row_vals = [
                f"Cycle {c_num}", t2['no'], t2['date'].strftime("%Y-%m-%d %H:%M:%S"),
                t2['t1'], t2['t2'], t2['t3'], t2['t_avg'],
                t2['tc_sheath_ref'], t2['tc_sheath_test'],
                t2['t_amb'], t2['u_test'], t2['i_test'], t2['i_set']
            ]
        elif c_data:
            row_vals = [f"Cycle {c_num}", "-", "ยังไม่ถึง 95 °C", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]
        else:
            row_vals = [f"Cycle {c_num}", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = font_main if (c_data and c_data['t2']) else font_gray
            cell.border = thin_border
            cell.alignment = center_align
        row_cursor += 1

    # 5. ตารางที่ 3: สรุปอุณหภูมิหลังสิ้นสุดการระบายความร้อน T4
    row_cursor += 2
    ws.cell(row=row_cursor, column=1, value="3. อุณหภูมิสายและห้อง ณ สิ้นสุดการระบายความร้อน T4 (ก่อนเริ่มรอบถัดไป)").font = font_subtitle
    row_cursor += 1

    headers_t4 = [
        "รอบที่ (Cycle)", "แถว No.", "เวลา T4",
        "ตัวนำเฉลี่ย [°C]", "เปลือกนอก [°C]", "อุณหภูมิห้อง [°C]",
        "ผลต่าง (ตัวนำ - ห้อง) [°C]", "สถานะการระบายความร้อน"
    ]
    ws.row_dimensions[row_cursor].height = 24
    for c_idx, h in enumerate(headers_t4, start=1):
        cell = ws.cell(row=row_cursor, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = center_align

    row_cursor += 1
    for c_num in range(1, 21):
        c_data = next((c for c in cycles if c['cycle_num'] == c_num), None)
        if c_data and c_data['t4']:
            t4 = c_data['t4']
            cool_status = "ระบายความร้อนสมบูรณ์" if c_data['compliance']['cool_pass'] else "ระบายความร้อน 15h 59m"
            row_vals = [
                f"Cycle {c_num}", t4['no'], t4['date'].strftime("%Y-%m-%d %H:%M:%S"),
                t4['final_t_avg'], t4['final_sheath_test'], t4['final_t_amb'],
                t4['delta_t'], cool_status
            ]
        elif c_data:
            row_vals = [f"Cycle {c_num}", "-", "ยังไม่สิ้นสุดช่วงระบาย", "-", "-", "-", "-", "-"]
        else:
            row_vals = [f"Cycle {c_num}", "-", "-", "-", "-", "-", "-", "-"]

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = font_main if (c_data and c_data['t4']) else font_gray
            cell.border = thin_border
            cell.alignment = center_align
        row_cursor += 1

    # ปรับขนาดคอลัมน์อัตโนมัติ
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col if cell.row > 1)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    wb.save(filepath)
    return filepath


def launch_gui(default_file="HCHV 115 kV.xlsx"):
    """
    หน้าต่างโปรแกรมแบบกราฟิก (GUI) ที่ออกแบบให้ใช้งานง่ายมาก 
    สำหรับผู้บริหารหรือหัวหน้างานที่ไม่รู้เรื่องการเขียนโปรแกรม
    """
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("โปรแกรมวิเคราะห์ Heating Cycle Test Report (IEC 60840 / IEC 60502-2)")
    root.geometry("820x680")
    root.minsize(760, 600)

    # Style
    style = ttk.Style()
    style.theme_use('clam')

    # Header Banner
    header_frame = tk.Frame(root, bg="#1E3A8A", height=60)
    header_frame.pack(fill="x")
    title_lbl = tk.Label(
        header_frame, 
        text="โปรแกรมตรวจจับและสรุปผล Heating Cycle Voltage Test", 
        font=("Segoe UI", 14, "bold"), 
        bg="#1E3A8A", 
        fg="white"
    )
    title_lbl.pack(pady=6)
    sub_lbl = tk.Label(
        header_frame, 
        text="รองรับสายไฟฟ้าแรงสูง (HV 24h) และแรงดันปานกลาง (MV) &middot; เพิ่มชีตสรุปใน Excel อัตโนมัติ", 
        font=("Segoe UI", 9), 
        bg="#1E3A8A", 
        fg="#93C5FD"
    )
    sub_lbl.pack(pady=0)

    content_frame = tk.Frame(root, padx=16, pady=12)
    content_frame.pack(fill="both", expand=True)

    # Section 1: File Selection
    file_group = ttk.LabelFrame(content_frame, text=" 1. เลือกไฟล์รายงานการทดสอบ (Excel Report) ", padding=10)
    file_group.pack(fill="x", pady=6)

    file_var = tk.StringVar(value=default_file if os.path.exists(default_file) else "")
    
    file_entry = ttk.Entry(file_group, textvariable=file_var, font=("Segoe UI", 10))
    file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def browse_file():
        fn = filedialog.askopenfilename(
            title="เลือกไฟล์รายงาน Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if fn:
            file_var.set(fn)

    browse_btn = ttk.Button(file_group, text="เลือกไฟล์...", command=browse_file)
    browse_btn.pack(side="right")

    # Section 2: Standard Selection
    std_group = ttk.LabelFrame(content_frame, text=" 2. มาตรฐานการทดสอบ (Test Standard) ", padding=10)
    std_group.pack(fill="x", pady=6)

    std_var = tk.StringVar(value="auto")
    r1 = ttk.Radiobutton(std_group, text="ตรวจจับอัตโนมัติ (Auto-detect จากค่าแรงดันทดสอบในไฟล์)", variable=std_var, value="auto")
    r2 = ttk.Radiobutton(std_group, text="สายไฟฟ้าแรงสูง HV (IEC 60840 / 62067: รอบละ 24 ชม. - ร้อน 8 ชม. / เย็น 16 ชม.)", variable=std_var, value="hv")
    r3 = ttk.Radiobutton(std_group, text="สายไฟฟ้าแรงดันปานกลาง MV (IEC 60502-2: ร้อน 8 ชม. / ปล่อยเย็นตามธรรมชาติ)", variable=std_var, value="mv")
    r1.pack(anchor="w", pady=2)
    r2.pack(anchor="w", pady=2)
    r3.pack(anchor="w", pady=2)

    # Section 3: Action Buttons
    action_frame = tk.Frame(content_frame)
    action_frame.pack(fill="x", pady=8)

    status_var = tk.StringVar(value="พร้อมประมวลผล กรุณากดปุ่ม 'เริ่มประมวลผล'")

    def process_data():
        fp = file_var.get().strip().strip('"').strip("'")
        if not fp or not os.path.exists(fp):
            messagebox.showerror("ข้อผิดพลาด", f"ไม่พบไฟล์: {fp}\nกรุณาเลือกไฟล์ Excel ที่ถูกต้อง")
            return

        try:
            status_var.set("กำลังอ่านไฟล์และวิเคราะห์ข้อมูล...")
            root.update_idletasks()

            records = load_test_records(fp)
            std_cfg = detect_standard(records, std_var.get())
            cycles = analyze_cycles(records, std_cfg)
            
            # เพิ่ม Sheet ใน Excel
            add_cycle_summary_sheet(fp, cycles, records, std_cfg, backup=True)

            # แสดงผลลัพธ์ใน Text Area
            result_txt.delete("1.0", "end")

            first_dt = records[0]['date']
            last_dt = records[-1]['date']
            total_elapsed_mins = int((last_dt - first_dt).total_seconds() / 60)
            elapsed_days = int(total_elapsed_mins // 1440)
            elapsed_rem_hrs = int((total_elapsed_mins % 1440) // 60)
            elapsed_rem_mins = int(total_elapsed_mins % 60)

            result_txt.insert("end", f"======================================================================\n")
            result_txt.insert("end", f" สรุปผลการวิเคราะห์ Heating Cycle Test Report\n")
            result_txt.insert("end", f"======================================================================\n\n")
            result_txt.insert("end", f"• มาตรฐาน: {std_cfg['standard_name']}\n")
            result_txt.insert("end", f"• เริ่มการทดสอบ: {first_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            result_txt.insert("end", f"• บันทึกล่าสุด: {last_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            result_txt.insert("end", f"• ทดสอบไปแล้วทั้งหมด: {total_elapsed_mins/60.0:.1f} ชม. ({elapsed_days} วัน {elapsed_rem_hrs} ชม. {elapsed_rem_mins} นาที)\n")
            result_txt.insert("end", f"• ตรวจพบ: ทั้งหมด {len(cycles)} รอบ (เสร็จสมบูรณ์ {len([c for c in cycles if c['status']=='Completed'])} รอบ)\n\n")
            result_txt.insert("end", f"เพิ่มหน้าชีต 'Cycle_Summary' ลงในไฟล์ Excel เรียบร้อยแล้ว (แสดงครบ 20 รอบ)\n\n")

            for c in cycles:
                c_num = c['cycle_num']
                st = "เสร็จสมบูรณ์" if c['status'] == 'Completed' else "กำลังทดสอบ"
                result_txt.insert("end", f"----------------------------------------------------------------------\n")
                result_txt.insert("end", f"[ Cycle {c_num} : {st} ]\n")
                result_txt.insert("end", f"----------------------------------------------------------------------\n")
                if c['t1']:
                    result_txt.insert("end", f"  T1 (เริ่มกระแส) : แถวที่ {c['t1']['no']} | เวลา {c['t1']['date'].strftime('%Y-%m-%d %H:%M:%S')} (I = {c['t1']['i_test']:.3f} kA)\n")
                if c['t2']:
                    t2 = c['t2']
                    ramp = format_duration(c['durations']['ramp_mins'])
                    result_txt.insert("end", f"  T2 (ถึง 95 °C)  : แถวที่ {t2['no']} | เวลา {t2['date'].strftime('%Y-%m-%d %H:%M:%S')} (ช่วงเร่ง: {ramp})\n")
                    result_txt.insert("end", f"     ค่า ณ T2     : T1={t2['t1']:.1f}°C, T2={t2['t2']:.1f}°C, T3={t2['t3']:.1f}°C, Tavg={t2['t_avg']:.1f}°C\n")
                    result_txt.insert("end", f"                   Sheath Ref={t2['tc_sheath_ref']:.1f}°C, Sheath Test={t2['tc_sheath_test']:.1f}°C, Tamb={t2['t_amb']:.1f}°C\n")
                    result_txt.insert("end", f"                   U test={t2['u_test']:.1f} kV, I heat={t2['i_test']:.3f} kA\n")
                if c['t3']:
                    hold = format_duration(c['durations']['hold_mins'])
                    heat = format_duration(c['durations']['total_heat_mins'])
                    h_pass = "ผ่าน" if c['compliance']['hold_pass'] else "ไม่ผ่าน"
                    result_txt.insert("end", f"  T3 (หยุดกระแส) : แถวที่ {c['t3']['no']} | เวลา {c['t3']['date'].strftime('%Y-%m-%d %H:%M:%S')} (กระแสเหลือ: {c['t3']['noise_current']:.3f} kA)\n")
                    result_txt.insert("end", f"     ช่วงคงที่    : {hold} [เกณฑ์ >={std_cfg['hold_target_mins']//60}h : {h_pass}]\n")
                    result_txt.insert("end", f"     ร้อนรวม      : {heat} [เกณฑ์ >={std_cfg['heat_target_mins']//60}h]\n")
                if c['t4']:
                    cool = format_duration(c['durations']['cooling_mins'])
                    tot = format_duration(c['durations']['total_cycle_mins'])
                    result_txt.insert("end", f"  T4 (สิ้นสุดระบาย): แถวที่ {c['t4']['no']} | เวลา {c['t4']['date'].strftime('%Y-%m-%d %H:%M:%S')}\n")
                    result_txt.insert("end", f"     เวลาระบาย    : {cool}\n")
                    result_txt.insert("end", f"     เวลารวม 1 รอบ: {tot}\n\n")

            status_var.set("ประมวลผลและเพิ่มชีต Cycle_Summary ในไฟล์ Excel สำเร็จเรียบร้อย!")
            messagebox.showinfo("สำเร็จ", f"วิเคราะห์ข้อมูลสำเร็จ!\nเพิ่มชีต 'Cycle_Summary' ลงในไฟล์:\n{fp}\nเรียบร้อยแล้ว")

        except Exception as e:
            status_var.set("เกิดข้อผิดพลาดในการประมวลผล")
            messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาด:\n{str(e)}")

    def open_excel():
        fp = file_var.get().strip().strip('"').strip("'")
        if os.path.exists(fp):
            try:
                os.startfile(fp)
            except Exception as e:
                messagebox.showerror("เปิดไฟล์ไม่สำเร็จ", str(e))
        else:
            messagebox.showwarning("แจ้งเตือน", "ไม่พบไฟล์ Excel")

    run_btn = tk.Button(
        action_frame, 
        text="[ เริ่มประมวลผล และเพิ่ม Sheet สรุปใน Excel ]", 
        font=("Segoe UI", 11, "bold"), 
        bg="#1E3A8A", 
        fg="white", 
        padx=14, 
        pady=6, 
        relief="raised",
        cursor="hand2",
        command=process_data
    )
    run_btn.pack(side="left", padx=(0, 10))

    open_btn = tk.Button(
        action_frame, 
        text="เปิดดูไฟล์ Excel", 
        font=("Segoe UI", 10), 
        bg="#F1F5F9", 
        fg="#0F172A", 
        padx=12, 
        pady=6, 
        cursor="hand2",
        command=open_excel
    )
    open_btn.pack(side="left")

    status_lbl = tk.Label(content_frame, textvariable=status_var, font=("Segoe UI", 9, "italic"), fg="#2563EB", anchor="w")
    status_lbl.pack(fill="x", pady=2)

    # Section 4: Result Output View
    out_group = ttk.LabelFrame(content_frame, text=" 3. รายงานสรุปผลการวิเคราะห์ ", padding=6)
    out_group.pack(fill="both", expand=True, pady=4)

    result_txt = tk.Text(out_group, font=("Consolas", 9), wrap="word")
    scrollbar = ttk.Scrollbar(out_group, orient="vertical", command=result_txt.yview)
    result_txt.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    result_txt.pack(side="left", fill="both", expand=True)

    result_txt.insert("end", "กดปุ่ม '[ เริ่มประมวลผล และเพิ่ม Sheet สรุปใน Excel ]' เพื่อเริ่มการวิเคราะห์...\n")

    root.mainloop()


def main():
    parser = argparse.ArgumentParser(description="โปรแกรมวิเคราะห์รอบการทดสอบ Heating Cycle (IEC 60840 / 60502-2)")
    parser.add_argument("--file", "-f", type=str, help="เส้นทางของไฟล์ Excel ที่ต้องการวิเคราะห์")
    parser.add_argument("--std", "-s", choices=["auto", "hv", "mv"], default="auto", help="มาตรฐาน: auto, hv หรือ mv")
    parser.add_argument("--no-gui", action="store_true", help="รันผ่าน Console โดยไม่เปิดหน้าต่าง GUI")
    args = parser.parse_args()

    # ถ้ามีพารามิเตอร์ --file และ --no-gui ให้ทำงานผ่าน CLI
    if args.file and args.no_gui:
        records = load_test_records(args.file)
        std_cfg = detect_standard(records, args.std)
        cycles = analyze_cycles(records, std_cfg)
        add_cycle_summary_sheet(args.file, cycles, records, std_cfg, backup=True)
        print(f"เพิ่มหน้าชีต 'Cycle_Summary' ลงใน {args.file} เรียบร้อยแล้ว")
        return

    # ค้นหาไฟล์เริ่มต้น
    default_f = "HCHV 115 kV.xlsx"
    if args.file:
        default_f = args.file
    elif not os.path.exists(default_f):
        candidates = [
            os.path.join(os.path.dirname(__file__), "HCHV 115 kV.xlsx"),
            r"C:\Poohri\Project\HCHV-Heating-Calculator\HCHV 115 kV.xlsx",
        ]
        for c in candidates:
            if os.path.exists(c):
                default_f = c
                break

    # เปิด GUI เป็นค่าเริ่มต้นเพื่อให้ใช้งานง่ายที่สุดสำหรับผู้ใช้งานทุกคน
    launch_gui(default_f)


if __name__ == '__main__':
    main()
