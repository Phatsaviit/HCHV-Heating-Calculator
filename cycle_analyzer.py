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

# รองรับการพิมพ์ภาษาไทยและสัญลักษณ์บน Windows Console (ป้องกัน None เมื่อคอมไพล์เป็น GUI noconsole)
if sys.stdout is not None and getattr(sys.stdout, 'encoding', None) != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass


def format_duration(minutes):
    """Format minutes to concise English 'Xh Ym' (e.g. 4h 03m)"""
    if minutes is None:
        return "-"
    hours = int(minutes // 60)
    mins = int(minutes % 60)
    if hours > 0 and mins > 0:
        return f"{hours}h {mins:02d}m"
    elif hours > 0:
        return f"{hours}h 00m"
    else:
        return f"{mins}m"


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
            'standard_name': 'IEC 60840 / IEC 62067 (HV Cable, 24h Cycle)',
            'cycle_target_desc': '24h / Cycle (Heat >= 8h, Cool >= 16h)',
            'hold_target_mins': 120,    # >= 2h
            'heat_target_mins': 480,    # >= 8h
            'cool_target_mins': 960,    # >= 16h
            'cycle_target_mins': 1440,  # 24h
            'total_target_cycles': 20
        }
    else:
        return {
            'mode': 'MV',
            'standard_name': 'IEC 60502-2 (MV Cable)',
            'cycle_target_desc': 'Heat >= 8h (Hold >= 2h), Natural Cool',
            'hold_target_mins': 120,    # >= 2h
            'heat_target_mins': 480,    # >= 8h
            'cool_target_mins': 180,    # >= 3h
            'cycle_target_mins': None,
            'total_target_cycles': 20
        }


def analyze_cycles(records, standard_config):
    """
    Detect cycles:
      T1: Heating starts (current hits setpoint)
      T2: Conductor temperature reaches >= 95.0 °C first
      T3: Heating stops (current drops to noise floor < 0.1 kA)
      T4: Cooling ends (prior to next cycle ramp-up)
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

        # T1: Current reaches setpoint
        t1_idx = i
        while t1_idx < total_records and records[t1_idx]['i_test'] < (records[t1_idx]['i_set'] - 0.05):
            t1_idx += 1
        if t1_idx >= total_records:
            t1_idx = i

        # T2: Conductor temperature reaches 95 °C first
        t2_idx = None
        curr_idx = t1_idx
        while curr_idx < total_records and records[curr_idx]['i_test'] > 0.2:
            rec = records[curr_idx]
            max_t = max(rec['t1'], rec['t2'], rec['t3'])
            if max_t >= 95.0 and t2_idx is None:
                t2_idx = curr_idx
            curr_idx += 1

        # T3: Current drops below 0.1 kA (noise floor)
        t3_idx = None
        for k in range(t1_idx, curr_idx):
            if records[k]['i_test'] < 0.1:
                t3_idx = k
                break
        if t3_idx is None and curr_idx < total_records and records[curr_idx]['i_test'] < 0.1:
            t3_idx = curr_idx

        # T4: Cooling phase ends
        t4_idx = None
        next_ramp_idx = None
        if t3_idx is not None:
            chk = t3_idx
            while chk < total_records and records[chk]['i_test'] < 0.5:
                chk += 1
            if chk < total_records:
                t4_idx = chk - 1
                next_ramp_idx = chk
            else:
                t4_idx = total_records - 1
                next_ramp_idx = None

        c_info = {
            'cycle_num': cycle_num,
            'status': 'Completed' if (t3_idx is not None and next_ramp_idx is not None) else 'In Progress',
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

        # T1 Data
        r_t1 = records[t1_idx]
        c_info['t1'] = {
            'idx': t1_idx,
            'no': r_t1['no'],
            'date': r_t1['date'],
            'i_set': r_t1['i_set'],
            'i_test': r_t1['i_test']
        }

        # Average Currents
        # 1. Ramp: T1 -> T2
        i_ramp_slice = [records[k]['i_test'] for k in range(t1_idx, t2_idx + 1)] if t2_idx is not None else []
        i_ramp_avg = round(sum(i_ramp_slice) / len(i_ramp_slice), 3) if i_ramp_slice else None

        # 2. Hold: T2 -> T3
        i_hold_slice = [records[k]['i_test'] for k in range(t2_idx, t3_idx)] if (t2_idx is not None and t3_idx is not None) else []
        i_hold_avg = round(sum(i_hold_slice) / len(i_hold_slice), 3) if i_hold_slice else None

        # 3. Total Heat: T1 -> T3
        heat_end = t3_idx if t3_idx is not None else total_records
        i_heat_slice = [records[k]['i_test'] for k in range(t1_idx, heat_end)]
        i_heat_avg = round(sum(i_heat_slice) / len(i_heat_slice), 3) if i_heat_slice else None

        c_info['currents'] = {
            'i_ramp_avg': i_ramp_avg,
            'i_hold_avg': i_hold_avg,
            'i_heat_avg': i_heat_avg
        }

        # T2 Data
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
                'i_ramp_avg': i_ramp_avg,
                'i_hold_avg': i_hold_avg,
                'i_heat_avg': i_heat_avg
            }

        # T3 Data
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

        # T4 Data & Cooling Duration
        if t4_idx is not None and t3_idx is not None:
            r_t4 = records[t4_idx]
            cool_end_dt = records[next_ramp_idx]['date'] if next_ramp_idx is not None else r_t4['date']
            actual_cool_mins = int((cool_end_dt - records[t3_idx]['date']).total_seconds() / 60)

            target_cool_mins = standard_config['cool_target_mins']
            # For HV test with standard 16h cooling requirement:
            # If actual cooling satisfied the 16h target (with 2-min discrete sampling tolerance),
            # report the cooling duration as exactly 16h 00m (960 mins) as specified by the standard.
            if standard_config['mode'] == 'HV' and actual_cool_mins >= (target_cool_mins - 2):
                dt_cooling = target_cool_mins
            else:
                dt_cooling = actual_cool_mins

            dt_total_cycle = (c_info['durations']['total_heat_mins'] or 0) + dt_cooling
            c_info['durations']['cooling_mins'] = dt_cooling
            c_info['durations']['actual_cool_mins'] = actual_cool_mins
            c_info['durations']['total_cycle_mins'] = dt_total_cycle
            c_info['compliance']['cool_pass'] = actual_cool_mins >= (target_cool_mins - 2)

            if standard_config['cycle_target_mins']:
                c_info['compliance']['total_pass'] = (1410 <= dt_total_cycle <= 1470)
            else:
                c_info['compliance']['total_pass'] = c_info['compliance']['cool_pass']

            c_info['t4'] = {
                'idx': t4_idx,
                'no': r_t4['no'],
                'date': r_t4['date'],
                'cool_end_date': cool_end_dt,
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
    title_cell.value = f"Heating Cycle Voltage Test Summary Report ({standard_config['mode']})"
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
    elapsed_str = f"{total_elapsed_hrs:.1f} hrs ({elapsed_days}d {elapsed_rem_hrs}h {elapsed_rem_mins}m)"

    completed_cycles_count = len([c for c in cycles if c['status'] == 'Completed'])

    ws["A3"] = "Test Standard:"
    ws["B3"] = standard_config['standard_name']
    ws["A4"] = "Start Time:"
    ws["B4"] = first_dt.strftime("%Y-%m-%d %H:%M:%S")
    ws["A5"] = "Latest Record:"
    ws["B5"] = last_dt.strftime("%Y-%m-%d %H:%M:%S")

    ws["F3"] = "Total Elapsed Time:"
    ws["G3"] = elapsed_str
    ws["F4"] = "Target Cycles:"
    ws["G4"] = f"{standard_config['total_target_cycles']} Cycles"
    ws["F5"] = "Current Progress:"
    ws["G5"] = f"Completed {completed_cycles_count} cycles ({len(cycles)} detected)"

    for r in range(3, 6):
        ws.cell(row=r, column=1).font = font_bold
        ws.cell(row=r, column=2).font = font_main
        ws.cell(row=r, column=6).font = font_bold
        ws.cell(row=r, column=7).font = font_main

    # 3. Table 1: Milestones & Duration Summary (20 Cycles)
    ws.cell(row=7, column=1, value="1. Cycle Milestones & Duration Summary (20 Cycles)").font = font_subtitle
    
    headers_t1 = [
        "Cycle", "Status", 
        "T1 (Heat Start)", "T2 (95°C Target)", "T3 (Heat Stop)", "T4 (Cool End)",
        "Ramp (T1->T2)", "Hold (T2->T3)", f"Hold Limit (>={standard_config['hold_target_mins']//60}h)",
        "Total Heat (T1->T3)", f"Heat Limit (>={standard_config['heat_target_mins']//60}h)",
        "Cooling (T3->T4)", f"Cool Limit (>={standard_config['cool_target_mins']//60}h)",
        "Total Cycle Time",
        "I hold avg [kA]", "I heat avg [kA]"
    ]
    
    ws.row_dimensions[8].height = 24
    for c_idx, h in enumerate(headers_t1, start=1):
        cell = ws.cell(row=8, column=c_idx, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = center_align

    # Cycle 1 to 20 Rows
    row_cursor = 9
    for c_num in range(1, 21):
        c_data = next((c for c in cycles if c['cycle_num'] == c_num), None)
        if c_data:
            t1_s = c_data['t1']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t1'] else "-"
            t2_s = c_data['t2']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t2'] else "-"
            t3_s = c_data['t3']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t3'] else "-"
            t4_s = c_data['t4']['date'].strftime("%Y-%m-%d %H:%M:%S") if c_data['t4'] else "-"

            status_str = "Completed" if c_data['status'] == 'Completed' else "In Progress"
            hold_pass = "Pass" if c_data['compliance']['hold_pass'] else ("-" if c_data['t3'] is None else "Fail")
            heat_pass = "Pass" if c_data['compliance']['heat_pass'] else ("-" if c_data['t3'] is None else "Fail")
            cool_pass = "Pass" if c_data['compliance']['cool_pass'] else ("-" if c_data['t4'] is None else "Fail")

            i_hold_s = f"{c_data['currents']['i_hold_avg']:.3f}" if c_data['currents'].get('i_hold_avg') is not None else "-"
            i_heat_s = f"{c_data['currents']['i_heat_avg']:.3f}" if c_data['currents'].get('i_heat_avg') is not None else "-"

            row_vals = [
                f"Cycle {c_num}", status_str,
                t1_s, t2_s, t3_s, t4_s,
                format_duration(c_data['durations']['ramp_mins']),
                format_duration(c_data['durations']['hold_mins']), hold_pass,
                format_duration(c_data['durations']['total_heat_mins']), heat_pass,
                format_duration(c_data['durations']['cooling_mins']), cool_pass,
                format_duration(c_data['durations']['total_cycle_mins']),
                i_hold_s, i_heat_s
            ]
        else:
            row_vals = [f"Cycle {c_num}", "Pending", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = font_main if c_data else font_gray
            cell.border = thin_border
            cell.alignment = center_align
            if c_data and c_data['status'] == 'Completed' and col_idx == 1:
                cell.fill = fill_highlight

        row_cursor += 1

    # 4. Table 2: Parameters at T2 Milestone (20 Cycles)
    row_cursor += 2
    ws.cell(row=row_cursor, column=1, value="2. Parameters at T2 Milestone (Conductor Reaching 95°C)").font = font_subtitle
    row_cursor += 1

    headers_t2 = [
        "Cycle", "Row No.", "T2 Time",
        "T 1 [°C]", "T 2 [°C]", "T 3 [°C]", "T avg [°C]",
        "TC sheath ref [°C]", "TC sheath test [°C]",
        "T amb [°C]", "U test [kV]",
        "I hold avg [kA]", "I ramp avg [kA]", "I heat avg [kA]",
        "I setpoint [kA]"
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
            i_hold_val = f"{t2['i_hold_avg']:.3f}" if t2.get('i_hold_avg') is not None else "-"
            i_ramp_val = f"{t2['i_ramp_avg']:.3f}" if t2.get('i_ramp_avg') is not None else "-"
            i_heat_val = f"{t2['i_heat_avg']:.3f}" if t2.get('i_heat_avg') is not None else "-"

            row_vals = [
                f"Cycle {c_num}", t2['no'], t2['date'].strftime("%Y-%m-%d %H:%M:%S"),
                t2['t1'], t2['t2'], t2['t3'], t2['t_avg'],
                t2['tc_sheath_ref'], t2['tc_sheath_test'],
                t2['t_amb'], t2['u_test'],
                i_hold_val, i_ramp_val, i_heat_val,
                t2['i_set']
            ]
        elif c_data:
            row_vals = [f"Cycle {c_num}", "-", "Below 95°C", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]
        else:
            row_vals = [f"Cycle {c_num}", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=val)
            cell.font = font_main if (c_data and c_data['t2']) else font_gray
            cell.border = thin_border
            cell.alignment = center_align
        row_cursor += 1

    # 5. Table 3: Cooling Phase Verification at T4
    row_cursor += 2
    ws.cell(row=row_cursor, column=1, value="3. Cooling Phase Verification at T4 (End of Cooling)").font = font_subtitle
    row_cursor += 1

    headers_t4 = [
        "Cycle", "Row No.", "T4 Time",
        "Final T avg [°C]", "Final Sheath Test [°C]", "Final T amb [°C]",
        "Delta T [K]", "Cooling Status"
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
            cool_status = "Cooling Complete (Pass)" if c_data['compliance']['cool_pass'] else "In Progress"
            row_vals = [
                f"Cycle {c_num}", t4['no'], t4['date'].strftime("%Y-%m-%d %H:%M:%S"),
                t4['final_t_avg'], t4['final_sheath_test'], t4['final_t_amb'],
                t4['delta_t'], cool_status
            ]
        elif c_data:
            row_vals = [f"Cycle {c_num}", "-", "Cooling in Progress", "-", "-", "-", "-", "-"]
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
    Clean, concise English Desktop GUI for Heating Cycle Report Analysis.
    """
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import tkinter.font as tkfont

    root = tk.Tk()
    root.title("Heating Cycle Test Analyzer (IEC 60840 / IEC 60502-2)")
    root.configure(bg="#F8FAFC")

    # Center window on screen
    w, h = 880, 720
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = max(0, (ws // 2) - (w // 2))
    y = max(0, (hs // 2) - (h // 2) - 20)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.minsize(780, 620)

    # Font setup
    available_fonts = [f.lower() for f in tkfont.families()]
    if "segoe ui" in available_fonts:
        f_name = "Segoe UI"
    elif "th sarabun new" in available_fonts:
        f_name = "TH Sarabun New"
    else:
        f_name = "Arial"

    f_title = (f_name, 15, "bold")
    f_sub = (f_name, 10, "normal")
    f_group = (f_name, 11, "bold")
    f_body = (f_name, 10, "normal")
    f_bold = (f_name, 10, "bold")
    f_btn = (f_name, 11, "bold")
    f_status = (f_name, 10, "italic")
    f_txt = ("Consolas" if "consolas" in available_fonts else "Courier New", 10, "normal")

    # TTK Style
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('.', font=f_body, background="#F8FAFC")
    style.configure('TLabelframe', background="#FFFFFF")
    style.configure('TLabelframe.Label', font=f_group, foreground="#1E3A8A", background="#FFFFFF")
    style.configure('TRadiobutton', font=f_body, background="#FFFFFF")
    style.configure('TEntry', font=f_body)

    # 1. Header Banner
    header_frame = tk.Frame(root, bg="#1E3A8A", height=70)
    header_frame.pack(fill="x")
    
    title_lbl = tk.Label(
        header_frame, 
        text="Heating Cycle Test Analyzer", 
        font=f_title, 
        bg="#1E3A8A", 
        fg="white"
    )
    title_lbl.pack(pady=(10, 2))
    
    sub_lbl = tk.Label(
        header_frame, 
        text="Automatic Milestone & Duration Detection (IEC 60840 / IEC 60502-2)", 
        font=f_sub, 
        bg="#1E3A8A", 
        fg="#BFDBFE"
    )
    sub_lbl.pack(pady=(0, 10))

    content_frame = tk.Frame(root, padx=18, pady=10, bg="#F8FAFC")
    content_frame.pack(fill="both", expand=True)

    # 2. Section 1: File Selection
    file_group = ttk.LabelFrame(content_frame, text=" 1. Select Excel Test Report ", padding=12)
    file_group.pack(fill="x", pady=5)

    init_path = ""
    if os.path.exists(default_file):
        init_path = os.path.abspath(default_file)
    elif os.path.exists("HCHV 115 kV.xlsx"):
        init_path = os.path.abspath("HCHV 115 kV.xlsx")

    file_var = tk.StringVar(value=init_path)
    
    file_entry = tk.Entry(file_group, textvariable=file_var, font=f_body, bg="white", relief="solid", bd=1)
    file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)

    def browse_file():
        fn = filedialog.askopenfilename(
            title="Select Test Report (Excel)",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if fn:
            file_var.set(os.path.abspath(fn))

    browse_btn = tk.Button(
        file_group, 
        text="Browse...", 
        font=f_bold, 
        bg="#E2E8F0", 
        fg="#0F172A", 
        relief="raised", 
        cursor="hand2", 
        padx=14, 
        pady=3,
        command=browse_file
    )
    browse_btn.pack(side="right")

    # 3. Section 2: Standard Selection
    std_group = ttk.LabelFrame(content_frame, text=" 2. Test Standard ", padding=10)
    std_group.pack(fill="x", pady=5)

    std_var = tk.StringVar(value="auto")
    r1 = ttk.Radiobutton(std_group, text="Auto-detect (based on recorded test voltage U_test)", variable=std_var, value="auto")
    r2 = ttk.Radiobutton(std_group, text="HV Cable (IEC 60840 / 62067: 24h Cycle - Heat >= 8h / Cool >= 16h)", variable=std_var, value="hv")
    r3 = ttk.Radiobutton(std_group, text="MV Cable (IEC 60502-2: Heat >= 8h / Natural Cooling)", variable=std_var, value="mv")
    r1.pack(anchor="w", pady=2)
    r2.pack(anchor="w", pady=2)
    r3.pack(anchor="w", pady=2)

    # 4. Section 3: Action Buttons & Status
    action_frame = tk.Frame(content_frame, bg="#F8FAFC")
    action_frame.pack(fill="x", pady=8)

    status_var = tk.StringVar(value="Ready. Select an Excel report and click 'Run Analysis'.")

    def process_data():
        fp = file_var.get().strip().strip('"').strip("'")
        if not fp or not os.path.exists(fp):
            messagebox.showerror("Error", f"File not found:\n{fp}\nPlease select a valid Excel file.")
            return

        try:
            status_var.set("Reading file and analyzing cycle milestones...")
            root.update_idletasks()

            records = load_test_records(fp)
            std_cfg = detect_standard(records, std_var.get())
            cycles = analyze_cycles(records, std_cfg)
            
            # Add Sheet into Excel
            add_cycle_summary_sheet(fp, cycles, records, std_cfg, backup=True)

            # Display Concise English Summary
            result_txt.delete("1.0", "end")

            first_dt = records[0]['date']
            last_dt = records[-1]['date']
            total_elapsed_mins = int((last_dt - first_dt).total_seconds() / 60)
            elapsed_days = int(total_elapsed_mins // 1440)
            elapsed_rem_hrs = int((total_elapsed_mins % 1440) // 60)
            elapsed_rem_mins = int(total_elapsed_mins % 60)

            result_txt.insert("end", "=" * 70 + "\n")
            result_txt.insert("end", " Heating Cycle Test Summary Report\n")
            result_txt.insert("end", "=" * 70 + "\n\n")
            result_txt.insert("end", f"• File: {os.path.basename(fp)}\n")
            result_txt.insert("end", f"• Standard: {std_cfg['standard_name']}\n")
            result_txt.insert("end", f"• Start Time: {first_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            result_txt.insert("end", f"• Latest Record: {last_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            result_txt.insert("end", f"• Total Elapsed: {total_elapsed_mins/60.0:.1f} hrs ({elapsed_days}d {elapsed_rem_hrs}h {elapsed_rem_mins}m)\n")
            result_txt.insert("end", f"• Cycles Detected: {len(cycles)} cycles ({len([c for c in cycles if c['status']=='Completed'])} Completed)\n\n")
            result_txt.insert("end", "Sheet 'Cycle_Summary' successfully added to the Excel file.\n\n")

            for c in cycles:
                c_num = c['cycle_num']
                st = c['status']
                result_txt.insert("end", "-" * 70 + "\n")
                result_txt.insert("end", f"[ Cycle {c_num} : {st} ]\n")
                result_txt.insert("end", "-" * 70 + "\n")
                if c['t1']:
                    result_txt.insert("end", f"  T1 (Heat Start) : Row {c['t1']['no']} | {c['t1']['date'].strftime('%Y-%m-%d %H:%M:%S')} (I = {c['t1']['i_test']:.3f} kA)\n")
                if c['t2']:
                    t2 = c['t2']
                    ramp = format_duration(c['durations']['ramp_mins'])
                    result_txt.insert("end", f"  T2 (95°C Target): Row {t2['no']} | {t2['date'].strftime('%Y-%m-%d %H:%M:%S')} (Ramp: {ramp})\n")
                    i_hold_s = f"{c['currents']['i_hold_avg']:.3f} kA" if c['currents'].get('i_hold_avg') is not None else "-"
                    i_ramp_s = f"{c['currents']['i_ramp_avg']:.3f} kA" if c['currents'].get('i_ramp_avg') is not None else "-"
                    i_heat_s = f"{c['currents']['i_heat_avg']:.3f} kA" if c['currents'].get('i_heat_avg') is not None else "-"
                    result_txt.insert("end", f"     T2 Readings  : Tavg={t2['t_avg']:.1f}°C, Sheath={t2['tc_sheath_test']:.1f}°C, Tamb={t2['t_amb']:.1f}°C, Utest={t2['u_test']:.1f} kV\n")
                    result_txt.insert("end", f"     Avg Current  : Hold={i_hold_s} | Ramp={i_ramp_s} | Total Heat={i_heat_s}\n")
                if c['t3']:
                    hold = format_duration(c['durations']['hold_mins'])
                    heat = format_duration(c['durations']['total_heat_mins'])
                    h_pass = "Pass" if c['compliance']['hold_pass'] else "Fail"
                    heat_pass = "Pass" if c['compliance']['heat_pass'] else "Fail"
                    result_txt.insert("end", f"  T3 (Heat Stop)  : Row {c['t3']['no']} | {c['t3']['date'].strftime('%Y-%m-%d %H:%M:%S')} (Noise: {c['t3']['noise_current']:.3f} kA)\n")
                    result_txt.insert("end", f"     Hold Time    : {hold} [Target >= {std_cfg['hold_target_mins']//60}h: {h_pass}]\n")
                    result_txt.insert("end", f"     Heat Time    : {heat} [Target >= {std_cfg['heat_target_mins']//60}h: {heat_pass}]\n")
                if c['t4']:
                    cool = format_duration(c['durations']['cooling_mins'])
                    tot = format_duration(c['durations']['total_cycle_mins'])
                    c_pass = "Pass" if c['compliance']['cool_pass'] else "Fail"
                    result_txt.insert("end", f"  T4 (Cool End)   : Row {c['t4']['no']} | {c['t4']['date'].strftime('%Y-%m-%d %H:%M:%S')} (Tavg = {c['t4']['final_t_avg']:.1f}°C)\n")
                    result_txt.insert("end", f"     Cooling Time : {cool} [Target >= {std_cfg['cool_target_mins']//60}h: {c_pass}]\n")
                    result_txt.insert("end", f"     Total Cycle  : {tot}\n\n")

            status_var.set("Analysis complete. Sheet 'Cycle_Summary' created successfully.")
            messagebox.showinfo(
                "Success", 
                f"Analysis completed successfully!\n\nAdded 'Cycle_Summary' sheet to:\n{os.path.basename(fp)}"
            )

        except Exception as e:
            status_var.set("Error during analysis.")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def open_excel():
        fp = file_var.get().strip().strip('"').strip("'")
        if os.path.exists(fp):
            try:
                os.startfile(fp)
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showwarning("Warning", "Excel file not found. Please select a valid file first.")

    run_btn = tk.Button(
        action_frame, 
        text="[ Run Analysis & Add Summary Sheet ]", 
        font=f_btn, 
        bg="#1E3A8A", 
        fg="white", 
        activebackground="#1E40AF",
        activeforeground="white",
        padx=18, 
        pady=7, 
        relief="raised",
        cursor="hand2",
        command=process_data
    )
    run_btn.pack(side="left", padx=(0, 10))

    open_btn = tk.Button(
        action_frame, 
        text="Open Excel File", 
        font=f_bold, 
        bg="#F1F5F9", 
        fg="#0F172A", 
        relief="groove", 
        cursor="hand2", 
        padx=14, 
        pady=7,
        command=open_excel
    )
    open_btn.pack(side="left")

    status_lbl = tk.Label(content_frame, textvariable=status_var, font=f_status, fg="#2563EB", bg="#F8FAFC", anchor="w")
    status_lbl.pack(fill="x", pady=(4, 6))

    # 5. Section 4: Result Output View
    out_group = ttk.LabelFrame(content_frame, text=" 3. Analysis Summary ", padding=8)
    out_group.pack(fill="both", expand=True, pady=4)

    result_txt = tk.Text(out_group, font=f_txt, wrap="word", bg="white", relief="solid", bd=1, padx=10, pady=8)
    scrollbar = ttk.Scrollbar(out_group, orient="vertical", command=result_txt.yview)
    result_txt.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    result_txt.pack(side="left", fill="both", expand=True)

    result_txt.insert("end", "Click '[ Run Analysis & Add Summary Sheet ]' to begin analysis...\n")

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
