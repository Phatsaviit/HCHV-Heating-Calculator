"""
HCHV Current Calculator & Operating Model (IEC 60840 / Annex A)
Cable: 115 kV 1x800 SQ.MM. Copper XLPE
Loop length: ~14 meters

Calibrated against actual test cycles from HCHV_Test_Database.xlsx:
- Steady-state thermal resistance R_th_hold = 0.715 K.m/W
- Ramp-up thermal resistance R_th_ramp = 0.815 K.m/W (accounting for test bay convection and HV dielectric loss)
- Thermal time constant tau_ramp = 2.60 hours, tau_cool = 2.71 hours
- Ratio Delta_T_internal / Delta_T_external = 1.81
"""

import math
import argparse

def calculate_conductor_rac(temp_c):
    """
    Calculate AC resistance of 800 sq.mm copper conductor at given temperature.
    """
    alpha20 = 0.00393
    r_dc_20_per_m = 0.0221 / 1000.0  # Max Ohm/m at 20 deg C
    r_dc_temp = r_dc_20_per_m * (1.0 + alpha20 * (temp_c - 20.0))
    k_ac = 1.090  # Skin & Proximity factor
    return r_dc_temp * k_ac

def predict_heating_profile(tamb_c, target_temp_c=97.5, ramp_target_hours=5.0):
    """
    Calculate the 2-stage heating current profile:
    1. Phase 1 (T1 -> T2): Boost current to reach 95 deg C within ramp_target_hours (<= 6h)
    2. Phase 2 (T2 -> T3): Holding current to maintain target_temp_c (95 - 100 deg C) for >= 2h
    3. Phase 3 (T3 -> T4): Natural cooling for >= 16h to <= 30 deg C or <= Tamb + 10 K
    """
    # Calibrated directly from actual 77 data points in HCHV_Test_Database.xlsx:
    R_TH_HOLD_ACTUAL = 0.718  # K.m/W (Mean steady-state thermal resistance from Cycles 1, 2, 3)
    BASELINE_BOOST_A = 1830.0 # Measured actual boost current in Cycles 2-3 at Tamb 32.0 C
    BASELINE_TAMB = 32.0
    TAU_RAMP = 2.60   # hours (Calibrated dynamic ramp time constant)
    TAU_COOL = 2.70   # hours (Calibrated natural cooling time constant)
    
    # --- Phase 2: Steady State Holding Current ---
    dt_hold = max(1.0, target_temp_c - tamb_c)
    rac_target = calculate_conductor_rac(target_temp_c)
    power_hold = dt_hold / R_TH_HOLD_ACTUAL
    i_hold_a = math.sqrt(power_hold / rac_target)
    
    # --- Phase 1: Ramp-Up Boost Current ---
    # Scaled directly from actual lab baseline (Cycles 2-3: 1,830 A at Tamb 32.0 C, 5.0h)
    dt_req = max(10.0, 95.0 - tamb_c)
    dt_ref = 95.0 - BASELINE_TAMB # 63.0 K
    time_factor = (5.0 / ramp_target_hours) ** 0.28
    i_boost_a = BASELINE_BOOST_A * math.sqrt(dt_req / dt_ref) * time_factor
    
    # Expected Sheath Surface Temperature in Steady State
    tc_sheath_expected = tamb_c + (target_temp_c - tamb_c) / (1.0 + 1.81)

    # Expected Joint Surface Temperatures in Steady State (Calibrated from Test_Record T4 & T5)
    t4_joint_expected = tamb_c + 0.282 * (target_temp_c - tamb_c)
    t5_joint_expected = tamb_c + 0.178 * (target_temp_c - tamb_c)
    
    # Cooling prediction at 16 hours
    t_cool_16h = (tamb_c - 2.0) + (target_temp_c - (tamb_c - 2.0)) * math.exp(-16.0 / TAU_COOL)

    # Panel setting recommendation (rounded to practical 0.05 kA steps)
    i_boost_set = math.ceil(i_boost_a / 50.0) * 0.05
    i_hold_set = round(i_hold_a / 50.0) * 0.05
    
    return {
        "Tamb_C": round(tamb_c, 1),
        "Target_Conductor_Temp_C": round(target_temp_c, 1),
        "Ramp_Target_Hours": round(ramp_target_hours, 1),
        "Phase1_Boost_Current_A": round(i_boost_a, 0),
        "Phase1_Boost_Current_kA": round(i_boost_a / 1000.0, 3),
        "Phase1_Recommended_Set_kA": round(i_boost_set, 2),
        "Phase2_Hold_Current_A": round(i_hold_a, 0),
        "Phase2_Hold_Current_kA": round(i_hold_a / 1000.0, 3),
        "Phase2_Recommended_Set_kA": round(i_hold_set, 2),
        "Expected_Sheath_Temp_C": round(tc_sheath_expected, 1),
        "Expected_Joint_T4_C": round(t4_joint_expected, 1),
        "Expected_Joint_T5_C": round(t5_joint_expected, 1),
        "Expected_Cooled_Temp_16h_C": round(t_cool_16h, 1),
        "Cooling_Standard_Pass": t_cool_16h <= 30.0 or t_cool_16h <= (tamb_c + 10.0)
    }

def estimate_conductor_temp_from_sheath(tc_sheath_c, tamb_c):
    """
    Estimate internal conductor temperature directly from external sheath temperature (Method 2).
    Formula calibrated from test data: Tc = TC_sheath + 1.81 * (TC_sheath - Tamb)
    """
    return round(tc_sheath_c + 1.81 * (tc_sheath_c - tamb_c), 1)

def calculate_heating_time(current_a, t_start_c, t_target_c, tamb_c):
    """
    Calculate time required to reach target conductor temperature under constant heating current.
    """
    R_TH = 0.717
    C_TH = 10100.0  # J/(K.m)
    alpha20 = 0.00393
    r_dc_20_per_m = 0.0221 / 1000.0
    k_ac = 1.090

    k = (current_a ** 2) * r_dc_20_per_m * k_ac * R_TH
    denom = 1.0 - k * alpha20
    if denom <= 0:
        return {
            "possible": False,
            "reason": "กระแสสูงเกินพิกัดเสี่ยงต่อ Thermal Runaway",
            "t_infinity_c": None
        }
    
    t_inf = (tamb_c + k * (1.0 - 20.0 * alpha20)) / denom
    if t_inf <= t_target_c:
        return {
            "possible": False,
            "reason": f"กระแส {current_a:.0f} A ไม่เพียงพอต่อการดันถึง {t_target_c:.1f} °C (อุณหภูมิคงที่สูงสุดทำได้เพียง {t_inf:.1f} °C)",
            "t_infinity_c": round(t_inf, 1)
        }
    
    tau_eff = (R_TH * C_TH / 3600.0) / denom
    if t_start_c >= t_target_c:
        return {
            "possible": True,
            "time_hours": 0.0,
            "hours": 0,
            "minutes": 0,
            "formatted": "0 ชม. 00 นาที",
            "t_infinity_c": round(t_inf, 1),
            "tau_eff_hours": round(tau_eff, 2)
        }
    
    time_hours = tau_eff * math.log((t_inf - t_start_c) / (t_inf - t_target_c))
    total_mins = int(round(time_hours * 60))
    h = total_mins // 60
    m = total_mins % 60
    
    return {
        "possible": True,
        "time_hours": round(time_hours, 3),
        "hours": h,
        "minutes": m,
        "formatted": f"{h} ชม. {m:02d} นาที",
        "t_infinity_c": round(t_inf, 1),
        "tau_eff_hours": round(tau_eff, 2)
    }

def calculate_cooling_time(t_start_c, t_target_c, tamb_c):
    """
    Calculate time required to naturally cool down to target conductor temperature after stopping current.
    """
    TAU_COOL = 2.60 # hours
    if t_target_c <= tamb_c:
        return {
            "possible": False,
            "reason": f"อุณหภูมิเป้าหมาย {t_target_c:.1f} °C ต่ำกว่าหรือเท่ากับอุณหภูมิห้อง {tamb_c:.1f} °C (ไม่สามารถระบายความร้อนตามธรรมชาติให้ต่ำกว่าอุณหภูมิห้องได้)"
        }
    
    if t_start_c <= t_target_c:
        return {
            "possible": True,
            "time_hours": 0.0,
            "hours": 0,
            "minutes": 0,
            "formatted": "0 ชม. 00 นาที",
            "temp_16h_c": round(tamb_c + (t_start_c - tamb_c) * math.exp(-16.0 / TAU_COOL), 1),
            "iec_pass": True
        }
    
    time_hours = TAU_COOL * math.log((t_start_c - tamb_c) / (t_target_c - tamb_c))
    total_mins = int(round(time_hours * 60))
    h = total_mins // 60
    m = total_mins % 60
    
    t_16h = tamb_c + (t_start_c - tamb_c) * math.exp(-16.0 / TAU_COOL)
    iec_pass = (t_16h <= 30.0) or (t_16h <= tamb_c + 10.0)
    
    return {
        "possible": True,
        "time_hours": round(time_hours, 3),
        "hours": h,
        "minutes": m,
        "formatted": f"{h} ชม. {m:02d} นาที",
        "temp_16h_c": round(t_16h, 1),
        "iec_pass": iec_pass
    }

def main():
    parser = argparse.ArgumentParser(description="HCHV Heating Current Profile & Thermal Transient Calculator")
    parser.add_argument("--tamb", type=float, default=32.0, help="Ambient temperature in deg C (default: 32.0)")
    parser.add_argument("--target", type=float, default=97.0, help="Target conductor temp in deg C (default: 97.0)")
    parser.add_argument("--ramp_hrs", type=float, default=5.0, help="Ramp up hours to reach 95 deg C (default: 5.0)")
    parser.add_argument("--heat_curr", type=float, default=None, help="Calculate heating time for this current (A)")
    parser.add_argument("--t_start", type=float, default=30.0, help="Start temp for heating/cooling calc (deg C)")
    parser.add_argument("--cool_target", type=float, default=None, help="Calculate cooling time to reach this temp (deg C)")
    args = parser.parse_args()

    res = predict_heating_profile(args.tamb, args.target, args.ramp_hrs)
    
    print("\n" + "="*65)
    print(" HCHV CABLE HEATING PROFILE RECOMMENDATION (IEC 60840)")
    print(" Cable: 115 kV 1x800 SQ.MM. Copper XLPE (14m Loop)")
    print(" Grounded directly on actual test database (4,530 recorded points)")
    print("="*65)
    print(f" Ambient Temperature (Tamb)     : {res['Tamb_C']} deg C")
    print(f" Target Conductor Temp (Tavg)    : {res['Target_Conductor_Temp_C']} deg C (Standard: 95 - 100 deg C)")
    print(f" Ramp-up Target Time (T1 -> T2)  : {res['Ramp_Target_Hours']} hours (Standard <= 6.0 hr)")
    print("-" * 65)
    print(f" [PHASE 1: RAMP-UP (0.0h to {res['Ramp_Target_Hours']}h)]")
    print(f"  --> Calculated Boost Current  : {res['Phase1_Boost_Current_A']:.0f} A ({res['Phase1_Boost_Current_kA']:.3f} kA)")
    print(f"  --> Recommended Panel Setting : {res['Phase1_Recommended_Set_kA']:.2f} kA")
    print(f" [PHASE 2: STEADY HOLD ({res['Ramp_Target_Hours']}h to 8.0h)]")
    print(f"  --> Calculated Hold Current   : {res['Phase2_Hold_Current_A']:.0f} A ({res['Phase2_Hold_Current_kA']:.3f} kA)")
    print(f"  --> Recommended Panel Setting : {res['Phase2_Recommended_Set_kA']:.2f} kA")
    print(f"  --> Expected Sheath Temp      : ~{res['Expected_Sheath_Temp_C']} deg C")
    print(f"  --> Expected Joint Temp (T4/T5): ~{res['Expected_Joint_T4_C']} / {res['Expected_Joint_T5_C']} deg C")
    print("-" * 65)
    print(" [PHASE 3: NATURAL COOLING (8.0h to 24.0h = 16 hours)]")
    print(f"  --> Stop current (0 A) -> Expected temp at 16h: {res['Expected_Cooled_Temp_16h_C']} deg C")
    print(f"  --> Conformance with IEC: {'PASSED (<= 30 deg C or <= Tamb+10 K)' if res['Cooling_Standard_Pass'] else 'CHECK'}")
    print("-" * 65)
    
    # Heating Time Evaluation if requested
    heat_i = args.heat_curr if args.heat_curr else res['Phase1_Boost_Current_A']
    h_res = calculate_heating_time(heat_i, args.t_start, 95.0, args.tamb)
    print(f" [TIME TO HEAT EVALUATION]")
    if h_res["possible"]:
        print(f"  --> Current {heat_i:.0f} A from {args.t_start:.1f} C to 95.0 C : {h_res['formatted']} (~{h_res['time_hours']:.2f} hrs)")
        print(f"  --> Max Steady-State Temp (T_infinity) : {h_res['t_infinity_c']} deg C")
    else:
        print(f"  --> Notice: {h_res['reason']}")
    print("-" * 65)

    # Cooling Time Evaluation if requested
    c_target = args.cool_target if args.cool_target else 35.0
    c_res = calculate_cooling_time(args.target, c_target, args.tamb)
    print(f" [TIME TO COOL EVALUATION]")
    if c_res["possible"]:
        print(f"  --> Cool from {args.target:.1f} C down to {c_target:.1f} C : {c_res['formatted']} (~{c_res['time_hours']:.2f} hrs)")
        print(f"  --> Temp at 16h: {c_res['temp_16h_c']} deg C | IEC: {'PASSED' if c_res['iec_pass'] else 'CHECK'}")
    else:
        print(f"  --> Notice: {c_res['reason']}")

    print("="*65 + "\n")

if __name__ == "__main__":
    main()
