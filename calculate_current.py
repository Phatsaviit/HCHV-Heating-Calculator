"""
HCHV Current Calculator & Operating Model (IEC 60840 / Annex A)
Cable: 115 kV 1x800 SQ.MM. Copper XLPE (BCC Spec H-26-0017R1)
Loop length: ~14 meters

Calibrated against actual test cycles from HCHV_Test_Database.xlsx:
- R_th_effective = 0.6923 K.m/W
- Thermal time constant tau = 2.71 hours
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
    R_TH_EFF = 0.6923  # Calibrated thermal resistance (K.m/W)
    TAU_HOURS = 2.71   # Calibrated thermal time constant (hours)
    
    # --- Phase 2: Steady State Holding Current ---
    rac_target = calculate_conductor_rac(target_temp_c)
    delta_t_hold = target_temp_c - tamb_c
    power_hold = delta_t_hold / R_TH_EFF  # W/m
    i_hold_a = math.sqrt(power_hold / rac_target)
    
    # --- Phase 1: Ramp-Up Boost Current ---
    # Transient equation: T(t) - Tamb = (T_ss - Tamb) * (1 - exp(-t / tau))
    # We want T(ramp_target_hours) = 95.0 deg C (entry into standard band)
    target_ramp_temp = 95.0
    exp_factor = 1.0 - math.exp(-ramp_target_hours / TAU_HOURS)
    delta_t_ss_req = (target_ramp_temp - tamb_c) / exp_factor
    t_ss_conductor = tamb_c + delta_t_ss_req
    
    # Average conductor temperature during ramp-up phase for Rac estimate
    t_avg_ramp = (tamb_c + target_ramp_temp) / 2.0
    rac_ramp = calculate_conductor_rac(t_avg_ramp)
    power_boost = delta_t_ss_req / R_TH_EFF
    i_boost_a = math.sqrt(power_boost / rac_ramp)
    
    # Expected Sheath Surface Temperature in Steady State
    # Delta_internal / Delta_external = 1.81
    # T_sheath = Tamb + (T_target - Tamb) / (1 + 1.81)
    tc_sheath_expected = tamb_c + (target_temp_c - tamb_c) / (1.0 + 1.81)
    
    # Cooling prediction at 16 hours
    t_cool_16h = tamb_c + (target_temp_c - tamb_c) * math.exp(-16.0 / TAU_HOURS)
    
    return {
        "Tamb_C": round(tamb_c, 1),
        "Target_Conductor_Temp_C": round(target_temp_c, 1),
        "Ramp_Target_Hours": round(ramp_target_hours, 1),
        "Phase1_Boost_Current_A": round(i_boost_a, 0),
        "Phase1_Boost_Current_kA": round(i_boost_a / 1000.0, 3),
        "Phase2_Hold_Current_A": round(i_hold_a, 0),
        "Phase2_Hold_Current_kA": round(i_hold_a / 1000.0, 3),
        "Expected_Sheath_Temp_C": round(tc_sheath_expected, 1),
        "Expected_Cooled_Temp_16h_C": round(t_cool_16h, 1),
        "Cooling_Standard_Pass": t_cool_16h <= 30.0 or t_cool_16h <= (tamb_c + 10.0)
    }

def estimate_conductor_temp_from_sheath(tc_sheath_c, tamb_c):
    """
    Estimate internal conductor temperature directly from external sheath temperature (Method 2).
    Formula calibrated from test data: Tc = Tamb + 2.81 * (TC_sheath - Tamb) = TC_sheath + 1.81 * (TC_sheath - Tamb)
    """
    return round(tc_sheath_c + 1.81 * (tc_sheath_c - tamb_c), 1)

def main():
    parser = argparse.ArgumentParser(description="HCHV Heating Current Profile Calculator")
    parser.add_argument("--tamb", type=float, default=32.0, help="Ambient temperature in deg C (default: 32.0)")
    parser.add_argument("--target", type=float, default=97.5, help="Target conductor temp in deg C (default: 97.5)")
    parser.add_argument("--ramp_hrs", type=float, default=5.0, help="Ramp up hours to reach 95 deg C (default: 5.0)")
    args = parser.parse_args()

    res = predict_heating_profile(args.tamb, args.target, args.ramp_hrs)
    
    print("\n" + "="*65)
    print(" HCHV CABLE HEATING PROFILE RECOMMENDATION (IEC 60840)")
    print(" Cable: BCC 115 kV 1x800 SQ.MM. (14m Loop)")
    print("="*65)
    print(f" Ambient Temperature (Tamb)     : {res['Tamb_C']} deg C")
    print(f" Target Conductor Temp (Tavg)    : {res['Target_Conductor_Temp_C']} deg C (Standard: 95 - 100 deg C)")
    print(f" Ramp-up Target Time (T1 -> T2)  : {res['Ramp_Target_Hours']} hours (Standard <= 6.0 hr)")
    print("-" * 65)
    print(f" [PHASE 1: RAMP-UP (0.0h to {res['Ramp_Target_Hours']}h)]")
    print(f"  --> SET CURRENT               : {res['Phase1_Boost_Current_A']:.0f} A ({res['Phase1_Boost_Current_kA']:.3f} kA)")
    print(f" [PHASE 2: STEADY HOLD ({res['Ramp_Target_Hours']}h to 8.0h)]")
    print(f"  --> SET CURRENT               : {res['Phase2_Hold_Current_A']:.0f} A ({res['Phase2_Hold_Current_kA']:.3f} kA)")
    print(f"  --> Expected Sheath Temp      : ~{res['Expected_Sheath_Temp_C']} deg C")
    print("-" * 65)
    print(" [PHASE 3: NATURAL COOLING (8.0h to 24.0h = 16 hours)]")
    print(f"  --> Stop current (0 A) -> Expected temp at 16h: {res['Expected_Cooled_Temp_16h_C']} deg C")
    print(f"  --> Conformance with IEC: {'PASSED (<= 30 deg C or <= Tamb+10 K)' if res['Cooling_Standard_Pass'] else 'CHECK'}")
    print("="*65 + "\n")

if __name__ == "__main__":
    main()
