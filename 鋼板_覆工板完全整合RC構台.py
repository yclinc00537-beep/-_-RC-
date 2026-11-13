import streamlit as st
import numpy as np
import math
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# =========================================================================================
# 1. 結構力學核心引擎 (v18.0 全參數動態連動版：ACI 318-19 Bischoff 開裂剛度 + 板/梁雙單剪)
# =========================================================================================
def get_statutory_min_days(span_max, support_type, code_mode):
    if code_mode != "自動依台灣規範 (跨距與支撐條件判定)":
        return int(code_mode.split(" ")[1])
    is_unshored = "無下方支撐" in support_type
    if span_max <= 3.0:
        return 7 if is_unshored else 4
    elif span_max <= 6.0:
        return 14 if is_unshored else 7
    else:
        return 21 if is_unshored else 14

def calculate_single_case(params, case_name):
    fc = params['fc']
    fy = params['fy']
    span_x = params['span_x'] 
    span_y = params['span_y'] 
    b_beam = params['b_beam']
    h_beam = params['h_beam']
    support_type = params['support_type']
    t_steel = params['t_steel']
    min_fc_ratio = params.get('min_fc_ratio', 75) / 100.0
    min_days_req = params.get('min_days_req', 21)
    
    # 讀取縱向鋼筋參數 (#4~#8, @10~20)
    long_bar_no = params['long_bar_no']
    long_bar_space = params['long_bar_space']
    bar_areas_all = {3: 0.713, 4: 1.267, 5: 1.986, 6: 2.865, 7: 3.871, 8: 5.067}
    as_long_single = bar_areas_all.get(long_bar_no, 1.986)
    As_per_meter = (as_long_single * 100.0) / float(long_bar_space)
    
    # 讀取柱頭幾何參數
    col_shape = params['col_shape']
    col_b = params['col_b']
    col_d = params['col_d']
    col_dia = params['col_dia']
    col_equiv_w = math.sqrt(col_b * col_d) if col_shape == "矩形柱 (b × d)" else col_dia
    
    # 大梁箍筋參數
    stirrup_legs = params['stirrup_legs'] 
    stirrup_no = params['stirrup_no']     
    stirrup_space = params['stirrup_space'] 
    
    days = [7, 10, 14, 21, 28]
    age_pct = {7: 0.70, 10: 0.78, 14: 0.85, 21: 0.93, 28: 1.00}
    
    a_b = bar_areas_all.get(stirrup_no, 1.267)
    Av = stirrup_legs * a_b
    
    beta_col = (max(col_b, col_d) / min(col_b, col_d)) if col_shape == "矩形柱 (b × d)" else 1.0
    
    # 📌【動態自重與荷重設定】：隨各區版厚 h_case 即時計算靜載重 (24 kgf/m²/cm + 100模板雜項)
    if case_name == "區域一：PC400 挖土機工作平台區":
        h_case = params['h_case_1']
        d_slab = h_case - 5.5
        w_dead = 24.0 * h_case + 100.0
        w_live = 1700.0 * 1.3  # 含動態衝擊係數 1.3
        has_stirrup = params['has_stirrup_1']
        alpha_s = 40.0
        
        mode = params['deck_mode_1']
        if mode == "單鋪 2層活動覆工板 (無銲接複合效益)":
            max_share_28 = 0.0
            cover_board_bonus = 0.10
            t_cm = 0.0
            steel_plate_info = "單鋪 2層活動覆工板 (僅發揮二維表面剛度均勻擴散)"
        elif mode == "單鋪 PL12mm 鋼板 (全面銲死錨碇複合斷面)":
            max_share_28 = 0.634
            cover_board_bonus = 0.0
            t_cm = t_steel / 10.0
            steel_plate_info = f"PL{t_steel}mm 鋼板銲死複合斷面 (28天硬化後分擔63.4%剪力)"
        else:
            max_share_28 = 0.634
            cover_board_bonus = 0.10
            t_cm = t_steel / 10.0
            steel_plate_info = f"PL{t_steel}mm 鋼板銲死 + 2層活動覆工板雙重均勻分散鋪面"
            
    elif case_name == "區域二：20T 土車行駛通道區":
        h_case = params['h_case_2']
        d_slab = h_case - 5.0
        w_dead = 24.0 * h_case + 100.0
        w_live = 1250.0 * 1.3  # 含動態衝擊係數 1.3
        alpha_s = 40.0
        
        max_share_28 = 0.50
        cover_board_bonus = 0.0
        has_stirrup = params['has_stirrup_2']
        t_cm = t_steel / 10.0
        steel_plate_info = f"PL{t_steel}mm 鋼板 / 車道動態抗滑移剛性鋪面"
        
    else:
        h_case = params['h_case_3']
        d_slab = h_case - 5.0
        w_dead = 24.0 * h_case + 100.0
        w_live = 350.0         # 常態施工活載重
        alpha_s = 30.0
        
        max_share_28 = 0.0    
        cover_board_bonus = 0.0
        has_stirrup = params['has_stirrup_3']
        t_cm = 0.0
        steel_plate_info = "無 (地下常態區邊平板依 ACI 318-19 αs=30 進行三面沖切折減檢核)"

    # 換算均佈使用荷重與因數化設計荷重 (kgf/cm²)
    w_service = (w_dead + w_live) / 10000.0
    w_factored = (1.4 * w_dead + 1.7 * w_live) / 10000.0

    # 柱頭臨界週長 b0 與核心面積計算
    if col_shape == "矩形柱 (b × d)":
        c1_eff = col_b + 2.0 * t_cm + d_slab
        c2_eff = col_d + 2.0 * t_cm + d_slab
        if case_name == "區域三：逆打地下一般平板 (常態無施工重載區)":
            b0_eff = 2.0 * (col_b + d_slab / 2.0) + (col_d + d_slab)
            core_area_cm2 = (col_b + d_slab / 2.0) * (col_d + d_slab)
        else:
            b0_eff = 2.0 * (c1_eff + c2_eff)
            core_area_cm2 = (col_b + d_slab) * (col_d + d_slab)
    else:
        d_eff_circle = col_dia + 2.0 * t_cm + d_slab
        if case_name == "區域三：逆打地下一般平板 (常態無施工重載區)":
            b0_eff = math.pi * (col_dia + d_slab) * 0.65
            core_area_cm2 = (math.pi / 4.0) * ((col_dia + d_slab) ** 2) * 0.65
        else:
            b0_eff = math.pi * d_eff_circle
            core_area_cm2 = (math.pi / 4.0) * ((col_dia + d_slab) ** 2)

    d_beam = h_beam - 8.0
    L_long = max(span_x, span_y) * 100.0
    L_short = min(span_x, span_y) * 100.0
    delta_limit = L_long / 240.0
    
    # ACI 318-19 尺寸效應折減係數 λs
    if d_slab <= 25.0:
        lambda_s_slab = 1.0
        size_effect_tag = f"d={d_slab}≤25cm(λs=1.0)"
    else:
        lambda_s_slab = math.sqrt(2.0 / (1.0 + 0.04 * d_slab))
        size_effect_tag = f"d={d_slab}>25cm(λs={round(lambda_s_slab,3)})"
        
    lambda_s_beam = 1.0 if d_beam <= 25.0 else math.sqrt(2.0 / (1.0 + 0.04 * d_beam))
    
    # 受拉鋼筋比 ρw
    rho_w_slab = As_per_meter / (100.0 * d_slab)
    rho_w_beam = max(0.008, min(0.025, rho_w_slab * 1.15))
    
    # 1. 雙向穿孔剪力原始總需求
    net_tributary_cm2 = max(10000.0, (span_x * span_y * 10000.0) - core_area_cm2)
    Vu_pure_total = w_factored * net_tributary_cm2

    # 2. 樓版單向剪力需求 (1m 寬短向版帶，距梁邊 d_slab 處臨界斷面)
    clear_span_short = max(100.0, L_short - b_beam)
    Vu_slab_1way_demand = w_factored * 100.0 * max(0.0, (clear_span_short / 2.0) - d_slab)

    # 3. 外圍大梁單向剪力需求 (梯形負擔面積 + 大梁自重，距柱面 d_beam 處臨界斷面)
    # 梯形等效均佈寬度 w_eq = w_factored * (L_short / 2) * (1 - L_short / (2 * L_long))
    w_trib_beam = w_factored * (L_short / 2.0) * (1.0 - (L_short / (2.0 * L_long)))
    w_beam_self_factored = 1.4 * (2400.0 * (b_beam / 100.0) * (h_beam / 100.0)) / 100.0  # kgf/cm
    w_beam_total_line = w_trib_beam + w_beam_self_factored
    Vu_beam_pure = w_beam_total_line * max(0.0, (L_long / 2.0) - (col_equiv_w / 2.0) - d_beam)

    # 樓版毛慣性矩 Ig (100cm 寬版帶) 與使用彎矩 Ma
    Ig_slab = 100.0 * (h_case ** 3) / 12.0
    yt_slab = h_case / 2.0
    # 雙向連續版短向跨中/支承等效彎矩係數 α ≈ 0.065
    Ma_slab = 0.065 * (w_service * 100.0) * (L_short ** 2)
    Es = 2.04e6  # 鋼筋彈性模數 kgf/cm²

    matrix = []
    earliest_safe_day = None
    
    for day in days:
        pct = age_pct[day]
        fci = fc * pct       
        Ec = 15100.0 * math.sqrt(fci)
        
        if case_name != "區域三：逆打地下一般平板 (常態無施工重載區)":
            if day == 7:     efficiency_factor = 0.30
            elif day == 10:  efficiency_factor = 0.45
            elif day == 14:  efficiency_factor = 0.65
            elif day == 21:  efficiency_factor = 0.85
            else:            efficiency_factor = 1.00
            actual_share = max_share_28 * efficiency_factor + cover_board_bonus * (pct**0.5)
        else:
            actual_share = 0.0
            
        Vu_final_demand = Vu_pure_total * (1.0 - actual_share)
        Vu_beam_demand = Vu_beam_pure * 0.5 if "有下方支撐" in support_type else Vu_beam_pure
        Vu_slab_1way_eff = Vu_slab_1way_demand * 0.5 if "有下方支撐" in support_type else Vu_slab_1way_demand

        # (A) ACI 318-19 Table 22.6.5.2 雙向穿孔剪力強度
        if has_stirrup == "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)":
            vc = 0.53 * lambda_s_slab * math.sqrt(fci)
            Vs = (Av * fy * d_slab) / float(stirrup_space)
            Vn = vc * b0_eff * d_slab + Vs
            Vn_max = 1.6 * math.sqrt(fci) * b0_eff * d_slab
            if Vn > Vn_max:
                Vn = Vn_max
            coeff_note = f"ACI-19: 0.53×{size_effect_tag}√fci' + Vs"
        else:
            vc_1 = 1.06 * lambda_s_slab * math.sqrt(fci)
            vc_2 = 0.53 * (1.0 + 2.0 / beta_col) * lambda_s_slab * math.sqrt(fci)
            vc_3 = 0.265 * (2.0 + (alpha_s * d_slab) / b0_eff) * lambda_s_slab * math.sqrt(fci)
            vc_aci19_base = min(vc_1, vc_2, vc_3)
            
            vc = vc_aci19_base * 0.65
            Vn = vc * b0_eff * d_slab
            eff_k = round((vc / math.sqrt(fci)), 3)
            coeff_note = f"ACI-19(β={round(beta_col,2)})×{size_effect_tag}×0.65 = {eff_k}√fci'"
            
        phi_Vn = 0.75 * Vn
        
        # 👇【請確認以下 (B) 與 (C) 兩段完整存在於 phi_Vn 與撓度計算之間】👇
        # (B) ACI 318-19 Table 22.5.5.1 樓版單向剪力容量 (100cm 寬無箍筋版帶)
        vc_slab_1way = 2.12 * lambda_s_slab * (rho_w_slab ** (1.0 / 3.0)) * math.sqrt(fci)
        vc_slab_1way = min(vc_slab_1way, 1.33 * math.sqrt(fci))
        phi_Vn_slab_1way = 0.75 * vc_slab_1way * 100.0 * d_slab

        # (C) ACI 318-19 Table 22.5.5.1 外圍大梁單向剪力容量
        Av_beam = 2.0 * a_b
        d_crit_beam = 800.0 * (rho_w_beam ** (2.0 / 3.0)) - 25.0
        vc_trad = 0.53 * math.sqrt(fci)
        vc_aci19 = 2.12 * lambda_s_beam * (rho_w_beam ** (1.0 / 3.0)) * math.sqrt(fci)
        
        if d_beam <= d_crit_beam:
            vc_beam_stress = max(vc_trad, vc_aci19)
        else:
            vc_beam_stress = min(vc_trad, vc_aci19)
        vc_beam_stress = min(vc_beam_stress, 1.33 * math.sqrt(fci))
            
        vc_beam = vc_beam_stress * b_beam * d_beam
        Vs_beam = (Av_beam * fy * d_beam) / float(stirrup_space)
        phi_Vn_beam = 0.75 * (vc_beam + Vs_beam)
        # 👆【補齊結束，下方緊接 ACI 318-19 標準撓度計算】👆

        # 📌【ACI 318-19 第 24.2.4 節標準撓度計算 (區分活載重瞬時 ΔL 與持續載重徐變+乾縮 λΔ·Δsus)】
        fr = 2.0 * math.sqrt(fci)
        Mcr = (fr * Ig_slab) / yt_slab
        
        n_mod = Es / Ec
        rho_n = rho_w_slab * n_mod
        k_cr = math.sqrt(2.0 * rho_n + (rho_n ** 2)) - rho_n
        Icr_slab = (100.0 * ((k_cr * d_slab) ** 3) / 3.0) + n_mod * As_per_meter * ((d_slab - k_cr * d_slab) ** 2)

        w_dead_cm = w_dead / 10000.0
        w_live_cm = w_live / 10000.0
        w_sus_cm = w_dead_cm + 0.20 * w_live_cm  # 持續性載重 = 100%靜載重 + 20%持續施工堆置活載
        
        # 1. 計算持續載重下的開裂有效慣性矩 Ie_sus (Bischoff 公式)
        Ma_sus = 0.065 * (w_sus_cm * 100.0) * (L_short ** 2)
        if Ma_sus <= (2.0 / 3.0) * Mcr:
            Ie_sus = Ig_slab
        else:
            ratio_m_sus = ((2.0 / 3.0) * Mcr) / Ma_sus
            denom_sus = 1.0 - (ratio_m_sus ** 2) * (1.0 - (Icr_slab / Ig_slab))
            Ie_sus = min(Ig_slab, Icr_slab / max(0.1, denom_sus))
            
        # 2. 計算全載重 (D+L) 下的開裂有效慣性矩 Ie_total (Bischoff 公式)
        if Ma_slab <= (2.0 / 3.0) * Mcr:
            Ie_total = Ig_slab
        else:
            ratio_m = ((2.0 / 3.0) * Mcr) / Ma_slab
            denom = 1.0 - (ratio_m ** 2) * (1.0 - (Icr_slab / Ig_slab))
            Ie_total = min(Ig_slab, Icr_slab / max(0.1, denom))
            
        # 3. 計算各階段瞬時撓度 (樓版短向版帶 + 外圍大梁變位貢獻)
        Ig_beam = b_beam * (h_beam ** 3) / 12.0
        Ie_beam = 0.55 * Ig_beam
        
        def calc_inst_deflection(w_load_cm, Ie_s):
            d_slab_i = 0.0020 * (w_load_cm * 100.0 * (L_short ** 4)) / (Ec * Ie_s)
            w_trib_b = w_load_cm * (L_short / 2.0) * (1.0 - (L_short / (2.0 * L_long)))
            d_beam_i = 0.0026 * (w_trib_b * (L_long ** 4)) / (Ec * Ie_beam)
            d_tot_i = d_slab_i + 0.5 * d_beam_i
            return d_tot_i * 0.5 if "有下方支撐" in support_type else d_tot_i

        delta_D_inst = calc_inst_deflection(w_dead_cm, Ie_sus)
        delta_DL_inst = calc_inst_deflection(w_service, Ie_total)
        delta_sus_inst = calc_inst_deflection(w_sus_cm, Ie_sus)
        
        # ACI 318-19 規定：純活載重瞬時撓度 ΔL = Δ(D+L) - ΔD
        delta_L_inst = max(0.0, delta_DL_inst - delta_D_inst)
        
        # 4. ACI 318-19 Eq. 24.2.4.1.1 徐變 + 乾縮綜合乘數 λΔ = ξ / (1 + 50 * ρ')
        rho_prime = 0.5 * rho_w_slab
        if day <= 10:
            xi_factor = 2.4
        elif day <= 14:
            xi_factor = 2.2
        else:
            xi_factor = 2.0
            
        lambda_delta = xi_factor / (1.0 + 50.0 * rho_prime)
        
        # 長期附加撓度 (徐變 + 乾縮) 與 總增量撓度 (ΔL + λΔ * Δsus)
        delta_cp_sh = lambda_delta * delta_sus_inst
        delta_total_aci = delta_L_inst + delta_cp_sh
        
        # 雙軌法規限值 (ACI 318-19 Table 24.2.2)
        limit_L360 = L_long / 360.0   # 純活載重瞬時撓度限值
        limit_L240 = L_long / 240.0   # 活載 + 徐變 + 乾縮 總撓度限值
        
        deflection_ok = (delta_L_inst <= limit_L360) and (delta_total_aci <= limit_L240)

        # 五道剛性拆模解鎖門檻檢核
        days_ok = day >= min_days_req
        strength_ok = pct >= min_fc_ratio
        shear_ok = phi_Vn > Vu_final_demand
        slab_1way_ok = phi_Vn_slab_1way > Vu_slab_1way_eff
        beam_shear_ok = phi_Vn_beam > Vu_beam_demand
        
        if days_ok and strength_ok and shear_ok and slab_1way_ok and beam_shear_ok and deflection_ok:
            is_safe = "🟢 准予拆模 (PASS)"
            strategy = f"已滿法定 {min_days_req} 天、強度達 {int(pct*100)}% (≥{int(min_fc_ratio*100)}%) 且雙向穿孔、板/梁單剪與雙軌撓度全數過關。"
            if earliest_safe_day is None:
                earliest_safe_day = day
        else:
            is_safe = "🔴 禁止拆模 (NG)"
            reasons = []
            if not days_ok:
                reasons.append(f"未滿法規最少養護 {min_days_req} 天")
            if not strength_ok:
                reasons.append(f"強度{int(pct*100)}%未達{int(min_fc_ratio*100)}%死線")
            if not shear_ok:
                reasons.append(f"穿孔({round(phi_Vn/1000.0,1)}T)<({round(Vu_final_demand/1000.0,1)}T)")
            if not slab_1way_ok:
                reasons.append(f"板單剪({round(phi_Vn_slab_1way/1000.0,1)}T/m)<({round(Vu_slab_1way_eff/1000.0,1)}T/m)")
            if not beam_shear_ok:
                reasons.append(f"梁單剪({round(phi_Vn_beam/1000.0,1)}T)<({round(Vu_beam_demand/1000.0,1)}T)")
            if not deflection_ok:
                if delta_L_inst > limit_L360:
                    reasons.append(f"活載瞬時撓度({round(delta_L_inst,2)}cm)>L/360({round(limit_L360,2)}cm)")
                if delta_total_aci > limit_L240:
                    reasons.append(f"L+徐變+乾縮總撓度({round(delta_total_aci,2)}cm)>L/240({round(limit_L240,2)}cm)")
            strategy = "嚴禁拆模！原因：" + "；".join(reasons) + "。"
        
        matrix.append({
            "day": day,
            "pct_str": f"{int(pct*100)}%",
            "fci": round(fci, 1),
            "wu_kgm2": round(w_factored * 10000.0, 0),
            "coeff_note": coeff_note,
            "phi_Vn": round(phi_Vn / 1000.0, 1),
            "Vu": round(Vu_final_demand / 1000.0, 1),
            "slab_1way": f"{round(phi_Vn_slab_1way/1000.0, 1)} / {round(Vu_slab_1way_eff/1000.0, 1)}",
            "beam_1way": f"{round(phi_Vn_beam/1000.0, 1)} / {round(Vu_beam_demand/1000.0, 1)}",
            "delta_info": f"ΔL={round(delta_L_inst,2)} | 總={round(delta_total_aci,2)} (λΔ={round(lambda_delta,2)})",
            "limit": f"{round(limit_L360,2)} / {round(limit_L240,2)}",
            "status": is_safe,
            "strategy": strategy,
            "steel_info": steel_plate_info,
            "thickness": int(h_case),
            "As_m": round(As_per_meter, 2)
        })
    return matrix, earliest_safe_day

# =========================================================================
# 2. Word 報告書自動化產出引擎
# =========================================================================
def set_cell_background(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def generate_cumulative_word_report(params, all_cases_results):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(10.0)
    
    title = doc.add_paragraph()
    run = title.add_run("逆打施工平台與地下平板結構拆模時間綜合核算報告書 (ACI 318-19 規範)")
    run.font.size = Pt(16)
    run.font.bold = True
    title.alignment = 1
    
    doc.add_heading("一、 結構工程基礎設計與拆模管制參數明細", level=2)
    table_in = doc.add_table(rows=1, cols=2)
    table_in.style = 'Table Grid'
    
    cell_title_l = table_in.cell(0, 0)
    cell_title_r = table_in.cell(0, 1)
    cell_title_l.paragraphs[0].add_run('基本工程參數項目').font.bold = True
    cell_title_r.paragraphs[0].add_run('設計規範數值').font.bold = True
    set_cell_background(cell_title_l, "4F81BD")
    set_cell_background(cell_title_r, "4F81BD")
        
    min_ratio = params.get('min_fc_ratio', 75)
    min_days = params.get('min_days_req', 21)
    
    if params['col_shape'] == "矩形柱 (b × d)":
        col_desc = f"矩形柱 {params['col_b']} cm × {params['col_d']} cm"
    else:
        col_desc = f"圓柱 直徑 D = {params['col_dia']} cm"
        
    bar_areas_all = {4: 1.267, 5: 1.986, 6: 2.865, 7: 3.871, 8: 5.067}
    as_m = round((bar_areas_all.get(params['long_bar_no'], 1.986) * 100.0) / float(params['long_bar_space']), 2)

    inputs_list = [
        ("混凝土設計強度 fc'", f"{params['fc']} kgf/cm²"),
        ("抗剪與撓度設計規範", "ACI 318-19 Table 22.5.5.1 / Table 22.6.5.2 / Table 24.2.2 (雙軌撓度)"),
        ("法定最少拆模養護時間", f"{min_days} 天（依台灣施工規範跨距與支撐條件判定）"),
        ("法定最低拆模強度門檻", f"{min_ratio}% fc' ({round(params['fc']*min_ratio/100.0, 1)} kgf/cm²)"),
        ("鋼筋屈服強度 fy", f"{params['fy']} kgf/cm²"),
        ("縱向主鋼筋配置 (As)", f"#{params['long_bar_no']} @ {params['long_bar_space']} cm (單米鋼筋量 {as_m} cm²/m)"),
        ("樓版淨幾何長寬", f"{params['span_x']} m × {params['span_y']} m"),
        ("四週外圍大樑尺寸", f"{params['b_beam']} cm × {params['h_beam']} cm"),
        ("柱頭幾何型式與尺寸", col_desc),
        ("區域一(PC400)設定版厚", f"{params['h_case_1']} cm ({params['has_stirrup_1']})"),
        ("區域二(土車道)設定版厚", f"{params['h_case_2']} cm ({params['has_stirrup_2']})"),
        ("區域三(常態室)設定版厚", f"{params['h_case_3']} cm ({params['has_stirrup_3']})"),
        ("外圍大樑專屬 Stirrups 配置", f"{params['stirrup_legs']}支 - #{params['stirrup_no']} @ {params['stirrup_space']} cm")
    ]
    
    for item in inputs_list:
        row_cells = table_in.add_row().cells
        row_cells[0].text = item[0]
        row_cells[1].text = item[1]
        
    doc.add_paragraph("\n")
    doc.add_heading("二、 各施工分區拆模時間與多齡期核算矩陣", level=2)
    
    for case_name, (matrix, earliest_day) in all_cases_results.items():
        day_str = f"最早可安全拆模時間：【第 {earliest_day} 天】" if earliest_day else "【28天內均未達標，必須維持滿堂回撐】"
        doc.add_heading(f"分區項目: {case_name} ── {day_str}", level=3)
        
        table_out = doc.add_table(rows=1, cols=11)
        table_out.style = 'Table Grid'
        headers = ["版厚\n(Wu)", "齡期\n(強度比)", "預估強度\n(fci')", "ACI 318-19\n穿孔係數", "穿孔剪力\n容量/需求(T)", "板單向剪力\n容量/需求(T/m)", "大梁單剪\n容量/需求(T)", "雙軌撓度(cm)\nΔL | 總撓度", "雙軌限值\nL/360 | L/240", "拆模判定", "施工管制與對策"]
        
        for i, title_text in enumerate(headers):
            target_cell = table_out.cell(0, i)
            run_out = target_cell.paragraphs[0].add_run(title_text)
            run_out.font.bold = True
            run_out.font.size = Pt(8.0)
            set_cell_background(target_cell, "366092")
            
        for item in matrix:
            row_cells = table_out.add_row().cells
            row_cells[0].text = f"{item['thickness']}cm\n({int(item['wu_kgm2'])})"
            row_cells[1].text = f"{item['day']}天\n({item['pct_str']})"
            row_cells[2].text = f"{item['fci']}"
            row_cells[3].text = item['coeff_note']
            row_cells[4].text = f"{item['phi_Vn']} / {item['Vu']}"
            row_cells[5].text = item['slab_1way']
            row_cells[6].text = item['beam_1way']
            row_cells[7].text = item['delta_info']
            row_cells[8].text = item['limit']
            row_cells[9].text = item['status']
            row_cells[10].text = item['strategy']
            
            if "禁止" in item['status']:
                for cell in row_cells:
                    set_cell_background(cell, "FDE9D9")
            else:
                for cell in row_cells:
                    set_cell_background(cell, "EBF1DE")
        doc.add_paragraph()
        
    doc.add_heading("三、 ACI 318-19 雙軌撓度與抗剪規範說明", level=2)
    doc.add_paragraph(
        f"""1. 撓度檢核嚴格遵循 ACI 318-19 第 24.2.4 節與 Table 24.2.2：區分「純活載重瞬時撓度 ΔL ≤ L/360」與「持續載重長期附加撓度（含徐變與乾縮，乘數 λΔ = ξ/(1+50ρ')）加上瞬時活載撓度 ΔL + λΔ·Δsus ≤ L/240」，並以 Bischoff 公式動態計算開裂有效慣性矩 Ie。
        2. 各分區因數化荷重 Wu 隨樓版厚度（24 kgf/m²/cm + 100 kgf/m²）動態計算，並同步驅動雙向穿孔剪力、樓版單向剪力（1m寬板帶）與外圍大梁梯形負擔面積單向剪力。
        3. 現場實際拆模解鎖必須滿足：(1) 養護達 {min_days} 天；(2) 強度達 {min_ratio}% fc'；(3) 穿孔、板單剪、梁單剪皆合格；(4) 雙軌撓度皆小於規範限值。"""
    )
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ==========================================
# 3. Streamlit 使用者前端網頁介面 (UI Window)
# ==========================================
st.set_page_config(page_title="逆打施工平台拆模時間綜合分析APP (ACI 318-19)", layout="wide")

st.title("逆打施工平台與地下平板 ACI 318-19 拆模時間自動化系統")
st.write(
    """本系統已升級為【v18.0 全參數動態連動版】。
各分區版厚與荷重即時連動「ACI 318-19 Bischoff 開裂慣性矩 Ie」、「樓版單向剪力（隨版厚變化）」與「大梁梯形負擔單向剪力（隨各區荷重變化）」，徹底消除固定常數。"""
)

st.sidebar.header("結構工程與拆模管制參數面版")

st.sidebar.subheader("--- 規範最少拆模時間與強度門檻 ---")
code_days_mode = st.sidebar.selectbox("法規最少拆模天數限制", [
    "自動依台灣規範 (跨距與支撐條件判定)",
    "強制 7 天 (短跨度或有連續回撐)",
    "強制 14 天 (中跨度或有連續回撐)",
    "強制 21 天 (大跨度 >6m 且逆打懸空無支撐)"
], index=0)

min_fc_ratio = st.sidebar.slider("法定最低拆模強度門檻 (% fc')", 70, 100, 75, step=5)

deck_mode_1 = st.sidebar.selectbox("區域一 (PC400區) 複合鋪面型式", [
    "單鋪 2層活動覆工板 (無銲接複合效益)",
    "單鋪 PL12mm 鋼板 (全面銲死錨碇複合斷面)",
    "PL12mm 鋼板銲死 + 2層活動覆工板 (最高規格雙重防線)"
], index=2)

st.sidebar.subheader("--- 縱向主鋼筋配置 (#4~#8, @10~20) ---")
col_lb1, col_lb2 = st.sidebar.columns(2)
with col_lb1:
    long_bar_no = st.selectbox("縱向鋼筋號數", [4, 5, 6, 7, 8], index=1)
with col_lb2:
    long_bar_space = st.selectbox("縱向鋼筋間距 (cm)", [10, 12, 15, 18, 20], index=2)

st.sidebar.subheader("--- 四週外圍大樑專屬 Stirrups 配置 ---")
stirrup_legs = st.sidebar.selectbox("A. 大梁箍筋配置肢數 (Legs)", [2, 3, 4], index=0) 
stirrup_no = st.sidebar.selectbox("B. 大梁箍筋鋼筋號數 (Bar No.)", [3, 4, 5], index=1)   
stirrup_space = st.sidebar.selectbox("C. 大梁箍筋綁紮間距 (Spacing cm)", [10, 12, 15, 20], index=0) 

st.sidebar.subheader("--- 柱頭幾何型式與尺寸 ---")
col_shape = st.sidebar.radio("柱頭斷面形狀", ["矩形柱 (b × d)", "圓柱 (直徑 D)"], horizontal=True)

col_b, col_d, col_dia = 90.0, 90.0, 90.0
if col_shape == "矩形柱 (b × d)":
    c_col1, c_col2 = st.sidebar.columns(2)
    with c_col1:
        col_b = st.number_input("柱寬 b (cm)", value=90.0, step=5.0)
    with c_col2:
        col_d = st.number_input("柱深 d (cm)", value=90.0, step=5.0)
else:
    col_dia = st.sidebar.number_input("圓柱直徑 D (cm)", value=90.0, step=5.0)

st.sidebar.subheader("--- 區域一：挖土機區獨立變數 ---")
h_case_1 = st.sidebar.slider("1-A. 區域一 樓版厚度 (cm)", 20, 50, 30)
has_stirrup_1 = st.sidebar.selectbox("1-B. 區域一 版內暗樑剪力補強", ["無版內剪力筋 (ACI 318-19 Table 22.6.5.2)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域二：土車行駛區獨立變數 ---")
h_case_2 = st.sidebar.slider("2-A. 區域二 樓版厚度 (cm)", 20, 50, 25)
has_stirrup_2 = st.sidebar.selectbox("2-B. 區域二 版內暗樑剪力補強", ["無版內剪力筋 (ACI 318-19 Table 22.6.5.2)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域三：常態地下室獨立變數 ---")
h_case_3 = st.sidebar.slider("3-A. 區域三 樓版厚度 (cm)", 20, 60, 40)
has_stirrup_3 = st.sidebar.selectbox("3-B. 區域三 版內暗樑剪力補強", ["無版內剪力筋 (ACI 318-19 Table 22.6.5.2)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 基礎材料與網格幾何參數 ---")
t_steel = st.sidebar.selectbox("1. 鋪設鋼板基本厚度 PL (mm)", [12, 16, 19, 22], index=0)
fc = st.sidebar.selectbox("2. 混凝土 28天設計強度 fc' (kgf/cm²)", [280, 350, 420], index=1)
fy = st.sidebar.selectbox("3. 鋼筋屈服強度 fy (kgf/cm²)", [2800, 4200, 4900], index=1)

support_type = st.sidebar.selectbox("4. 施工支撐壓力傳遞模式", [
    "逆打工法完全懸空模式 (無下方支撐/傳遞)", 
    "有下方支撐傳遞 (多層連續回撐體系)"
], index=0)

col1, col2 = st.sidebar.columns(2)
with col1:
    span_x = st.number_input("版淨長度 X (m)", value=6.6, step=0.1)
    b_beam = st.number_input("大樑寬度 b (cm)", value=60.0, step=5.0)
with col2:
    span_y = st.number_input("版淨寬度 Y (m)", value=4.9, step=0.1)
    h_beam = st.number_input("大樑高度 h (cm)", value=80.0, step=5.0)

span_max = max(span_x, span_y)
min_days_req = get_statutory_min_days(span_max, support_type, code_days_mode)
st.sidebar.info(f"📌 依目前跨距 ({span_max}m) 與支撐條件，台灣施工規範最少拆模天數為：【{min_days_req} 天】")

params = {
    'fc': fc, 'fy': fy, 'span_x': span_x, 'span_y': span_y,
    'b_beam': b_beam, 'h_beam': h_beam,
    'col_shape': col_shape, 'col_b': col_b, 'col_d': col_d, 'col_dia': col_dia,
    'long_bar_no': long_bar_no, 'long_bar_space': long_bar_space,
    'support_type': support_type, 't_steel': t_steel, 
    'min_fc_ratio': min_fc_ratio,
    'min_days_req': min_days_req,
    'stirrup_legs': stirrup_legs, 'stirrup_no': stirrup_no, 'stirrup_space': stirrup_space,
    'deck_mode_1': deck_mode_1,
    'h_case_1': h_case_1, 'has_stirrup_1': has_stirrup_1,
    'h_case_2': h_case_2, 'has_stirrup_2': has_stirrup_2,
    'h_case_3': h_case_3, 'has_stirrup_3': has_stirrup_3
}

cases_list = [
    "區域一：PC400 挖土機工作平台區",
    "區域二：20T 土車行駛通道區",
    "區域三：逆打地下一般平板 (常態無施工重載區)"
]

all_cases_results = {}
for case in cases_list:
    all_cases_results[case] = calculate_single_case(params, case)

for case_name, (matrix, earliest_day) in all_cases_results.items():
    st.subheader(f"分析分區: {case_name}")
    if earliest_day:
        st.success(f"⏱️ **系統精算建議：本區最早可安全拆模（開挖）時間為【第 {earliest_day} 天】**（已滿足規範最少 {min_days_req} 天、強度 ≥ {min_fc_ratio}% fc'、雙向穿孔、板/梁單剪與開裂撓度死線）")
    else:
        st.error("⚠️ **系統精算警告：本區在 28 天內均未同時滿足規範天數、強度與開裂撓度死線，必須維持原狀土模或立設滿堂重型回撐！**")
    
    display_data = []
    for r in matrix:
        display_data.append({
            "版厚 (Wu kgf/m²)": f"{r['thickness']} cm ({int(r['wu_kgm2'])})",
            "齡期 (強度比)": f"{r['day']} 天 ({r['pct_str']})",
            "預估強度 (kgf/cm²)": r['fci'],
            "ACI 318-19 穿孔係數": r['coeff_note'],
            "穿孔剪力 容量/需求 (T)": f"{r['phi_Vn']} / {r['Vu']}",
            "板單向剪力 容量/需求 (T/m)": r['slab_1way'],
            "大梁單剪 容量/需求 (T)": r['beam_1way'],
            "總撓度(cm) [Bischoff Ie]": r['delta_info'],
            "限值 L/240": r['limit'],
            "拆模判定": r['status'],
            "施工管制與未達標原因說明": r['strategy']
        })
    st.table(display_data)

st.subheader("綜合申報書一鍵匯出面版")
word_file = generate_cumulative_word_report(params, all_cases_results)

st.download_button(
    label="導出【三大區域拆模時間與 ACI 318-19 核算】標準 Word 報告書 (.docx)",
    data=word_file,
    file_name="ACI_318_19_Formwork_Removal_Report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    key="fixed_final_download_key_panel"
)

st.stop()
