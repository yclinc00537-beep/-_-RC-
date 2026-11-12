import streamlit as st
import numpy as np
import math
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# =========================================================================================
# 1. 結構力學核心引擎 (v14.0 整合台灣規範最少拆模天數 + 75% fc' + 開裂剛度折減四重死線)
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
    col_w = params['col_w']
    support_type = params['support_type']
    t_steel = params['t_steel']
    min_fc_ratio = params.get('min_fc_ratio', 75) / 100.0
    min_days_req = params.get('min_days_req', 21)
    
    stirrup_legs = params['stirrup_legs'] 
    stirrup_no = params['stirrup_no']     
    stirrup_space = params['stirrup_space'] 
    
    days = [7, 10, 14, 21, 28]
    age_pct = {7: 0.70, 10: 0.78, 14: 0.85, 21: 0.93, 28: 1.00}
    
    bar_areas = {3: 0.713, 4: 1.267}
    a_b = bar_areas.get(stirrup_no, 0.713)
    Av = stirrup_legs * a_b
    
    if case_name == "區域一：PC400 挖土機工作平台區 (#8@12)":
        h_case = params['h_case_1']
        d_slab = h_case - 5.5  
        w_service = (750.0 + 1700.0) / 10000.0  
        w_factored = (1.4 * 750.0 + 1.7 * 1700.0 * 1.3) / 10000.0 
        has_stirrup = params['has_stirrup_1']
        
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
            
        Vu_beam_pure = 11.2 * 1000.0
        b0_eff = (col_w + (2.0 * t_cm) + d_slab) * 4.0
        
    elif case_name == "區域二：20T 土車行駛通道區 (#5@10)":
        h_case = params['h_case_2']
        d_slab = h_case - 5.0 
        w_service = (625.0 + 1250.0) / 10000.0
        w_factored = (1.4 * 625.0 + 1.7 * 1250.0 * 1.3) / 10000.0
        
        max_share_28 = 0.50
        cover_board_bonus = 0.0
        has_stirrup = params['has_stirrup_2']
        
        t_cm = t_steel / 10.0
        b0_eff = (col_w + (2.0 * t_cm) + d_slab) * 4.0
        Vu_beam_pure = 9.5 * 1000.0
        steel_plate_info = f"PL{t_steel}mm 鋼板 / 車道動態抗滑移剛性鋪面"
        
    else:
        h_case = params['h_case_3']
        d_slab = h_case - 5.0 
        w_service = (1000.0 + 350.0) / 10000.0
        w_factored = (1.4 * 1000.0 + 1.7 * 350.0) / 10000.0
        
        max_share_28 = 0.0    
        cover_board_bonus = 0.0
        has_stirrup = params['has_stirrup_3']
        
        b0_eff = (2.0 * col_w + col_w + d_slab) * 2.0  
        Vu_beam_pure = 10.5 * 1000.0
        steel_plate_info = "無 (地下常態區邊平板依 ACI 進行三面沖切折減檢核)"

    d_beam = h_beam - 8.0  
    L = max(span_x, span_y) * 100.0  
    delta_limit = L / 240.0 
    
    lambda_s = min(1.0, math.sqrt(2.0 / (1.0 + 0.04 * d_slab)))
    Vu_pure_total = w_factored * ((span_x * span_y * 10000.0) - (col_w + d_slab)**2)

    matrix = []
    earliest_safe_day = None
    
    for day in days:
        pct = age_pct[day]
        fci = fc * pct       
        Ec = 15100.0 * math.sqrt(fci)
        
        if case_name != "區域三：逆打地下一般平板 (常態無施工重載區 #5@24)":
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

        # ACI 318-19 穿孔剪力容量
        if has_stirrup == "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)":
            vc = 0.53 * math.sqrt(fci)
            Vs = (Av * fy * d_slab) / float(stirrup_space)
            Vn = vc * b0_eff * d_slab + Vs
            Vn_max = 1.6 * math.sqrt(fci) * b0_eff * d_slab
            if Vn > Vn_max:
                Vn = Vn_max
            coeff_note = "0.53√fci' + Vs"
        else:
            eff_coeff = round(1.06 * lambda_s * 0.65, 3)
            vc = (1.06 * lambda_s * 0.65) * math.sqrt(fci)
            Vn = vc * b0_eff * d_slab
            coeff_note = f"1.06×λs({round(lambda_s,2)})×0.65 = {eff_coeff}√fci'"
            
        phi_Vn = 0.75 * Vn
        
        # 單向大樑端部抗剪容量
        vc_beam = 0.53 * math.sqrt(fci) * b_beam * d_beam
        Vs_beam = ((2.0 * a_b) * fy * d_beam) / float(stirrup_space)
        phi_Vn_beam = 0.75 * (vc_beam + Vs_beam)
        
        # 早期開裂剛度折減與加速徐變撓度模型
        if day <= 10:
            ie_ratio = 0.25        
            creep_multiplier = 3.0 
        elif day <= 14:
            ie_ratio = 0.45
            creep_multiplier = 2.5
        else:
            ie_ratio = 0.80 if case_name == "區域一：PC400 挖土機工作平台區 (#8@12)" else 0.50
            creep_multiplier = 2.0 
            
        w_line = w_service * b_beam  
        I_e = ie_ratio * (b_beam * (h_case**3) / 12.0)
        
        delta_instant = 0.0026 * (w_line * (L**4)) / (Ec * I_e)
        if "有下方支撐" in support_type:
            delta_instant *= 0.5
            
        delta_total = delta_instant * creep_multiplier

        # 四道剛性拆模解鎖門檻檢核
        days_ok = day >= min_days_req
        strength_ok = pct >= min_fc_ratio
        shear_ok = phi_Vn > Vu_final_demand
        beam_shear_ok = phi_Vn_beam > Vu_beam_demand
        deflection_ok = delta_total <= delta_limit
        
        if days_ok and strength_ok and shear_ok and beam_shear_ok and deflection_ok:
            is_safe = "🟢 准予拆模 (PASS)"
            strategy = f"已滿法定 {min_days_req} 天、強度達 {int(pct*100)}% (≥{int(min_fc_ratio*100)}%) 且應力與開裂撓度全數過關，可進行分區拆模開挖。"
            if earliest_safe_day is None:
                earliest_safe_day = day
        else:
            is_safe = "🔴 禁止拆模 (NG)"
            reasons = []
            if not days_ok:
                reasons.append(f"未滿法規最少養護 {min_days_req} 天限制")
            if not strength_ok:
                reasons.append(f"強度僅{int(pct*100)}%未達{int(min_fc_ratio*100)}%死線")
            if not shear_ok:
                reasons.append(f"穿孔容量({round(phi_Vn/1000.0,1)}T)<需求({round(Vu_final_demand/1000.0,1)}T)")
            if not beam_shear_ok:
                reasons.append(f"大梁單剪({round(phi_Vn_beam/1000.0,1)}T)<需求({round(Vu_beam_demand/1000.0,1)}T)")
            if not deflection_ok:
                reasons.append(f"開裂徐變撓度({round(delta_total,2)}cm)>限值({round(delta_limit,2)}cm)[Ie={ie_ratio}Ig]")
            strategy = "嚴禁拆模開挖！原因：" + "；".join(reasons) + "。維持原狀土模或滿堂回撐。"
        
        matrix.append({
            "day": day,
            "pct_str": f"{int(pct*100)}%",
            "fci": round(fci, 1),
            "coeff_note": coeff_note,
            "phi_Vn": round(phi_Vn / 1000.0, 1),
            "Vu": round(Vu_final_demand / 1000.0, 1), 
            "phi_Vn_beam": round(phi_Vn_beam / 1000.0, 1),
            "Vu_beam": round(Vu_beam_demand / 1000.0, 1),
            "delta_info": f"{round(delta_total, 2)} (Ie={ie_ratio})",
            "limit": round(delta_limit, 2),
            "status": is_safe,
            "strategy": strategy,
            "steel_info": steel_plate_info,
            "thickness": int(h_case)
        })
    return matrix, earliest_safe_day

# =========================================================================
# 2. Word 報告書自動化產出引擎
# =========================================================================================
def set_cell_background(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def generate_cumulative_word_report(params, all_cases_results):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(10.5)
    
    title = doc.add_paragraph()
    run = title.add_run("逆打施工平台與地下平板結構拆模時間綜合核算報告書")
    run.font.size = Pt(18)
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
    inputs_list = [
        ("混凝土設計強度 fc'", f"{params['fc']} kgf/cm²"),
        ("法定最少拆模養護時間", f"{min_days} 天（依台灣施工規範跨距與支撐條件判定）"),
        ("法定最低拆模強度門檻", f"{min_ratio}% fc' ({round(params['fc']*min_ratio/100.0, 1)} kgf/cm²)"),
        ("鋼筋屈服強度 fy", f"{params['fy']} kgf/cm²"),
        ("樓版淨幾何長寬", f"{params['span_x']} m × {params['span_y']} m"),
        ("四週外圍大樑尺寸", f"{params['b_beam']} cm × {params['h_beam']} cm"),
        ("柱頭方形核心尺寸", f"{params['col_w']} cm × {params['col_w']} cm"),
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
        table_out = doc.add_table(rows=1, cols=10)
        table_out.style = 'Table Grid'
        headers = ["版厚", "齡期\n(強度比)", "預估強度\n(fci')", "抗剪係數修正\n(ACI 318-19)", "穿孔剪力\n容量/需求(T)", "大梁剪力\n容量/需求(T)", "總撓度(cm)\n(開裂剛度)", "限值\n(L/240)", "拆模判定", "施工管制與對策"]
        
        for i, title_text in enumerate(headers):
            target_cell = table_out.cell(0, i)
            run_out = target_cell.paragraphs[0].add_run(title_text)
            run_out.font.bold = True
            run_out.font.size = Pt(8.5)
            set_cell_background(target_cell, "366092")
        for item in matrix:
            row_cells = table_out.add_row().cells
            row_cells[0].text = f"{item['thickness']} cm"
            row_cells[1].text = f"{item['day']}天\n({item['pct_str']})"
            row_cells[2].text = f"{item['fci']}"
            row_cells[3].text = item['coeff_note']
            row_cells[4].text = f"{item['phi_Vn']} / {item['Vu']}"
            row_cells[5].text = f"{item['phi_Vn_beam']} / {item['Vu_beam']}"
            row_cells[6].text = item['delta_info']
            row_cells[7].text = f"{item['limit']}"
            row_cells[8].text = item['status']
            row_cells[9].text = item['strategy']
            
            if "禁止" in item['status']:
                for cell in row_cells:
                    set_cell_background(cell, "FDE9D9")
            else:
                for cell in row_cells:
                    set_cell_background(cell, "EBF1DE")    
        doc.add_paragraph()
        
        doc.add_heading("三、 拆模時間雙軌解鎖與品管管制死線說明", level=2)
        doc.add_paragraph(
            f"""1. 本系統嚴格整合四道拆模解鎖死線：(1) 養護齡期必須達到台灣施工規範最少拆模天數（本案判定為 {min_days} 天）；(2) 混凝土抗壓強度必須達到設計強度之 {min_ratio}% 以上；(3) ACI 318-19 尺寸效應 λs 與施工偏載 0.65 折減後之穿孔與單向抗剪容量大於需求；(4) 計入早期開裂慣性矩折減（7~10天 Ie=0.25Ig、14天 Ie=0.45Ig）與加速徐變後之總撓度小於 L/240 限值。
            2. 未同時滿足上述四道門檻前，平台下方土模及邊界連續壁側櫬墊絕對禁止盲目掏空或解體拆模。
            3. 現場實際拆模解鎖必須以同條件養護試體抗壓試驗報告與監測數據為唯一憑證，經結構技師核可後方得施作。"""
        )
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 3. Streamlit 使用者前端網頁介面 (UI Window)
# ==========================================
st.set_page_config(page_title="逆打施工平台拆模時間綜合分析APP", layout="wide")

st.title("逆打施工平台與地下平板 ACI 318-19 拆模時間自動化系統")
st.write(
    """本系統已升級為【v14.0 台灣規範最少拆模時間完全合規版】。
    全面整合「規範最少拆模天數限制」、「75% fc' 強度死線」、「應力安全門檻」與「早期開裂剛度折減（Ie=0.25~0.45Ig）及加速徐變撓度限制」，自動推算各區最早可安全拆模天數。"""
)

st.sidebar.header("結構工程與拆模管制參數面版")

# 📌 區域一特化鋪面下拉選單
st.sidebar.subheader("--- 規範最少拆模時間與強度門檻 ---")
code_days_mode = st.sidebar.selectbox("法規最少拆模天數限制", [
"自動依台灣規範 (跨距與支撐條件判定)",
"強制 7 天 (短跨度或有連續回撐)",
"強制 14 天 (中跨度或有連續回撐)",
"強制 21 天 (大跨度 >6m 且逆打懸空無支撐)"
], index=0)

min_fc_ratio = st.sidebar.slider("法定最低拆模強度門檻 (% fc')", 70, 100, 75, step=5)

deck_mode_1 = st.sidebar.selectbox("區域一(PC400區)複合鋪面型式", [
    "單鋪 2層活動覆工板 (無銲接複合效益)",
    "單鋪 PL12mm 鋼板 (全面銲死錨碇複合斷面)",
    "PL12mm 鋼板銲死 + 2層活動覆工板 (最高規格雙重防線)"
], index=2)

st.sidebar.subheader("--- 四週外圍大樑專屬 Stirrups 配置 ---")
stirrup_legs = st.sidebar.selectbox("A. 大梁箍筋配置肢數 (Legs)", [1, 2, 4], index=1) 
stirrup_no = st.sidebar.selectbox("B. 大梁箍筋鋼筋號數 (Bar No.)", [3, 4], index=1)   
stirrup_space = st.sidebar.selectbox("C. 大梁箍筋綁紮間距 (Spacing cm)", [10, 12, 15], index=0) 

# 📌 開放三大分區獨立版厚、配筋銷栓與剪力筋聯動選單
st.sidebar.subheader("--- 區域一：挖土機區獨立變數 ---")
h_case_1 = st.sidebar.slider("1-A. 區域一 樓版厚度 (cm)", 20, 50, 30)
has_stirrup_1 = st.sidebar.selectbox("1-B. 區域一 版內暗樑剪力補強", ["無版內剪力筋 (1.06×λs×0.65 ≈ 0.69√fc')", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域二：土車行駛區獨立變數 ---")
h_case_2 = st.sidebar.slider("2-A. 區域二 樓版厚度 (cm)", 20, 50, 25)
has_stirrup_2 = st.sidebar.selectbox("2-B. 區域二 版內暗樑剪力補強", ["無版內剪力筋 (1.06×λs×0.65 ≈ 0.69√fc')", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域三：常態地下室獨立變數 ---")
h_case_3 = st.sidebar.slider("3-A. 區域三 樓版厚度 (cm)", 20, 60, 40)
has_stirrup_3 = st.sidebar.selectbox("3-B. 區域三 版內暗樑剪力補強", ["無版內剪力筋 (1.06×λs×0.65 ≈ 0.69√fc')", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 基礎材料與幾何參數 ---")
t_steel = st.sidebar.selectbox("1. 鋪設鋼板基本厚度 PL (mm)", [12, 16, 19, 22], index=0)
fc = st.sidebar.selectbox("2. 混凝土 28天設計強度 fc' (kgf/cm²)", [280, 350, 420], index=1)
fy = st.sidebar.selectbox("3. 鋼筋屈服強度 fy (kgf/cm²)", [2800, 4200], index=1)

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

col_w = st.sidebar.number_input("柱頭方形尺寸 (cm)", value=90.0, step=5.0)

span_max = max(span_x, span_y)
min_days_req = get_statutory_min_days(span_max, support_type, code_days_mode)
st.sidebar.info(f"📌 依目前跨距 ({span_max}m) 與支撐條件，台灣施工規範最少拆模天數為：【{min_days_req} 天】")

params = {
    'fc': fc, 'fy': fy, 'span_x': span_x, 'span_y': span_y,
    'b_beam': b_beam, 'h_beam': h_beam, 'col_w': col_w,
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
    "區域一：PC400 挖土機工作平台區 (#8@12)",
    "區域二：20T 土車行駛通道區 (#5@10)",
    "區域三：逆打地下一般平板 (常態無施工重載區 #5@24)"
]

all_cases_results = {}
for case in cases_list:
    all_cases_results[case] = calculate_single_case(params, case)

# 📌【核心修正】：正確解包 (matrix, earliest_day)，徹底消除 NameError
for case_name, (matrix, earliest_day) in all_cases_results.items():
    st.subheader(f"分析分區: {case_name}")
    if earliest_day:
        st.success(f"⏱️ 系統精算建議：本區最早可安全拆模（開挖）時間為【第 {earliest_day} 天】（已滿足規範最少 {min_days_req} 天、強度 ≥ {min_fc_ratio}% fc'、抗剪應力與開裂撓度四道死線）")
    else:
        st.error("⚠️ **系統精算警告：本區在 28 天內均未同時滿足強度與開裂撓度死線，必須維持原狀土模或立設滿堂重型回撐！**")
    
    display_data = []
    for r in matrix:
        display_data.append({
            "版厚": f"{r['thickness']} cm",
            "齡期 (強度比)": f"{r['day']} 天 ({r['pct_str']})",
            "預估強度 (kgf/cm²)": r['fci'],
            "抗剪係數修正 (ACI 318-19)": r['coeff_note'],
            "穿孔剪力 容量/需求 (T)": f"{r['phi_Vn']} / {r['Vu']}",
            "大梁單剪 容量/需求 (T)": f"{r['phi_Vn_beam']} / {r['Vu_beam']}",
            "總撓度(cm) [開裂剛度]": r['delta_info'],
            "限值 L/240": r['limit'],
            "拆模判定": r['status'],
            "施工管制與未達標原因說明": r['strategy']
        })
    st.table(display_data)

st.subheader("綜合申報書一鍵匯出面版")
word_file = generate_cumulative_word_report(params, all_cases_results)

st.download_button(
    label="導出【三大區域拆模時間與多齡期核算】標準 Word 報告書 (.docx)",
    data=word_file,
    file_name="ACI_318_19_Formwork_Removal_Report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    key="fixed_final_download_key_panel"
)

st.stop()
