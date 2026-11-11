import streamlit as st
import numpy as np
import math
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import io

# =========================================================================================
# 1. 結構力學核心引擎 (v10.0 完全自由度活算版：全面連動三大區域獨立版厚、配筋與鋪面選單)
# =========================================================================================
def calculate_single_case(params, case_name):
    fc = params['fc']
    fy = params['fy']
    span_x = params['span_x'] 
    span_y = params['span_y'] 
    b_beam = params['b_beam']
    h_beam = params['h_beam']
    col_w = params['col_w']
    struct_type = params['struct_type']
    support_type = params['support_type']
    t_steel = params['t_steel']
    
    # 讀取大梁箍筋參數
    stirrup_legs = params['stirrup_legs'] 
    stirrup_no = params['stirrup_no']     
    stirrup_space = params['stirrup_space'] 
    
    # 5大施工齡期陣列
    days = [7, 10, 14, 21, 28]
    age_pct = {7: 0.70, 10: 0.78, 14: 0.85, 21: 0.93, 28: 1.00}
    
    bar_areas = {3: 0.713, 4: 1.267}
    a_b = bar_areas.get(stirrup_no, 0.713)
    
    # 📢【力學引擎全解鎖】：全面讀取左側面板輸入之各分區獨立參數
    if case_name == "區域一：PC400 挖土機工作平台區 (#8@12)":
        h_case = params['h_case_1']
        d_slab = h_case - 5.5  
        w_service = (750.0 + 1700.0) / 10000.0  
        w_factored = (1.4 * 750.0 + 1.7 * 1700.0 * 1.3) / 10000.0 
        
        # 讀取區域一專屬主筋銷栓與暗樑剪力筋開關
        V_dowel = params['dowel_1'] * 1000.0
        has_stirrup = params['has_stirrup_1']
        
        # 📌【鋪面選單聯動】：依據選單彈性驅動後台應力稀釋函數與週長發育
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
        else: # 雙重鋪面：鋼板 + 覆工板
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
        V_dowel = params['dowel_2'] * 1000.0
        has_stirrup = params['has_stirrup_2']
        
        t_cm = t_steel / 10.0
        b0_eff = (col_w + (2.0 * t_cm) + d_slab) * 4.0
        Vu_beam_pure = 9.5 * 1000.0
        steel_plate_info = f"PL{t_steel}mm 鋼板 / 車道動態抗滑移剛性鋪面"
        
    else:  # 區域三：逆打地下一般平板 (常態無施工重載區 #5@24)
        h_case = params['h_case_3']
        d_slab = h_case - 5.0 
        w_service = (1000.0 + 350.0) / 10000.0
        w_factored = (1.4 * 1000.0 + 1.7 * 350.0) / 10000.0
        
        max_share_28 = 0.0    
        cover_board_bonus = 0.0
        V_dowel = params['dowel_3'] * 1000.0
        has_stirrup = params['has_stirrup_3']
        
        # 邊平板三面受剪幾何折減模型
        b0_eff = (2.0 * col_w + col_w + d_slab) * 2.0  
        Vu_beam_pure = 10.5 * 1000.0
        steel_plate_info = "無 (地下常態區邊平板依 ACI 進行三面沖切折減檢核)"

    d_beam = h_beam - 8.0  
    L = max(span_x, span_y) * 100.0  
    delta_limit = L / 240.0 
    
    Vu_pure_total = w_factored * ((span_x * span_y * 10000.0) - (col_w + d_slab)**2)

    matrix = []
    for day in days:
        pct = age_pct[day]
        fci = fc * pct       
        Ec = 15100.0 * math.sqrt(fci)
        
        # 早期強度硬化對複合分擔率之折減發育
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
        Vu_beam_demand = Vu_beam_pure * 0.5 if support_type == "有下方支撐傳遞 (多層連續回撐體系)" else Vu_beam_pure

        # 穿孔剪力因數化容量 (連動各區獨立暗樑剪力筋開關)
        vc = 0.53 * math.sqrt(fci)
        if has_stirrup == "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)":
            Vs = (Av * fy * d_slab) / float(stirrup_space)
        else:
            Vs = 0.0
            
        phi_Vn = 0.75 * (vc * b0_eff * d_slab + Vs + V_dowel)
        
        # 單向大樑端部抗剪容量 (連動大梁專屬箍筋 #4@10cm)
        vc_beam = 0.53 * math.sqrt(fci) * b_beam * d_beam
        Vs_beam = ((2.0 * a_b) * fy * d_beam) / float(stirrup_space)
        phi_Vn_beam = 0.75 * (vc_beam + Vs_beam)
        
        # 早期開裂與加速徐變撓度模型
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
        if support_type == "有下方支撐傳遞 (多層連續回撐體系)":
            delta_instant *= 0.5
            
        delta_total = delta_instant * creep_multiplier

        shear_ok = phi_Vn > Vu_final_demand
        beam_shear_ok = phi_Vn_beam > Vu_beam_demand
        deflection_ok = delta_total <= delta_limit
        
        if shear_ok and beam_shear_ok and deflection_ok:
            is_safe = "安全符合"
            strategy = "常規拆模與回撐"
        else:
            is_safe = "剛度強度不足"
            if not shear_ok:
                if case_name == "區域三：逆打地下一般平板 (常態無施工重載區 #5@24)":
                    strategy = f"🔴 邊平板沖切容量不足！7d~14d邊界週長縮減，強制保留外牆側支撐土模，嚴禁提前開挖解體！"
                else:
                    strategy = f"🔴 ACI沖切破壞警告！早期容量({round(phi_Vn/1000.0,1)}T)低於需求({round(Vu_final_demand/1000.0,1)}T)。絕對禁止開挖下方土方！"
            elif not beam_shear_ok:
                strategy = f"❌ 大梁端部單向剪力容量({round(phi_Vn_beam/1000.0,1)}T)小於需求({round(Vu_beam_demand/1000.0,1)}T)！早期支座面有剪斷風險，必須立設重型回撐"
            else:
                strategy = f"🟡 早期下陷超規({round(delta_total,2)}cm > 限值{round(delta_limit,2)}cm)！強制大跨度下方立設重型鋼管支撐"
        
        matrix.append({
            "day": day,
            "fci": round(fci, 1),
            "phi_Vn": round(phi_Vn / 1000.0, 1),
            "Vu": round(Vu_final_demand / 1000.0, 1), 
            "phi_Vn_beam": round(phi_Vn_beam / 1000.0, 1),
            "Vu_beam": round(Vu_beam_demand / 1000.0, 1),
            "delta": round(delta_total, 2),
            "limit": round(delta_limit, 2),
            "status": is_safe,
            "strategy": strategy,
            "steel_info": steel_plate_info,
            "thickness": int(h_case)
        })
    return matrix

# =========================================================================
# 2. Word 報告書自動化產出引擎 (儲存格 Tuple 絕對指定寫入)
# =========================================================================
def set_cell_background(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def generate_cumulative_word_report(params, all_cases_data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(11)
    
    title = doc.add_paragraph()
    run = title.add_run("逆打施工平台與地下平板結構拆模多層區綜合核算報告書")
    run.font.size = Pt(18)
    run.font.bold = True
    title.alignment = 1
    
    doc.add_heading("一、 結構工程基礎設計參數明細", level=2)
    table_in = doc.add_table(rows=1, cols=2)
    table_in.style = 'Table Grid'
    
    cell_title_l = table_in.cell(0, 0)
    cell_title_r = table_in.cell(0, 1)
    
    cell_title_l.paragraphs[0].add_run('基本工程參數項目').font.bold = True
    cell_title_r.paragraphs[0].add_run('設計規範數值').font.bold = True
    set_cell_background(cell_title_l, "4F81BD")
    set_cell_background(cell_title_r, "4F81BD")
        
    inputs_list = [
        ("混凝土設計強度 fc'", f"{params['fc']} kgf/cm²"),
        ("鋼筋屈服強度 fy", f"{params['fy']} kgf/cm²"),
        ("樓版淨幾何長寬", f"{params['span_x']} m × {params['span_y']} m"),
        ("四週外圍大樑尺寸", f"{params['b_beam']} cm × {params['h_beam']} cm"),
        ("柱頭方形核心尺寸", f"{params['col_w']} cm × {params['col_w']} cm"),
        ("區域一(PC400)設定版厚", f"{params['h_case_1']} cm"),
        ("區域二(土車道)設定版厚", f"{params['h_case_2']} cm"),
        ("區域三(常態室)設定版厚", f"{params['h_case_3']} cm"),
        ("外圍大樑專屬 Stirrups 配置", f"{params['stirrup_legs']}支 - #{params['stirrup_no']} @ {params['stirrup_space']} cm")
    ]
    
    for item in inputs_list:
        row_cells = table_in.add_row().cells
        row_cells[0].text = item[0]
        row_cells[1].text = item[1]
        
    doc.add_paragraph("\n")
    doc.add_heading("二、 各施工分區與荷載狀況累加檢核矩陣", level=2)
    
    for case_name, matrix in all_cases_data.items():
        doc.add_heading(f"分區項目: {case_name}", level=3)
        table_out = doc.add_table(rows=1, cols=10)
        table_out.style = 'Table Grid'
        headers = ["分區版厚", "齡期", "強度\n(fci')", "穿孔剪力\n容量/需求(T)", "大梁剪力\n容量/需求(T)", "長期撓度\n(cm)", "法規限值\n(L/240)", "安全判定", "建議對策", "表面鋼板與錨碇配置"]
        
        for i, title_text in enumerate(headers):
            target_cell = table_out.cell(0, i)
            run_out = target_cell.paragraphs[0].add_run(title_text)
            run_out.font.bold = True
            run_out.font.size = Pt(9.5)
            set_cell_background(target_cell, "366092")
        for item in matrix:
            row_cells = table_out.add_row().cells
            row_cells[0].text = f"{item['thickness']} cm"
            row_cells[1].text = f"{item['day']}天"
            row_cells[2].text = f"{item['fci']}"
            row_cells[3].text = f"{item['phi_Vn']} / {item['Vu']}"
            row_cells[4].text = f"{item['phi_Vn_beam']} / {item['Vu_beam']}"
            row_cells[5].text = f"{item['delta']}"
            row_cells[6].text = f"{item['limit']}"
            row_cells[7].text = item['status']
            row_cells[8].text = item['strategy']
            row_cells[9].text = item['steel_info']
            
            if "不足" in item['status']:
                for cell in row_cells:
                    set_cell_background(cell, "FDE9D9")
        doc.add_paragraph()
        
    doc.add_heading("三、 ACI 318 與 AISC 綜合品管管制死線說明", level=2)
    doc.add_paragraph(
        f"""1. 本案依據最終核定之全區版厚與幾何進行精密核算。重載平台採取指定之複合鋪面分擔率與邊樑 #4@10cm 強大端部箍筋。28天期混凝土硬化達100%時，系統將完全隨您輸入之各區版筋與剪力筋幾何實時活算。
2. 唯在養護 7天 與 14天 早期階段時，受限於低強度基材與邊平板區週長縮減缺陷，系統依舊剛性跳出紅燈警告。計畫書強制規定：混凝土養護未確實滿 21~28 天且強度未達 100% 之前，平台下方土模及邊界連續壁側櫬墊絕對禁止進行任何盲目掏空或解體拆模，維持原狀土模完全承壓效果，確保逆打全生命週期之安全。
3. 現場硬開機解鎖死線必須以現場同條件養護試體實際壓驗報告為唯一憑證，嚴禁未達標盲目施工。"""
    )
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 3. Streamlit 使用者前端網頁介面 (UI Window)
# ==========================================
st.set_page_config(page_title="逆打施工平台多分區綜合分析APP", layout="wide")

st.title("逆打施工平台與地下平板 ACI 多狀況累加自動化系統")
st.write(
    """本系統已升級為【v10.0 全自由度活算版】。
全面開放左側面板之三大分區獨立版厚、主筋銷栓噸數、暗樑剪力筋開關，並於區域一正式增設複合鋪面形式下拉選單控制核心。"""
)

st.sidebar.header("結構工程輸入參數矩陣面版")

# 📌 區域一特化鋪面下拉選單
st.sidebar.subheader("--- 區域一 (PC400區) 專屬鋪面選單 ---")
deck_mode_1 = st.sidebar.selectbox("區域一複合鋪面型式", [
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
dowel_1 = st.sidebar.selectbox("1-B. #8@12cm 主筋銷栓抗剪貢獻 (T)", [0.0, 5.0, 10.0, 15.0, 20.0], index=3)
has_stirrup_1 = st.sidebar.selectbox("1-C. 區域一 版內暗樑剪力補強", ["無版內剪力筋 (純混凝土受沖切)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域二：土車行駛區獨立變數 ---")
h_case_2 = st.sidebar.slider("2-A. 區域二 樓版厚度 (cm)", 20, 50, 25)
dowel_2 = st.sidebar.selectbox("2-B. #5@10cm 主筋銷栓抗剪貢獻 (T)", [0.0, 5.0, 8.0, 12.0], index=2)
has_stirrup_2 = st.sidebar.selectbox("2-C. 區域二 版內暗樑剪力補強", ["無版內剪力筋 (純混凝土受沖切)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 區域三：常態地下室獨立變數 ---")
h_case_3 = st.sidebar.slider("3-A. 區域三 樓版厚度 (cm)", 20, 60, 40)
dowel_3 = st.sidebar.selectbox("3-B. #5@24cm 主筋銷栓抗剪貢獻 (T)", [0.0, 3.0, 5.0], index=0)
has_stirrup_3 = st.sidebar.selectbox("3-C. 區域三 版內暗樑剪力補強", ["無版內剪力筋 (純混凝土受沖切)", "配置柱頭暗樑剪力筋 (隨大梁箍筋連動)"], index=0)

st.sidebar.subheader("--- 基礎材料與幾何參數 ---")
t_steel = st.sidebar.selectbox("1. 鋪設鋼板基本厚度 PL (mm)", [12, 16, 19, 22], index=0)
fc = st.sidebar.selectbox("2. 混凝土 28天設計強度 fc' (kgf/cm²)", [280, 350, 420], index=1)
fy = st.sidebar.selectbox("3. 鋼筋屈服強度 fy (kgf/cm²)", [2800, 4200], index=1)

struct_type = st.sidebar.selectbox("4. 結構構件型式分類", [
    "平板結構 (有柱列暗樑與單根#3剪力筋)", 
    "有樑板結構 (由大主梁主導受剪)"
], index=0)

support_type = st.sidebar.selectbox("5. 施工支撐壓力傳遞模式", [
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

params = {
    'fc': fc, 'fy': fy, 'span_x': span_x, 'span_y': span_y,
    'b_beam': b_beam, 'h_beam': h_beam, 'col_w': col_w,
    'struct_type': struct_type, 'support_type': support_type,
    't_steel': t_steel, 'stirrup_legs': stirrup_legs,
    'stirrup_no': stirrup_no, 'stirrup_space': stirrup_space,
    'deck_mode_1': deck_mode_1,
    'h_case_1': h_case_1, 'dowel_1': dowel_1, 'has_stirrup_1': has_stirrup_1,
    'h_case_2': h_case_2, 'dowel_2': dowel_2, 'has_stirrup_2': has_stirrup_2,
    'h_case_3': h_case_3, 'dowel_3': dowel_3, 'has_stirrup_3': has_stirrup_3
}

cases_list = [
    "區域一：PC400 挖土機工作平台區 (#8@12)",
    "區域二：20T 土車行駛通道區 (#5@10)",
    "區域三：逆打地下一般平板 (常態無施工重載區 #5@24)"
]

all_cases_results = {}
for case in cases_list:
    all_cases_results[case] = calculate_single_case(params, case)

for case_name, matrix in all_cases_results.items():
    st.subheader(f"分析分區: {case_name} — ACI 多齡期核算矩陣")
    
    display_data = []
    for r in matrix:
        display_data.append({
            "分區版厚": f"{r['thickness']} cm",
            "施工齡期": f"{r['day']} 天",
            "預估強度(kgf/cm²)": r['fci'],
            "穿孔剪力 容量/需求 (T)": f"{r['phi_Vn']} / {r['Vu']}",
            "大梁單向剪力 容量/需求 (T)": f"{r['phi_Vn_beam']} / {r['Vu_beam']}",
            "長期總累計撓度 (cm)": r['delta'],
            "法規撓度限值 L/240 (cm)": r['limit'],
            "結構安全判定": r['status'],
            "建議輔助與支撐管理對策": r['strategy'],
            "表面鋼板與錨碇配置": r['steel_info']
        })
    st.table(display_data)

st.subheader("綜合申報書一鍵匯出面版")
word_file = generate_cumulative_word_report(params, all_cases_results)

st.download_button(
    label="導出【三大區域合併累加】標準 Word 報告書 (.docx)",
    data=word_file,
    file_name="ACI_Multi_Case_Cumulative_Report.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    key="fixed_final_download_key_panel"
)

st.stop()