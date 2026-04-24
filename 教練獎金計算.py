import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font

def generate_perfect_salary_report(uploaded_files):
    # 教練單價表
    COACH_PRICING = {
        "林意潔": {"團1-2人": 800, "團3人": 900, "團4人": 900, "團5人": 950, "團6人": 1000, "1對2(1.5hr)": 1275, "1對2": 850, "1對1(1.5hr)": 1275, "1對1": 850},
        "陳秀蓉": {"團1-2人": 800, "團3人": 850, "團4人": 850, "團5人": 900, "團6人": 950, "1對2(1.5hr)": 1200, "1對2": 800, "1對1(1.5hr)": 1200, "1對1": 800},
        "陳怡廷": {"團1-2人": 800, "團3人": 850, "團4人": 900, "團5人": 950, "團6人": 1000, "1對2(1.5hr)": 1275, "1對2": 850, "1對1(1.5hr)": 1275, "1對1": 850},
        "鍾佳蓁 Rita": {"團1-2人": 900, "團3人": 950, "團4人": 1000, "團5人": 1050, "團6人": 1100, "1對2(1.5hr)": 1425, "1對2": 950, "1對1(1.5hr)": 1425, "1對1": 950},
        "黃宛婷": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "楊子慧(小在)": {"團1-2人": 650, "團3人": 800, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1050, "1對2": 700, "1對1(1.5hr)": 1050, "1對1": 700},
        "許力尹 LOUIS": {"團1-2人": 600, "團3人": 800, "團4人": 800, "團5人": 800, "團6人": 800, "1對2(1.5hr)": 1200, "1對2": 800, "1對1(1.5hr)": 1200, "1對1": 800},
        "顥顥": {"團1-2人": 0, "團3人": 1000, "團4人": 1000, "團5人": 1100, "團6人": 1100, "1對2(1.5hr)": 1500, "1對2": 1000, "1對1(1.5hr)": 1500, "1對1": 1000},
        "洪睿絃": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "紀儒蓁": {"團1-2人": 700, "團3人": 800, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "李翎瑋": {"團1-2人": 800, "團3人": 850, "團4人": 900, "團5人": 950, "團6人": 1000, "1對2(1.5hr)": 1275, "1對2": 850, "1對1(1.5hr)": 1275, "1對1": 850},
        "郭奕伶": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "郭品均": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "邴妍語": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "張鈞弼": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 850, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "蕭竣升": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "紀萃文": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "李函豫": {"團1-2人": 600, "團3人": 650, "團4人": 700, "團5人": 750, "團6人": 800, "1對2(1.5hr)": 975, "1對2": 650, "1對1(1.5hr)": 975, "1對1": 650},
        "尤子綺": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "張楷翌": {"團1-2人": 600, "團3人": 650, "團4人": 700, "團5人": 750, "團6人": 800, "1對2(1.5hr)": 975, "1對2": 650, "1對1(1.5hr)": 975, "1對1": 650},
        "侯懿庭": {"團1-2人": 700, "團3人": 750, "團4人": 800, "團5人": 850, "團6人": 900, "1對2(1.5hr)": 1125, "1對2": 750, "1對1(1.5hr)": 1125, "1對1": 750},
        "謝俐池": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "黃姿菁": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "籃郁雯": {"團1-2人": 550, "團3人": 600, "團4人": 700, "團5人": 700, "團6人": 700, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "徐漫": {"團1-2人": 500, "團3人": 550, "團4人": 600, "團5人": 650, "團6人": 700, "1對2(1.5hr)": 825, "1對2": 550, "1對1(1.5hr)": 825, "1對1": 550},
        "鄭筠馨": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "高舒涵": {"團1-2人": 550, "團3人": 600, "團4人": 650, "團5人": 700, "團6人": 750, "1對2(1.5hr)": 900, "1對2": 600, "1對1(1.5hr)": 900, "1對1": 600},
        "邱靜瑜": {"團1-2人": 600, "團3人": 650, "團4人": 700, "團5人": 750, "團6人": 800, "1對2(1.5hr)": 975, "1對2": 650, "1對1(1.5hr)": 975, "1對1": 650}
    }

    NAME_CONVERSION = {
        "意潔": "林意潔", "Vivi": "陳秀蓉", "怡廷": "陳怡廷", "佳蓁": "鍾佳蓁 Rita", "宛婷": "黃宛婷",
        "小在": "楊子慧(小在)", "LOUIS": "許力尹 LOUIS", "顥顥": "顥顥", "睿絃": "洪睿絃", "儒蓁": "紀儒蓁",
        "翎瑋": "李翎瑋", "奕伶": "郭奕伶", "品均": "郭品均", "妍語": "邴妍語", "鈞弼": "張鈞弼",
        "竣升": "蕭竣升", "萃萃": "紀萃文", "函豫": "李函豫", "子綺": "尤子綺", "楷翌": "張楷翌",
        "懿庭": "侯懿庭", "俐池": "謝俐池", "姿菁": "黃姿菁", "郁雯": "籃郁雯", "徐漫": "徐漫",
        "筠馨": "鄭筠馨", "舒涵": "高舒涵", "靜瑜": "邱靜瑜"
    }
    SENIORITY_DATA = {
        "林意潔": "兩年以上", "陳秀蓉": "兩年以上", "陳怡廷": "兩年以上", "鍾佳蓁 Rita": "兩年以上",
        "黃宛婷": "一年以上至兩年以下", "楊子慧(小在)": "一年以上至兩年以下", "許力尹 LOUIS": "一年以上至兩年以下",
        "顥顥": "一年以上至兩年以下", "洪睿絃": "一年以上至兩年以下", "紀儒蓁": "兩年以上",
        "李翎瑋": "一年以上至兩年以下", "郭奕伶": "一年以上至兩年以下", "郭品均": "一年以上至兩年以下",
        "邴妍語": "一年以上至兩年以下", "張鈞弼": "一年以上至兩年以下", "蕭竣升": "一年以上至兩年以下",
        "紀萃文": "一年以上至兩年以下", "李函豫": "一年以上至兩年以下", "尤子綺": "一年以下",
        "張楷翌": "一年以上至兩年以下", "侯懿庭": "一年以下", "謝俐池": "一年以下",
        "黃姿菁": "一年以下", "籃郁雯": "一年以下", "徐漫": "一年以下",
        "鄭筠馨": "一年以下", "高舒涵": "一年以下", "邱靜瑜": "一年以下"
    }
    course_types = ["團1-2人", "團3人", "團4人", "團5人", "團6人", "1對2(1.5hr)", "1對2", "1對1(1.5hr)", "1對1"]
    columns = [
        "教練", "課程", "單價", "義昌館堂數", "義昌館金額", "高美館堂數", "高美館金額", "中山館堂數", "中山館金額",
        "堂數達標獎金", "三館總堂數", "三館總金額", "應付金額", "堂數達標獎金(加項)", "總計", "執行業務(扣款)", "補充保費(扣款)", "應付薪資", "年資"
    ]

    all_rows = []
    for coach in COACH_PRICING.keys():
        for course in course_types:
            row = {col: 0 for col in columns}
            row["教練"], row["課程"], row["單價"], row["年資"] = coach, course, COACH_PRICING[coach].get(course, 0), ""
            all_rows.append(row)
    df_master = pd.DataFrame(all_rows)

    coach_seniority = {}
    for file in uploaded_files:
        df_info = pd.read_excel(file, sheet_name='統計總表', nrows=1, header=None)
        loc_name = str(df_info.iloc[0, 1]).strip()
        file.seek(0)
        df_stats = pd.read_excel(file, sheet_name='統計總表', skiprows=2).fillna(0)
        
        s_col, a_col = f"{loc_name}堂數", f"{loc_name}金額"
        for _, s_row in df_stats.iterrows():
            raw_name = str(s_row['姓名']).strip().split(' ')[0]
            mapped_name = NAME_CONVERSION.get(raw_name, raw_name)
            if '年資' in s_row and s_row['年資'] != 0:
                coach_seniority[mapped_name] = str(s_row['年資']).strip()
            mask = (df_master["教練"] == mapped_name)
            for course in course_types:
                if course in s_row:
                    val = s_row[course]
                    df_master.loc[mask & (df_master["課程"] == course), s_col] = val
                    df_master.loc[mask & (df_master["課程"] == course), a_col] = val * COACH_PRICING.get(mapped_name, {}).get(course, 0)

    final_dfs = []
    for coach in COACH_PRICING.keys():
        c_df = df_master[df_master["教練"] == coach].copy()
        seniority_val = SENIORITY_DATA.get(coach, coach_seniority.get(coach, ""))
        c_df["三館總堂數"] = c_df[["義昌館堂數", "高美館堂數", "中山館堂數"]].sum(axis=1)
        c_df["三館總金額"] = c_df[["義昌館金額", "高美館金額", "中山館金額"]].sum(axis=1)
        base_pay = c_df["三館總金額"].sum()
        total_sum = base_pay 
        tax_10 = int(total_sum * 0.1) if total_sum >= 20000 else 0
        health_211 = int(total_sum * 0.0211) if total_sum >= 20000 else 0
        final_pay = total_sum - tax_10 - health_211
        c_df["年資"] = seniority_val
        for col, val in zip(["應付金額", "總計", "執行業務(扣款)", "補充保費(扣款)", "應付薪資"], [base_pay, total_sum, tax_10, health_211, final_pay]):
            c_df[col] = val
        final_dfs.append(c_df)
        sub_row = {col: 0 for col in columns}
        sub_row["教練"], sub_row["課程"], sub_row["應付金額"], sub_row["總計"], sub_row["應付薪資"], sub_row["年資"] = coach, "小計", base_pay, total_sum, final_pay, seniority_val
        final_dfs.append(pd.DataFrame([sub_row]))

    df_final = pd.concat(final_dfs, ignore_index=True)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='薪資清單')
        ws, thin = writer.sheets['薪資清單'], Side(style='thin')
        border, center = Border(top=thin, left=thin, right=thin, bottom=thin), Alignment(horizontal='center', vertical='center')
        for cell in ws[1]:
            cell.fill, cell.font, cell.alignment, cell.border = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"), Font(bold=True), center, border
        start_r, pay_idx, senior_idx = 2, df_final.columns.get_loc("應付金額") + 1, df_final.columns.get_loc("年資") + 1
        for r in range(2, ws.max_row + 1):
            coach_val, next_coach = ws.cell(row=r, column=1).value, ws.cell(row=r + 1, column=1).value if r < ws.max_row else None
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).border, ws.cell(row=r, column=c).alignment = border, center
                if ws.cell(row=r, column=2).value == "小計": ws.cell(row=r, column=c).fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            if next_coach != coach_val:
                ws.merge_cells(start_row=start_r, start_column=1, end_row=r, end_column=1)
                for c in range(pay_idx, senior_idx + 1):
                    ws.merge_cells(start_row=start_r, start_column=c, end_row=r, end_column=c)
                start_r = r + 1
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 15
    return output.getvalue()

st.title("教練薪資結算系統")
uploaded_files = st.file_uploader("上傳預約統計表 (.xlsx)", type=["xlsx"], accept_multiple_files=True)
if uploaded_files:
    if st.button("開始計算薪資並排版"):
        final_excel_data = generate_perfect_salary_report(uploaded_files)
        st.download_button(label="下載薪資明細表", data=final_excel_data, file_name="教練薪資明細_最終版.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
