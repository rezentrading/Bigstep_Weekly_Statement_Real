import streamlit as st
import pandas as pd
import re
import math
import io

# === 1. 함수 정의 ===
def normalize_name(name):
    """이름 정규화"""
    if pd.isna(name): return ""
    name = str(name)
    name = re.sub(r'\d+', '', name)
    name = re.sub(r'\(.*?\)', '', name)
    return name.strip().replace(" ", "")

def clean_num(x):
    """숫자 변환"""
    if pd.isna(x) or x == '': return 0
    try:
        return float(str(x).replace(',', ''))
    except:
        return 0

def find_col_idx(headers, keyword, exclude_keyword=None):
    for i, h in enumerate(headers):
        if keyword in str(h):
            if exclude_keyword and exclude_keyword in str(h):
                continue
            return i
    return -1

# === 2. 스트림릿 화면 구성 ===
st.set_page_config(page_title="빅스텝 주차 정산기", layout="wide")

st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
    }
    .stButton>button {
        width: 100%;
        margin-top: 20px;
        background-color: #FF4B4B;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 빅스텝 통합 주차 정산서 생성기")
st.markdown("### 1. 쿠팡 & 배민 엑셀 파일 업로드")
st.info("쿠팡 파일과 배민 파일 2개를 동시에 선택해서 드래그하거나 업로드하세요.")

# 파일 업로더
uploaded_files = st.file_uploader("엑셀 파일 2개 업로드", accept_multiple_files=True, type=['xlsx'])

if len(uploaded_files) == 2:
    st.success(f"📂 파일 2개가 준비되었습니다.")
    
    if st.button("🚀 정산서 생성하기"):
        try:
            # === 파일 구분 로직 ===
            coupang_file = None
            baemin_file = None
            
            for f in uploaded_files:
                f.seek(0)
                df_temp = pd.read_excel(f, header=None, engine='openpyxl', nrows=50)
                
                header_row_idx = -1
                for i, row in df_temp.iterrows():
                    row_str = row.astype(str).values
                    if '기사부담 고용보험' in str(row_str) or '라이더부담\n고용보험료' in str(row_str):
                        header_row_idx = i
                        break
                
                is_coupang = False
                if header_row_idx != -1:
                    header_list = df_temp.iloc[header_row_idx].astype(str).tolist()
                    col_idx = -1
                    for idx, h in enumerate(header_list):
                        if '고용보험' in h and ('기사' in h or '라이더' in h):
                            col_idx = idx
                            break
                    
                    if col_idx != -1:
                        for k in range(header_row_idx + 1, min(header_row_idx + 6, len(df_temp))):
                            val = clean_num(df_temp.iloc[k, col_idx])
                            if val < 0:
                                is_coupang = True
                                break
                
                f.seek(0)
                if is_coupang:
                    coupang_file = f
                else:
                    baemin_file = f

            if not coupang_file or not baemin_file:
                # 시트명으로 2차 시도
                for f in uploaded_files:
                    f.seek(0)
                    xl = pd.ExcelFile(f, engine='openpyxl')
                    if '종합' in xl.sheet_names:
                        coupang_file = f
                    elif any('을지' in s for s in xl.sheet_names):
                        baemin_file = f
                    f.seek(0)

            if not coupang_file or not baemin_file:
                st.error("❌ 파일을 구분할 수 없습니다. 쿠팡/배민 파일이 맞는지 확인해주세요.")
                st.stop()

            # === 데이터 처리 ===
            all_data = {}

            # [쿠팡 처리]
            coupang_file.seek(0)
            df_c = pd.read_excel(coupang_file, sheet_name='종합', header=None, engine='openpyxl')
            header_row = df_c.iloc[8].astype(str).tolist()
            
            idx_name = 2
            idx_orders = 5
            idx_total_1 = find_col_idx(header_row, '총 정산금액')
            idx_total_2 = find_col_idx(header_row, '정산금액', exclude_keyword='총')
            idx_emp_rider = find_col_idx(header_row, '기사부담 고용보험')
            idx_ind_rider = find_col_idx(header_row, '기사부담 산재보험')
            idx_hourly = find_col_idx(header_row, '시간제보험')
            idx_retro = find_col_idx(header_row, '보험료 소급')

            for i in range(16, len(df_c)):
                row = df_c.iloc[i]
                name = normalize_name(row[idx_name])
                if not name or name == 'nan': continue
                
                orders = clean_num(row[idx_orders])
                raw_total = clean_num(row[idx_total_1])
                if raw_total == 0 and orders > 0 and idx_total_2 != -1:
                    raw_total = clean_num(row[idx_total_2])
                
                # 쿠팡: 수수료 차감 없음
                net_total = raw_total 
                emp_rider = abs(clean_num(row[idx_emp_rider]))
                ind_rider = abs(clean_num(row[idx_ind_rider]))
                hourly = abs(clean_num(row[idx_hourly]))
                retro = abs(clean_num(row[idx_retro]))
                
                if name not in all_data: all_data[name] = {}
                all_data[name].update({
                    'c_orders': orders,
                    'c_total': net_total,
                    'c_emp': emp_rider,
                    'c_ind': ind_rider,
                    'c_hourly': hourly,
                    'c_retro': retro
                })

            # [배민 처리]
            baemin_file.seek(0)
            df_b = pd.read_excel(baemin_file, sheet_name='을지_협력사 소속 라이더 정산 확인용', header=None, engine='openpyxl')
            header_row = df_b.iloc[17].astype(str).tolist()
            
            idx_orders = find_col_idx(header_row, '처리건수')
            idx_total = find_col_idx(header_row, 'C(A+B)')
            idx_emp_rider = find_col_idx(header_row, '라이더부담\n고용보험료')
            idx_ind_rider = find_col_idx(header_row, '라이더부담\n산재보험료')
            idx_hourly = find_col_idx(header_row, '시간제보험료')
            idx_retro_f = find_col_idx(header_row, '(F)')
            idx_retro_g = find_col_idx(header_row, '(G)')
            
            for i in range(19, len(df_b)):
                row = df_b.iloc[i]
                name = normalize_name(row[2])
                if not name or name == 'nan': continue
                
                orders = clean_num(row[idx_orders])
                raw_total = clean_num(row[idx_total])
                
                # 배민: 수수료(100원) 차감
                fee = orders * 100
                net_total = raw_total - fee
                
                emp_rider = clean_num(row[idx_emp_rider])
                ind_rider = clean_num(row[idx_ind_rider])
                hourly = clean_num(row[idx_hourly])
                retro = abs(clean_num(row[idx_retro_f]) + clean_num(row[idx_retro_g]))
                
                if name not in all_data: all_data[name] = {}
                all_data[name].update({
                    'b_orders': orders,
                    'b_total': net_total,
                    'b_emp': emp_rider,
                    'b_ind': ind_rider,
                    'b_hourly': hourly,
                    'b_retro': retro
                })

            # === 엑셀 생성 (v8 로직) ===
            final_rows = []
            sorted_names = sorted(all_data.keys())

            for name in sorted_names:
                d = all_data[name]
                c_total = d.get('c_total', 0)
                b_total = d.get('b_total', 0)
                c_promo, b_promo, reward = 0, 0, 0
                
                final_sum = c_total + b_total + c_promo + b_promo + reward
                tax = math.floor(final_sum * 0.03 / 10) * 10
                local_tax = math.floor(final_sum * 0.003 / 10) * 10
                total_retro = d.get('c_retro', 0) + d.get('b_retro', 0)
                
                ins_sum = (d.get('c_emp', 0) + d.get('b_emp', 0) + d.get('c_ind', 0) + 
                           d.get('b_ind', 0) + d.get('c_hourly', 0) + d.get('b_hourly', 0))
                
                final_pay = final_sum - ins_sum + total_retro - tax - local_tax

                final_rows.append({
                    '성함': name,
                    '쿠팡 오더수': d.get('c_orders', 0),
                    '배민 오더수': d.get('b_orders', 0),
                    '쿠팡 총금액': c_total,
                    '배민 총금액': b_total,
                    '쿠팡 프로모션': c_promo,
                    '배민 프로모션': b_promo,
                    '리워드': reward,
                    '최종합산': final_sum,
                    '쿠팡 고용보험': d.get('c_emp', 0),
                    '쿠팡 산재보험': d.get('c_ind', 0),
                    '배민 고용보험': d.get('b_emp', 0),
                    '배민 산재보험': d.get('b_ind', 0),
                    '쿠팡 시간제 보험': d.get('c_hourly', 0),
                    '배민 시간제 보험': d.get('b_hourly', 0),
                    '보험료 환급(소급)': total_retro,
                    '소득세': tax,
                    '지방소득세': local_tax,
                    '선지급차감': 0,
                    '최종지급(액)': final_pay
                })

            df_out = pd.DataFrame(final_rows)

            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df_out.to_excel(writer, index=False, sheet_name='정산서')

            wb = writer.book
            ws = writer.sheets['정산서']
            fmt_num = wb.add_format({'num_format': '#,##0'})
            fmt_hide_zero = wb.add_format({'num_format': '#,##0;-#,##0;""'})

            # 컬럼 너비 및 서식
            ws.set_column('A:A', 12) 
            ws.set_column('B:E', 14, fmt_num)
            ws.set_column('F:H', 14, fmt_hide_zero)
            ws.set_column('I:R', 14, fmt_num)
            ws.set_column('S:S', 14, fmt_hide_zero)
            ws.set_column('T:T', 14, fmt_num)

            # 수식 적용
            for i in range(len(df_out)):
                row = i + 2
                val_sum = df_out.iloc[i]['최종합산']
                val_tax = df_out.iloc[i]['소득세']
                val_local = df_out.iloc[i]['지방소득세']
                val_final = df_out.iloc[i]['최종지급(액)']

                # I=D+E+F+G+H
                ws.write_formula(f'I{row}', f'=D{row}+E{row}+F{row}+G{row}+H{row}', fmt_num, val_sum)
                # Q=I*0.03
                ws.write_formula(f'Q{row}', f'=ROUNDDOWN(I{row}*0.03, -1)', fmt_num, val_tax)
                # R=I*0.003
                ws.write_formula(f'R{row}', f'=ROUNDDOWN(I{row}*0.003, -1)', fmt_num, val_local)
                # T=I-(J+K+L+M+N+O)+P-(Q+R)-S
                ws.write_formula(f'T{row}', f'=I{row}-(J{row}+K{row}+L{row}+M{row}+N{row}+O{row})+P{row}-(Q{row}+R{row})-S{row}', fmt_num, val_final)

            writer.close()
            output.seek(0)

            st.write("---")
            st.success("🎉 정산서 생성이 완료되었습니다!")
            st.download_button(
                label="📥 엑셀 파일 다운로드 (Click)",
                data=output,
                file_name='빅스텝_통합_주차정산서.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")

elif len(uploaded_files) > 0:
    st.warning("⚠️ 쿠팡 파일과 배민 파일, 총 2개를 모두 업로드해주세요.")