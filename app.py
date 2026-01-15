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

def classify_file(file_obj):
    """파일 정밀 분석 (시트명 -> 내용 순서로 확인)"""
    try:
        file_obj.seek(0)
        # 1단계: 시트 이름으로 구분 (가장 확실)
        xl = pd.ExcelFile(file_obj, engine='openpyxl')
        sheet_names = xl.sheet_names
        
        if '종합' in sheet_names: 
            return 'coupang'
        if any('을지' in s for s in sheet_names): 
            return 'baemin'

        # 2단계: 시트 이름으로 안 되면 내용으로 구분
        file_obj.seek(0)
        df_temp = pd.read_excel(file_obj, header=None, engine='openpyxl', nrows=30)
        file_str = df_temp.astype(str).to_string()

        # 쿠팡만의 키워드
        if '총 정산 오더수' in file_str and '기사부담 고용보험' in file_str:
            return 'coupang'
        
        # 배민만의 키워드 (을지, 갑지 등)
        if 'C(A+B)' in file_str or '라이더부담' in file_str:
            return 'baemin'
            
        return None
    except Exception:
        return None

# === 2. 스트림릿 화면 구성 ===
st.set_page_config(page_title="빅스텝 주차 정산기", layout="wide")

st.markdown("""
<style>
    .main > div { padding-top: 2rem; }
    .stButton>button { width: 100%; margin-top: 20px; background-color: #FF4B4B; color: white; font-size: 18px; padding: 10px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 빅스텝 통합 주차 정산서 생성기")
st.markdown("### 엑셀 파일 업로드 (자동 분류)")
st.info("쿠팡, 배민 파일을 개수 상관없이 드래그해서 넣어주세요. (파일명이 달라도 내용 보고 알아서 분류합니다)")

# 파일 업로더
uploaded_files = st.file_uploader("엑셀 파일들을 이곳에 놓으세요", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    # 1. 파일 분류 단계
    coupang_files = []
    baemin_files = []
    unknown_files = []

    for f in uploaded_files:
        ftype = classify_file(f)
        f.seek(0) # 커서 초기화
        if ftype == 'coupang':
            coupang_files.append(f)
        elif ftype == 'baemin':
            baemin_files.append(f)
        else:
            unknown_files.append(f.name)

    # 2. 분류 결과 표시
    col1, col2 = st.columns(2)
    with col1:
        st.success(f"📦 **쿠팡 파일 ({len(coupang_files)}개)**")
        for cf in coupang_files: st.caption(f"- {cf.name}")
    with col2:
        st.info(f"🛵 **배민 파일 ({len(baemin_files)}개)**")
        for bf in baemin_files: st.caption(f"- {bf.name}")
    
    if unknown_files:
        st.warning(f"⚠️ 인식 불가 파일 (건너뜀): {unknown_files}")

    # 3. 정산 버튼
    if coupang_files or baemin_files:
        if st.button("🚀 정산서 통합 생성하기"):
            try:
                all_data = {}

                # --- [A] 쿠팡 파일들 처리 ---
                for c_file in coupang_files:
                    c_file.seek(0)
                    # 시트 이름이 '종합'인지 확인하고 읽기, 없으면 첫 번째 시트 읽기
                    try:
                        df = pd.read_excel(c_file, sheet_name='종합', header=None, engine='openpyxl')
                    except:
                        df = pd.read_excel(c_file, header=None, engine='openpyxl')

                    # 헤더 위치 동적 탐색
                    header_row_idx = -1
                    for i, row in df.iterrows():
                        if '기사부담 고용보험' in str(row.values):
                            header_row_idx = i
                            break
                    
                    if header_row_idx == -1: continue # 헤더 못 찾으면 건너뜀

                    header_row = df.iloc[header_row_idx].astype(str).tolist()
                    
                    idx_name = 2 # 이름은 보통 2번째
                    idx_orders = find_col_idx(header_row, '오더수') # '총 정산 오더수' 등
                    idx_total_1 = find_col_idx(header_row, '총 정산금액')
                    idx_total_2 = find_col_idx(header_row, '정산금액', exclude_keyword='총')
                    idx_emp = find_col_idx(header_row, '기사부담 고용보험')
                    idx_ind = find_col_idx(header_row, '기사부담 산재보험')
                    idx_hourly = find_col_idx(header_row, '시간제보험')
                    idx_retro = find_col_idx(header_row, '보험료 소급')

                    for i in range(header_row_idx + 1, len(df)): # 헤더 다음 줄부터
                        row = df.iloc[i]
                        name = normalize_name(row[idx_name])
                        if not name or name == 'nan': continue
                        
                        orders = clean_num(row[idx_orders]) if idx_orders != -1 else 0
                        
                        raw_total = 0
                        if idx_total_1 != -1: raw_total = clean_num(row[idx_total_1])
                        if raw_total == 0 and orders > 0 and idx_total_2 != -1:
                            raw_total = clean_num(row[idx_total_2])
                        
                        net_total = raw_total 
                        
                        emp = abs(clean_num(row[idx_emp])) if idx_emp != -1 else 0
                        ind = abs(clean_num(row[idx_ind])) if idx_ind != -1 else 0
                        hourly = abs(clean_num(row[idx_hourly])) if idx_hourly != -1 else 0
                        retro = abs(clean_num(row[idx_retro])) if idx_retro != -1 else 0

                        if name not in all_data: 
                            all_data[name] = {'c_orders':0, 'c_total':0, 'c_emp':0, 'c_ind':0, 'c_hourly':0, 'c_retro':0,
                                              'b_orders':0, 'b_total':0, 'b_emp':0, 'b_ind':0, 'b_hourly':0, 'b_retro':0}
                        
                        all_data[name]['c_orders'] += orders
                        all_data[name]['c_total'] += net_total
                        all_data[name]['c_emp'] += emp
                        all_data[name]['c_ind'] += ind
                        all_data[name]['c_hourly'] += hourly
                        all_data[name]['c_retro'] += retro

                # --- [B] 배민 파일들 처리 ---
                for b_file in baemin_files:
                    b_file.seek(0)
                    try:
                        df = pd.read_excel(b_file, sheet_name='을지_협력사 소속 라이더 정산 확인용', header=None, engine='openpyxl')
                    except:
                        # 시트명 다르면 '을지'가 포함된 시트 찾기
                        xl = pd.ExcelFile(b_file, engine='openpyxl')
                        target_sheet = next((s for s in xl.sheet_names if '을지' in s), xl.sheet_names[0])
                        df = pd.read_excel(b_file, sheet_name=target_sheet, header=None, engine='openpyxl')

                    # 헤더 위치 동적 탐색
                    header_row_idx = -1
                    for i, row in df.iterrows():
                        if 'C(A+B)' in str(row.values) or '처리건수' in str(row.values):
                            header_row_idx = i
                            break
                    
                    if header_row_idx == -1: continue

                    header_row = df.iloc[header_row_idx].astype(str).tolist()
                    
                    idx_orders = find_col_idx(header_row, '처리건수')
                    idx_total = find_col_idx(header_row, 'C(A+B)')
                    idx_emp = find_col_idx(header_row, '라이더부담\n고용보험료')
                    idx_ind = find_col_idx(header_row, '라이더부담\n산재보험료')
                    idx_hourly = find_col_idx(header_row, '시간제보험료')
                    idx_retro_f = find_col_idx(header_row, '(F)')
                    idx_retro_g = find_col_idx(header_row, '(G)')
                    
                    # 배민 이름 컬럼 찾기 (보통 '라이더명')
                    idx_name_b = find_col_idx(header_row, '라이더명')
                    if idx_name_b == -1: idx_name_b = 2 # 못 찾으면 기본값

                    for i in range(header_row_idx + 1, len(df)):
                        row = df.iloc[i]
                        name = normalize_name(row[idx_name_b])
                        if not name or name == 'nan': continue
                        
                        orders = clean_num(row[idx_orders]) if idx_orders != -1 else 0
                        raw_total = clean_num(row[idx_total]) if idx_total != -1 else 0
                        
                        fee = orders * 100
                        net_total = raw_total - fee
                        
                        emp = clean_num(row[idx_emp]) if idx_emp != -1 else 0
                        ind = clean_num(row[idx_ind]) if idx_ind != -1 else 0
                        hourly = clean_num(row[idx_hourly]) if idx_hourly != -1 else 0
                        
                        retro_f = clean_num(row[idx_retro_f]) if idx_retro_f != -1 else 0
                        retro_g = clean_num(row[idx_retro_g]) if idx_retro_g != -1 else 0
                        retro = abs(retro_f + retro_g)

                        if name not in all_data: 
                            all_data[name] = {'c_orders':0, 'c_total':0, 'c_emp':0, 'c_ind':0, 'c_hourly':0, 'c_retro':0,
                                              'b_orders':0, 'b_total':0, 'b_emp':0, 'b_ind':0, 'b_hourly':0, 'b_retro':0}

                        all_data[name]['b_orders'] += orders
                        all_data[name]['b_total'] += net_total
                        all_data[name]['b_emp'] += emp
                        all_data[name]['b_ind'] += ind
                        all_data[name]['b_hourly'] += hourly
                        all_data[name]['b_retro'] += retro

                # === 엑셀 생성 ===
                final_rows = []
                sorted_names = sorted(all_data.keys())

                for name in sorted_names:
                    d = all_data[name]
                    
                    c_total = d['c_total']
                    b_total = d['b_total']
                    c_promo, b_promo, reward = 0, 0, 0
                    
                    final_sum = c_total + b_total + c_promo + b_promo + reward
                    tax = math.floor(final_sum * 0.03 / 10) * 10
                    local_tax = math.floor(final_sum * 0.003 / 10) * 10
                    total_retro = d['c_retro'] + d['b_retro']
                    
                    ins_sum = (d['c_emp'] + d['b_emp'] + d['c_ind'] + d['b_ind'] + d['c_hourly'] + d['b_hourly'])
                    final_pay = final_sum - ins_sum + total_retro - tax - local_tax

                    final_rows.append({
                        '성함': name,
                        '쿠팡 오더수': d['c_orders'],
                        '배민 오더수': d['b_orders'],
                        '쿠팡 총금액': c_total,
                        '배민 총금액': b_total,
                        '쿠팡 프로모션': c_promo,
                        '배민 프로모션': b_promo,
                        '리워드': reward,
                        '최종합산': final_sum,
                        '쿠팡 고용보험': d['c_emp'],
                        '쿠팡 산재보험': d['c_ind'],
                        '배민 고용보험': d['b_emp'],
                        '배민 산재보험': d['b_ind'],
                        '쿠팡 시간제 보험': d['c_hourly'],
                        '배민 시간제 보험': d['b_hourly'],
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

                ws.set_column('A:A', 12)
                ws.set_column('B:E', 14, fmt_num)
                ws.set_column('F:H', 14, fmt_hide_zero)
                ws.set_column('I:R', 14, fmt_num)
                ws.set_column('S:S', 14, fmt_hide_zero)
                ws.set_column('T:T', 14, fmt_num)

                for i in range(len(df_out)):
                    row = i + 2
                    val_sum = df_out.iloc[i]['최종합산']
                    val_tax = df_out.iloc[i]['소득세']
                    val_local = df_out.iloc[i]['지방소득세']
                    val_final = df_out.iloc[i]['최종지급(액)']

                    ws.write_formula(f'I{row}', f'=D{row}+E{row}+F{row}+G{row}+H{row}', fmt_num, val_sum)
                    ws.write_formula(f'Q{row}', f'=ROUNDDOWN(I{row}*0.03, -1)', fmt_num, val_tax)
                    ws.write_formula(f'R{row}', f'=ROUNDDOWN(I{row}*0.003, -1)', fmt_num, val_local)
                    ws.write_formula(f'T{row}', f'=I{row}-(J{row}+K{row}+L{row}+M{row}+N{row}+O{row})+P{row}-(Q{row}+R{row})-S{row}', fmt_num, val_final)

                writer.close()
                output.seek(0)

                st.write("---")
                st.success(f"🎉 정산서 통합 생성이 완료되었습니다! (총 {len(final_rows)}명)")
                st.download_button(
                    label="📥 엑셀 파일 다운로드 (Click)",
                    data=output,
                    file_name='빅스텝_통합_주차정산서_최종.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")

elif uploaded_files:
    st.info("파일을 분석 중입니다...")