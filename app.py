import streamlit as st
import pandas as pd
import re
import math
import io
import msoffice_crypto_tool

# === 고정 비밀번호 설정 (사업자번호) ===
FILE_PASSWORD = "2598801569"

# === 1. 함수 정의 ===
def normalize_name(name):
    """이름 정규화 (숫자, 괄호 제거)"""
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

def decrypt_file(file_obj):
    """파일이 암호화되어 있다면 해제하여 반환"""
    file_obj.seek(0)
    try:
        # 1. 암호화된 파일인지 확인 및 해제 시도
        office_file = msoffice_crypto_tool.OfficeFile(file_obj)
        office_file.load_key(password=FILE_PASSWORD)
        
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        decrypted.name = file_obj.name # 원래 파일명 유지
        return decrypted
    except Exception:
        # 2. 암호화되지 않았거나(일반 파일), 다른 오류라면 원본 그대로 반환
        file_obj.seek(0)
        return file_obj

def find_header_row(df):
    """
    데이터프레임에서 실제 헤더가 있는 행 번호를 찾는다.
    쿠팡: '기사부담 고용보험' 또는 '성함'과 '총 정산금액'이 있는 줄
    배민: '라이더명'과 '처리건수'가 있는 줄
    """
    for i, row in df.iterrows():
        row_str = row.astype(str).values
        row_joined = " ".join(row_str)
        
        # 쿠팡 헤더 특징
        if '기사부담' in row_joined and '고용보험' in row_joined:
            return i, 'coupang'
        if '성함' in row_joined and '총 정산금액' in row_joined:
            return i, 'coupang'
            
        # 배민 헤더 특징
        if '라이더명' in row_joined and '처리건수' in row_joined:
            return i, 'baemin'
        if '라이더명' in row_joined and 'C(A+B)' in row_joined:
            return i, 'baemin'
            
    return -1, None

# === 2. 스트림릿 화면 구성 ===
st.set_page_config(page_title="빅스텝 주차 정산기", layout="wide")

st.markdown("""
<style>
    .main > div { padding-top: 2rem; }
    .stButton>button { width: 100%; margin-top: 20px; background-color: #FF4B4B; color: white; font-size: 18px; padding: 10px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 빅스텝 통합 주차 정산서 생성기")
st.markdown(f"### 엑셀 파일 업로드 (비밀번호 자동해제)")
st.info(f"비밀번호(`{FILE_PASSWORD}`)가 걸린 파일도 그대로 올리시면 됩니다. (개수 무제한, 자동 분류)")

# 파일 업로더 (여러 파일 허용)
uploaded_files = st.file_uploader("엑셀 파일들을 이곳에 놓으세요", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    # 0. 파일 전처리 (암호 해제) 및 분류
    coupang_files = []
    baemin_files = []
    unknown_files = []
    
    # 처리된 파일 객체들을 저장할 리스트 (나중에 다시 읽기 위함)
    processed_files_map = [] # (file_obj, file_type, header_idx)

    for f in uploaded_files:
        # 암호 해제 시도
        unlocked_f = decrypt_file(f)
        
        # 일단 읽어서 분류
        try:
            df_raw = pd.read_excel(unlocked_f, header=None, engine='openpyxl')
            header_idx, ftype = find_header_row(df_raw)
            
            if header_idx != -1:
                processed_files_map.append((unlocked_f, ftype, header_idx))
                if ftype == 'coupang':
                    coupang_files.append(unlocked_f)
                else:
                    baemin_files.append(unlocked_f)
            else:
                unknown_files.append(f.name)
        except Exception as e:
            unknown_files.append(f"{f.name} (읽기 실패)")

    # 2. 분류 결과 표시
    col1, col2 = st.columns(2)
    with col1:
        st.success(f"📦 **쿠팡 파일 ({len(coupang_files)}개)**")
        for cf in coupang_files: st.caption(f"- {cf.name}")
    with col2:
        st.info(f"🛵 **배민 파일 ({len(baemin_files)}개)**")
        for bf in baemin_files: st.caption(f"- {bf.name}")
    
    if unknown_files:
        st.warning(f"⚠️ 인식 불가 파일: {unknown_files}")

    # 3. 정산 버튼
    if coupang_files or baemin_files:
        if st.button("🚀 정산서 통합 생성하기"):
            try:
                all_data = {}

                # 분류된 파일들을 순회하며 데이터 추출
                for f_obj, ftype, h_idx in processed_files_map:
                    f_obj.seek(0)
                    df = pd.read_excel(f_obj, header=None, engine='openpyxl')
                    header_row = df.iloc[h_idx].astype(str).tolist()

                    if ftype == 'coupang':
                        # --- [A] 쿠팡 처리 ---
                        idx_name = find_col_idx(header_row, '성함')
                        if idx_name == -1: idx_name = 2
                        
                        idx_orders = find_col_idx(header_row, '오더수')
                        idx_total_1 = find_col_idx(header_row, '총 정산금액')
                        idx_total_2 = find_col_idx(header_row, '정산금액', exclude_keyword='총')
                        idx_emp = find_col_idx(header_row, '기사부담 고용보험')
                        idx_ind = find_col_idx(header_row, '기사부담 산재보험')
                        idx_hourly = find_col_idx(header_row, '시간제보험')
                        idx_retro = find_col_idx(header_row, '보험료 소급')

                        for i in range(h_idx + 1, len(df)):
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

                    elif ftype == 'baemin':
                        # --- [B] 배민 처리 ---
                        idx_orders = find_col_idx(header_row, '처리건수')
                        idx_total = find_col_idx(header_row, 'C(A+B)')
                        idx_emp = find_col_idx(header_row, '라이더부담\n고용보험료')
                        idx_ind = find_col_idx(header_row, '라이더부담\n산재보험료')
                        idx_hourly = find_col_idx(header_row, '시간제보험료')
                        idx_retro_f = find_col_idx(header_row, '(F)')
                        idx_retro_g = find_col_idx(header_row, '(G)')
                        
                        idx_name_b = find_col_idx(header_row, '라이더명')
                        if idx_name_b == -1: idx_name_b = 2

                        for i in range(h_idx + 1, len(df)):
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

                # 메모리에 엑셀 저장
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df_out.to_excel(writer, index=False, sheet_name='정산서')

                wb = writer.book
                ws = writer.sheets['정산서']
                fmt_num = wb.add_format({'num_format': '#,##0'})
                fmt_hide_zero = wb.add_format({'num_format': '#,##0;-#,##0;""'})

                # 서식 및 수식 적용 (v8 동일)
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
    # 안내 메시지 (파일 올리는 중)
    st.info("파일을 분석 중입니다... 잠시만 기다려주세요.")