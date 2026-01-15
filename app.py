import streamlit as st
import pandas as pd
import re
import math
import io
import msoffcrypto
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# ==========================================
# [설정] 구글 시트 주소
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pKrWaGlrAZP1nJLsKFFnUlgOOasCmiKqpovA_t5k2qA/edit?gid=0#gid=0"
# ==========================================

# 고정 설정
FILE_PASSWORD = "2598801569"
LOGIN_PASSWORD = "2598801569"

# === 1. 구글 시트 기록 함수 ===
def log_to_sheet(c_count, b_count):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(SHEET_URL).sheet1
        
        now = datetime.now()
        total_income = (c_count + b_count) * 10
        sheet.append_row([
            now.strftime("%Y-%m-%d"), 
            now.strftime("%H:%M:%S"), 
            "지인(사용자)", 
            c_count, 
            b_count, 
            total_income
        ])
        return True
    except Exception as e:
        st.error(f"⚠️ 구글 시트 기록 실패: {e}")
        return False

# === 2. 유틸리티 함수 ===
def normalize_name(name):
    if pd.isna(name): return ""
    name = str(name)
    name = re.sub(r'\d+', '', name)
    name = re.sub(r'\(.*?\)', '', name)
    return name.strip().replace(" ", "")

def clean_num(x):
    if pd.isna(x) or x == '': return 0
    try: return float(str(x).replace(',', ''))
    except: return 0

def decrypt_file(file_obj):
    file_obj.seek(0)
    try:
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file_obj)
        office_file.load_key(password=FILE_PASSWORD)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        decrypted.name = file_obj.name
        return decrypted
    except:
        file_obj.seek(0)
        return file_obj

def get_sheet_data(file_obj):
    """시트 이름으로 배민/쿠팡 구분"""
    try:
        xl = pd.ExcelFile(file_obj, engine='openpyxl')
        sheet_names = xl.sheet_names
        
        # 1. 배민: '을지' 우선
        for sheet in sheet_names:
            if '을지' in sheet:
                return xl.parse(sheet, header=None), 'baemin'
        
        # 2. 쿠팡: '종합' 우선
        if '종합' in sheet_names:
            return xl.parse('종합', header=None), 'coupang'
            
        return xl.parse(0, header=None), None
    except:
        return pd.DataFrame(), None

def analyze_headers_type(df, detected_type):
    """헤더 행 위치 찾기"""
    for i in range(min(len(df) - 1, 40)):
        row_curr = " ".join(df.iloc[i].astype(str).values).replace(" ", "")
        row_next = " ".join(df.iloc[i+1].astype(str).values).replace(" ", "")
        
        # 쿠팡
        if '총정산오더수' in row_curr and '기사부담' in row_next: return i, i+1, 'coupang'
        if '총정산오더수' in row_curr and '기사부담' in row_curr: return i, i, 'coupang'
        
        # 배민
        if ('라이더명' in row_curr or '성명' in row_curr) and ('처리건수' in row_curr or '배달료' in row_curr): 
            return i, i, 'baemin'
            
    if detected_type == 'baemin':
        for i in range(min(len(df), 40)):
            row_str = " ".join(df.iloc[i].astype(str).values)
            if '라이더명' in row_str or '성명' in row_str:
                return i, i, 'baemin'

    return -1, -1, None

def find_col_in_list(header_list, keywords, exclude=None):
    """특정 리스트 안에서 키워드 찾기"""
    clean_keywords = [k.replace(" ", "") for k in keywords]
    clean_exclude = [e.replace(" ", "") for e in exclude] if exclude else []
    
    for i, val in enumerate(header_list):
        val_str = str(val).replace(" ", "").replace("\n", "")
        
        if all(k in val_str for k in clean_keywords):
            if clean_exclude and any(e in val_str for e in clean_exclude):
                continue
            return i
    return -1

# === 3. 화면 구성 ===
st.set_page_config(page_title="빅스텝 정산 시스템", layout="wide")

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'processed_data' not in st.session_state: st.session_state['processed_data'] = None

# [A] 로그인
if not st.session_state['logged_in']:
    st.title("🔒 빅스텝 정산 시스템")
    pwd = st.text_input("접속 암호", type="password")
    if st.button("로그인"):
        if pwd == LOGIN_PASSWORD:
            st.session_state['logged_in'] = True
            st.rerun()
        else: st.error("암호가 틀렸습니다.")
    st.stop()

# [B] 메인
st.title("📊 빅스텝 통합 주차 정산서 생성기")
st.info("쿠팡, 배민 파일을 모두 드래그해서 넣어주세요. (비밀번호 자동 해제)")

uploaded_files = st.file_uploader("파일 업로드", accept_multiple_files=True, type=['xlsx'], label_visibility="collapsed")

if uploaded_files:
    if st.button("🚀 정산서 분석 및 생성 (1차 확인)"):
        processed_files_map = []
        
        # 1. 파일 분석
        for f in uploaded_files:
            unlocked = decrypt_file(f)
            try:
                df_raw, detected_type = get_sheet_data(unlocked)
                if not df_raw.empty:
                    m_idx, s_idx, ftype = analyze_headers_type(df_raw, detected_type)
                    if m_idx != -1:
                        processed_files_map.append((df_raw, ftype, m_idx, s_idx))
            except: pass
        
        if not processed_files_map:
            st.error("❌ 유효한 정산 파일을 찾지 못했습니다.")
        else:
            all_data = {}
            total_c, total_b = 0, 0
            
            for df, ftype, m_idx, s_idx in processed_files_map:
                data_start = s_idx + 1 
                
                h_main = df.iloc[m_idx].astype(str).tolist()
                h_sub = df.iloc[s_idx].astype(str).tolist()

                if ftype == 'coupang':
                    # [A] 쿠팡 로직 (총금액 찾기 강화: 윗줄/아랫줄 모두 검색)
                    idx_nm = find_col_in_list(h_main, ['성함']); idx_nm = 2 if idx_nm == -1 else idx_nm
                    idx_od = find_col_in_list(h_main, ['총', '정산', '오더수'])
                    if idx_od == -1: idx_od = find_col_in_list(h_main, ['오더수'])
                    
                    # ★ 핵심 수정: 총금액을 윗줄(h_main)과 아랫줄(h_sub) 모두에서 찾음
                    idx_net = find_col_in_list(h_main, ['수수료', '차감'])
                    if idx_net == -1: idx_net = find_col_in_list(h_sub, ['수수료', '차감'])
                    if idx_net == -1: idx_net = find_col_in_list(h_main, ['총', '정산금액'], exclude=['오더'])
                    if idx_net == -1: idx_net = find_col_in_list(h_sub, ['총', '정산금액'], exclude=['오더'])

                    # 보험료는 주로 아랫줄
                    idx_emp = find_col_in_list(h_sub, ['기사부담', '고용보험'])
                    idx_ind = find_col_in_list(h_sub, ['기사부담', '산재보험'])
                    idx_hr = find_col_in_list(h_sub, ['시간제보험'])
                    idx_ret = find_col_in_list(h_sub, ['보험료', '소급'])
                    
                    for i in range(data_start, len(df)):
                        row = df.iloc[i]
                        nm = normalize_name(row[idx_nm])
                        if not nm or nm == 'nan': continue
                        
                        od = clean_num(row[idx_od]) if idx_od != -1 else 0
                        total_c += od
                        
                        # 총금액
                        rt = clean_num(row[idx_net]) if idx_net != -1 else 0
                        
                        ep = abs(clean_num(row[idx_emp])) if idx_emp != -1 else 0
                        id_ = abs(clean_num(row[idx_ind])) if idx_ind != -1 else 0
                        hr = abs(clean_num(row[idx_hr])) if idx_hr != -1 else 0
                        ret = abs(clean_num(row[idx_ret])) if idx_ret != -1 else 0
                        
                        if nm not in all_data: all_data[nm] = {'c_od':0,'c_tot':0,'c_ep':0,'c_id':0,'c_hr':0,'c_ret':0,'b_od':0,'b_tot':0,'b_ep':0,'b_id':0,'b_hr':0,'b_ret':0}
                        all_data[nm]['c_od']+=od; all_data[nm]['c_tot']+=rt; all_data[nm]['c_ep']+=ep; all_data[nm]['c_id']+=id_; all_data[nm]['c_hr']+=hr; all_data[nm]['c_ret']+=ret

                elif ftype == 'baemin':
                    # [B] 배민 로직 (유지)
                    idx_nm = find_col_in_list(h_main, ['라이더명'])
                    if idx_nm == -1: idx_nm = find_col_in_list(h_main, ['성명'])
                    if idx_nm == -1: idx_nm = 2
                    
                    idx_od = find_col_in_list(h_main, ['처리건수'])
                    if idx_od == -1: idx_od = find_col_in_list(h_main, ['배달건수'])

                    idx_tot = find_col_in_list(h_main, ['배달료', 'A']) 
                    if idx_tot == -1: idx_tot = find_col_in_list(h_main, ['배달료']) 

                    idx_ep = find_col_in_list(h_main, ['라이더부담', '고용', '②'])
                    if idx_ep == -1: idx_ep = find_col_in_list(h_main, ['라이더부담', '고용']) 
                    
                    idx_id = find_col_in_list(h_main, ['라이더부담', '산재', '④'])
                    if idx_id == -1: idx_id = find_col_in_list(h_main, ['라이더부담', '산재'])

                    idx_hr = find_col_in_list(h_main, ['시간제', '(D)'])
                    if idx_hr == -1: idx_hr = find_col_in_list(h_main, ['시간제'])
                    
                    # 배민 소급은 로직상 0 (사용자 입력)
                    idx_retro = find_col_in_list(h_main, ['소급'])
                    
                    for i in range(data_start, len(df)):
                        row = df.iloc[i]
                        nm = normalize_name(row[idx_nm])
                        if not nm or nm == 'nan': continue
                        
                        od = clean_num(row[idx_od]) if idx_od != -1 else 0
                        total_b += od
                        
                        raw_tot = clean_num(row[idx_tot]) if idx_tot != -1 else 0
                        fee = od * 100 
                        nt = raw_tot - fee
                        
                        ep = clean_num(row[idx_ep]) if idx_ep != -1 else 0
                        id_ = clean_num(row[idx_id]) if idx_id != -1 else 0
                        hr = clean_num(row[idx_hr]) if idx_hr != -1 else 0
                        
                        if nm not in all_data: all_data[nm] = {'c_od':0,'c_tot':0,'c_ep':0,'c_id':0,'c_hr':0,'c_ret':0,'b_od':0,'b_tot':0,'b_ep':0,'b_id':0,'b_hr':0,'b_ret':0}
                        all_data[nm]['b_od']+=od; all_data[nm]['b_tot']+=nt; all_data[nm]['b_ep']+=ep; all_data[nm]['b_id']+=id_; all_data[nm]['b_hr']+=hr; all_data[nm]['b_ret']+=0 # 소급 0

            # 3. 엑셀 생성
            final_rows = []
            for nm in sorted(all_data.keys()):
                d = all_data[nm]
                f_sum = d['c_tot'] + d['b_tot']
                tax = math.floor(f_sum * 0.03 / 10) * 10
                ltax = math.floor(f_sum * 0.003 / 10) * 10
                t_ret = d['c_ret']
                ins = d['c_ep']+d['b_ep']+d['c_id']+d['b_id']+d['c_hr']+d['b_hr']
                
                final_rows.append({
                    '성함': nm, '쿠팡 오더수': d['c_od'], '배민 오더수': d['b_od'],
                    '쿠팡 총금액': d['c_tot'], '배민 총금액': d['b_tot'],
                    '쿠팡 프로모션': 0, '배민 프로모션': 0, '리워드': 0,
                    '최종합산': f_sum,
                    '쿠팡 고용보험': d['c_ep'], '쿠팡 산재보험': d['c_id'],
                    '배민 고용보험': d['b_ep'], '배민 산재보험': d['b_id'],
                    '쿠팡 시간제 보험': d['c_hr'], '배민 시간제 보험': d['b_hr'],
                    '오배달차감': '', 
                    '보험료 환급(소급)': t_ret, # 쿠팡 소급만 자동입력
                    '소득세': tax, '지방소득세': ltax, '선지급차감': 0, '최종지급(액)': 0
                })
            
            df_out = pd.DataFrame(final_rows)
            out = io.BytesIO()
            writer = pd.ExcelWriter(out, engine='xlsxwriter')
            df_out.to_excel(writer, index=False, sheet_name='정산서')
            
            wb = writer.book
            ws = writer.sheets['정산서']
            fmt_num = wb.add_format({'num_format': '#,##0'})
            fmt_hide = wb.add_format({'num_format': '#,##0;-#,##0;""'})
            
            ws.set_column('A:A', 12); ws.set_column('B:E', 14, fmt_num)
            ws.set_column('F:H', 14, fmt_hide); ws.set_column('I:U', 14, fmt_num)

            for i in range(len(df_out)):
                r = i + 2
                ws.write_formula(f'I{r}', f'=D{r}+E{r}+F{r}+G{r}+H{r}', fmt_num, df_out.iloc[i]['최종합산'])
                ws.write_formula(f'R{r}', f'=ROUNDDOWN(I{r}*0.03, -1)', fmt_num, df_out.iloc[i]['소득세'])
                ws.write_formula(f'S{r}', f'=ROUNDDOWN(I{r}*0.003, -1)', fmt_num, df_out.iloc[i]['지방소득세'])
                formula_final = f'=I{r}-(J{r}+K{r}+L{r}+M{r}+N{r}+O{r})-P{r}+Q{r}-(R{r}+S{r})-T{r}'
                ws.write_formula(f'U{r}', formula_final, fmt_num, df_out.iloc[i]['최종지급(액)'])
            
            writer.close()
            out.seek(0)

            st.session_state['processed_data'] = {
                'excel_data': out.getvalue(),
                'c_cnt': total_c,
                'b_cnt': total_b
            }
            st.rerun()

if st.session_state['processed_data']:
    data = st.session_state['processed_data']
    st.markdown("---")
    st.success(f"✅ **정산서 생성 완료!** (쿠팡: {data['c_cnt']}건 / 배민: {data['b_cnt']}건)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📥 1. 엑셀 다운로드", data['excel_data'], '빅스텝_통합_주차정산서.xlsx')
    with col2:
        if st.button("💸 2. 최종 확정 및 전송"):
            if log_to_sheet(data['c_cnt'], data['b_cnt']):
                st.toast("✅ 구글 시트 기록 완료!")
                st.balloons()
                st.session_state['processed_data'] = None
                st.rerun()