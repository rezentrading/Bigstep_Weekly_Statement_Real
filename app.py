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
    """
    이름 정규화: 숫자(ID)만 제거하고, 괄호(구분자)는 유지
    예: 김기열(서구) -> 김기열(서구) 유지 / 홍길동123 -> 홍길동
    """
    if pd.isna(name): return ""
    name = str(name)
    name = re.sub(r'\d+', '', name) # 숫자 제거
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
    try:
        xl = pd.ExcelFile(file_obj, engine='openpyxl')
        sheet_names = xl.sheet_names
        
        for sheet in sheet_names:
            if '을지' in sheet: return xl.parse(sheet, header=None), 'baemin'
        
        if '종합' in sheet_names: return xl.parse('종합', header=None), 'coupang'
            
        return xl.parse(0, header=None), None
    except:
        return pd.DataFrame(), None

def analyze_headers_type(df, detected_type):
    for i in range(min(len(df) - 1, 40)):
        row_curr = " ".join(df.iloc[i].astype(str).values).replace(" ", "")
        row_next = " ".join(df.iloc[i+1].astype(str).values).replace(" ", "")
        
        if '총정산오더수' in row_curr and '기사부담' in row_next: return i, i+1, 'coupang'
        if '총정산오더수' in row_curr and '기사부담' in row_curr: return i, i, 'coupang'
        
        if ('라이더명' in row_curr or '성명' in row_curr) and ('처리건수' in row_curr or '배달료' in row_curr): 
            return i, i, 'baemin'
            
    if detected_type == 'baemin':
        for i in range(min(len(df), 40)):
            row_str = " ".join(df.iloc[i].astype(str).values)
            if '라이더명' in row_str or '성명' in row_str: return i, i, 'baemin'

    return -1, -1, None

def find_col_index_global(df, keywords, exclude=None, max_rows=30):
    clean_keywords = [k.replace(" ", "") for k in keywords]
    clean_exclude = [e.replace(" ", "") for e in exclude] if exclude else []

    for c in range(len(df.columns)):
        for r in range(min(len(df), max_rows)):
            val = str(df.iloc[r, c]).replace(" ", "").replace("\n", "")
            if all(k in val for k in clean_keywords):
                if clean_exclude and any(e in val for e in clean_exclude): continue
                return c 
    return -1

def find_col_in_list(header_list, keywords, exclude=None):
    clean_keywords = [k.replace(" ", "") for k in keywords]
    clean_exclude = [e.replace(" ", "") for e in exclude] if exclude else []
    
    for i, val in enumerate(header_list):
        val_str = str(val).replace(" ", "").replace("\n", "")
        if all(k in val_str for k in clean_keywords):
            if clean_exclude and any(e in val_str for e in clean_exclude): continue
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
        for f in uploaded_files:
            unlocked = decrypt_file(f)
            try:
                df_raw, detected_type = get_sheet_data(unlocked)
                if not df_raw.empty:
                    m_idx, s_idx, ftype = analyze_headers_type(df_raw, detected_type)
                    if m_idx != -1: processed_files_map.append((df_raw, ftype, m_idx, s_idx))
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
                    idx_nm = find_col_index_global(df, ['성함']); idx_nm = 2 if idx_nm == -1 else idx_nm
                    idx_od = find_col_index_global(df, ['총', '정산', '오더수'])
                    if idx_od == -1: idx_od = find_col_index_global(df, ['오더수'])
                    
                    idx_net = find_col_index_global(df, ['수수료', '차감'])
                    if idx_net == -1: idx_net = find_col_index_global(df, ['차감', '금액'])
                    if idx_net == -1: idx_net = find_col_index_global(df, ['총', '정산금액'], exclude=['오더'])

                    idx_emp = find_col_index_global(df, ['기사부담', '고용보험'])
                    idx_ind = find_col_index_global(df, ['기사부담', '산재보험'])
                    idx_hr = find_col_index_global(df, ['시간제보험'])
                    idx_ret = find_col_index_global(df, ['보험료', '소급'])
                    
                    for i in range(data_start, len(df)):
                        row = df.iloc[i]
                        nm = normalize_name(row[idx_nm])
                        if not nm or nm == 'nan': continue
                        
                        od = clean_num(row[idx_od]) if idx_od != -1 else 0
                        total_c += od
                        rt = clean_num(row[idx_net]) if idx_net != -1 else 0
                        
                        ep = abs(clean_num(row[idx_emp])) if idx_emp != -1 else 0
                        id_ = abs(clean_num(row[idx_ind])) if idx_ind != -1 else 0
                        hr = abs(clean_num(row[idx_hr])) if idx_hr != -1 else 0
                        ret = abs(clean_num(row[idx_ret])) if idx_ret != -1 else 0
                        
                        if nm not in all_data: all_data[nm] = {'c_od':0,'c_tot':0,'c_ep':0,'c_id':0,'c_hr':0,'c_ret':0,'b_od':0,'b_tot':0,'b_ep':0,'b_id':0,'b_hr':0,'b_ret':0}
                        all_data[nm]['c_od']+=od; all_data[nm]['c_tot']+=rt; all_data[nm]['c_ep']+=ep; all_data[nm]['c_id']+=id_; all_data[nm]['c_hr']+=hr; all_data[nm]['c_ret']+=ret

                elif ftype == 'baemin':
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
                        all_data[nm]['b_od']+=od; all_data[nm]['b_tot']+=nt; all_data[nm]['b_ep']+=ep; all_data[nm]['b_id']+=id_; all_data[nm]['b_hr']+=hr; all_data[nm]['b_ret']+=0

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
                    '오배달차감': 0, '보험료 환급(소급)': t_ret,
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
                'df_out': df_out, # 데이터프레임 저장
                'c_cnt': total_c,
                'b_cnt': total_b
            }
            st.rerun()

# [C] 결과 확인 및 개별 발송 화면
if st.session_state['processed_data']:
    data = st.session_state['processed_data']
    st.markdown("---")
    st.success(f"✅ **정산서 생성 완료!** (쿠팡: {data['c_cnt']}건 / 배민: {data['b_cnt']}건)")
    
    # 탭 구성: 전체 다운로드 vs 개별 정산서
    tab1, tab2 = st.tabs(["📥 엑셀 전체 다운로드", "🧑‍✈️ 라이더별 개별 정산서 보내기"])
    
    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 1. 엑셀 다운로드", data['excel_data'], '빅스텝_통합_주차정산서.xlsx')
        with col2:
            if st.button("💸 2. 최종 확정 및 전송 (구글시트 기록)"):
                if log_to_sheet(data['c_cnt'], data['b_cnt']):
                    st.toast("✅ 구글 시트 기록 완료!")
                    st.balloons()
                    st.session_state['processed_data'] = None
                    st.rerun()

    with tab2:
        st.info("👇 아래에서 라이더 이름을 선택하면 '카톡 전송용 텍스트'와 '표'가 생성됩니다.")
        
        df_res = data['df_out']
        # 라이더 이름 목록
        rider_list = df_res['성함'].tolist()
        
        selected_rider = st.selectbox("라이더 선택", rider_list)
        
        if selected_rider:
            # 해당 라이더 데이터 추출
            rider_data = df_res[df_res['성함'] == selected_rider].iloc[0]
            
            # --- [1] 카톡 전송용 텍스트 생성 ---
            # 계산식 직접 적용 (엑셀 수식은 파이썬에서 안보이므로)
            final_pay = rider_data['최종합산'] - (
                rider_data['쿠팡 고용보험'] + rider_data['쿠팡 산재보험'] + 
                rider_data['배민 고용보험'] + rider_data['배민 산재보험'] + 
                rider_data['쿠팡 시간제 보험'] + rider_data['배민 시간제 보험']
            ) + rider_data['보험료 환급(소급)'] - (rider_data['소득세'] + rider_data['지방소득세'])
            
            msg_template = f"""[빅스텝 정산 알림]
            
안녕하세요 {selected_rider}님, 이번 주 정산 내역입니다.

■ 운행 요약
- 쿠팡 오더: {rider_data['쿠팡 오더수']}건 / {int(rider_data['쿠팡 총금액']):,}원
- 배민 오더: {rider_data['배민 오더수']}건 / {int(rider_data['배민 총금액']):,}원
------------------------------------
💰 총 매출: {int(rider_data['최종합산']):,}원

■ 차감 내역 (보험료/세금)
- 고용/산재: {int(rider_data['쿠팡 고용보험']+rider_data['쿠팡 산재보험']+rider_data['배민 고용보험']+rider_data['배민 산재보험']):,}원
- 시간제보험: {int(rider_data['쿠팡 시간제 보험']+rider_data['배민 시간제 보험']):,}원
- 소득세/지방세: {int(rider_data['소득세']+rider_data['지방소득세']):,}원
------------------------------------
👉 실 지급액: {int(final_pay):,}원

*오배달/소급 등 추가 변동 사항은 별도 문의 바랍니다.
"""
            st.text_area("복사해서 카톡에 붙여넣으세요!", value=msg_template, height=300)
            
            # --- [2] 표(영수증) 보기 ---
            st.markdown("### 🧾 상세 정산 내역 (캡처용)")
            
            # 보기 좋게 전치(Transpose)하여 표시
            disp_df = pd.DataFrame(rider_data).reset_index()
            disp_df.columns = ['항목', '금액/건수']
            
            # 숫자 포맷팅 (천단위 콤마)
            def format_currency(x):
                try: return f"{int(x):,}"
                except: return x
            disp_df['금액/건수'] = disp_df['금액/건수'].apply(format_currency)
            
            st.dataframe(disp_df, use_container_width=True, hide_index=True)