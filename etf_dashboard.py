import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import imaplib
import email
from email.header import decode_header
import zipfile
import io
import datetime
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import requests
from datetime import timedelta

# ----------------------------------------------------------------------
# 페이지 기본 설정
# ----------------------------------------------------------------------
st.set_page_config(page_title="ETF Market Monitoring (v9.0)", layout="wide")
st.title("📊 ETF Market Monitoring Dashboard (통합판)")

# 전역 변수 설정
CSV_FILE_ID = "1kvm2KgIlGMTIN3IOpLOUwEirKWlid8aP"
EXCEL_FILE_ID = "1xdKEXMRXf0TECNRvUedJ4jU9Pz29cRo4"

# ----------------------------------------------------------------------
# 1. 지메일(Gmail) 첨부파일 자동 수신 및 CSV 추출 함수
# ----------------------------------------------------------------------
def fetch_csvs_from_gmail():
    gmail_user = st.secrets.get("GMAIL_USER")
    gmail_pass = st.secrets.get("GMAIL_APP_PASSWORD")

    if not gmail_user or not gmail_pass:
        return pd.DataFrame()

    fetched_dfs = []
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_pass)
        mail.select("inbox")

        since_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%d-%b-%Y")
        search_query = f'(SUBJECT "[KRX]" SINCE "{since_date}")'
        status, messages = mail.search(None, search_query)

        if status == "OK" and messages[0]:
            email_ids = messages[0].split()

            for e_id in email_ids:
                res, msg_data = mail.fetch(e_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8", errors="replace")

                        if "[KRX]" in subject and "구독형 데이터" in subject:
                            for part in msg.walk():
                                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                                    continue
                                filename = part.get_filename()
                                if filename:
                                    fn_decoded, fn_enc = decode_header(filename)[0]
                                    if isinstance(fn_decoded, bytes):
                                        filename = fn_decoded.decode(fn_enc if fn_enc else "utf-8", errors="replace")

                                    if filename.endswith(".zip") and "[ETF]" in filename:
                                        zip_bytes = part.get_payload(decode=True)
                                        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                                            for csv_name in z.namelist():
                                                if csv_name.endswith(".csv"):
                                                    with z.open(csv_name) as csv_file:
                                                        df_temp = pd.read_csv(csv_file, encoding='cp949', thousands=',')
                                                        fetched_dfs.append(df_temp)
        mail.logout()
    except Exception as e:
        st.warning(f"⚠️ 지메일 수신 실패 (기존 구글 드라이브 데이터 사용): {e}")
        return pd.DataFrame()

    if fetched_dfs:
        return pd.concat(fetched_dfs, ignore_index=True)
    return pd.DataFrame()


# ----------------------------------------------------------------------
# 2. 구글 드라이브 원본 CSV 업데이트(업로드) 함수 (수정본)
# ----------------------------------------------------------------------
def update_drive_csv(df, file_id):
    try:
        creds_info = dict(st.secrets["gcp_service_account"])
        credentials = service_account.Credentials.from_service_account_info(
            creds_info, scopes=['https://www.googleapis.com/auth/drive']
        )
        service = build('drive', 'v3', credentials=credentials)

        df_export = df.copy()
        df_export['거래일자'] = df_export['거래일자'].dt.strftime('%Y%m%d')

        # 💡 원본 순수 CSV 컬럼 목록만 정의 (파생/계산 컬럼 제외)
        original_cols = [
            '거래일자', '상품그룹ID', '종목코드', '종목명', '회원사명', 
            'LP매도거래량', 'LP매도거래대금', 'LP매수거래량', 'LP매수거래대금'
        ]
        
        # 실제 존재하는 원본 컬럼만 필터링하여 저장
        valid_cols = [col for col in original_cols if col in df_export.columns]
        df_export = df_export[valid_cols]

        csv_buffer = io.BytesIO()
        df_export.to_csv(csv_buffer, index=False, encoding='cp949')
        csv_buffer.seek(0)

        media = MediaIoBaseUpload(csv_buffer, mimetype='text/csv', resumable=True)
        service.files().update(fileId=file_id, media_body=media).execute()
        st.toast("✅ 구글 드라이브 원본 CSV가 깔끔하게 업데이트되었습니다!", icon="💾")
        return True
    except Exception as e:
        st.error(f"구글 드라이브 업데이트 실패: {e}")
        return False

# ----------------------------------------------------------------------
# 3. 메인 데이터 로드 및 전처리 (지메일 수신 자동 포함)
# ----------------------------------------------------------------------
@st.cache_data(ttl=3600)
def load_data():
    df_excel = None
    df_base = None

    if "gcp_service_account" in st.secrets:
        try:
            creds_info = dict(st.secrets["gcp_service_account"])
            credentials = service_account.Credentials.from_service_account_info(
                creds_info, scopes=['https://www.googleapis.com/auth/drive']
            )
            service = build('drive', 'v3', credentials=credentials)

            try:
                request = service.files().export_media(fileId=EXCEL_FILE_ID, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                excel_bytes = io.BytesIO(request.execute())
            except:
                request = service.files().get_media(fileId=EXCEL_FILE_ID)
                excel_bytes = io.BytesIO(request.execute())
            df_excel = pd.read_excel(excel_bytes)

            request_csv = service.files().get_media(fileId=CSV_FILE_ID)
            csv_bytes = io.BytesIO(request_csv.execute())
            df_base = pd.read_csv(csv_bytes, encoding='cp949', thousands=',')
        except Exception as e:
            pass

    if df_excel is None:
        try:
            excel_url = f"https://docs.google.com/spreadsheets/d/{EXCEL_FILE_ID}/export?format=xlsx"
            df_excel = pd.read_excel(excel_url)
        except:
            excel_url = f"https://drive.google.com/uc?export=download&id={EXCEL_FILE_ID}"
            df_excel = pd.read_excel(excel_url)

    if df_base is None:
        csv_url = f"https://drive.google.com/uc?export=download&id={CSV_FILE_ID}"
        df_base = pd.read_csv(csv_url, encoding='cp949', thousands=',')

    base_row_count = len(df_base)

    master_db = {}
    for _, row in df_excel.iterrows():
        std_code = str(row.iloc[0]).strip().upper().replace(" ", "")
        if pd.isna(std_code) or std_code == 'NAN': continue

        raw_short = str(row.iloc[1]).strip().upper().replace(" ", "")
        if (not raw_short or raw_short == 'NAN') and len(std_code) >= 9:
            raw_short = std_code[3:9]
        a_code = "A" + raw_short if not raw_short.startswith("A") else raw_short

        name = str(row.iloc[3])
        if pd.isna(name) or name == 'nan': name = str(row.iloc[2])
        if pd.isna(name) or name == 'nan': name = str(row.iloc[1])

        market = "국내" if str(row.iloc[10]).replace(" ", "") == "국내" else "해외"
        asset = "주식" if str(row.iloc[11]).replace(" ", "") == "주식" else "그외"
        base_index = str(row.iloc[6]).upper().replace(" ", "")
        rep_keys = ["코스피200", "코스닥150", "S&P500", "나스닥100", "NASDAQ100"]
        is_rep = "대표지수" if any(k in base_index for k in rep_keys) else "그외"
        deriv = "일반" if str(row.iloc[8]).replace(" ", "") == "일반" else "파생"
        tracking = "패시브" if str(row.iloc[9]).replace(" ", "") in ["실물(패시브)", "합성(패시브)"] else "액티브"
        amc_raw = str(row.iloc[13]).replace(" ", "")
        amc = amc_raw if amc_raw and amc_raw != 'nan' else "기타운용사"
        cat_key = f"{market} | {asset} | {is_rep} | {deriv} | {tracking}"

        master_db[std_code] = {
            'a_code': a_code, 'name': name, 'market': market, 'asset': asset,
            'is_rep': is_rep, 'deriv': deriv, 'tracking': tracking, 'category_key': cat_key, 'amc': amc
        }

    # 지메일 백그라운드 자동 수신
    try:
        df_gmail = fetch_csvs_from_gmail()
    except:
        df_gmail = pd.DataFrame()

    if not df_gmail.empty:
        df = pd.concat([df_base, df_gmail], ignore_index=True)
    else:
        df = df_base.copy()

    df = df[df['상품그룹ID'].str.upper() == 'ETF'].copy()

    clean_date_str = (
        df['거래일자']
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.replace(r'\D', '', regex=True)
        .str.strip()
    )
    df['거래일자'] = pd.to_datetime(clean_date_str, format='%Y%m%d', errors='coerce')
    df = df.dropna(subset=['거래일자']).copy()

    df = df.drop_duplicates(subset=['거래일자', '종목코드', '회원사명'], keep='last').reset_index(drop=True)

    # 💡 자동 업데이트 로직(update_drive_csv) 제거됨 (버튼으로 분리)
    has_new_data = len(df) > base_row_count

    df['LP매도거래대금'] = df['LP매도거래대금'].fillna(0)
    df['LP매수거래대금'] = df['LP매수거래대금'].fillna(0)
    if 'LP매도거래량' in df.columns: df['LP매도거래량'] = df['LP매도거래량'].fillna(0)
    if 'LP매수거래량' in df.columns: df['LP매수거래량'] = df['LP매수거래량'].fillna(0)

    df['총LP거래대금'] = df['LP매도거래대금'] + df['LP매수거래대금']
    df['LP순매수대금'] = df['LP매수거래대금'] - df['LP매도거래대금']

    master_df = pd.DataFrame.from_dict(master_db, orient='index')
    df['종목코드'] = df['종목코드'].str.strip().str.upper().str.replace(' ', '')
    df['a_code'] = df['종목코드'].map(master_df['a_code']).fillna(df['종목코드'])

    master_names = df['종목코드'].map(master_df['name'])
    df['종목명'] = np.where(master_names.isna() | (master_names == ''), df['종목명'], master_names)

    categories = ['market', 'asset', 'is_rep', 'deriv', 'tracking', 'amc', 'category_key']
    for cat in categories:
        df[cat] = df['종목코드'].map(master_df[cat]).fillna('미분류')

    return df, master_db, has_new_data

df, master_db, has_new_data = load_data()
if df.empty:
    st.error("⚠️ 데이터 로드에 실패하였거나 표시할 ETF 데이터가 없습니다.")
    st.stop()


# ----------------------------------------------------------------------
# 4. KRX Open API 연동 함수들
# ----------------------------------------------------------------------
@st.cache_data(ttl=86400, show_spinner=False)
def get_krx_open_api_market_data(tickers, start_date, end_date):
    api_key = st.secrets.get("KRX_API_KEY")
    if not api_key: return pd.DataFrame()

    url = "https://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd"
    headers = {"AUTH_KEY": api_key, "AUTH-KEY": api_key}
    date_list = pd.date_range(start=start_date, end=end_date, freq='B')
    
    all_data = []
    clean_tickers = [str(t).strip().zfill(6) for t in tickers]
    
    for dt in date_list:
        basDd = dt.strftime('%Y%m%d')
        params = {"basDd": basDd, "AUTH_KEY": api_key}
        
        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if "OutBlock_1" in data:
                    df_day = pd.DataFrame(data["OutBlock_1"])
                    if not df_day.empty and 'ISU_CD' in df_day.columns:
                        df_day['short_code'] = df_day['ISU_CD'].astype(str).str.extract(r'(\d{6})')[0]
                        filtered_df = df_day[df_day['short_code'].isin(clean_tickers)]
                        
                        if not filtered_df.empty and 'ACC_TRDVAL' in filtered_df.columns:
                            trd_vals = pd.to_numeric(filtered_df['ACC_TRDVAL'].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                            daily_total = trd_vals.sum()
                            all_data.append({'거래일자': pd.to_datetime(dt.date()), '시장거래대금': daily_total})
        except Exception:
            pass
        time.sleep(0.15)
            
    if all_data:
        result_df = pd.DataFrame(all_data)
        result_df.set_index('거래일자', inplace=True)
        result_df = result_df.replace([np.inf, -np.inf], np.nan).fillna(0)
        return result_df
    else:
        return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def get_krx_snapshot(target_date):
    api_key = st.secrets.get("KRX_API_KEY")
    if not api_key: return pd.DataFrame()
    
    url = "https://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd"
    headers = {"AUTH_KEY": api_key, "AUTH-KEY": api_key}
    
    for i in range(7):
        dt = pd.to_datetime(target_date) - timedelta(days=i)
        if dt.weekday() >= 5: continue
            
        basDd = dt.strftime('%Y%m%d')
        params = {"basDd": basDd, "AUTH_KEY": api_key}
        
        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if "OutBlock_1" in data and len(data["OutBlock_1"]) > 0:
                    df = pd.DataFrame(data["OutBlock_1"])
                    df['query_date'] = dt.date()
                    return df
        except Exception:
            pass
        time.sleep(0.1)
        
    return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def get_krx_daily_nav_shares(ticker, start_date, end_date):
    api_key = st.secrets.get("KRX_API_KEY")
    if not api_key: return pd.DataFrame()
    
    url = "https://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd"
    headers = {"AUTH_KEY": api_key, "AUTH-KEY": api_key}
    
    extended_start = pd.to_datetime(start_date) - timedelta(days=10) 
    date_list = pd.date_range(start=extended_start, end=end_date, freq='B')
    
    clean_ticker = str(ticker).strip().zfill(6)
    all_data = []
    
    for dt in date_list:
        basDd = dt.strftime('%Y%m%d')
        params = {"basDd": basDd, "AUTH_KEY": api_key}
        
        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if "OutBlock_1" in data:
                    df_day = pd.DataFrame(data["OutBlock_1"])
                    if not df_day.empty and 'ISU_CD' in df_day.columns:
                        df_day['short_code'] = df_day['ISU_CD'].astype(str).str.extract(r'(\d{6})')[0]
                        row = df_day[df_day['short_code'] == clean_ticker]
                        
                        if not row.empty:
                            nav_col = next((c for c in ['NAV', 'TDD_NAV', 'IDX_NAV'] if c in row.columns), None)
                            share_col = next((c for c in ['LIST_SHRS', 'LST_SHRS', 'LST_STK_VL'] if c in row.columns), None)
                            
                            nav = pd.to_numeric(str(row.iloc[0][nav_col]).replace(',', ''), errors='coerce') if nav_col else 0
                            shares = pd.to_numeric(str(row.iloc[0][share_col]).replace(',', ''), errors='coerce') if share_col else 0
                            
                            all_data.append({
                                '날짜': pd.to_datetime(dt.date()),
                                'NAV': nav,
                                'SHARES': shares
                            })
        except Exception:
            pass
        time.sleep(0.15)
        
    if all_data:
        df_res = pd.DataFrame(all_data).sort_values('날짜').drop_duplicates('날짜').reset_index(drop=True)
        return df_res
    return pd.DataFrame()


# ----------------------------------------------------------------------
# 사이드바 (업데이트 버튼 및 Date Filter)
# ----------------------------------------------------------------------
st.sidebar.header("⚙️ 데이터 관리")
if has_new_data:
    st.sidebar.info("💡 **새로운 메일 데이터가 병합되었습니다.** 드라이브에 원본을 저장하시겠습니까?")
if st.sidebar.button("💾 구글 드라이브 원본 덮어쓰기", type="primary"):
    with st.spinner("병합된 최신 데이터를 구글 드라이브에 저장 중입니다..."):
        if update_drive_csv(df, CSV_FILE_ID):
            st.cache_data.clear()  # 기존 캐시 메모리 초기화
            st.sidebar.success("저장 완료! 화면을 새로고침합니다.")
            time.sleep(1)          # 성공 메시지를 사용자가 읽을 수 있도록 1초 대기
            st.rerun()             # 화면 즉시 새로고침 (새로운 데이터 상태 반영)

st.sidebar.divider()

st.sidebar.header("🗓️ 데이터 날짜 설정")

min_val = df['거래일자'].min()
max_val = df['거래일자'].max()

min_date = min_val.date() if hasattr(min_val, 'date') else pd.to_datetime(min_val).date()
max_date = max_val.date() if hasattr(max_val, 'date') else pd.to_datetime(max_val).date()

search_type = st.sidebar.radio(
    "조회 방식을 선택하세요",
    ["기간 조회", "특정 일자 조회"],
    horizontal=True
)

if search_type == "기간 조회":
    date_selection = st.sidebar.date_input(
        "조회 기간을 선택하세요", 
        [min_date, max_date], 
        min_value=min_date,
        max_value=max_date
    )
    if len(date_selection) == 2:
        start_date, end_date = date_selection
    else:
        start_date = end_date = date_selection[0]
        
else: 
    single_date = st.sidebar.date_input(
        "조회할 일자를 선택하세요",
        max_date, 
        min_value=min_date,
        max_value=max_date
    )
    start_date = single_date
    end_date = single_date

df_filtered = df[(df['거래일자'].dt.date >= start_date) & (df['거래일자'].dt.date <= end_date)].copy()


# ----------------------------------------------------------------------
# UI Tabs 구성 (9번 탭 추가)
# ----------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
    "1. 종합 대시보드", "2. ETF 구분별 분석", "3. LP사 다각도 분석", 
    "4. ETF별 주력 LP 분석", "5. 운용사별 주력 ETF 분석",
    "6. 종목 집중도 분석", "7. 시장 전체 거래대금 (KRX)", "8. 설정/환매 추이 (추정)", "9. NAV 괴리율 조회"
])

# ==========================================
# Tab 1 ~ Tab 7 생략 없이 그대로 포함 (동일)
# ==========================================
with tab1:
    st.subheader("📊 시장 핵심 지표 및 추이 (KPI)")
    trend_type = st.radio("🔍 데이터 필터링 관점 선택 (아래 모든 지표에 적용됩니다)", ["시장 전체 (Total Market)", "특정 LP사 (Specific LP)", "특정 운용사 (Specific AMC)"], horizontal=True, key='t1_trend')
    target_lps = []
    target_amcs = []
    if trend_type == "특정 LP사 (Specific LP)":
        top_lp = df_filtered.groupby('회원사명')['총LP거래대금'].sum().idxmax() if not df_filtered.empty else None
        target_lps = st.multiselect("비교 분석할 LP사 선택 (다중 선택 가능)", sorted(df_filtered['회원사명'].unique()), default=[top_lp] if top_lp else None, key='t1_lp_sel')
    elif trend_type == "특정 운용사 (Specific AMC)":
        amc_list = sorted([a for a in df_filtered['amc'].unique() if a != '미분류'])
        top_amc = df_filtered[df_filtered['amc'] != '미분류'].groupby('amc')['총LP거래대금'].sum().idxmax() if not df_filtered.empty else None
        target_amcs = st.multiselect("비교 분석할 운용사 선택 (다중 선택 가능)", amc_list, default=[top_amc] if top_amc else None, key='t1_amc_sel')
    
    st.divider()
    st.write("▼ ETF 섹터 세부 필터")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t1 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t1_mkt')
    ast_filter_t1 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t1_ast')
    rep_filter_t1 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t1_rep')
    drv_filter_t1 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t1_drv')
    trk_filter_t1 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t1_trk')
    
    df_t1 = df_filtered.copy()
    if mkt_filter_t1 != "전체": df_t1 = df_t1[df_t1['market'] == mkt_filter_t1]
    if ast_filter_t1 != "전체": df_t1 = df_t1[df_t1['asset'] == ast_filter_t1]
    if rep_filter_t1 != "전체": df_t1 = df_t1[df_t1['is_rep'] == rep_filter_t1]
    if drv_filter_t1 != "전체": df_t1 = df_t1[df_t1['deriv'] == drv_filter_t1]
    if trk_filter_t1 != "전체": df_t1 = df_t1[df_t1['tracking'] == trk_filter_t1]
    
    is_data_empty = False
    if trend_type == "특정 LP사 (Specific LP)":
        if target_lps: df_t1 = df_t1[df_t1['회원사명'].isin(target_lps)]
        else: is_data_empty = True 
    elif trend_type == "특정 운용사 (Specific AMC)":
        if target_amcs: df_t1 = df_t1[df_t1['amc'].isin(target_amcs)]
        else: is_data_empty = True 
        
    st.write("") 
    
    if is_data_empty or df_t1.empty:
        st.warning("조회할 데이터가 없습니다. 상단에서 대상을 선택하거나 필터 조건을 변경해 주세요.")
    else:
        total_amt = df_t1['총LP거래대금'].sum() / 100_000_000
        unique_days = df_t1['거래일자'].nunique()
        daily_avg = total_amt / unique_days if unique_days > 0 else 0
        active_etf_cnt = df_t1[df_t1['총LP거래대금'] > 0]['종목코드'].nunique()
        active_lp_cnt = df_t1[df_t1['총LP거래대금'] > 0]['회원사명'].nunique()
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("기간 총 거래대금", f"{total_amt:,.0f} 억원")
        col2.metric("일평균 거래대금", f"{daily_avg:,.0f} 억원")
        col3.metric("유효 거래 종목 수", f"{active_etf_cnt:,} 개")
        col4.metric("활동 LP 회원사 수", f"{active_lp_cnt:,} 사")
        
        st.divider()
        lp_total = df_t1.groupby('회원사명')['총LP거래대금'].sum().sort_values(ascending=False) / 100_000_000
        lp_total = lp_total[lp_total > 0]
        if not lp_total.empty:
            fig = px.bar(lp_total, x=lp_total.index, y=lp_total.values, title="선택된 조건 내 LP사별 총 거래대금 (억원)", labels={'y': '거래대금(억)', '회원사명': '증권사'}, color_discrete_sequence=['#4A90E2'])
            st.plotly_chart(fig, use_container_width=True)
            
        st.divider()
        st.subheader("📉 시계열(Time-Series) 일별 거래대금 추이")
        if trend_type == "시장 전체 (Total Market)":
            daily_vol = df_t1.groupby('거래일자')['총LP거래대금'].sum().reset_index()
            daily_vol['거래대금(억)'] = daily_vol['총LP거래대금'] / 100_000_000
            fig_t = px.line(daily_vol, x='거래일자', y='거래대금(억)', title="전체 시장 일별 LP 거래대금 추이", markers=True)
            st.plotly_chart(fig_t, use_container_width=True)
        elif trend_type == "특정 LP사 (Specific LP)":
            daily_vol = df_t1.groupby(['거래일자', '회원사명'])['총LP거래대금'].sum().reset_index()
            daily_vol['거래대금(억)'] = daily_vol['총LP거래대금'] / 100_000_000
            fig_t = px.line(daily_vol, x='거래일자', y='거래대금(억)', color='회원사명', title="선택 LP사별 일별 거래대금 추이", markers=True)
            st.plotly_chart(fig_t, use_container_width=True)
        elif trend_type == "특정 운용사 (Specific AMC)":
            daily_vol = df_t1.groupby(['거래일자', 'amc'])['총LP거래대금'].sum().reset_index()
            daily_vol['거래대금(억)'] = daily_vol['총LP거래대금'] / 100_000_000
            fig_t = px.line(daily_vol, x='거래일자', y='거래대금(억)', color='amc', title="선택 운용사별 일별 거래대금 추이", markers=True)
            st.plotly_chart(fig_t, use_container_width=True)

        st.divider()
        st.subheader("🏢 운용사별 LP사 시장 점유율 (M/S %) 매트릭스")
        st.write("각 운용사(AMC)의 전체 거래대금 중 특정 LP사가 차지하는 비중(%)을 보여줍니다. (세로열 기준 합산 100%)")

        amc_totals = df_t1.groupby('amc')['총LP거래대금'].sum().sort_values(ascending=False)
        lp_totals = df_t1.groupby('회원사명')['총LP거래대금'].sum().sort_values(ascending=False)
        
        valid_amcs = amc_totals[amc_totals > 0].index.tolist()
        valid_lps = lp_totals[lp_totals > 0].index.tolist()

        if valid_amcs and valid_lps:
            pivot_vol = df_t1.groupby(['회원사명', 'amc'])['총LP거래대금'].sum().unstack(fill_value=0)
            pivot_ms = pivot_vol.div(pivot_vol.sum(axis=0), axis=1) * 100
            pivot_ms = pivot_ms.reindex(index=valid_lps, columns=valid_amcs).fillna(0)
            
            st.dataframe(
                pivot_ms.style.format("{:.1f}%").background_gradient(cmap='Blues', axis=0),
                use_container_width=True
            )
        else:
            st.warning("선택하신 필터 조건에 해당하는 점유율 데이터가 없습니다.")



with tab2:
    st.subheader("📈 ETF 섹터 필터링을 통한 점유율 및 성향 분석")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t2 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t2_mkt')
    ast_filter_t2 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t2_ast')
    rep_filter_t2 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t2_rep')
    drv_filter_t2 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t2_drv')
    trk_filter_t2 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t2_trk')
    df_t2 = df_filtered.copy()
    if mkt_filter_t2 != "전체": df_t2 = df_t2[df_t2['market'] == mkt_filter_t2]
    if ast_filter_t2 != "전체": df_t2 = df_t2[df_t2['asset'] == ast_filter_t2]
    if rep_filter_t2 != "전체": df_t2 = df_t2[df_t2['is_rep'] == rep_filter_t2]
    if drv_filter_t2 != "전체": df_t2 = df_t2[df_t2['deriv'] == drv_filter_t2]
    if trk_filter_t2 != "전체": df_t2 = df_t2[df_t2['tracking'] == trk_filter_t2]
    t2_etf = df_t2.groupby(['회원사명', '종목명'])[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', 'LP매도거래량', 'LP매수거래량']].sum().reset_index()
    t2_etf['평균매도단가'] = np.where(t2_etf['LP매도거래량'] > 0, t2_etf['LP매도거래대금'] / t2_etf['LP매도거래량'], 0)
    t2_etf['평균매수단가'] = np.where(t2_etf['LP매수거래량'] > 0, t2_etf['LP매수거래대금'] / t2_etf['LP매수거래량'], 0)
    t2_etf['체결수량(min)'] = t2_etf[['LP매도거래량', 'LP매수거래량']].min(axis=1)
    t2_etf['추정매매이익'] = (t2_etf['평균매도단가'] - t2_etf['평균매수단가']) * t2_etf['체결수량(min)']
    agg_df = t2_etf.groupby('회원사명')[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', '추정매매이익']].sum().reset_index()
    agg_df = agg_df[agg_df['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)
    total_vol = agg_df['총LP거래대금'].sum()
    agg_df['점유율(%)'] = (agg_df['총LP거래대금'] / total_vol) * 100 if total_vol > 0 else 0
    agg_df['거래대금(억)'] = agg_df['총LP거래대금'] / 100_000_000
    agg_df['추정매매손익(백만)'] = agg_df['추정매매이익'] / 1_000_000
    st.write("### 1️⃣ 타겟 섹터 회원사 점유율")
    col1, col2 = st.columns([1, 1])
    with col1:
        show_df = agg_df[['회원사명', '거래대금(억)', '추정매매손익(백만)', '점유율(%)']].copy()
        show_df['순위'] = range(1, len(show_df) + 1)
        st.dataframe(show_df.set_index('순위').style.format({'거래대금(억)': '{:,.0f}', '추정매매손익(백만)': '{:,.0f}', '점유율(%)': '{:.1f}%'}), use_container_width=True)
    with col2:
        if not agg_df.empty:
            fig2 = px.pie(agg_df.head(10), values='거래대금(억)', names='회원사명', hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)
    st.divider()
    st.write("### 2️⃣ LP사 매수/매도 스탠스 (순매수 현황)")
    imb_df = agg_df.copy()
    imb_df['순매수비율(%)'] = np.where(imb_df['총LP거래대금'] > 0, (imb_df['LP순매수대금'] / imb_df['총LP거래대금']) * 100, 0)
    imb_df = imb_df.sort_values('LP순매수대금', ascending=False)
    fig_imb = px.bar(imb_df, x='회원사명', y='LP순매수대금', title="필터링된 섹터 내 순매수/순매도 누적대금", labels={'LP순매수대금': '순매수대금(원)', '회원사명': '증권사'}, color='LP순매수대금', color_continuous_scale=px.colors.diverging.RdBu_r, text=imb_df['LP순매수대금'] / 100_000_000)
    fig_imb.update_traces(texttemplate='%{text:,.0f}억', textposition='outside')
    st.plotly_chart(fig_imb, use_container_width=True)
    st.write("▼ 세부 금액 데이터")
    imb_df['매도대금(억)'] = imb_df['LP매도거래대금'] / 100_000_000
    imb_df['매수대금(억)'] = imb_df['LP매수거래대금'] / 100_000_000
    imb_df['순매수대금(억)'] = imb_df['LP순매수대금'] / 100_000_000
    st.dataframe(imb_df[['회원사명', '매도대금(억)', '매수대금(억)', '순매수대금(억)', '거래대금(억)', '순매수비율(%)']].style.format({'매도대금(억)': '{:,.0f}', '매수대금(억)': '{:,.0f}', '순매수대금(억)': '{:,.0f}', '거래대금(억)': '{:,.0f}', '순매수비율(%)': '{:.2f}%'}), use_container_width=True, hide_index=True)

with tab3:
    lp_list = df_filtered.groupby('회원사명')['총LP거래대금'].sum().sort_values(ascending=False).index.tolist()
    target_lp = st.selectbox("📌 분석 대상 LP사 선택", lp_list)
    if target_lp:
        df_lp = df_filtered[df_filtered['회원사명'] == target_lp]
        lp_total_amt = df_lp['총LP거래대금'].sum()
        sys_total_amt = df_filtered['총LP거래대금'].sum()
        ms = (lp_total_amt / sys_total_amt) * 100 if sys_total_amt > 0 else 0
        st.info(f"**{target_lp}** | 해당 기간 LP 총 거래대금: {lp_total_amt/100_000_000:,.0f} 억원 | 전체 시장 점유율(M/S): {ms:.2f}%")
        c1, c2 = st.columns(2)
        with c1:
            st.write("1️⃣ 섹터별 상세 거래 내역 (5단계 분류)")
            market_sector = df_filtered.groupby('category_key')['총LP거래대금'].sum()
            lp_sector = df_lp.groupby('category_key')['총LP거래대금'].sum().reset_index()
            lp_sector['내부비중(%)'] = (lp_sector['총LP거래대금'] / lp_total_amt) * 100
            lp_sector['섹터내_MS(%)'] = lp_sector.apply(lambda r: (r['총LP거래대금'] / market_sector.get(r['category_key'], 1)) * 100, axis=1)
            lp_sector['대금(억)'] = lp_sector['총LP거래대금'] / 100_000_000
            st.dataframe(lp_sector[lp_sector['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)[['category_key', '대금(억)', '내부비중(%)', '섹터내_MS(%)']].style.format({'대금(억)': '{:,.0f}', '내부비중(%)': '{:.1f}%', '섹터내_MS(%)': '{:.1f}%'}), use_container_width=True, hide_index=True)
        with c2:
            st.write("2️⃣ 운용사(AMC)별 커버리지 및 충성도 (섹터 필터링)")
            f1, f2, f3 = st.columns(3)
            mkt_filter_t3 = f1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t3_mkt')
            ast_filter_t3 = f2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t3_ast')
            rep_filter_t3 = f3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t3_rep')
            f4, f5 = st.columns(2)
            drv_filter_t3 = f4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t3_drv')
            trk_filter_t3 = f5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t3_trk')
            df_c2_lp = df_lp.copy()
            df_c2_mkt = df_filtered.copy()
            if mkt_filter_t3 != "전체": 
                df_c2_lp = df_c2_lp[df_c2_lp['market'] == mkt_filter_t3]
                df_c2_mkt = df_c2_mkt[df_c2_mkt['market'] == mkt_filter_t3]
            if ast_filter_t3 != "전체": 
                df_c2_lp = df_c2_lp[df_c2_lp['asset'] == ast_filter_t3]
                df_c2_mkt = df_c2_mkt[df_c2_mkt['asset'] == ast_filter_t3]
            if rep_filter_t3 != "전체": 
                df_c2_lp = df_c2_lp[df_c2_lp['is_rep'] == rep_filter_t3]
                df_c2_mkt = df_c2_mkt[df_c2_mkt['is_rep'] == rep_filter_t3]
            if drv_filter_t3 != "전체": 
                df_c2_lp = df_c2_lp[df_c2_lp['deriv'] == drv_filter_t3]
                df_c2_mkt = df_c2_mkt[df_c2_mkt['deriv'] == drv_filter_t3]
            if trk_filter_t3 != "전체": 
                df_c2_lp = df_c2_lp[df_c2_lp['tracking'] == trk_filter_t3]
                df_c2_mkt = df_c2_mkt[df_c2_mkt['tracking'] == trk_filter_t3]
            filtered_lp_total = df_c2_lp['총LP거래대금'].sum()
            if filtered_lp_total > 0:
                amc_lp = df_c2_lp.groupby('amc')['총LP거래대금'].sum().reset_index()
                amc_mkt = df_c2_mkt.groupby('amc')['총LP거래대금'].sum()
                amc_lp['대금(억)'] = amc_lp['총LP거래대금'] / 100_000_000
                amc_lp['내부비중(%)'] = (amc_lp['총LP거래대금'] / filtered_lp_total) * 100
                amc_lp['AMC내_MS(%)'] = amc_lp.apply(lambda r: (r['총LP거래대금'] / amc_mkt.get(r['amc'], 1)) * 100 if amc_mkt.get(r['amc'], 0) > 0 else 0, axis=1)
                st.dataframe(amc_lp[amc_lp['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)[['amc', '대금(억)', '내부비중(%)', 'AMC내_MS(%)']].style.format({'대금(억)': '{:,.0f}', '내부비중(%)': '{:.1f}%', 'AMC내_MS(%)': '{:.1f}%'}), use_container_width=True, hide_index=True)
            else:
                st.warning("선택하신 필터 조건에 해당하는 LP 거래 내역이 없습니다.")
        st.divider()
        st.subheader(f"3️⃣ [{target_lp}] 전체 거래 종목 상세 분석 (섹터 필터링 & 추정매매손익)")
        fd1, fd2, fd3, fd4, fd5 = st.columns(5)
        mkt_filter_detail = fd1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='detail_mkt')
        ast_filter_detail = fd2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='detail_ast')
        rep_filter_detail = fd3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='detail_rep')
        drv_filter_detail = fd4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='detail_drv')
        trk_filter_detail = fd5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='detail_trk')
        df_detail = df_lp.copy()
        if mkt_filter_detail != "전체": df_detail = df_detail[df_detail['market'] == mkt_filter_detail]
        if ast_filter_detail != "전체": df_detail = df_detail[df_detail['asset'] == ast_filter_detail]
        if rep_filter_detail != "전체": df_detail = df_detail[df_detail['is_rep'] == rep_filter_detail]
        if drv_filter_detail != "전체": df_detail = df_detail[df_detail['deriv'] == drv_filter_detail]
        if trk_filter_detail != "전체": df_detail = df_detail[df_detail['tracking'] == trk_filter_detail]
        total_detail_vol = df_detail['총LP거래대금'].sum()
        if total_detail_vol > 0:
            detail_etfs = df_detail.groupby('종목명')[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', 'LP매도거래량', 'LP매수거래량']].sum().reset_index()
            detail_etfs = detail_etfs[detail_etfs['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False).reset_index(drop=True)
            detail_etfs['순위'] = detail_etfs.index + 1
            detail_etfs['평균매도단가'] = np.where(detail_etfs['LP매도거래량'] > 0, detail_etfs['LP매도거래대금'] / detail_etfs['LP매도거래량'], 0)
            detail_etfs['평균매수단가'] = np.where(detail_etfs['LP매수거래량'] > 0, detail_etfs['LP매수거래대금'] / detail_etfs['LP매수거래량'], 0)
            detail_etfs['체결수량(min)'] = detail_etfs[['LP매도거래량', 'LP매수거래량']].min(axis=1)
            detail_etfs['추정매매이익'] = (detail_etfs['평균매도단가'] - detail_etfs['평균매수단가']) * detail_etfs['체결수량(min)']
            detail_etfs['거래대금(억)'] = detail_etfs['총LP거래대금'] / 100_000_000
            detail_etfs['매도대금(억)'] = detail_etfs['LP매도거래대금'] / 100_000_000
            detail_etfs['매수대금(억)'] = detail_etfs['LP매수거래대금'] / 100_000_000
            detail_etfs['순매수대금(억)'] = detail_etfs['LP순매수대금'] / 100_000_000
            detail_etfs['추정매매손익(백만)'] = detail_etfs['추정매매이익'] / 1_000_000
            detail_etfs['비중(%)'] = (detail_etfs['총LP거래대금'] / total_detail_vol) * 100
            total_row = pd.DataFrame([{'순위': 0, '종목명': '📊 [총 합계]', '거래대금(억)': detail_etfs['거래대금(억)'].sum(), '매도대금(억)': detail_etfs['매도대금(억)'].sum(), '매수대금(억)': detail_etfs['매수대금(억)'].sum(), '순매수대금(억)': detail_etfs['순매수대금(억)'].sum(), '추정매매손익(백만)': detail_etfs['추정매매손익(백만)'].sum(), '비중(%)': 100.0}])
            detail_etfs = pd.concat([total_row, detail_etfs], ignore_index=True)
            st.write(f"해당 필터 조건 거래 종목 수: **{len(detail_etfs)-1:,}개** | 기간 총 거래대금: **{total_detail_vol/100_000_000:,.0f}억원**")
            show_cols = ['순위', '종목명', '거래대금(억)', '매도대금(억)', '매수대금(억)', '순매수대금(억)', '추정매매손익(백만)', '비중(%)']
            st.dataframe(detail_etfs[show_cols].set_index('순위').style.format({'거래대금(억)': '{:,.0f}', '매도대금(억)': '{:,.0f}', '매수대금(억)': '{:,.0f}', '순매수대금(억)': '{:,.0f}', '추정매매손익(백만)': '{:,.0f}', '비중(%)': '{:.2f}%'}), use_container_width=True)
        else:
            st.warning("선택하신 필터 조건에 해당하는 종목 거래 내역이 없습니다.")

with tab4:
    st.subheader("🔍 특정 ETF 종목의 LP 점유율 파악")
    etf_list = df_filtered.groupby(['a_code', '종목명'])['총LP거래대금'].sum().sort_values(ascending=False).reset_index()
    etf_options = [f"[{row['a_code']}] {row['종목명']}" for _, row in etf_list.iterrows()]
    selected_etf_str = st.selectbox("종목 검색 (거래대금 순 배열):", etf_options)
    if selected_etf_str:
        a_code_target = selected_etf_str.split("]")[0][1:]
        df_target = df_filtered[df_filtered['a_code'] == a_code_target]
        tot_target = df_target['총LP거래대금'].sum()
        st.write(f"**해당 ETF 기간 총 거래대금:** {tot_target/100_000_000:,.0f} 억원")
        target_lp_df = df_target.groupby('회원사명')[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', 'LP매도거래량', 'LP매수거래량']].sum().reset_index()
        target_lp_df = target_lp_df[target_lp_df['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)
        target_lp_df['평균매도단가'] = np.where(target_lp_df['LP매도거래량'] > 0, target_lp_df['LP매도거래대금'] / target_lp_df['LP매도거래량'], 0)
        target_lp_df['평균매수단가'] = np.where(target_lp_df['LP매수거래량'] > 0, target_lp_df['LP매수거래대금'] / target_lp_df['LP매수거래량'], 0)
        target_lp_df['체결수량(min)'] = target_lp_df[['LP매도거래량', 'LP매수거래량']].min(axis=1)
        target_lp_df['추정매매이익'] = (target_lp_df['평균매도단가'] - target_lp_df['평균매수단가']) * target_lp_df['체결수량(min)']
        target_lp_df['대금(억)'] = target_lp_df['총LP거래대금'] / 100_000_000
        target_lp_df['매도대금(억)'] = target_lp_df['LP매도거래대금'] / 100_000_000
        target_lp_df['매수대금(억)'] = target_lp_df['LP매수거래대금'] / 100_000_000
        target_lp_df['순매수대금(억)'] = target_lp_df['LP순매수대금'] / 100_000_000
        target_lp_df['추정매매이익(백만)'] = target_lp_df['추정매매이익'] / 1_000_000
        target_lp_df['점유율(%)'] = (target_lp_df['총LP거래대금'] / tot_target) * 100
        target_lp_df['누적점유율(%)'] = target_lp_df['점유율(%)'].cumsum()
        show_cols = ['회원사명', '대금(억)', '매도대금(억)', '매수대금(억)', '순매수대금(억)', '추정매매이익(백만)', '점유율(%)', '누적점유율(%)']
        st.dataframe(target_lp_df[show_cols].style.format({'대금(억)': '{:,.0f}', '매도대금(억)': '{:,.0f}', '매수대금(억)': '{:,.0f}', '순매수대금(억)': '{:,.0f}', '추정매매이익(백만)': '{:,.0f}', '점유율(%)': '{:.1f}%', '누적점유율(%)': '{:.1f}%'}), use_container_width=True, hide_index=True)
        

# ==========================================
# Tab 5: 운용사별 주력 ETF 분석
# ==========================================
with tab5:
    st.subheader("🏢 운용사(AMC)별 ETF 및 파트너 LP 분석")
    
    st.write("▼ **필터 조건 설정**")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t5 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t5_mkt')
    ast_filter_t5 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t5_ast')
    rep_filter_t5 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t5_rep')
    drv_filter_t5 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t5_drv')
    trk_filter_t5 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t5_trk')
    
    # 필터 적용
    df_t5 = df_filtered.copy()
    if mkt_filter_t5 != "전체": df_t5 = df_t5[df_t5['market'] == mkt_filter_t5]
    if ast_filter_t5 != "전체": df_t5 = df_t5[df_t5['asset'] == ast_filter_t5]
    if rep_filter_t5 != "전체": df_t5 = df_t5[df_t5['is_rep'] == rep_filter_t5]
    if drv_filter_t5 != "전체": df_t5 = df_t5[df_t5['deriv'] == drv_filter_t5]
    if trk_filter_t5 != "전체": df_t5 = df_t5[df_t5['tracking'] == trk_filter_t5]

    # 필터링된 데이터 기반으로 운용사 리스트 갱신
    amc_list = df_t5.groupby('amc')['총LP거래대금'].sum().sort_values(ascending=False).index.tolist()
    target_amc = st.selectbox("운용사 선택", amc_list, key='t5_amc_sel')
    
    if target_amc:
        df_amc = df_t5[df_t5['amc'] == target_amc]
        amc_tot_vol = df_amc['총LP거래대금'].sum()
        st.success(f"**{target_amc}** (선택된 필터 기준) 총 거래대금: {amc_tot_vol/100_000_000:,.0f} 억원")
        
        if amc_tot_vol > 0:
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.write("### 🥇 LP사별 거래대금 및 M/S 비중")
                lp_ms_df = df_amc.groupby('회원사명')['총LP거래대금'].sum().reset_index()
                lp_ms_df = lp_ms_df[lp_ms_df['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)
                lp_ms_df['거래대금(억)'] = lp_ms_df['총LP거래대금'] / 100_000_000
                lp_ms_df['M/S(%)'] = (lp_ms_df['총LP거래대금'] / amc_tot_vol) * 100
                
                st.dataframe(
                    lp_ms_df[['회원사명', '거래대금(억)', 'M/S(%)']].style.format({
                        '거래대금(억)': '{:,.0f}',
                        'M/S(%)': '{:.2f}%'
                    }),
                    use_container_width=True, hide_index=True
                )
                
            with col2:
                st.write("### 📊 개별 ETF 종목별 핵심 파트너 LP (1~3위)")
                result_data = []
                for (a_code, name), group in df_amc.groupby(['a_code', '종목명']):
                    tot_vol = group['총LP거래대금'].sum()
                    if tot_vol == 0: continue
                    lp_rank = group.groupby('회원사명')['총LP거래대금'].sum().sort_values(ascending=False)
                    row = {'단축코드': a_code, '종목명': name, '총대금(억)': tot_vol / 100_000_000}
                    for i in range(3):
                        if i < len(lp_rank):
                            row[f'{i+1}위_LP'] = lp_rank.index[i]
                            row[f'{i+1}위_비중(%)'] = (lp_rank.values[i] / tot_vol) * 100
                        else:
                            row[f'{i+1}위_LP'] = '-'
                            row[f'{i+1}위_비중(%)'] = 0
                    result_data.append(row)
                    
                st.dataframe(
                    pd.DataFrame(result_data).sort_values('총대금(억)', ascending=False).style.format({
                        '총대금(억)': '{:,.0f}', 
                        '1위_비중(%)': '{:.1f}%', 
                        '2위_비중(%)': '{:.1f}%', 
                        '3위_비중(%)': '{:.1f}%'
                    }), 
                    use_container_width=True, hide_index=True
                )
        else:
            st.warning("선택하신 필터 조건에 해당하는 거래 내역이 없습니다.")

with tab6:
    st.subheader("🎯 종목 집중도 (HHI 및 Top-N 의존도)")
    
    # 1. 5가지 ETF 섹터 필터 추가
    st.write("▼ **필터 조건 설정**")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t6 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t6_mkt')
    ast_filter_t6 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t6_ast')
    rep_filter_t6 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t6_rep')
    drv_filter_t6 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t6_drv')
    trk_filter_t6 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t6_trk')
    
    # 필터 적용
    df_t6 = df_filtered.copy()
    if mkt_filter_t6 != "전체": df_t6 = df_t6[df_t6['market'] == mkt_filter_t6]
    if ast_filter_t6 != "전체": df_t6 = df_t6[df_t6['asset'] == ast_filter_t6]
    if rep_filter_t6 != "전체": df_t6 = df_t6[df_t6['is_rep'] == rep_filter_t6]
    if drv_filter_t6 != "전체": df_t6 = df_t6[df_t6['deriv'] == drv_filter_t6]
    if trk_filter_t6 != "전체": df_t6 = df_t6[df_t6['tracking'] == trk_filter_t6]

    st.divider()

    # 2. 필터 적용된 데이터로 기존 HHI 지수 계산
    st.write("### 1️⃣ 필터 조건 적용 HHI 및 Top 3 종목 의존도")
    lp_vol = df_t6.groupby('회원사명')['총LP거래대금'].sum()
    lp_etf_vol = df_t6.groupby(['회원사명', '종목명'])[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금']].sum().reset_index()
    
    records = []
    for lp, lp_tot in lp_vol.items():
        if lp_tot == 0: continue
        etfs = lp_etf_vol[lp_etf_vol['회원사명'] == lp].sort_values('총LP거래대금', ascending=False)
        t1_name = etfs.iloc[0]['종목명'] if len(etfs) > 0 else "-"
        t1_vol = etfs.iloc[0]['총LP거래대금'] if len(etfs) > 0 else 0
        t2_name = etfs.iloc[1]['종목명'] if len(etfs) > 1 else "-"
        t2_vol = etfs.iloc[1]['총LP거래대금'] if len(etfs) > 1 else 0
        t3_name = etfs.iloc[2]['종목명'] if len(etfs) > 2 else "-"
        t3_vol = etfs.iloc[2]['총LP거래대금'] if len(etfs) > 2 else 0
        top3_vol_sum = t1_vol + t2_vol + t3_vol
        shares = (etfs['총LP거래대금'] / lp_tot) * 100
        hhi = (shares ** 2).sum()
        records.append({'회원사명': lp, '총대금(억)': lp_tot / 100_000_000, '1위 종목명': t1_name, '1위 비중(%)': (t1_vol / lp_tot) * 100, '2위 종목명': t2_name, '2위 비중(%)': (t2_vol / lp_tot) * 100, '3위 종목명': t3_name, '3위 비중(%)': (t3_vol / lp_tot) * 100, 'Top 3 누적비중(%)': (top3_vol_sum / lp_tot) * 100, 'HHI 지수': hhi})
    
    conc_df = pd.DataFrame(records).sort_values('HHI 지수', ascending=False)
    st.info("💡 **HHI (허핀달-허쉬만 지 지수)**: 포트폴리오 내 개별 종목 점유율의 제곱합 (0~10,000). 숫자가 클수록 소수 특정 종목에 거래가 기형적으로 집중되어 있음을 의미합니다.")
    st.dataframe(conc_df.style.format({'총대금(억)': '{:,.0f}', '1위 비중(%)': '{:.1f}%', '2위 비중(%)': '{:.1f}%', '3위 비중(%)': '{:.1f}%', 'Top 3 누적비중(%)': '{:.1f}%', 'HHI 지수': '{:,.0f}'}), use_container_width=True, hide_index=True)

    st.divider()

    # 3. 새로운 기능: LP사별 상위 10개 종목 합집합 매트릭스 (순위 기준)
    st.write("### 2️⃣ LP사별 주력 종목 비교 매트릭스 (Top 10 합집합 기준)")
    st.write("선택한 LP사들이 각각 가장 많이 거래한 상위 10개 종목들을 모아, 해당 종목이 각 LP사 내에서 **몇 위**인지 보여줍니다. (Top 10 밖인 경우 빈칸 처리됩니다.)")

    # 거래대금이 존재하는 전체 LP사 목록 (거래대금 순)
    all_active_lps = lp_vol.sort_values(ascending=False).index.tolist()
    
    # 다 지우면 전체가 보이도록 기본값(default)을 전체 LP로 설정
    selected_lps = st.multiselect(
        "비교할 LP사 선택 (선택을 지우면 전체가 조회됩니다)", 
        all_active_lps, 
        default=all_active_lps, 
        key='t6_lp_sel'
    )
    
    target_lps_t6 = selected_lps if selected_lps else all_active_lps
        
    if target_lps_t6:
        rank_records = []
        top_etfs_union = set()
        
        # 선택된 각 LP사별 상위 10개 종목 추출 및 순위 기록
        for lp in target_lps_t6:
            lp_etfs = lp_etf_vol[lp_etf_vol['회원사명'] == lp].sort_values('총LP거래대금', ascending=False)
            top10 = lp_etfs.head(10).copy()
            top10['순위'] = range(1, len(top10) + 1)
            
            for _, row in top10.iterrows():
                rank_records.append({
                    '종목명': row['종목명'],
                    '회원사명': row['회원사명'],
                    '순위': row['순위']
                })
                top_etfs_union.add(row['종목명'])
        
        if rank_records:
            rank_df = pd.DataFrame(rank_records)
            
            # Pivot 생성 (행: 종목명, 열: 회원사명, 값: 순위)
            pivot_t6 = rank_df.pivot(index='종목명', columns='회원사명', values='순위')
            
            # 정렬을 위해 합집합 종목들의 총 거래대금 합계 계산 (내부 정렬용)
            matrix_df = df_t6[
                (df_t6['회원사명'].isin(target_lps_t6)) & 
                (df_t6['종목명'].isin(top_etfs_union))
            ]
            vol_sum = matrix_df.groupby('종목명')['총LP거래대금'].sum()
            
            # 선택한 LP사 순서대로 열 정렬
            cols_ordered = [lp for lp in target_lps_t6 if lp in pivot_t6.columns]
            pivot_t6 = pivot_t6[cols_ordered]
            
            # 임시 컬럼을 추가해 전체 거래대금 기준으로 내림차순 정렬 후, 임시 컬럼 삭제
            pivot_t6['sort_vol'] = vol_sum
            pivot_t6 = pivot_t6.sort_values('sort_vol', ascending=False)
            pivot_t6 = pivot_t6.drop(columns=['sort_vol'])
            
            st.success(f"📌 {len(target_lps_t6)}개 LP사의 Top 10 종목을 병합하여 총 **{len(pivot_t6)}개**의 고유 ETF가 도출되었습니다.")
            
            # 순위를 보여주는 포맷 함수 (NaN인 경우 아예 빈 텍스트 출력)
            def format_rank(val):
                if pd.isna(val): return ""
                return f"{int(val)}위"
                
            # 결측치(NaN) 배경과 글씨를 모두 흰색으로 고정하는 함수
            def highlight_nan(val):
                if pd.isna(val):
                    return 'background-color: white; color: white;'
                return ''
            
            # 스타일 적용: Blues_r을 사용해 값이 작을수록(1위) 진하게 설정
            styled_pivot = pivot_t6.style \
                .format(format_rank, subset=cols_ordered) \
                .background_gradient(cmap='Blues_r', vmin=1, vmax=10, subset=cols_ordered)
            
            # Pandas 버전에 따라 applymap 또는 map 호출하여 빈칸 숨기기 적용
            if hasattr(styled_pivot, "map"):
                styled_pivot = styled_pivot.map(highlight_nan, subset=cols_ordered)
            else:
                styled_pivot = styled_pivot.applymap(highlight_nan, subset=cols_ordered)
                
            st.dataframe(styled_pivot, use_container_width=True)
        else:
            st.warning("선택하신 조건 및 LP사에 해당하는 거래 내역이 없습니다.")

with tab7:
    st.subheader("🇰🇷 KRX 공식 시장 거래대금 분석 (Open API 연동)")
    st.write("선택한 ETF 섹터의 **KRX 공식 전체 거래대금**을 조회하고, LP 거래대금과의 비율(**LP 관여율**)을 정확하게 분석합니다.")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t7 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t7_mkt')
    ast_filter_t7 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t7_ast')
    rep_filter_t7 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t7_rep')
    drv_filter_t7 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t7_drv')
    trk_filter_t7 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t7_trk')
    df_t7 = df_filtered.copy()
    if mkt_filter_t7 != "전체": df_t7 = df_t7[df_t7['market'] == mkt_filter_t7]
    if ast_filter_t7 != "전체": df_t7 = df_t7[df_t7['asset'] == ast_filter_t7]
    if rep_filter_t7 != "전체": df_t7 = df_t7[df_t7['is_rep'] == rep_filter_t7]
    if drv_filter_t7 != "전체": df_t7 = df_t7[df_t7['deriv'] == drv_filter_t7]
    if trk_filter_t7 != "전체": df_t7 = df_t7[df_t7['tracking'] == trk_filter_t7]
    target_etfs = [code.replace('A', '') for code in df_t7['a_code'].unique()]
    total_found_cnt = len(target_etfs)
    st.info(f"선택된 필터 조건 대상 ETF 종목 수: **{total_found_cnt} 개**")
    if st.button("📊 KRX Open API 데이터로 정확하게 분석하기", type="primary", key="btn_t7"):
        if "KRX_API_KEY" not in st.secrets:
            st.error("⚠️ `.streamlit/secrets.toml` 파일에 `KRX_API_KEY` 설정이 필요합니다.")
        elif total_found_cnt == 0:
            st.warning("선택된 종목이 없습니다. 필터를 변경해주세요.")
        else:
            st.warning("⏳ **API 밴(Ban) 방지를 위해 데이터를 안전한 속도로 가져오고 있습니다. 로딩 중에 버튼을 여러 번 누르지 마세요! (최대 30초 소요)**")
            with st.spinner("KRX Open API에서 데이터를 실시간 수집 중입니다... (🚨 과부하 방지를 위해 연타 금지)"):
                try:
                    tickers_tuple = tuple(target_etfs)
                    daily_market_val = get_krx_open_api_market_data(tickers_tuple, start_date, end_date)
                    if daily_market_val.empty:
                        st.error("데이터를 불러오지 못했습니다. API Key가 올바른지, 혹은 해당 기간에 영업일이 포함되었는지 확인해 주세요.")
                    else:
                        daily_market_val['시장거래대금'] = pd.to_numeric(daily_market_val['시장거래대금'], errors='coerce').fillna(0)
                        daily_market_val['시장거래대금(억)'] = daily_market_val['시장거래대금'] / 100_000_000
                        lp_daily = df_t7.groupby(df_t7['거래일자'])['총LP거래대금'].sum().to_frame(name='LP거래대금')
                        lp_daily['LP거래대금'] = pd.to_numeric(lp_daily['LP거래대금'], errors='coerce').fillna(0)
                        lp_daily['LP거래대금'] = lp_daily['LP거래대금'] / 2
                        lp_daily['LP거래대금(억)'] = lp_daily['LP거래대금'] / 100_000_000
                        merged_df = daily_market_val.join(lp_daily, how='outer').fillna(0)
                        merged_df['LP관여율(%)'] = np.where(merged_df['시장거래대금'] > 0, (merged_df['LP거래대금'] / merged_df['시장거래대금']) * 100, 0)
                        merged_df = merged_df.reset_index().rename(columns={'거래일자': '날짜', 'index': '날짜'}) 
                        merged_df['날짜'] = pd.to_datetime(merged_df['날짜']).dt.date
                        fig_t7 = px.line(merged_df, x='날짜', y=['시장거래대금(억)', 'LP거래대금(억)'], title="선택 섹터 KRX 공식 시장 거래대금 vs LP 총 거래대금(조정됨) 추이 (단위: 억원)", markers=True, labels={'value': '거래대금(억)', 'variable': '구분'})
                        st.plotly_chart(fig_t7, use_container_width=True)
                        fig_t7_ratio = px.bar(merged_df, x='날짜', y='LP관여율(%)', title="KRX 공식 시장 거래대금 대비 LP 관여율 (%)", text=merged_df['LP관여율(%)'].apply(lambda x: f"{x:.1f}%"), color_discrete_sequence=['#ff9f43'])
                        fig_t7_ratio.update_traces(textposition='outside')
                        st.plotly_chart(fig_t7_ratio, use_container_width=True)
                        st.dataframe(merged_df[['날짜', '시장거래대금(억)', 'LP거래대금(억)', 'LP관여율(%)']].style.format({'시장거래대금(억)': '{:,.0f}', 'LP거래대금(억)': '{:,.0f}', 'LP관여율(%)': '{:.2f}%'}), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")


# ==========================================
# Tab 8: 설정/환매 추이 추정
# ==========================================
with tab8:
    st.subheader("🔄 ETF별 설정/환매 추이 추정")
    st.write("선택하신 기준 기간 동안의 **(상장좌수 변동) × (평균 순자산가치)** 공식을 활용하여 종목별 순설정 자금 규모를 추정합니다.")
    
    st.write("▼ **필터 조건 설정**")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t8 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t8_mkt')
    ast_filter_t8 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t8_ast')
    rep_filter_t8 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t8_rep')
    drv_filter_t8 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t8_drv')
    trk_filter_t8 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t8_trk')
    
    amc_list_t8 = sorted([a for a in df_filtered['amc'].unique() if a != '미분류'])
    target_amcs_t8 = st.multiselect("운용사(AMC) 다중 선택 (비워두면 조건 내 전체 종목 조회)", amc_list_t8, key='t8_amc')
    
    df_t8 = df_filtered.copy()
    if mkt_filter_t8 != "전체": df_t8 = df_t8[df_t8['market'] == mkt_filter_t8]
    if ast_filter_t8 != "전체": df_t8 = df_t8[df_t8['asset'] == ast_filter_t8]
    if rep_filter_t8 != "전체": df_t8 = df_t8[df_t8['is_rep'] == rep_filter_t8]
    if drv_filter_t8 != "전체": df_t8 = df_t8[df_t8['deriv'] == drv_filter_t8]
    if trk_filter_t8 != "전체": df_t8 = df_t8[df_t8['tracking'] == trk_filter_t8]
    if target_amcs_t8: df_t8 = df_t8[df_t8['amc'].isin(target_amcs_t8)]
    
    target_etfs_t8 = [code.replace('A', '') for code in df_t8['a_code'].unique()]
    
    st.divider()
    
    if search_type == "특정 일자 조회":
        calc_start_date = start_date - timedelta(days=1)
        while calc_start_date.weekday() >= 5: calc_start_date -= timedelta(days=1)
        calc_end_date = end_date
        st.info(f"📅 **단일 일자 선택됨:** 직전 영업일인 **{calc_start_date}**부터 조회일인 **{calc_end_date}**까지의 변동을 계산합니다.")
    else:
        calc_start_date = start_date
        calc_end_date = end_date
        st.info(f"📅 **기간 조회 선택됨:** **{calc_start_date}**부터 **{calc_end_date}**까지의 변동을 계산합니다.")
    
    if st.button("📊 순설정 자금 추정 테이블 생성", type="primary", key="btn_t8"):
        if "KRX_API_KEY" not in st.secrets:
            st.error("⚠️ `.streamlit/secrets.toml` 파일에 `KRX_API_KEY` 설정이 필요합니다.")
        elif len(target_etfs_t8) == 0:
            st.warning("선택하신 조건에 해당하는 종목이 없습니다. 필터를 변경해 주세요.")
        else:
            st.warning("⏳ **데이터 스냅샷을 조회 중입니다. 로딩 중에 버튼을 연타하지 마세요! (약 5초 소요)**")
            with st.spinner("KRX API에서 상장좌수 및 NAV 스냅샷을 수집 중입니다..."):
                df_start = get_krx_snapshot(calc_start_date)
                df_end = get_krx_snapshot(calc_end_date)
                
                if df_start.empty or df_end.empty:
                    st.error("해당 기간의 데이터를 KRX에서 불러오지 못했습니다. 장 휴장일이거나 범위를 벗어났을 수 있습니다.")
                else:
                    def extract_snap_data(df_snap):
                        df_snap['ISU_CD'] = df_snap['ISU_CD'].astype(str)
                        df_snap['short_code'] = df_snap['ISU_CD'].str.extract(r'(\d{6})')[0]
                        nav_col = next((c for c in ['NAV', 'TDD_NAV', 'IDX_NAV'] if c in df_snap.columns), None)
                        share_col = next((c for c in ['LIST_SHRS', 'LST_SHRS', 'LST_STK_VL'] if c in df_snap.columns), None)
                        
                        df_res = pd.DataFrame()
                        df_res['short_code'] = df_snap['short_code']
                        df_res['ISU_NM'] = df_snap.get('ISU_NM', df_snap.get('ISU_ABBRV', ''))
                        df_res['NAV'] = pd.to_numeric(df_snap[nav_col].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0) if nav_col else 0
                        df_res['SHARES'] = pd.to_numeric(df_snap[share_col].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0) if share_col else 0
                        
                        return df_res.drop_duplicates('short_code').set_index('short_code')
                        
                    ds_start = extract_snap_data(df_start)
                    ds_end = extract_snap_data(df_end)
                    common_idx = [t for t in target_etfs_t8 if t in ds_start.index and t in ds_end.index]
                    
                    if not common_idx:
                        st.warning("선택하신 종목들에 대한 API 스냅샷 데이터를 찾을 수 없습니다.")
                    else:
                        ds_start = ds_start.loc[common_idx]
                        ds_end = ds_end.loc[common_idx]
                        
                        res_df = pd.DataFrame(index=common_idx)
                        res_df['종목명'] = ds_end['ISU_NM']
                        res_df['좌수증감'] = ds_end['SHARES'] - ds_start['SHARES']
                        res_df['평균NAV'] = (ds_start['NAV'] + ds_end['NAV']) / 2
                        res_df['추정대금(원)'] = res_df['좌수증감'] * res_df['평균NAV']
                        res_df['추정대금(억)'] = res_df['추정대금(원)'] / 100_000_000
                        
                        # 순설정액만 단독으로 유지
                        res_df['순설정(억)'] = res_df['추정대금(억)']
                        
                        res_df = res_df.sort_values('순설정(억)', ascending=False).reset_index()
                        res_df.rename(columns={'index': '단축코드'}, inplace=True)
                        
                        st.success(f"✅ 데이터 조회 완료! (시작 기준일: {df_start['query_date'].iloc[0]} / 종료 기준일: {df_end['query_date'].iloc[0]})")
                        st.write(f"조회된 종목 수: **{len(res_df)} 개**")
                        
                        def highlight_net(val):
                            if val > 0: return 'color: #10b981; font-weight: bold'
                            elif val < 0: return 'color: #ef4444; font-weight: bold'
                            return ''
                            
                        st.dataframe(
                            res_df[['단축코드', '종목명', '순설정(억)']].style
                            .format({'순설정(억)': '{:,.0f}'})
                            .map(highlight_net, subset=['순설정(억)']),
                            use_container_width=True, hide_index=True
                        )

    st.divider()
    st.subheader("🔍 특정 ETF 일자별 설정/환매 상세 분석")
    etf_options_t8 = df_t8[['a_code', '종목명']].drop_duplicates()
    etf_options_list = [f"[{row['a_code'].replace('A', '')}] {row['종목명']}" for _, row in etf_options_t8.iterrows()]
    selected_etf_t8 = st.selectbox("일자별 분석을 진행할 ETF를 선택하세요:", ["선택 안함"] + etf_options_list)
    
    if selected_etf_t8 != "선택 안함":
        target_code_t8 = selected_etf_t8.split("]")[0][1:]
        target_name_t8 = selected_etf_t8.split("] ")[1]
        
        if st.button(f"'{target_name_t8}' 상세 분석 실행", type="primary", key="btn_t8_detail"):
            st.warning("⏳ **해당 종목의 과거 일별 데이터를 수집 중입니다. (10~30초 소요)**")
            with st.spinner(f"[{target_name_t8}] 일별 변동 내역을 가져오는 중입니다..."):
                df_daily = get_krx_daily_nav_shares(target_code_t8, calc_start_date, calc_end_date)
                if df_daily.empty:
                    st.error("해당 종목의 기간 내 데이터를 찾을 수 없습니다.")
                else:
                    df_daily['prev_SHARES'] = df_daily['SHARES'].shift(1)
                    df_daily['prev_NAV'] = df_daily['NAV'].shift(1)
                    df_daily = df_daily[df_daily['날짜'].dt.date >= calc_start_date].copy()
                    
                    if df_daily.empty:
                        st.error("유효한 영업일 데이터가 없습니다.")
                    else:
                        df_daily['좌수증감'] = df_daily['SHARES'] - df_daily['prev_SHARES']
                        df_daily['평균NAV'] = (df_daily['NAV'] + df_daily['prev_NAV']) / 2
                        df_daily['추정대금(원)'] = df_daily['좌수증감'] * df_daily['평균NAV']
                        df_daily['추정대금(억)'] = df_daily['추정대금(원)'] / 100_000_000
                        df_daily['순설정(억)'] = df_daily['추정대금(억)']
                        
                        total_net = df_daily['순설정(억)'].sum()
                        
                        st.metric("선택 기간 순설정 합계", f"{total_net:,.0f} 억원")
                        
                        df_daily['날짜_str'] = df_daily['날짜'].dt.strftime('%Y-%m-%d')
                        df_daily['구분'] = np.where(df_daily['순설정(억)'] >= 0, '설정(+)', '환매(-)')
                        
                        fig = px.bar(
                            df_daily, x='날짜_str', y='순설정(억)', color='구분',
                            color_discrete_map={'설정(+)': '#10b981', '환매(-)': '#ef4444'},
                            title=f"[{target_name_t8}] 일자별 순설정대금 추이 (단위: 억원)",
                            labels={'날짜_str': '영업일', '순설정(억)': '순설정 대금(억)'}
                        )
                        st.plotly_chart(fig, use_container_width=True)


# ==========================================
# Tab 9: NAV 괴리율 조회 (신규 탭)
# ==========================================
with tab9:
    st.subheader("📐 ETF NAV vs 종가 괴리(Disparity) 분석")
    st.write("조회하고 싶은 **특정 단일 일자**를 선택하여, KRX 공식 **NAV와 종가의 차이(`NAV - 종가`)** 및 괴리율을 분석합니다.")
    
    # 9번 탭 전용 단일 날짜 선택기
    target_snap_date = st.date_input(
        "괴리율을 조회할 기준 일자를 선택하세요",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
        key='t9_date_picker'
    )
    
    st.write("▼ **필터 조건 설정**")
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t9 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t9_mkt')
    ast_filter_t9 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t9_ast')
    rep_filter_t9 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t9_rep')
    drv_filter_t9 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t9_drv')
    trk_filter_t9 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t9_trk')
    
    amc_list_t9 = sorted([a for a in df_filtered['amc'].unique() if a != '미분류'])
    target_amcs_t9 = st.multiselect("운용사(AMC) 다중 선택 (비워두면 조건 내 전체 종목 조회)", amc_list_t9, key='t9_amc')
    
    df_t9 = df_filtered.copy()
    if mkt_filter_t9 != "전체": df_t9 = df_t9[df_t9['market'] == mkt_filter_t9]
    if ast_filter_t9 != "전체": df_t9 = df_t9[df_t9['asset'] == ast_filter_t9]
    if rep_filter_t9 != "전체": df_t9 = df_t9[df_t9['is_rep'] == rep_filter_t9]
    if drv_filter_t9 != "전체": df_t9 = df_t9[df_t9['deriv'] == drv_filter_t9]
    if trk_filter_t9 != "전체": df_t9 = df_t9[df_t9['tracking'] == trk_filter_t9]
    if target_amcs_t9: df_t9 = df_t9[df_t9['amc'].isin(target_amcs_t9)]
    
    target_etfs_t9 = [code.replace('A', '') for code in df_t9['a_code'].unique()]
    
    st.divider()
    
    if st.button("📊 NAV 괴리 분석 데이터 불러오기", type="primary", key="btn_t9"):
        if "KRX_API_KEY" not in st.secrets:
            st.error("⚠️ `.streamlit/secrets.toml` 파일에 `KRX_API_KEY` 설정이 필요합니다.")
        elif len(target_etfs_t9) == 0:
            st.warning("선택하신 조건에 해당하는 종목이 없습니다. 필터를 변경해 주세요.")
        else:
            st.warning(f"⏳ **{target_snap_date} 기준 KRX 데이터를 조회 중입니다. (약 3초 소요)**")
            
            with st.spinner("KRX API에서 최신 NAV 및 종가 데이터를 수집 중입니다..."):
                df_snap = get_krx_snapshot(target_snap_date)
                
                if df_snap.empty:
                    st.error("해당 일자의 데이터를 가져오지 못했습니다. 장 휴장일이거나 유효하지 않은 날짜입니다.")
                else:
                    df_snap['ISU_CD'] = df_snap['ISU_CD'].astype(str)
                    df_snap['short_code'] = df_snap['ISU_CD'].str.extract(r'(\d{6})')[0]
                    
                    nav_col = next((c for c in ['NAV', 'TDD_NAV', 'IDX_NAV'] if c in df_snap.columns), None)
                    # 💡 실제 응답 컬럼인 'TDD_CLSPRC' 우선 탐색 반영
                    close_col = next((c for c in ['TDD_CLSPRC', 'TDD_CLPR', 'CLPR', 'STCK_PRC'] if c in df_snap.columns), None)
                    
                    if not nav_col or not close_col:
                        st.error(f"API 응답에서 NAV 또는 종가 컬럼을 찾을 수 없습니다. (발견된 컬럼: {list(df_snap.columns)})")
                    else:
                        res9 = pd.DataFrame()
                        res9['단축코드'] = df_snap['short_code']
                        res9['종목명'] = df_snap.get('ISU_NM', df_snap.get('ISU_ABBRV', ''))
                        res9['NAV'] = pd.to_numeric(df_snap[nav_col].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                        res9['종가'] = pd.to_numeric(df_snap[close_col].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                        
                        res9 = res9[res9['단축코드'].isin(target_etfs_t9)].drop_duplicates('단축코드').copy()
                        
                        if res9.empty:
                            st.warning("조건에 매칭되는 종목 데이터가 없습니다.")
                        else:
                            res9['괴리금액(NAV-종가)'] = res9['NAV'] - res9['종가']
                            res9['괴리율(%)'] = np.where(res9['NAV'] > 0, ((res9['NAV'] - res9['종가']) / res9['NAV']) * 100, 0)
                            
                            st.success(f"✅ 분석 완료! (기준일: {df_snap['query_date'].iloc[0]})")
                            
                            top_10 = res9.sort_values('괴리금액(NAV-종가)', ascending=False).head(10).reset_index(drop=True)
                            bot_10 = res9.sort_values('괴리금액(NAV-종가)', ascending=True).head(10).reset_index(drop=True)
                            
                            col_l, col_r = st.columns(2)
                            with col_l:
                                st.write("### 📈 NAV - 종가 큰 종목 상위 10개 (고평가/괴리 확대)")
                                st.dataframe(
                                    top_10[['단축코드', '종목명', 'NAV', '종가', '괴리금액(NAV-종가)', '괴리율(%)']].style.format({
                                        'NAV': '{:,.0f}원', '종가': '{:,.0f}원', 
                                        '괴리금액(NAV-종가)': '{:+,.0f}원', '괴리율(%)': '{:+.2f}%'
                                    }), use_container_width=True, hide_index=True
                                )
                                
                            with col_r:
                                st.write("### 📉 NAV - 종가 작은 종목 하위 10개 (저평가/마이너스 괴리)")
                                st.dataframe(
                                    bot_10[['단축코드', '종목명', 'NAV', '종가', '괴리금액(NAV-종가)', '괴리율(%)']].style.format({
                                        'NAV': '{:,.0f}원', '종가': '{:,.0f}원', 
                                        '괴리금액(NAV-종가)': '{:+,.0f}원', '괴리율(%)': '{:+.2f}%'
                                    }), use_container_width=True, hide_index=True
                                )
                                
                            st.divider()
                            st.subheader("📋 필터링된 전체 종목 NAV 괴리 상세 목록")
                            
                            full_table = res9.sort_values('괴리금액(NAV-종가)', ascending=False).reset_index(drop=True)
                            st.dataframe(
                                full_table[['단축코드', '종목명', 'NAV', '종가', '괴리금액(NAV-종가)', '괴리율(%)']].style.format({
                                    'NAV': '{:,.0f}원', '종가': '{:,.0f}원', 
                                    '괴리금액(NAV-종가)': '{:+,.0f}원', '괴리율(%)': '{:+.2f}%'
                                }), use_container_width=True, hide_index=True
                            )
