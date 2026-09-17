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
st.set_page_config(page_title="ETF Market Monitoring (v8.0)", layout="wide")
st.title("📊 ETF Market Monitoring Dashboard (통합판)")

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

        since_date = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%d-%b-%Y")
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
# 2. 구글 드라이브 원본 CSV 업데이트(업로드) 함수
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

        csv_buffer = io.BytesIO()
        df_export.to_csv(csv_buffer, index=False, encoding='cp949')
        csv_buffer.seek(0)

        media = MediaIoBaseUpload(csv_buffer, mimetype='text/csv', resumable=True)
        service.files().update(fileId=file_id, media_body=media).execute()
        st.toast("✅ 신규 데이터를 구글 드라이브 원본 파일에 성공적으로 업데이트했습니다!", icon="💾")
        
    except Exception as e:
        st.error(f"구글 드라이브 업데이트 실패: {e}")

# ----------------------------------------------------------------------
# 3. 메인 데이터 로드 및 전처리
# ----------------------------------------------------------------------
@st.cache_data(ttl=3600)
def load_data():
    csv_file_id = "1aJ5x-GYsdZzNOwmsV0RBDObWxBV-8Hcb"
    excel_file_id = "1xdKEXMRXf0TECNRvUedJ4jU9Pz29cRo4"

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
                request = service.files().export_media(fileId=excel_file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                excel_bytes = io.BytesIO(request.execute())
            except:
                request = service.files().get_media(fileId=excel_file_id)
                excel_bytes = io.BytesIO(request.execute())
            df_excel = pd.read_excel(excel_bytes)

            request_csv = service.files().get_media(fileId=csv_file_id)
            csv_bytes = io.BytesIO(request_csv.execute())
            df_base = pd.read_csv(csv_bytes, encoding='cp949', thousands=',')
        except Exception as e:
            pass

    if df_excel is None:
        try:
            excel_url = f"https://docs.google.com/spreadsheets/d/{excel_file_id}/export?format=xlsx"
            df_excel = pd.read_excel(excel_url)
        except:
            excel_url = f"https://drive.google.com/uc?export=download&id={excel_file_id}"
            df_excel = pd.read_excel(excel_url)

    if df_base is None:
        csv_url = f"https://drive.google.com/uc?export=download&id={csv_file_id}"
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

    if len(df) > base_row_count:
        if "gcp_service_account" in st.secrets:
            update_drive_csv(df, csv_file_id)

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

    return df, master_db

df, master_db = load_data()
if df.empty:
    st.error("⚠️ 데이터 로드에 실패하였거나 표시할 ETF 데이터가 없습니다.")
    st.stop()


# ----------------------------------------------------------------------
# 4. KRX Open API 연동 함수들
# ----------------------------------------------------------------------
@st.cache_data(ttl=86400, show_spinner=False)
def get_krx_open_api_market_data(tickers, start_date, end_date):
    """
    KRX Open API(etf_bydd_trd)를 통해 선택한 기간과 종목들의 일별 거래대금 합계를 가져옵니다.
    """
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
    """
    지정된 날짜 기준 가장 최근 영업일의 KRX ETF 전종목 스냅샷 데이터(NAV, 상장좌수 등)를 가져옵니다.
    """
    api_key = st.secrets.get("KRX_API_KEY")
    if not api_key: return pd.DataFrame()
    
    url = "https://data-dbg.krx.co.kr/svc/apis/etp/etf_bydd_trd"
    headers = {"AUTH_KEY": api_key, "AUTH-KEY": api_key}
    
    # 휴장일을 대비해 최대 7일 전까지 탐색하여 최초로 응답받는 영업일 데이터를 반환
    for i in range(7):
        dt = pd.to_datetime(target_date) - timedelta(days=i)
        if dt.weekday() >= 5: # 토, 일요일 패스
            continue
            
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


# ----------------------------------------------------------------------
# 사이드바 (Global Date Filter)
# ----------------------------------------------------------------------
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
# UI Tabs 구성 (8번 탭 추가)
# ----------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "1. 종합 대시보드", "2. ETF 구분별 분석", "3. LP사 다각도 분석", 
    "4. ETF별 주력 LP 분석", "5. 운용사별 주력 ETF 분석",
    "6. 종목 집중도 분석", "7. 시장 전체 거래대금 (KRX)", "8. 설정/환매 추이 (추정)"
])

# ==========================================
# Tab 1: 종합 대시보드
# ==========================================
with tab1:
    st.subheader("📊 시장 핵심 지표 및 추이 (KPI)")

    trend_type = st.radio(
        "🔍 데이터 필터링 관점 선택 (아래 모든 지표에 적용됩니다)", 
        ["시장 전체 (Total Market)", "특정 LP사 (Specific LP)", "특정 운용사 (Specific AMC)"], 
        horizontal=True, 
        key='t1_trend'
    )

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
        if target_lps:
            df_t1 = df_t1[df_t1['회원사명'].isin(target_lps)]
        else:
            is_data_empty = True 

    elif trend_type == "특정 운용사 (Specific AMC)":
        if target_amcs:
            df_t1 = df_t1[df_t1['amc'].isin(target_amcs)]
        else:
            is_data_empty = True 

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
            fig = px.bar(
                lp_total, x=lp_total.index, y=lp_total.values, 
                title="선택된 조건 내 LP사별 총 거래대금 (억원)",
                labels={'y': '거래대금(억)', '회원사명': '증권사'},
                color_discrete_sequence=['#4A90E2']
            )
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

# ==========================================
# Tab 2: ETF 구분별 분석
# ==========================================
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
        st.dataframe(show_df.set_index('순위').style.format({
            '거래대금(억)': '{:,.0f}', '추정매매손익(백만)': '{:,.0f}', '점유율(%)': '{:.1f}%'
        }), use_container_width=True)

    with col2:
        if not agg_df.empty:
            fig2 = px.pie(agg_df.head(10), values='거래대금(억)', names='회원사명', hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    st.write("### 2️⃣ LP사 매수/매도 스탠스 (순매수 현황)")
    imb_df = agg_df.copy()
    imb_df['순매수비율(%)'] = np.where(imb_df['총LP거래대금'] > 0, (imb_df['LP순매수대금'] / imb_df['총LP거래대금']) * 100, 0)
    imb_df = imb_df.sort_values('LP순매수대금', ascending=False)

    fig_imb = px.bar(
        imb_df, x='회원사명', y='LP순매수대금',
        title="필터링된 섹터 내 순매수/순매도 누적대금",
        labels={'LP순매수대금': '순매수대금(원)', '회원사명': '증권사'},
        color='LP순매수대금', color_continuous_scale=px.colors.diverging.RdBu_r,
        text=imb_df['LP순매수대금'] / 100_000_000
    )
    fig_imb.update_traces(texttemplate='%{text:,.0f}억', textposition='outside')
    st.plotly_chart(fig_imb, use_container_width=True)

    st.write("▼ 세부 금액 데이터")
    imb_df['매도대금(억)'] = imb_df['LP매도거래대금'] / 100_000_000
    imb_df['매수대금(억)'] = imb_df['LP매수거래대금'] / 100_000_000
    imb_df['순매수대금(억)'] = imb_df['LP순매수대금'] / 100_000_000
    
    st.dataframe(
        imb_df[['회원사명', '매도대금(억)', '매수대금(억)', '순매수대금(억)', '거래대금(억)', '순매수비율(%)']].style.format({
            '매도대금(억)': '{:,.0f}', '매수대금(억)': '{:,.0f}', 
            '순매수대금(억)': '{:,.0f}', '거래대금(억)': '{:,.0f}', '순매수비율(%)': '{:.2f}%'
        }), use_container_width=True, hide_index=True
    )

# ==========================================
# Tab 3: LP사 다각도 분석
# ==========================================
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
            st.dataframe(
                lp_sector[lp_sector['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)
                [['category_key', '대금(억)', '내부비중(%)', '섹터내_MS(%)']].style.format({'대금(억)': '{:,.0f}', '내부비중(%)': '{:.1f}%', '섹터내_MS(%)': '{:.1f}%'}), 
                use_container_width=True, hide_index=True
            )

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
                amc_lp['AMC내_MS(%)'] = amc_lp.apply(
                    lambda r: (r['총LP거래대금'] / amc_mkt.get(r['amc'], 1)) * 100 if amc_mkt.get(r['amc'], 0) > 0 else 0, axis=1
                )
                
                st.dataframe(
                    amc_lp[amc_lp['총LP거래대금'] > 0].sort_values('총LP거래대금', ascending=False)
                    [['amc', '대금(억)', '내부비중(%)', 'AMC내_MS(%)']].style.format({'대금(억)': '{:,.0f}', '내부비중(%)': '{:.1f}%', 'AMC내_MS(%)': '{:.1f}%'}),
                    use_container_width=True, hide_index=True
                )
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
            detail_etfs = df_detail.groupby('종목명')[
                ['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', 'LP매도거래량', 'LP매수거래량']
            ].sum().reset_index()

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

            total_row = pd.DataFrame([{
                '순위': 0,
                '종목명': '📊 [총 합계]',
                '거래대금(억)': detail_etfs['거래대금(억)'].sum(),
                '매도대금(억)': detail_etfs['매도대금(억)'].sum(),
                '매수대금(억)': detail_etfs['매수대금(억)'].sum(),
                '순매수대금(억)': detail_etfs['순매수대금(억)'].sum(),
                '추정매매손익(백만)': detail_etfs['추정매매손익(백만)'].sum(),
                '비중(%)': 100.0
            }])

            detail_etfs = pd.concat([total_row, detail_etfs], ignore_index=True)

            st.write(f"해당 필터 조건 거래 종목 수: **{len(detail_etfs)-1:,}개** | 기간 총 거래대금: **{total_detail_vol/100_000_000:,.0f}억원**")

            show_cols = ['순위', '종목명', '거래대금(억)', '매도대금(억)', '매수대금(억)', '순매수대금(억)', '추정매매손익(백만)', '비중(%)']
            
            st.dataframe(
                detail_etfs[show_cols].set_index('순위').style.format({
                    '거래대금(억)': '{:,.0f}',
                    '매도대금(억)': '{:,.0f}',
                    '매수대금(억)': '{:,.0f}',
                    '순매수대금(억)': '{:,.0f}',
                    '추정매매손익(백만)': '{:,.0f}',
                    '비중(%)': '{:.2f}%'
                }), 
                use_container_width=True
            )
        else:
            st.warning("선택하신 필터 조건에 해당하는 종목 거래 내역이 없습니다.")

# ==========================================
# Tab 4: ETF별 주력 LP 분석
# ==========================================
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

        target_lp_df = df_target.groupby('회원사명')[
            ['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금', 'LP매도거래량', 'LP매수거래량']
        ].sum().reset_index()
        
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

        show_cols = [
            '회원사명', '대금(억)', '매도대금(억)', '매수대금(억)', '순매수대금(억)', 
            '추정매매이익(백만)', '점유율(%)', '누적점유율(%)'
        ]

        st.dataframe(
            target_lp_df[show_cols].style.format({
                '대금(억)': '{:,.0f}',
                '매도대금(억)': '{:,.0f}',
                '매수대금(억)': '{:,.0f}',
                '순매수대금(억)': '{:,.0f}',
                '추정매매이익(백만)': '{:,.0f}',
                '점유율(%)': '{:.1f}%', 
                '누적점유율(%)': '{:.1f}%'
            }), use_container_width=True, hide_index=True
        )

# ==========================================
# Tab 5: 운용사별 주력 ETF 분석
# ==========================================
with tab5:
    st.subheader("🏢 운용사(AMC)별 ETF 및 1~3위 핵심 파트너 LP")
    amc_list = df_filtered.groupby('amc')['총LP거래대금'].sum().sort_values(ascending=False).index.tolist()
    target_amc = st.selectbox("운용사 선택", amc_list)

    if target_amc:
        df_amc = df_filtered[df_filtered['amc'] == target_amc]
        st.success(f"**{target_amc}** 총 거래대금: {df_amc['총LP거래대금'].sum()/100_000_000:,.0f} 억원")

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
                '총대금(억)': '{:,.0f}', '1위_비중(%)': '{:.1f}%', '2위_비중(%)': '{:.1f}%', '3위_비중(%)': '{:.1f}%'
            }), use_container_width=True, hide_index=True
        )

# ==========================================
# Tab 6: 종목 집중도 분석
# ==========================================
with tab6:
    st.subheader("🎯 종목 집중도 (HHI 및 Top-N 의존도)")

    lp_vol = df_filtered.groupby('회원사명')['총LP거래대금'].sum()
    lp_etf_vol = df_filtered.groupby(['회원사명', '종목명'])[['총LP거래대금', 'LP매도거래대금', 'LP매수거래대금', 'LP순매수대금']].sum().reset_index()

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

        records.append({
            '회원사명': lp, 
            '총대금(억)': lp_tot / 100_000_000,
            '1위 종목명': t1_name,
            '1위 비중(%)': (t1_vol / lp_tot) * 100,
            '2위 종목명': t2_name,
            '2위 비중(%)': (t2_vol / lp_tot) * 100,
            '3위 종목명': t3_name,
            '3위 비중(%)': (t3_vol / lp_tot) * 100,
            'Top 3 누적비중(%)': (top3_vol_sum / lp_tot) * 100,
            'HHI 지수': hhi
        })

    conc_df = pd.DataFrame(records).sort_values('HHI 지수', ascending=False)

    st.info("💡 **HHI (허핀달-허쉬만 지수)**: 포트폴리오 내 개별 종목 점유율의 제곱합 (0~10,000). 숫자가 클수록 소수 특정 종목에 거래가 기형적으로 집중되어 있음을 의미합니다.")
    st.dataframe(
        conc_df.style.format({
            '총대금(억)': '{:,.0f}', 
            '1위 비중(%)': '{:.1f}%', 
            '2위 비중(%)': '{:.1f}%', 
            '3위 비중(%)': '{:.1f}%', 
            'Top 3 누적비중(%)': '{:.1f}%',
            'HHI 지수': '{:,.0f}'
        }), use_container_width=True, hide_index=True
    )
    
# ==========================================
# Tab 7: 시장 전체 거래대금 (KRX Open API 연동)
# ==========================================
with tab7:
    st.subheader("🇰🇷 KRX 공식 시장 거래대금 분석 (Open API 연동)")
    st.write("선택한 ETF 섹터의 **KRX 공식 전체 거래대금**을 조회하고, LP 거래대금과의 비율(**LP 관여율**)을 정확하게 분석합니다.")

    # 5개 섹터 필터링 UI
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t7 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t7_mkt')
    ast_filter_t7 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t7_ast')
    rep_filter_t7 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t7_rep')
    drv_filter_t7 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t7_drv')
    trk_filter_t7 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t7_trk')

    # 필터 적용
    df_t7 = df_filtered.copy()
    if mkt_filter_t7 != "전체": df_t7 = df_t7[df_t7['market'] == mkt_filter_t7]
    if ast_filter_t7 != "전체": df_t7 = df_t7[df_t7['asset'] == ast_filter_t7]
    if rep_filter_t7 != "전체": df_t7 = df_t7[df_t7['is_rep'] == rep_filter_t7]
    if drv_filter_t7 != "전체": df_t7 = df_t7[df_t7['deriv'] == drv_filter_t7]
    if trk_filter_t7 != "전체": df_t7 = df_t7[df_t7['tracking'] == trk_filter_t7]
    
    # KRX 종목코드 형식 변환 ('A069500' -> '069500')
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
                        
                        merged_df['LP관여율(%)'] = np.where(
                            merged_df['시장거래대금'] > 0, 
                            (merged_df['LP거래대금'] / merged_df['시장거래대금']) * 100, 
                            0
                        )
                        
                        merged_df = merged_df.reset_index().rename(columns={'거래일자': '날짜', 'index': '날짜'}) 
                        merged_df['날짜'] = pd.to_datetime(merged_df['날짜']).dt.date
                        
                        fig_t7 = px.line(
                            merged_df, x='날짜', y=['시장거래대금(억)', 'LP거래대금(억)'],
                            title="선택 섹터 KRX 공식 시장 거래대금 vs LP 총 거래대금(조정됨) 추이 (단위: 억원)",
                            markers=True,
                            labels={'value': '거래대금(억)', 'variable': '구분'}
                        )
                        st.plotly_chart(fig_t7, use_container_width=True)
                        
                        fig_t7_ratio = px.bar(
                            merged_df, x='날짜', y='LP관여율(%)',
                            title="KRX 공식 시장 거래대금 대비 LP 관여율 (%)",
                            text=merged_df['LP관여율(%)'].apply(lambda x: f"{x:.1f}%"),
                            color_discrete_sequence=['#ff9f43']
                        )
                        fig_t7_ratio.update_traces(textposition='outside')
                        st.plotly_chart(fig_t7_ratio, use_container_width=True)
                        
                        st.dataframe(
                            merged_df[['날짜', '시장거래대금(억)', 'LP거래대금(억)', 'LP관여율(%)']].style.format({
                                '시장거래대금(억)': '{:,.0f}',
                                'LP거래대금(억)': '{:,.0f}',
                                'LP관여율(%)': '{:.2f}%'
                            }), use_container_width=True, hide_index=True
                        )
                
                except Exception as e:
                    st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")

# ==========================================
# Tab 8: 설정/환매 추이 추정 (신규 탭)
# ==========================================
with tab8:
    st.subheader("🔄 ETF별 설정/환매 추이 추정")
    st.write("선택하신 기준 기간 동안의 **(상장좌수 변동) × (평균 순자산가치)** 공식을 활용하여 종목별 펀드 설정 및 환매 자금 규모를 추정합니다.")
    
    st.write("▼ **필터 조건 설정**")
    
    c1, c2, c3, c4, c5 = st.columns(5)
    mkt_filter_t8 = c1.selectbox("국내/해외", ["전체"] + list(df_filtered['market'].unique()), key='t8_mkt')
    ast_filter_t8 = c2.selectbox("주식/그외", ["전체"] + list(df_filtered['asset'].unique()), key='t8_ast')
    rep_filter_t8 = c3.selectbox("대표지수", ["전체"] + list(df_filtered['is_rep'].unique()), key='t8_rep')
    drv_filter_t8 = c4.selectbox("일반/파생", ["전체"] + list(df_filtered['deriv'].unique()), key='t8_drv')
    trk_filter_t8 = c5.selectbox("패시브/액티브", ["전체"] + list(df_filtered['tracking'].unique()), key='t8_trk')
    
    amc_list_t8 = sorted([a for a in df_filtered['amc'].unique() if a != '미분류'])
    target_amcs_t8 = st.multiselect("운용사(AMC) 다중 선택 (비워두면 조건 내 전체 종목 조회)", amc_list_t8, key='t8_amc')
    
    # 필터 적용
    df_t8 = df_filtered.copy()
    if mkt_filter_t8 != "전체": df_t8 = df_t8[df_t8['market'] == mkt_filter_t8]
    if ast_filter_t8 != "전체": df_t8 = df_t8[df_t8['asset'] == ast_filter_t8]
    if rep_filter_t8 != "전체": df_t8 = df_t8[df_t8['is_rep'] == rep_filter_t8]
    if drv_filter_t8 != "전체": df_t8 = df_t8[df_t8['deriv'] == drv_filter_t8]
    if trk_filter_t8 != "전체": df_t8 = df_t8[df_t8['tracking'] == trk_filter_t8]
    if target_amcs_t8: df_t8 = df_t8[df_t8['amc'].isin(target_amcs_t8)]
    
    target_etfs_t8 = [code.replace('A', '') for code in df_t8['a_code'].unique()]
    
    st.divider()
    
    # 시작일/종료일 스마트 매핑
    if search_type == "특정 일자 조회":
        # 단일 날짜 선택 시, 시작일을 '직전 영업일'로 계산
        calc_start_date = start_date - timedelta(days=1)
        while calc_start_date.weekday() >= 5:  # 주말 건너뛰기
            calc_start_date -= timedelta(days=1)
        calc_end_date = end_date
        st.info(f"📅 **단일 일자 선택됨:** 직전 영업일인 **{calc_start_date}**부터 조회일인 **{calc_end_date}**까지의 변동을 계산합니다.")
    else:
        calc_start_date = start_date
        calc_end_date = end_date
        st.info(f"📅 **기간 조회 선택됨:** **{calc_start_date}**부터 **{calc_end_date}**까지의 변동을 계산합니다.")
    
    if st.button("📊 설정/환매 자금 추정 테이블 생성", type="primary", key="btn_t8"):
        if "KRX_API_KEY" not in st.secrets:
            st.error("⚠️ `.streamlit/secrets.toml` 파일에 `KRX_API_KEY` 설정이 필요합니다.")
        elif len(target_etfs_t8) == 0:
            st.warning("선택하신 조건에 해당하는 종목이 없습니다. 필터를 변경해 주세요.")
        else:
            st.warning("⏳ **데이터 스냅샷을 조회 중입니다. 로딩 중에 버튼을 연타하지 마세요! (약 5초 소요)**")
            with st.spinner("KRX API에서 상장좌수 및 NAV 스냅샷을 수집 중입니다..."):
                
                # 시작일과 종료일 단 2번의 API 호출 (최적화)
                df_start = get_krx_snapshot(calc_start_date)
                df_end = get_krx_snapshot(calc_end_date)
                
                if df_start.empty or df_end.empty:
                    st.error("해당 기간의 데이터를 KRX에서 불러오지 못했습니다. 장 휴장일이거나 범위를 벗어났을 수 있습니다.")
                else:
                    def extract_snap_data(df_snap):
                        df_snap['ISU_CD'] = df_snap['ISU_CD'].astype(str)
                        df_snap['short_code'] = df_snap['ISU_CD'].str.extract(r'(\d{6})')[0]
                        
                        # API 리턴 키값의 유연성을 위한 대체 탐색 (NAV, LIST_SHRS 등)
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
                    
                    # 대상 종목 교집합 필터링
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
                        
                        # 대금 계산 로직 적용
                        res_df['추정대금(원)'] = res_df['좌수증감'] * res_df['평균NAV']
                        res_df['추정대금(억)'] = res_df['추정대금(원)'] / 100_000_000
                        
                        # 설정액(+), 환매액(- -> +절대값), 순설정액 구분
                        res_df['설정(억)'] = np.where(res_df['추정대금(억)'] > 0, res_df['추정대금(억)'], 0)
                        res_df['환매(억)'] = np.where(res_df['추정대금(억)'] < 0, np.abs(res_df['추정대금(억)']), 0)
                        res_df['순설정(억)'] = res_df['추정대금(억)']
                        
                        # 테이블 정렬 및 정리
                        res_df = res_df.sort_values('순설정(억)', ascending=False).reset_index()
                        res_df.rename(columns={'index': '단축코드'}, inplace=True)
                        
                        st.success(f"✅ 데이터 조회 완료! (시작 기준일: {df_start['query_date'].iloc[0]} / 종료 기준일: {df_end['query_date'].iloc[0]})")
                        st.write(f"조회된 종목 수: **{len(res_df)} 개**")
                        
                        # Pandas Styler를 통한 조건부 서식 지정
                        def highlight_net(val):
                            if val > 0:
                                return 'color: #10b981; font-weight: bold' # 그린
                            elif val < 0:
                                return 'color: #ef4444; font-weight: bold' # 레드
                            return ''
                            
                        st.dataframe(
                            res_df[['단축코드', '종목명', '설정(억)', '환매(억)', '순설정(억)']].style
                            .format({
                                '설정(억)': '{:,.0f}',
                                '환매(억)': '{:,.0f}',
                                '순설정(억)': '{:,.0f}'
                            })
                            .map(highlight_net, subset=['순설정(억)']),
                            use_container_width=True, hide_index=True
                        )
