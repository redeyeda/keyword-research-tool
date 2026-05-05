#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
키워드 리서치 툴 v3.0
네이버 광고 API + 검색 API + 데이터랩(트렌드 + 쇼핑인사이트) 통합
"""

import streamlit as st
import requests
import hashlib
import hmac
import base64
import time
import json
import re
import io
import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────
#  설정 파일 저장/불러오기
# ──────────────────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def load_config():
    """
    API 키 우선순위:
    1순위 - Streamlit Cloud Secrets
    2순위 - 로컬 config.json
    """
    try:
        if hasattr(st, "secrets") and "naver_api_key" in st.secrets:
            return {
                "naver_api_key":       st.secrets.get("naver_api_key", ""),
                "naver_secret_key":    st.secrets.get("naver_secret_key", ""),
                "naver_customer_id":   st.secrets.get("naver_customer_id", ""),
                "naver_client_id":     st.secrets.get("naver_client_id", ""),
                "naver_client_secret": st.secrets.get("naver_client_secret", ""),
                "claude_api_key":      st.secrets.get("claude_api_key", ""),
                "_source": "cloud",
            }
    except:
        pass
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                d["_source"] = "local"
                return d
        except:
            pass
    return {}

def save_config(data: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        return False

# ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="키워드 리서치 툴 v3", page_icon="🔍", layout="wide")
st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:800;
  background:linear-gradient(90deg,#2E5090,#4A90D9);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;}
.sub-title{color:#666;font-size:.9rem;margin-bottom:1.5rem;}
.stat-box{background:#F0F4FF;border-radius:12px;padding:1rem 1.2rem;
  text-align:center;border-left:4px solid #2E5090;margin-bottom:.5rem;}
.stat-num{font-size:1.8rem;font-weight:800;color:#2E5090;}
.stat-lbl{font-size:.8rem;color:#666;}
.api-group{background:#F8F9FA;border-radius:8px;padding:.8rem;margin-bottom:.5rem;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🔍 키워드 리서치 툴 v3.0</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">네이버 광고API + 검색API + 데이터랩(트렌드·쇼핑인사이트) 통합 분석</div>', unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
#  사이드바
# ──────────────────────────────────────────────────────────────
cfg = load_config()   # 저장된 키 불러오기

with st.sidebar:
    st.header("⚙️ API 키 설정")

    st.markdown('<div class="api-group">', unsafe_allow_html=True)
    st.markdown("**📊 네이버 검색광고 API**")
    naver_api_key     = st.text_input("ACCESS_LICENSE", value=cfg.get("naver_api_key",""),     type="password", key="ad_key")
    naver_secret_key  = st.text_input("SECRET_KEY",     value=cfg.get("naver_secret_key",""),  type="password", key="ad_secret")
    naver_customer_id = st.text_input("CUSTOMER_ID",    value=cfg.get("naver_customer_id",""), type="password", key="ad_cid")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="api-group">', unsafe_allow_html=True)
    st.markdown("**🔎 네이버 오픈API (검색·데이터랩)**")
    st.caption("developers.naver.com → 내 애플리케이션")
    naver_client_id     = st.text_input("Client ID",     value=cfg.get("naver_client_id",""),     type="password", key="open_id")
    naver_client_secret = st.text_input("Client Secret", value=cfg.get("naver_client_secret",""), type="password", key="open_secret")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="api-group">', unsafe_allow_html=True)
    st.markdown("**🤖 Claude API**")
    claude_api_key = st.text_input("API Key (sk-ant-...)", value=cfg.get("claude_api_key",""), type="password", key="claude_key")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 저장 버튼 (로컬 전용 / 클라우드는 Secrets 사용)
    if cfg.get("_source") == "cloud":
        st.success("✅ Streamlit Cloud Secrets에서 자동 로드됨")
        st.caption("키 변경은 Streamlit Cloud 대시보드 → Secrets에서 수정하세요.")
    else:
        if st.button("💾 API 키 저장 (로컬)", use_container_width=True, type="primary"):
            if not any([naver_api_key, naver_secret_key, naver_customer_id, claude_api_key]):
                st.warning("저장할 키를 먼저 입력하세요.")
            else:
                ok = save_config({
                    "naver_api_key":      naver_api_key,
                    "naver_secret_key":   naver_secret_key,
                    "naver_customer_id":  naver_customer_id,
                    "naver_client_id":    naver_client_id,
                    "naver_client_secret":naver_client_secret,
                    "claude_api_key":     claude_api_key,
                })
                if ok:
                    st.success("✅ 저장 완료!")

        if os.path.exists(CONFIG_FILE):
            mtime = datetime.fromtimestamp(os.path.getmtime(CONFIG_FILE))
            st.caption(f"📁 로컬 저장: {mtime.strftime('%Y-%m-%d %H:%M')}")
            if st.button("🗑️ 저장된 키 삭제", use_container_width=True):
                os.remove(CONFIG_FILE)
                st.rerun()
        else:
            st.caption("📁 저장된 키 없음")

    st.divider()
    st.markdown("**📌 API 키 위치**")
    st.caption("광고API: searchad.naver.com → API 관리")
    st.caption("오픈API: developers.naver.com → 내 애플리케이션")
    st.caption("Claude: console.anthropic.com → API Keys")


# ══════════════════════════════════════════════════════════════
#  API 함수
# ══════════════════════════════════════════════════════════════

def _ad_sig(secret_key, timestamp, method, path):
    msg = f"{timestamp}.{method}.{path}"
    h   = hmac.new(secret_key.encode(), msg.encode(), hashlib.sha256)
    return base64.b64encode(h.digest()).decode()

def _clean_key(val):
    """API 키에서 공백·줄바꿈·따옴표 제거"""
    if not val: return ""
    return str(val).strip().strip('"').strip("'").strip()

def _make_ad_headers(api_key, secret_key, cid, path):
    """매 요청마다 새 타임스탬프와 서명 생성"""
    ts = str(int(time.time() * 1000))
    return {
        "X-Timestamp": ts,
        "X-API-KEY":   _clean_key(api_key),
        "X-Customer":  cid,
        "X-Signature": _ad_sig(_clean_key(secret_key), ts, "GET", path),
    }

def get_naver_keyword_stats(keywords, api_key, secret_key, customer_id):
    BASE = "https://api.searchad.naver.com"
    path = "/keywordstool"

    # customer_id 숫자만 추출
    cid = re.sub(r"[^0-9]", "", str(customer_id).strip())
    if not cid:
        st.error("CUSTOMER_ID가 비어있습니다. Secrets에서 naver_customer_id 확인하세요.")
        return []

    # 키워드 정제 — 한글/영문/숫자/공백만 허용
    clean_kw = []
    for k in keywords:
        if not k:
            continue
        k = str(k).strip()
        # 한글, 영문, 숫자, 공백만 허용
        k = re.sub(r"[^가-힣a-zA-Z0-9 ]", "", k)
        k = " ".join(k.split()).strip()
        if len(k) >= 2 and k not in clean_kw:
            clean_kw.append(k)

    if not clean_kw:
        st.warning("유효한 키워드가 없습니다.")
        return []

    results = []
    fail_count = 0

    # 키워드 1개씩 개별 호출 (쉼표 조인 방식 오류 방지)
    for kw in clean_kw:
        hdrs = _make_ad_headers(api_key, secret_key, cid, path)
        try:
            r = requests.get(
                BASE + path,
                headers=hdrs,
                params={"hintKeywords": kw, "showDetail": "1"},
                timeout=15,
            )
            if r.status_code == 200:
                results.extend(r.json().get("keywordList", []))
            else:
                fail_count += 1
                if len(results) == 0 and fail_count == 1:
                    st.error(f"네이버 광고API 오류[{r.status_code}]: {r.text[:300]}")
                    return []
        except Exception as e:
            fail_count += 1
            if len(results) == 0 and fail_count == 1:
                st.error(f"네이버 광고API 연결 오류: {e}")
                return []
        time.sleep(0.15)  # 과호출 방지

    if fail_count > 0:
        st.warning(f"일부 키워드 조회 실패: {fail_count}건 / 성공: {len(results)}건")

    return results


def get_naver_suggest(keyword):
    try:
        r = requests.get("https://ac.search.naver.com/nx/ac",
                         params={"q": keyword, "con": "1", "frm": "nv", "ans": "2",
                                 "r_format": "json", "r_enc": "UTF-8",
                                 "l_enc": "UTF-8", "st": "100"}, timeout=10)
        if r.status_code == 200:
            out = []
            for group in r.json().get("items", []):
                for item in group:
                    if item and isinstance(item, list) and item[0]:
                        out.append(item[0])
            return out[:20]
    except: pass
    return []

def get_google_suggest(keyword):
    try:
        r = requests.get("https://suggestqueries.google.com/complete/search",
                         params={"client": "firefox", "hl": "ko", "q": keyword},
                         headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        if r.status_code == 200:
            d = r.json()
            return d[1][:20] if len(d) > 1 else []
    except: pass
    return []

def get_claude_longtail(main_keyword, api_key):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        prompt = f"""당신은 한국 SEO 전문 마케터입니다.
메인 키워드: "{main_keyword}"

[네이버 블로그 최적화 키워드 20개]
- 3단어 이상, 8자 이상, 정보탐색 의도
- 추천/방법/후기/비교/효과/종류/가격/차이/순위 등 포함

[구글/티스토리 최적화 키워드 20개]
- 4단어 이상, 10자 이상, 질문형/정보형
- 블로그가 공식사이트보다 상위 노출될 만한 롱테일

반드시 JSON만:
{{"naver_keywords":["키워드",...20개],"google_keywords":["키워드",...20개]}}"""
        msg   = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt.encode("utf-8").decode("utf-8")}]
        )
        text  = msg.content[0].text
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if match:
            d = json.loads(match.group())
            return d.get("naver_keywords", []), d.get("google_keywords", [])
    except Exception as e:
        st.warning(f"Claude API 오류: {e}")
    return [], []

def get_blog_doc_count(keyword, client_id, client_secret):
    try:
        r = requests.get("https://openapi.naver.com/v1/search/blog.json",
                         headers={"X-Naver-Client-Id": client_id,
                                  "X-Naver-Client-Secret": client_secret},
                         params={"query": keyword, "display": 1, "start": 1, "sort": "sim"},
                         timeout=10)
        if r.status_code == 200:
            return r.json().get("total", 0)
    except: pass
    return -1

def get_search_trend(keywords_batch, client_id, client_secret):
    """데이터랩 검색어 트렌드 — 최근 6개월 방향 분석"""
    end_date   = datetime.today()
    start_date = end_date - timedelta(days=180)
    # 키워드 그룹명과 실제 키워드 매핑 보존
    kw_list = [k.strip() for k in keywords_batch[:5] if k and k.strip()]
    if not kw_list:
        return {}
    body = {
        "startDate":     start_date.strftime("%Y-%m-%d"),
        "endDate":       end_date.strftime("%Y-%m-%d"),
        "timeUnit":      "month",
        "keywordGroups": [{"groupName": kw, "keywords": [kw]} for kw in kw_list],
    }
    try:
        r = requests.post(
            "https://openapi.naver.com/v1/datalab/search",
            headers={
                "X-Naver-Client-Id":     _clean_key(client_id),
                "X-Naver-Client-Secret": _clean_key(client_secret),
                "Content-Type":          "application/json",
            },
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
            timeout=15,
        )
        if r.status_code != 200:
            st.warning(f"트렌드 API 오류[{r.status_code}]: {r.text[:200]}")
            return {}
        result = {}
        for group in r.json().get("results", []):
            kw   = group.get("title", "")
            data = group.get("data", [])
            if not data or len(data) < 2:
                result[kw] = "→ 유지"
                continue
            h      = max(len(data) // 2, 1)
            first  = sum(d.get("ratio", 0) for d in data[:h]) / h
            second = sum(d.get("ratio", 0) for d in data[h:]) / max(len(data) - h, 1)
            if first == 0:
                result[kw] = "→ 유지"
                continue
            chg = (second - first) / first * 100
            result[kw] = (
                "↑ 급상승" if chg >= 25  else
                "↗ 상승"   if chg >= 10  else
                "→ 유지"   if chg >= -5  else
                "↘ 하락"   if chg >= -20 else
                "↓ 급하락"
            )
        return result
    except Exception as e:
        st.warning(f"트렌드 API 오류: {e}")
        return {}

# 쇼핑 의도 키워드 (API 대신 키워드 분석으로 판별)
SHOPPING_WORDS = [
    "구매","구입","주문","배송","할인","세일","가격","최저가","쿠폰","무료배송",
    "추천","후기","리뷰","사용기","구매후기","별점","평점","장단점",
    "어디서","사는곳","파는곳","구하는곳","온라인","오프라인","매장",
    "브랜드","정품","인기","베스트","신상","신제품","한정","품절",
    "효과","성분","사용법","용량","ml","g","개","세트","묶음",
]

def get_shopping_insight(keywords_batch, client_id, client_secret):
    """
    쇼핑 전환 가능성 — 키워드 의미 분석 기반 (API 대체)
    쇼핑인사이트 API가 일반 계정에서 제한되어 키워드 분석 방식으로 대체
    반환: {keyword: ratio(0~40)}
    """
    result = {}
    for kw in keywords_batch:
        if not kw:
            continue
        score = 0
        # 쇼핑 의도 단어 포함 여부
        for word in SHOPPING_WORDS:
            if word in kw:
                score += 8
        # 키워드 길이 — 짧을수록 구매 의도 높음
        if len(kw) <= 5:
            score += 10
        elif len(kw) <= 8:
            score += 6
        # 숫자 포함 (용량·가격 등)
        if any(c.isdigit() for c in kw):
            score += 5
        result[kw] = min(score, 40)
    return result


# ══════════════════════════════════════════════════════════════
#  점수 계산
# ══════════════════════════════════════════════════════════════
INFO_WORDS = [
    "추천","방법","후기","비교","효과","종류","가격","차이","이유","원인",
    "해결","선택","고르는","좋은","최고","리뷰","사용법","부작용","성분",
    "기간","순위","하는법","이란","효능","주의사항","vs","특징","선택방법",
]

def _parse_count(val):
    if val is None: return 5
    s = str(val).strip()
    if s.startswith("<"): return 5
    try: return int(float(s))
    except: return 5

def _blog_label(count):
    if count < 0:     return "확인불가"
    if count < 3000:  return "🟢 낮음"
    if count < 15000: return "🟡 중간"
    return                    "🔴 높음"

def _blog_score(count):
    if count < 0:     return 2
    if count < 3000:  return 1
    if count < 15000: return 2
    return                    3

def calculate_scores(keyword_data, claude_naver, claude_google,
                     blog_counts, trend_map, shopping_map):
    naver_set  = set(claude_naver)
    google_set = set(claude_google)
    rows = []
    trend_bonus_map = {"↑ 급상승":15,"↗ 상승":8,"→ 유지":0,"↘ 하락":-5,"↓ 급하락":-10}

    for kw in keyword_data:
        keyword = kw.get("relKeyword","").strip()
        if not keyword: continue

        pc     = _parse_count(kw.get("monthlyPcQcCnt"))
        mobile = _parse_count(kw.get("monthlyMobileQcCnt"))
        total  = pc + mobile

        comp_map = {"낮음":1,"중간":2,"높음":3,"low":1,"medium":2,"high":3}
        comp_raw = str(kw.get("compIdx",""))
        comp     = comp_map.get(comp_raw, 2)
        comp_kor = {"낮음":"낮음","중간":"중간","높음":"높음",
                    "low":"낮음","medium":"중간","high":"높음"}.get(comp_raw, comp_raw)

        char_len  = len(keyword)
        has_info  = any(w in keyword for w in INFO_WORDS)
        blog_cnt  = blog_counts.get(keyword, -1)
        bc        = _blog_score(blog_cnt)
        trend     = trend_map.get(keyword, "")
        shop_ratio= shopping_map.get(keyword, None)
        tb        = trend_bonus_map.get(trend, 0)

        # 네이버 블로그 적합도
        n  = 0
        n += 25 if 500<=total<=10000 else (18 if 100<=total<500 else (10 if total>10000 else 6))
        n += {1:15,2:8,3:2}[comp]
        n += {1:20,2:12,3:3}.get(bc, 10)   # 블로그 실경쟁도 (핵심)
        n += 20 if char_len>=8 else (10 if char_len>=5 else 0)
        n += 15 if has_info else 0
        n += tb
        if keyword in naver_set: n = min(n+8, 100)

        # 구글/티스토리 적합도
        g  = 0
        g += 22 if 100<=total<=5000 else (10 if total>5000 else (15 if total>=30 else 4))
        g += {1:20,2:12,3:3}[comp]
        g += {1:18,2:10,3:2}.get(bc, 8)
        g += 25 if char_len>=11 else (15 if char_len>=8 else (6 if char_len>=5 else 0))
        g += 15 if has_info else 0
        g += tb
        if keyword in google_set: g = min(g+8, 100)

        # 쇼핑 전환 점수
        shop_score = min(int(shop_ratio*3),100) if shop_ratio is not None else None

        # 추천 용도
        uses = []
        if min(n,100)>=60:                         uses.append("📝 네이버블로그")
        if min(g,100)>=60:                         uses.append("🔍 구글티스토리")
        if shop_ratio is not None and shop_ratio>10: uses.append("🛒 쇼핑몰연결")
        if trend in ("↑ 급상승","↗ 상승"):         uses.append("⚡ 트렌드")
        recommend = " / ".join(uses) if uses else "참고용"

        rows.append({
            "키워드":              keyword,
            "월간검색(PC)":        pc,
            "월간검색(모바일)":    mobile,
            "월간검색(합계)":      total,
            "광고경쟁도":          comp_kor,
            "블로그_문서수":       blog_cnt if blog_cnt>=0 else "미조회",
            "블로그_실경쟁도":     _blog_label(blog_cnt),
            "트렌드(6개월)":       trend if trend else "미조회",
            "쇼핑전환점수":        shop_score if shop_score is not None else "미조회",
            "네이버블로그_적합도": min(max(n,0),100),
            "구글티스토리_적합도": min(max(g,0),100),
            "정보성":              "✅" if has_info else "",
            "롱테일":              "✅" if char_len>=8 else "",
            "Claude추천_네이버":   "✅" if keyword in naver_set else "",
            "Claude추천_구글":     "✅" if keyword in google_set else "",
            "추천_용도":           recommend,
        })
    return rows


# ══════════════════════════════════════════════════════════════
#  엑셀 저장
# ══════════════════════════════════════════════════════════════
def to_excel_bytes(rows):
    df_all    = pd.DataFrame(rows)
    df_naver  = df_all[df_all["네이버블로그_적합도"]>=60].sort_values("네이버블로그_적합도",ascending=False).reset_index(drop=True)
    df_google = df_all[df_all["구글티스토리_적합도"]>=60].sort_values("구글티스토리_적합도",ascending=False).reset_index(drop=True)
    df_shop   = df_all[df_all["쇼핑전환점수"].apply(lambda x: isinstance(x,int) and x>=15)].sort_values("쇼핑전환점수",ascending=False).reset_index(drop=True)
    df_trend  = df_all[df_all["트렌드(6개월)"].isin(["↑ 급상승","↗ 상승"])].sort_values("네이버블로그_적합도",ascending=False).reset_index(drop=True)
    df_full   = df_all.sort_values("월간검색(합계)",ascending=False).reset_index(drop=True)

    buf = io.BytesIO()
    sheets = [
        (df_naver,  "📝 네이버블로그 추천"),
        (df_google, "🔍 구글_티스토리 추천"),
        (df_shop,   "🛒 쇼핑몰 연결 추천"),
        (df_trend,  "⚡ 트렌드 상승 키워드"),
        (df_full,   "📊 전체 키워드"),
    ]
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for df, sheet in sheets:
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws    = writer.sheets[sheet]
            hfill = PatternFill("solid", fgColor="2E5090")
            hfont = Font(bold=True, color="FFFFFF", size=10)
            ctr   = Alignment(horizontal="center", vertical="center")
            thin  = Side(style="thin", color="CCCCCC")
            bdr   = Border(left=thin,right=thin,top=thin,bottom=thin)
            for ci in range(1, ws.max_column+1):
                c = ws.cell(row=1,column=ci)
                c.fill=hfill; c.font=hfont; c.alignment=ctr; c.border=bdr
            for ri in range(2, ws.max_row+1):
                fill = PatternFill("solid", fgColor=("F5F8FF" if ri%2==0 else "FFFFFF"))
                for ci in range(1, ws.max_column+1):
                    c=ws.cell(row=ri,column=ci); c.fill=fill; c.alignment=ctr; c.border=bdr
            for ci, col in enumerate(df.columns,1):
                ml = max(len(str(col)), df.iloc[:,ci-1].astype(str).map(len).max() if len(df)>0 else 0)
                ws.column_dimensions[get_column_letter(ci)].width = max(14,ml+4)
            ws.freeze_panes = "A2"
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
#  메인 UI
# ══════════════════════════════════════════════════════════════
col1, col2 = st.columns([3, 1])
with col1:
    main_keyword = st.text_input("🎯 메인 키워드", placeholder="예: 세라마이드 크림")
with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    run_btn = st.button("🚀 분석 시작", use_container_width=True, type="primary")

with st.expander("⚙️ 고급 옵션"):
    top_n = st.slider("트렌드·쇼핑 API 호출 키워드 수 (상위 N개)", 10, 50, 30, 5)
    st.caption("전체 키워드 중 검색량 상위 N개에만 트렌드·쇼핑 API를 호출합니다. (속도·API 비용 최적화)")

st.divider()

if run_btn:
    missing = []
    if not main_keyword.strip():  missing.append("메인 키워드")
    if not naver_api_key:         missing.append("광고API - ACCESS_LICENSE")
    if not naver_secret_key:      missing.append("광고API - SECRET_KEY")
    if not naver_customer_id:     missing.append("광고API - CUSTOMER_ID")
    if missing:
        st.error(f"⛔ 입력 필요: {', '.join(missing)}")
        st.stop()

    use_open = bool(naver_client_id and naver_client_secret)
    if not use_open:
        st.info("ℹ️ 오픈API 키 미입력 → 블로그 문서수·트렌드·쇼핑 분석을 건너뜁니다.")

    prog = st.progress(0, text="분석 준비 중...")
    log  = st.empty()

    # ① 자동완성
    log.info("① 네이버 연관 키워드 수집 중...")
    naver_sug = get_naver_suggest(main_keyword.strip())
    prog.progress(10, text=f"네이버 자동완성 {len(naver_sug)}개")

    log.info("② 구글 연관 키워드 수집 중...")
    google_sug = get_google_suggest(main_keyword.strip())
    prog.progress(18, text=f"구글 자동완성 {len(google_sug)}개")

    # ② Claude 롱테일
    log.info("③ Claude AI 롱테일 생성 중...")
    claude_naver, claude_google = [], []
    if claude_api_key:
        claude_naver, claude_google = get_claude_longtail(main_keyword.strip(), claude_api_key)
        if claude_naver or claude_google:
            prog.progress(30, text=f"AI 롱테일 네이버 {len(claude_naver)}개 / 구글 {len(claude_google)}개")
        else:
            prog.progress(30, text="AI 롱테일 생성 실패 (크레딧 확인) → 자동완성 키워드로 진행")
    else:
        prog.progress(30, text="Claude API 키 없음 → 자동완성 키워드로 진행")

    # ③ 통합
    all_kw = list(dict.fromkeys(
        [main_keyword.strip()] + naver_sug + google_sug + claude_naver + claude_google
    ))
    prog.progress(35, text=f"키워드 통합 {len(all_kw)}개 (중복제거)")

    # ④ 광고 API (검색량·경쟁도)
    log.info("④ 네이버 검색량 & 광고 경쟁도 조회 중...")
    naver_stats = get_naver_keyword_stats(all_kw, naver_api_key, naver_secret_key, naver_customer_id)
    prog.progress(55, text=f"광고API {len(naver_stats)}개 완료")

    if not naver_stats:
        st.error("⛔ 네이버 광고API 데이터 없음. API 키를 확인하세요.")
        st.stop()

    # 상위 N개 선별
    sorted_stats = sorted(naver_stats, key=lambda k: _parse_count(k.get("monthlyPcQcCnt",0))+_parse_count(k.get("monthlyMobileQcCnt",0)), reverse=True)
    top_kws = [k.get("relKeyword","") for k in sorted_stats[:top_n] if k.get("relKeyword")]

    blog_counts  = {}
    trend_map    = {}
    shopping_map = {}

    if use_open and top_kws:
        # ⑤ 블로그 문서수
        log.info(f"⑤ 블로그 실경쟁도 조회 중 (상위 {len(top_kws)}개)...")
        for idx, kw in enumerate(top_kws):
            blog_counts[kw] = get_blog_doc_count(kw, naver_client_id, naver_client_secret)
            prog.progress(55 + int((idx+1)/len(top_kws)*15), text=f"블로그 문서수 ({idx+1}/{len(top_kws)})")
            time.sleep(0.2)

        # ⑥ 트렌드 (5개씩 배치)
        log.info("⑥ 검색어 트렌드 조회 중...")
        for i in range(0, len(top_kws), 5):
            batch = top_kws[i:i+5]
            trend_map.update(get_search_trend(batch, naver_client_id, naver_client_secret))
            prog.progress(min(70+int((i+5)/len(top_kws)*10),80), text=f"트렌드 조회 중 ({min(i+5,len(top_kws))}/{len(top_kws)})")
            time.sleep(0.5)

        # ⑦ 쇼핑 인사이트 (5개씩 배치)
        log.info("⑦ 쇼핑 인사이트 조회 중...")
        for i in range(0, len(top_kws), 5):
            batch = top_kws[i:i+5]
            shopping_map.update(get_shopping_insight(batch, naver_client_id, naver_client_secret))
            prog.progress(min(80+int((i+5)/len(top_kws)*10),90), text=f"쇼핑 인사이트 ({min(i+5,len(top_kws))}/{len(top_kws)})")
            time.sleep(0.5)

    # ⑧ 최종 점수
    log.info("⑧ 최종 점수 계산 중...")
    scored = calculate_scores(naver_stats, claude_naver, claude_google,
                              blog_counts, trend_map, shopping_map)
    excel_data = to_excel_bytes(scored)
    prog.progress(100, text="✅ 분석 완료!")
    log.empty()

    # 요약
    n_rec = sum(1 for r in scored if r["네이버블로그_적합도"]>=60)
    g_rec = sum(1 for r in scored if r["구글티스토리_적합도"]>=60)
    s_rec = sum(1 for r in scored if isinstance(r["쇼핑전환점수"],int) and r["쇼핑전환점수"]>=15)
    t_rec = sum(1 for r in scored if r["트렌드(6개월)"] in ("↑ 급상승","↗ 상승"))

    c1,c2,c3,c4,c5 = st.columns(5)
    for col,num,lbl in [(c1,len(scored),"전체 키워드"),(c2,n_rec,"📝 네이버 추천"),
                        (c3,g_rec,"🔍 구글 추천"),(c4,s_rec,"🛒 쇼핑 연결"),(c5,t_rec,"⚡ 트렌드 상승")]:
        col.markdown(f'<div class="stat-box"><div class="stat-num">{num}</div>'
                     f'<div class="stat-lbl">{lbl}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    df_all    = pd.DataFrame(scored)
    df_naver  = df_all[df_all["네이버블로그_적합도"]>=60].sort_values("네이버블로그_적합도",ascending=False).reset_index(drop=True)
    df_google = df_all[df_all["구글티스토리_적합도"]>=60].sort_values("구글티스토리_적합도",ascending=False).reset_index(drop=True)
    df_shop   = df_all[df_all["쇼핑전환점수"].apply(lambda x: isinstance(x,int) and x>=15)].sort_values("쇼핑전환점수",ascending=False).reset_index(drop=True)
    df_trend  = df_all[df_all["트렌드(6개월)"].isin(["↑ 급상승","↗ 상승"])].sort_values("네이버블로그_적합도",ascending=False).reset_index(drop=True)

    tab1,tab2,tab3,tab4,tab5 = st.tabs(["📝 네이버 블로그","🔍 구글/티스토리","🛒 쇼핑몰 연결","⚡ 트렌드 상승","📊 전체"])
    with tab1: st.dataframe(df_naver,  use_container_width=True, height=420)
    with tab2: st.dataframe(df_google, use_container_width=True, height=420)
    with tab3: st.dataframe(df_shop,   use_container_width=True, height=420)
    with tab4: st.dataframe(df_trend,  use_container_width=True, height=420)
    with tab5: st.dataframe(df_all.sort_values("월간검색(합계)",ascending=False).reset_index(drop=True), use_container_width=True, height=420)

    st.markdown("<br>", unsafe_allow_html=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_kw = re.sub(r'[\\/:*?"<>|]', "_", main_keyword.strip())
    st.download_button(
        label="📥 엑셀 다운로드 (5개 시트)",
        data=excel_data,
        file_name=f"keyword_research_{safe_kw}_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )
