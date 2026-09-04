#!/usr/bin/env python3
"""
event_pnl_dashboard.html에 붙여넣을 JS 상수를 엑셀 원본에서 자동 계산하는 스크립트.

README.md 6번 "데이터 갱신 워크플로우"에서 지금까지 사람이 매번 손으로 하던 계산 —
  - 판매일자 기준 채널×월별 매출/이익 집계
  - 행사 일정 + 판매현황을 매장·기간으로 매칭해 행사별 실적/진행중 잠정치 계산
  - 매장별 1:1 매칭(그리디, "계절성 기준" 날짜 거리)으로 전년동기 비교값 채우기
을 자동화한다. 산출물은 대시보드 파일을 직접 덮어쓰지 않고 검토용 .js 스니펫 +
요약 리포트로만 낸다 — 실제 대시보드에 반영하기 전에 사람이 한 번 확인하고
붙여넣는 것을 전제로 한다 (README에 나온 "정산완료 vs 진행중 판정", "판매일자 vs
주문번호 기준" 등 과거에 실제로 여러 번 버그가 났던 지점들이라, 자동 반영보다
검토 단계를 두는 쪽이 안전하다고 판단).

현재 다루는 범위 (v1):
  - CHANNEL_TOTAL_REVENUE / CHANNEL_MARGIN (+ _PRIOR) : 채널×월 전체 매출/마진
  - EVENTS : 행사 일정 × 판매현황을 매장/기간으로 매칭한 행사별 실적
  - 매장별 1:1 매칭(그리디, 계절성 기준)으로 백화점/라운지/밀로티의 revenuePlan/
    profitPlan(=전년동기 실적) 자동 계산
  - 취소 건 진단 리포트 (README 5번 버그#3 규칙 그대로 구현: 주문번호에서 정정
    접미사 제거 + 품목명으로 묶어서, 그룹 순수량<=0이면 취소)

아직 다루지 않는 범위 (다음 단계로 남겨둠 - 잘못 짐작해서 자동화하면 오히려
위험하다고 판단한 부분들):
  - 온라인/라운지/밀로티의 "계획(사전시뮬)" 데이터 — 실적 기반이 아니라 별도
    사전시뮬 자료가 출처라 이 스크립트의 입력(판매현황/행사일정)만으로는 계산 불가.
    온라인 채널의 revenuePlan/profitPlan은 항상 null로 남기고, 기존처럼 수기 입력.
  - PRODUCT_FULL_BY_STORE / PRODUCT_BY_MONTH_CHANNEL (품목별 성과) — 품목명에서
    "모델명만 추출"하는 정확한 규칙이 README에 명시돼 있지 않아, 잘못 추측한 정규식으로
    자동 생성하면 매장명이나 숫자가 조용히 틀려도 티가 안 날 위험이 있음. 다음 단계로
    미룸(스크립트 상단 TODO 참고).
  - PRODUCT_MATCH_LOTTESHOPPING (사전시뮬 vs 실적 매칭) — 사전시뮬 파일 포맷 미확정.

사용 예:
  python3 scripts/update_dashboard_data.py \\
      --sales 품목별판매현황.xlsx \\
      --schedule 행사일정.xlsx \\
      --dashboard event_pnl_dashboard.html \\
      --year 2026 --asof 2026-09-04 \\
      --out data_update_output.js --report data_update_report.md
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

SALES_REQUIRED_COLS = [
    "그룹명", "품목코드", "품목명", "주문번호", "매입처", "판매점",
    "고객명", "판매일자", "수량", "매출액", "실매출액", "입고가", "이익금", "이익율", "채널",
]
SCHEDULE_REQUIRED_COLS = ["매장", "시작일", "종료일"]

# 전년동기 매칭을 적용할 채널(전년동기 비교형). 온라인은 계획(사전시뮬) 기준이라 제외.
YOY_CHANNELS = {"백화점", "라운지", "밀로티"}

EOK = 100_000_000  # 억원


def read_excel_validated(path: str, required_cols: list[str], label: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise SystemExit(
            f"[{label}] 필수 컬럼이 없습니다: {missing}\n"
            f"  실제 컬럼: {list(df.columns)}\n"
            f"  README.md 2번 '데이터 구조' 항목의 컬럼명과 일치하는지 확인하세요."
        )
    return df


def load_store_channel_map(dashboard_html_path: str) -> dict[str, str]:
    """대시보드 HTML의 STORE_LISTS 상수를 그대로 읽어 매장→채널 매핑을 만든다.
    (채널 목록/매장 목록을 이 스크립트에 따로 하드코딩하면 대시보드와 어긋날 수 있어서,
    항상 대시보드 파일 자체를 단일 출처로 삼는다.)"""
    text = Path(dashboard_html_path).read_text(encoding="utf-8")
    m = re.search(r"const STORE_LISTS\s*=\s*\{(.*?)\n\};", text, re.S)
    if not m:
        raise SystemExit("대시보드 HTML에서 STORE_LISTS 상수를 찾지 못했습니다.")
    body = "{" + m.group(1) + "}"
    # JS 객체 리터럴(트레일링 콤마 포함)을 JSON으로 바꿔서 파싱
    body = re.sub(r",\s*([\]}])", r"\1", body)
    # "온라인": ALL_MALLS 처럼 변수 참조가 섞여 있어 STORE_LISTS 자체는 못 쓰고,
    # 배열 리터럴로 된 채널만 우선 파싱한 뒤, ALL_MALLS/OWN_MALL/EXTERNAL_MALLS는 별도 처리.
    own_mall_m = re.search(r'const OWN_MALL\s*=\s*"([^"]+)"', text)
    ext_malls_m = re.search(r"const EXTERNAL_MALLS\s*=\s*\[(.*?)\];", text, re.S)
    store_map: dict[str, str] = {}
    if own_mall_m and ext_malls_m:
        own_mall = own_mall_m.group(1)
        ext_malls = re.findall(r'"([^"]+)"', ext_malls_m.group(1))
        for m2 in [own_mall, *ext_malls]:
            store_map[m2] = "온라인"
    for ch_match in re.finditer(r'"([^"]+)"\s*:\s*\[(.*?)\]', body, re.S):
        ch, arr_body = ch_match.group(1), ch_match.group(2)
        if ch == "온라인":
            continue  # 위에서 ALL_MALLS로 이미 처리
        for store in re.findall(r'"([^"]+)"', arr_body):
            store_map[store] = ch
    if not store_map:
        raise SystemExit("STORE_LISTS에서 매장을 하나도 못 찾았습니다 - 파일 구조가 바뀌었는지 확인하세요.")
    return store_map


def month_key(d: pd.Timestamp) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def compute_channel_totals(sales: pd.DataFrame) -> dict:
    """채널×월별 매출/마진 전체 집계. CHANNEL_TOTAL_REVENUE/CHANNEL_MARGIN 두 딕셔너리를 만든다."""
    s = sales.copy()
    s["_month"] = s["판매일자"].apply(month_key)
    s["_year"] = s["판매일자"].dt.year
    out = {}
    for (year, ch), g in s.groupby(["_year", "채널"]):
        by_month = {}
        for month, mg in g.groupby("_month"):
            rev = mg["실매출액"].sum() / EOK
            profit = mg["이익금"].sum()
            margin = (profit / mg["실매출액"].sum() * 100) if mg["실매출액"].sum() else 0.0
            by_month[month] = {"revenue": round(rev, 4), "margin": round(margin, 2)}
        yr_rev = g["실매출액"].sum() / EOK
        yr_profit = g["이익금"].sum()
        yr_margin = (yr_profit / g["실매출액"].sum() * 100) if g["실매출액"].sum() else 0.0
        by_month["전체"] = {"revenue": round(yr_rev, 4), "margin": round(yr_margin, 2)}
        out.setdefault(int(year), {})[ch] = by_month
    return out


def cancel_diagnostics(sales: pd.DataFrame) -> pd.DataFrame:
    """README 5번 버그#3 규칙: 주문번호에서 정정 접미사(-2,-3,...) 제거 + 품목명으로 묶어서,
    그 그룹의 순수량이 0 이하면 취소로 판정. 매장 등 실 반영은 하지 않고 진단용 리포트만 만든다."""
    s = sales.copy()
    s["_base_order"] = s["주문번호"].astype(str).str.replace(r"-\d+$", "", regex=True)
    grp = s.groupby(["_base_order", "품목명"])["수량"].sum().reset_index()
    grp["취소"] = grp["수량"] <= 0
    return grp[grp["취소"]]


@dataclass
class EventRecord:
    name: str
    mall: str
    channel: str
    start: pd.Timestamp
    end: pd.Timestamp
    revenue: float  # 억원
    profit: float  # 억원
    order_count: int
    settled: bool


def build_event_records(schedule: pd.DataFrame, sales: pd.DataFrame, store_channel: dict) -> list[EventRecord]:
    records: list[EventRecord] = []
    unknown_stores = set()
    for _, row in schedule.iterrows():
        mall = str(row["매장"]).strip()
        channel = str(row["채널"]).strip() if "채널" in schedule.columns and pd.notna(row.get("채널")) else store_channel.get(mall)
        if channel is None:
            unknown_stores.add(mall)
            continue
        start = pd.to_datetime(row["시작일"])
        end = pd.to_datetime(row["종료일"])
        mask = (sales["판매점"] == mall) & (sales["채널"] == channel) & \
               (sales["판매일자"] >= start) & (sales["판매일자"] <= end)
        sub = sales[mask]
        revenue = sub["실매출액"].sum() / EOK
        profit = sub["이익금"].sum() / EOK
        order_count = sub["주문번호"].nunique()
        name = row["행사명"] if "행사명" in schedule.columns and pd.notna(row.get("행사명")) else \
            f"{mall} 행사({start:%m-%d}~{end:%m-%d})"
        records.append(EventRecord(
            name=name, mall=mall, channel=channel, start=start, end=end,
            revenue=round(revenue, 4), profit=round(profit, 4),
            order_count=int(order_count), settled=True,
        ))
    if unknown_stores:
        print(f"[경고] 채널을 알 수 없는 매장 {len(unknown_stores)}개는 건너뜀: {sorted(unknown_stores)}", file=sys.stderr)
    return records


def seasonal_distance(a: pd.Timestamp, b: pd.Timestamp) -> int:
    """연도를 무시하고 월-일만으로 계절적 거리(일수)를 계산. 연말/연초 경계도 짧게 잡히도록 원형 거리 사용."""
    doy_a = a.replace(year=2001).dayofyear if not (a.month == 2 and a.day == 29) else 60
    doy_b = b.replace(year=2001).dayofyear if not (b.month == 2 and b.day == 29) else 60
    diff = abs(doy_a - doy_b)
    return min(diff, 365 - diff)


def match_yoy_baseline(current: list[EventRecord], prior: list[EventRecord]) -> dict[int, EventRecord]:
    """매장별로, 이번해 행사를 시작일 순으로 순회하며 남아있는 전년 후보 중 계절적으로
    가장 가까운 것을 그리디로 배정한다(재사용 없음) - README 2번 '매장별 1:1 매칭' 규칙 그대로."""
    matches: dict[int, EventRecord] = {}
    by_store: dict[str, list[int]] = {}
    for i, ev in enumerate(prior):
        by_store.setdefault(ev.mall, []).append(i)
    current_sorted = sorted(range(len(current)), key=lambda i: current[i].start)
    for i in current_sorted:
        ev = current[i]
        pool = by_store.get(ev.mall, [])
        if not pool:
            continue
        best_j = min(pool, key=lambda j: seasonal_distance(ev.start, prior[j].start))
        matches[i] = prior[best_j]
        pool.remove(best_j)
    return matches


def to_event_dicts(current: list[EventRecord], baseline: dict[int, EventRecord],
                    asof: pd.Timestamp, sales: pd.DataFrame, store_channel: dict) -> list[dict]:
    out = []
    for i, ev in enumerate(current):
        base = baseline.get(i)
        d = {
            "name": ev.name,
            "month": month_key(ev.end),
            "channel": ev.channel,
            "mall": ev.mall,
            "startDate": ev.start.strftime("%Y-%m-%d"),
            "endDate": ev.end.strftime("%Y-%m-%d"),
        }
        if ev.channel in YOY_CHANNELS and base is not None:
            d["revenuePlan"] = base.revenue
            d["profitPlan"] = base.profit
        else:
            d["revenuePlan"] = None
            d["profitPlan"] = None
        if ev.end <= asof:
            d["profitActual"] = ev.profit
            d["revenueActual"] = ev.revenue
        elif ev.start > asof:
            # 아직 시작 전 - isSettled()는 profitActual이 없고 partialRevenue도 없으면
            # endDate<=now로 최종 판정하므로, 시작 전 행사는 그냥 실적 필드를 비워둔다.
            d["profitActual"] = None
            d["revenueActual"] = None
        else:
            mask = (sales["판매점"] == ev.mall) & (sales["채널"] == ev.channel) & \
                   (sales["판매일자"] >= ev.start) & (sales["판매일자"] <= asof)
            sub = sales[mask]
            d["profitActual"] = None
            d["revenueActual"] = None
            d["partialRevenue"] = round(sub["실매출액"].sum() / EOK, 4)
            d["partialProfit"] = round(sub["이익금"].sum() / EOK, 4)
            d["partialDays"] = (asof - ev.start).days + 1
            d["totalDays"] = (ev.end - ev.start).days + 1
            d["partialTx"] = int(sub["주문번호"].nunique())
        out.append(d)
    return out


def js_value(v):
    if v is None:
        return "null"
    if hasattr(v, "item"):  # numpy int64/float64 등을 파이썬 기본 타입으로
        v = v.item()
    if isinstance(v, float):
        return json.dumps(round(v, 4))
    return json.dumps(v, ensure_ascii=False)


def emit_events_js(events: list[dict]) -> str:
    lines = ["let EVENTS = ["]  # 대시보드 쪽이 "let"(구글시트 연동 시 재할당)이라 맞춰줌
    key_order = ["name", "month", "channel", "mall", "startDate", "endDate",
                 "profitPlan", "profitActual", "revenuePlan", "revenueActual",
                 "partialRevenue", "partialProfit", "partialDays", "totalDays", "partialTx"]
    for e in events:
        parts = [f"{k}: {js_value(e[k])}" for k in key_order if k in e]
        lines.append("  { " + ", ".join(parts) + " },")
    lines.append("];")
    return "\n".join(lines)


def emit_channel_totals_js(totals: dict, year: int, var_name: str, field: str) -> str:
    lines = [f"const {var_name} = {{"]
    for ch, by_month in totals.get(year, {}).items():
        pairs = ", ".join(f'"{m}": {v[field]}' for m, v in by_month.items())
        lines.append(f'  "{ch}": {{ {pairs} }},')
    lines.append("};")
    return "\n".join(lines)


def parse_channel_month_const(dashboard_text: str, var_name: str) -> dict:
    """대시보드에 이미 박혀있는 CHANNEL_TOTAL_REVENUE류 상수를 읽어온다 (덮어쓰기 전 급락/0
    같은 이상치를 잡아내는 sanity check용). 못 찾으면 빈 dict."""
    m = re.search(re.escape(f"const {var_name} = {{") + r"(.*?)\n\};", dashboard_text, re.S)
    if not m:
        return {}
    out = {}
    for ch_m in re.finditer(r'"([^"]+)"\s*:\s*\{([^}]*)\}', m.group(1)):
        ch, body = ch_m.group(1), ch_m.group(2)
        out[ch] = {mm.group(1): float(mm.group(2)) for mm in re.finditer(r'"([^"]+)"\s*:\s*([\d.]+)', body)}
    return out


def sanity_check(dashboard_path: str, totals: dict, current_year: int) -> list[str]:
    """새로 계산한 채널×월 매출이 기존 대시보드 값 대비 이상하게 튀는지(0으로 사라짐,
    반토막 등) 확인. 자동 반영 전 마지막 안전장치 - 여기서 경고가 나오면 watch_and_update.py는
    자동 반영하지 않고 사람 확인을 기다린다."""
    text = Path(dashboard_path).read_text(encoding="utf-8")
    old = parse_channel_month_const(text, "CHANNEL_TOTAL_REVENUE")
    new = totals.get(current_year, {})
    warnings = []
    for ch, old_by_month in old.items():
        old_total = old_by_month.get("전체")
        new_total = (new.get(ch) or {}).get("전체", {}).get("revenue")
        if old_total is None or not old_total:
            continue
        if new_total is None:
            warnings.append(f"'{ch}' 채널이 새 데이터에서 통째로 사라짐 (기존 {old_total}억)")
        elif new_total < old_total * 0.5:
            warnings.append(f"'{ch}' 채널 {current_year}년 전체 매출이 {old_total}억 → {new_total}억으로 급감(50%+ 감소)")
    return warnings


def splice_dashboard(dashboard_text: str, js_blocks: dict[str, str]) -> str:
    """새로 계산된 EVENTS/CHANNEL_* 블록으로 대시보드 안의 기존 상수 선언을 교체한다.
    js_blocks 키: EVENTS, CHANNEL_TOTAL_REVENUE, CHANNEL_TOTAL_REVENUE_PRIOR, CHANNEL_MARGIN,
    CHANNEL_MARGIN_PRIOR. 못 찾은 키는 조용히 건너뛰지 않고 예외를 던짐(자동화 중 대시보드
    구조가 바뀌어 조용히 반영 안 되는 사고를 막기 위해)."""
    text = dashboard_text
    patterns = {
        # 대시보드의 실제 EVENTS는 `let EVENTS = [...].map(e => ({...}));` 형태로,
        # 레거시 데이터(revenuePlan/revenueActual 명시 안 된 항목)를 위한 MARGIN_ASSUMPTION
        # 역산 보정이 뒤에 붙어있음. 새로 생성하는 EVENTS는 모든 필드를 항상 명시적으로
        # 채우기 때문에 이 보정이 필요 없어서, .map() 래퍼까지 통째로 새 배열로 교체한다.
        # (한 번 이 스크립트로 갱신되고 나면 .map() 래퍼 없는 단순 배열이 되므로, 둘 다 매치되게 함)
        "EVENTS": r"let EVENTS = \[.*?\n\](?:\.map\(e => \(\{.*?\}\)\))?;",
        "CHANNEL_TOTAL_REVENUE": r"const CHANNEL_TOTAL_REVENUE = \{.*?\n\};",
        "CHANNEL_TOTAL_REVENUE_PRIOR": r"const CHANNEL_TOTAL_REVENUE_PRIOR = \{.*?\n\};",
        "CHANNEL_MARGIN": r"const CHANNEL_MARGIN = \{.*?\n\};",
        "CHANNEL_MARGIN_PRIOR": r"const CHANNEL_MARGIN_PRIOR = \{.*?\n\};",
    }
    for key, new_block in js_blocks.items():
        pattern = patterns[key]
        if not re.search(pattern, text, re.S):
            raise RuntimeError(f"대시보드에서 '{key}' 블록을 찾지 못했습니다 - 파일 구조가 바뀐 것 같습니다. 자동 반영을 중단합니다.")
        text = re.sub(pattern, lambda m: new_block, text, count=1, flags=re.S)
    return text


def compute_all(sales_path: str, schedule_path: str, dashboard_path: str,
                 year: int | None = None, asof: str | None = None) -> dict:
    """엑셀 2개 → EVENTS/채널 총계/취소진단까지 한번에 계산. CLI(main)와
    watch_and_update.py(자동 감시)가 같이 쓰는 핵심 로직."""
    sales = read_excel_validated(sales_path, SALES_REQUIRED_COLS, "판매현황")
    sales["판매일자"] = pd.to_datetime(sales["판매일자"])
    schedule = read_excel_validated(schedule_path, SCHEDULE_REQUIRED_COLS, "행사일정")
    schedule["시작일"] = pd.to_datetime(schedule["시작일"])
    schedule["종료일"] = pd.to_datetime(schedule["종료일"])

    store_channel = load_store_channel_map(dashboard_path)
    current_year = year or int(schedule["종료일"].dt.year.max())
    asof_ts = pd.to_datetime(asof) if asof else pd.Timestamp(dt.date.today())

    all_events = build_event_records(schedule, sales, store_channel)
    current = [e for e in all_events if e.end.year == current_year]
    prior = [e for e in all_events if e.end.year == current_year - 1]
    baseline = match_yoy_baseline(current, prior)
    unmatched_yoy = sum(1 for i, e in enumerate(current) if e.channel in YOY_CHANNELS and i not in baseline)

    event_dicts = to_event_dicts(current, baseline, asof_ts, sales, store_channel)
    totals = compute_channel_totals(sales)
    cancels = cancel_diagnostics(sales)

    js_blocks = {
        "EVENTS": emit_events_js(event_dicts),
        "CHANNEL_TOTAL_REVENUE": emit_channel_totals_js(totals, current_year, "CHANNEL_TOTAL_REVENUE", "revenue"),
        "CHANNEL_TOTAL_REVENUE_PRIOR": emit_channel_totals_js(totals, current_year - 1, "CHANNEL_TOTAL_REVENUE_PRIOR", "revenue"),
        "CHANNEL_MARGIN": emit_channel_totals_js(totals, current_year, "CHANNEL_MARGIN", "margin"),
        "CHANNEL_MARGIN_PRIOR": emit_channel_totals_js(totals, current_year - 1, "CHANNEL_MARGIN_PRIOR", "margin"),
    }
    return {
        "event_dicts": event_dicts,
        "totals": totals,
        "cancels": cancels,
        "unmatched_yoy": unmatched_yoy,
        "current_year": current_year,
        "asof": asof_ts,
        "js_blocks": js_blocks,
    }


def build_report(events: list[dict], totals: dict, current_year: int, cancels: pd.DataFrame,
                  unmatched_yoy: int) -> str:
    settled = sum(1 for e in events if e.get("profitActual") is not None)
    pending = len(events) - settled
    lines = [
        "# 데이터 갱신 리포트",
        "",
        f"- 생성된 행사 수: {len(events)}건 (정산완료 {settled}건 / 진행중·예정 {pending}건)",
        f"- 전년동기 매칭 실패(백화점/라운지/밀로티인데 짝을 못 찾음): {unmatched_yoy}건 → revenuePlan/profitPlan이 null로 나갑니다.",
        f"- {current_year}년 채널×월 집계 채널 수: {len(totals.get(current_year, {}))}",
        f"- {current_year-1}년(전년) 채널×월 집계 채널 수: {len(totals.get(current_year-1, {}))}",
        f"- 취소 의심 건(주문번호+품목명 그룹 순수량<=0): {len(cancels)}건 (data_update_cancels.csv 참고)",
        "",
        "## 다음 단계 (이 스크립트가 자동화하지 않는 부분)",
        "- 온라인/라운지/밀로티 계획(사전시뮬) 수치는 여전히 수기 입력이 필요합니다.",
        "- 품목별 성과(PRODUCT_FULL_BY_STORE 등)는 이번 버전에서 다루지 않습니다.",
        "- 아래 .js 스니펫은 참고용입니다 — 대시보드에 붙여넣기 전에 반드시 눈으로 검토하세요.",
    ]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sales", required=True, help="품목별판매현황 엑셀 (여러 연도 데이터 포함 가능)")
    ap.add_argument("--schedule", required=True, help="행사 일정 엑셀 (매장·시작일·종료일, 여러 연도 누적 가능)")
    ap.add_argument("--dashboard", default="event_pnl_dashboard.html", help="STORE_LISTS를 읽어올 대시보드 HTML 경로")
    ap.add_argument("--year", type=int, default=None, help="이번해로 취급할 연도 (기본: 행사일정 중 최댓값)")
    ap.add_argument("--asof", default=None, help="정산 기준일 YYYY-MM-DD (기본: 오늘)")
    ap.add_argument("--out", default="data_update_output.js", help="생성된 JS 스니펫 출력 경로")
    ap.add_argument("--report", default="data_update_report.md", help="요약 리포트 출력 경로")
    ap.add_argument("--cancels-csv", default="data_update_cancels.csv", help="취소 의심 건 리포트 출력 경로")
    args = ap.parse_args()

    result = compute_all(args.sales, args.schedule, args.dashboard, args.year, args.asof)
    event_dicts, totals = result["event_dicts"], result["totals"]
    cancels, unmatched_yoy, current_year = result["cancels"], result["unmatched_yoy"], result["current_year"]

    js_chunks = [
        result["js_blocks"]["EVENTS"], "",
        result["js_blocks"]["CHANNEL_TOTAL_REVENUE"], "",
        result["js_blocks"]["CHANNEL_TOTAL_REVENUE_PRIOR"], "",
        result["js_blocks"]["CHANNEL_MARGIN"], "",
        result["js_blocks"]["CHANNEL_MARGIN_PRIOR"],
    ]
    Path(args.out).write_text("\n".join(js_chunks) + "\n", encoding="utf-8")
    Path(args.report).write_text(
        build_report(event_dicts, totals, current_year, cancels, unmatched_yoy), encoding="utf-8"
    )
    cancels.to_csv(args.cancels_csv, index=False, encoding="utf-8-sig")

    print(f"완료: {args.out}, {args.report}, {args.cancels_csv} 생성됨")
    print(f"행사 {len(event_dicts)}건, 전년동기 매칭 실패 {unmatched_yoy}건, 취소 의심 {len(cancels)}건")


if __name__ == "__main__":
    main()
