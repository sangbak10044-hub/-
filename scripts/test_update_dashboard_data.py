"""update_dashboard_data.py에 대한 자체 검증 스크립트.
합성 표본 엑셀(2025/2026, 백화점 1개 매장 + 온라인 1개 매장)을 만들어 스크립트를 돌려보고,
- STORE_LISTS 파싱이 실제 대시보드 파일에서 매장을 찾아내는지
- 채널×월 집계 숫자가 손계산과 맞는지
- 매장별 1:1 계절 매칭이 올바른 전년 행사를 골랐는지
- 진행중 행사의 partial* 필드가 올바르게 계산되는지
- 생성된 JS 스니펫이 문법적으로 유효한지 (node --check)
를 assert로 확인한다. 실행: python3 scripts/test_update_dashboard_data.py
"""
import subprocess
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import update_dashboard_data as udd

HERE = Path(__file__).parent
TMP = HERE / "_test_tmp"


def make_fixtures():
    TMP.mkdir(exist_ok=True)
    sales_rows = [
        # 롯데노원(백화점) 2025년 행사 - 계절적으로 이번해 7/10~7/19와 가장 가까움
        dict(그룹명="소파", 품목코드="A1", 품목명="빈체로 소파", 주문번호="ORD1", 매입처="X", 판매점="롯데노원",
             고객명="홍길동", 판매일자="2025-07-12", 수량=1, 매출액=10_000_000, 실매출액=10_000_000,
             입고가=6_000_000, 이익금=4_000_000, 이익율=40.0, 채널="백화점"),
        # 롯데노원 2025년 행사 - 계절적으로 더 먼 것(비교용, 매칭되면 안 됨)
        dict(그룹명="소파", 품목코드="A2", 품목명="빈체로 소파", 주문번호="ORD2", 매입처="X", 판매점="롯데노원",
             고객명="김철수", 판매일자="2025-02-01", 수량=1, 매출액=5_000_000, 실매출액=5_000_000,
             입고가=3_000_000, 이익금=2_000_000, 이익율=40.0, 채널="백화점"),
        # 롯데노원 2026년 이번해 정산완료 행사(7/10~7/19)
        dict(그룹명="소파", 품목코드="A3", 품목명="빈체로 소파", 주문번호="ORD3", 매입처="X", 판매점="롯데노원",
             고객명="박영희", 판매일자="2026-07-15", 수량=1, 매출액=13_000_000, 실매출액=13_000_000,
             입고가=7_000_000, 이익금=6_000_000, 이익율=46.0, 채널="백화점"),
        # 취소 건: 정정 접미사(-2)로 원 주문 상쇄 -> 순수량 0
        dict(그룹명="소파", 품목코드="A4", 품목명="테스트소파", 주문번호="ORD9", 매입처="X", 판매점="롯데노원",
             고객명="최영수", 판매일자="2026-07-16", 수량=1, 매출액=1_000_000, 실매출액=1_000_000,
             입고가=500_000, 이익금=500_000, 이익율=50.0, 채널="백화점"),
        dict(그룹명="소파", 품목코드="A4", 품목명="테스트소파", 주문번호="ORD9-2", 매입처="X", 판매점="롯데노원",
             고객명="최영수", 판매일자="2026-07-17", 수량=-1, 매출액=-1_000_000, 실매출액=-1_000_000,
             입고가=-500_000, 이익금=-500_000, 이익율=50.0, 채널="백화점"),
        # 온라인몰 - 진행중 행사(9/1~9/13) 중 일부만 판매 발생 (asof=9/4 기준)
        dict(그룹명="소파", 품목코드="B1", 품목명="루소 소파", 주문번호="ORD5", 매입처="Y", 판매점="온라인몰",
             고객명="이몽룡", 판매일자="2026-09-02", 수량=1, 매출액=2_000_000, 실매출액=2_000_000,
             입고가=1_200_000, 이익금=800_000, 이익율=40.0, 채널="온라인"),
    ]
    pd.DataFrame(sales_rows).to_excel(TMP / "sales.xlsx", index=False)

    schedule_rows = [
        dict(매장="롯데노원", 시작일="2025-07-08", 종료일="2025-07-18"),
        dict(매장="롯데노원", 시작일="2025-01-25", 종료일="2025-02-05"),
        dict(매장="롯데노원", 시작일="2026-07-10", 종료일="2026-07-19"),
        dict(매장="온라인몰", 시작일="2026-09-01", 종료일="2026-09-13"),
    ]
    pd.DataFrame(schedule_rows).to_excel(TMP / "schedule.xlsx", index=False)


def run_script():
    out = TMP / "output.js"
    report = TMP / "report.md"
    cancels = TMP / "cancels.csv"
    dashboard = HERE.parent / "event_pnl_dashboard.html"
    cmd = [
        sys.executable, str(HERE / "update_dashboard_data.py"),
        "--sales", str(TMP / "sales.xlsx"),
        "--schedule", str(TMP / "schedule.xlsx"),
        "--dashboard", str(dashboard),
        "--year", "2026", "--asof", "2026-09-04",
        "--out", str(out), "--report", str(report), "--cancels-csv", str(cancels),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        raise SystemExit("스크립트 실행 실패")
    return out.read_text(encoding="utf-8"), report.read_text(encoding="utf-8"), cancels


def main():
    make_fixtures()

    # 1) STORE_LISTS 파싱 단독 테스트
    store_map = udd.load_store_channel_map(str(HERE.parent / "event_pnl_dashboard.html"))
    assert store_map.get("롯데노원") == "백화점", store_map.get("롯데노원")
    assert store_map.get("온라인몰") == "온라인", store_map.get("온라인몰")
    print("[OK] STORE_LISTS 파싱")

    js_text, report_text, cancels_path = run_script()

    # 2) JS 문법 검증
    js_path = TMP / "output_for_check.js"
    js_path.write_text(js_text, encoding="utf-8")
    check = subprocess.run(["node", "--check", str(js_path)], capture_output=True, text=True)
    assert check.returncode == 0, check.stderr
    print("[OK] 생성된 JS 문법 유효")

    # 3) 채널 총계 숫자 검증 (백화점 2026-07: 1300만+100만-100만(취소 상쇄) = 1300만원 = 0.13억)
    assert '"2026-07": 0.13' in js_text, js_text
    print("[OK] 채널×월 매출 집계 값 (취소 상쇄 포함)")

    # 4) 매장별 1:1 계절 매칭: 2026-07-10~19 행사가 2025-07-08~18(가까운 쪽)과 매칭되어야 함 (2025-01-25~02-05가 아니라)
    assert 'revenuePlan: 0.1,' in js_text, "1:1 매칭이 계절적으로 가까운 행사를 고르지 못함(0.1억=7월 행사가 아니라 다른 후보가 선택됨):\n" + js_text
    print("[OK] 매장별 1:1 계절 매칭 (가까운 전년 행사를 선택)")

    # 5) 진행중 행사(온라인몰 9/1~13, asof 9/4)의 partial 필드 검증
    assert "partialRevenue" in js_text and "partialDays" in js_text
    assert "partialDays: 4" in js_text, js_text  # 9/1~9/4 = 4일째
    assert "totalDays: 13" in js_text, js_text
    print("[OK] 진행중 행사 partial* 필드")

    # 6) 취소 진단: ORD9/ORD9-2가 순수량 0이라 취소로 잡혀야 함
    cancels_df = pd.read_csv(cancels_path)
    assert (cancels_df["_base_order"] == "ORD9").any(), cancels_df
    print("[OK] 취소 건 진단(주문번호 정정접미사 제거 + 품목명 그룹)")

    print("\n모든 자체 검증 통과.")


if __name__ == "__main__":
    main()
