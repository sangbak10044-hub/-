"""watch_and_update.py 자체 검증 스크립트.
실제 git 저장소는 건드리지 않고, 임시 폴더에 대시보드 사본 + 합성 엑셀을 두고
run_once()를 돌려서:
  - 정상 케이스: 대시보드가 실제로 갱신되고, 백업이 남고, 상태파일이 기록되는지
  - 같은 파일 재실행: "변경 없음"으로 스킵하고 아무것도 안 건드리는지
  - 이상치 케이스: 채널 매출이 0으로 사라지면 자동 반영 안 하고 검토용 파일만 남기는지
를 assert로 확인한다. 실행: python3 scripts/test_watch_and_update.py
"""
import json
import shutil
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import watch_and_update as wu  # noqa: E402

HERE = Path(__file__).parent
TMP = HERE / "_test_watch_tmp"


# update_dashboard_data.py가 실제로 다루는 두 가지 구조적 특징(EVENTS 뒤에 .map() 보정이
# 붙는 것, STORE_LISTS로 매장→채널을 읽는 것)만 갖춘 축소판 대시보드. 진짜 event_pnl_dashboard.html을
# 그대로 베이스라인으로 쓰면 거기 박힌 큰 실적 숫자들 때문에, 작은 합성 픽스처를 넣을 때마다
# "채널이 통째로 사라짐" sanity check가 항상 걸려서(정상 케이스 테스트가 안 됨) 별도로 만든다.
MINI_DASHBOARD = """<!DOCTYPE html>
<html><head><title>t</title></head><body>
<script>
const OWN_MALL = "온라인몰";
const EXTERNAL_MALLS = ["롯데홈쇼핑"];
const ALL_MALLS = [OWN_MALL, ...EXTERNAL_MALLS];
const STORE_LISTS = {
  "온라인": ALL_MALLS,
  "백화점": ["롯데노원"],
};
const MARGIN_ASSUMPTION = 0.15;
let EVENTS = [
  {name:"작년 기준행사", month:"2025-07", channel:"백화점", mall:"롯데노원", startDate:"2025-07-08", endDate:"2025-07-18", profitPlan:null, profitActual:0.04, revenuePlan:null, revenueActual:0.1},
].map(e => ({
  ...e,
  revenuePlan: e.revenuePlan !== undefined ? e.revenuePlan : +(e.profitPlan / MARGIN_ASSUMPTION).toFixed(4),
  revenueActual: (e.revenueActual !== undefined) ? e.revenueActual :
                 (e.profitActual === null || e.profitActual === undefined) ? null :
                 +(e.profitActual / MARGIN_ASSUMPTION).toFixed(4),
}));
const CHANNEL_TOTAL_REVENUE = {
  "백화점": {"2026-07":0.13, "전체":0.13},
};
const CHANNEL_TOTAL_REVENUE_PRIOR = {
  "백화점": {"2025-07":0.1, "전체":0.1},
};
const CHANNEL_MARGIN = {
  "백화점": {"2026-07":46.15, "전체":46.15},
};
const CHANNEL_MARGIN_PRIOR = {
  "백화점": {"2025-07":40.0, "전체":40.0},
};
</script>
</body></html>
"""


def setup_repo():
    if TMP.exists():
        shutil.rmtree(TMP)
    repo_dir = TMP / "repo"
    downloads = TMP / "downloads"
    repo_dir.mkdir(parents=True)
    downloads.mkdir(parents=True)
    (repo_dir / "event_pnl_dashboard.html").write_text(MINI_DASHBOARD, encoding="utf-8")
    return repo_dir, downloads


def write_fixtures(downloads: Path, tag: str, revenue: int):
    sales_rows = [
        dict(그룹명="소파", 품목코드="A1", 품목명="빈체로 소파", 주문번호="ORD1", 매입처="X", 판매점="롯데노원",
             고객명="홍길동", 판매일자="2026-07-15", 수량=1, 매출액=revenue, 실매출액=revenue,
             입고가=revenue // 2, 이익금=revenue // 3, 이익율=33.0, 채널="백화점"),
    ]
    pd.DataFrame(sales_rows).to_excel(downloads / f"품목별판매현황_{tag}.xlsx", index=False)
    schedule_rows = [
        dict(매장="롯데노원", 시작일="2026-07-10", 종료일="2026-07-19"),
    ]
    pd.DataFrame(schedule_rows).to_excel(downloads / f"행사일정_{tag}.xlsx", index=False)


def write_config(repo_dir: Path, downloads: Path) -> Path:
    config = {
        "watch_folder": str(downloads),
        "repo_dir": str(repo_dir),
        "dashboard_filename": "event_pnl_dashboard.html",
        "sales_glob": "*품목별판매현황*.xlsx",
        "schedule_glob": "*행사일정*.xlsx",
        "year": 2026,
        "auto_push": False,
    }
    config_path = TMP / "watch_config.json"
    config_path.write_text(json.dumps(config), encoding="utf-8")
    return config_path


def main():
    repo_dir, downloads = setup_repo()
    write_fixtures(downloads, "v1", 13_000_000)
    config_path = write_config(repo_dir, downloads)

    dashboard_path = repo_dir / "event_pnl_dashboard.html"
    original_text = dashboard_path.read_text(encoding="utf-8")

    # 1) 정상 케이스: 처음 실행하면 반영되어야 함
    wu.run_once(config_path)
    updated_text = dashboard_path.read_text(encoding="utf-8")
    assert updated_text != original_text, "대시보드가 갱신되지 않음"
    assert "롯데노원 행사(07-10~07-19)" in updated_text, "새 행사가 EVENTS에 안 들어감"
    assert (repo_dir / "scripts" / "백업").exists(), "백업 폴더가 안 만들어짐"
    assert (repo_dir / "scripts" / ".watch_state.json").exists(), "상태 파일이 안 만들어짐"
    print("[OK] 정상 케이스: 새 파일 감지 -> 대시보드 자동 갱신 + 백업 + 상태 저장")

    # 2) 같은 파일로 재실행 -> 스킵돼야 함 (내용 변화 없음, 백업 추가 안 됨)
    backups_before = list((repo_dir / "scripts" / "백업").iterdir())
    wu.run_once(config_path)
    backups_after = list((repo_dir / "scripts" / "백업").iterdir())
    assert len(backups_before) == len(backups_after), "변경 없는데 또 백업이 생김 (중복 처리)"
    print("[OK] 같은 파일 재실행 -> 스킵(변경 없음)")

    # 3) 이상치 케이스: 새 파일이 기존 대비 채널 매출을 극단적으로 반토막 냄 -> 자동 반영 보류
    write_fixtures(downloads, "v2", 1_000)  # 13,000,000원 -> 1,000원으로 사실상 0에 가깝게 급감
    text_before_v2 = dashboard_path.read_text(encoding="utf-8")
    wu.run_once(config_path)
    text_after_v2 = dashboard_path.read_text(encoding="utf-8")
    assert text_before_v2 == text_after_v2, "이상치인데도 대시보드가 그냥 덮어써짐 (안전장치 미작동)"
    review_dir = repo_dir / "scripts" / "검토필요"
    assert review_dir.exists() and list(review_dir.iterdir()), "이상치를 검토용 파일로 안 남김"
    print("[OK] 이상치(매출 급감) 케이스 -> 자동 반영 보류하고 검토용 파일만 남김")

    print("\n모든 watch_and_update 자체 검증 통과.")
    shutil.rmtree(TMP)


if __name__ == "__main__":
    main()
