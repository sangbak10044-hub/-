#!/usr/bin/env python3
"""
사람이 엑셀을 "다운로드"까지만 해두면, 그 다음(계산 → 대시보드 반영 → git 푸시)을
자동으로 처리하는 스크립트. Windows 작업 스케줄러가 이 스크립트를 1시간마다 한 번씩
실행해주는 걸 전제로 만들어짐 (등록 방법은 scripts/register_task.bat 참고).

동작 순서 (한 번 실행될 때마다):
  1. watch_config.json에 적어둔 폴더(보통 다운로드 폴더)에서 판매현황/행사일정 엑셀 중
     "가장 최근에 수정된" 파일을 하나씩 찾음.
  2. 지난번에 이미 처리했던 파일과 같으면(경로+수정시각+크기 동일) 아무것도 안 하고 종료.
  3. update_dashboard_data.compute_all()로 계산.
  4. sanity_check()로 "채널 매출이 갑자기 0되거나 반토막" 같은 이상치가 없는지 확인하고,
     생성된 JS를 실제로 대시보드 파일에 끼워넣은 뒤 node --check로 문법도 검증.
  5. 문제 없으면: 대시보드 파일 덮어쓰기(교체 전 백업 남김) + git commit/push까지 자동.
     문제 있으면: 대시보드는 그대로 두고, "검토가 필요합니다" 리포트만 남김
     (README에 적힌 과거 버그들이 전부 "조용히 잘못된 숫자가 반영"되는 유형이라, 자동화한다고
     이 안전장치까지 건너뛰면 안 된다고 판단함).

사용법(수동 1회 실행 - 등록 전 테스트용):
  python scripts\\watch_and_update.py --config scripts\\watch_config.json
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from update_dashboard_data import compute_all, sanity_check, splice_dashboard  # noqa: E402

DEFAULT_CONFIG = {
    "watch_folder": "C:\\Users\\<사용자이름>\\Downloads",
    "repo_dir": "C:\\Users\\<사용자이름>\\Documents\\다우닝대시보드",
    "dashboard_filename": "event_pnl_dashboard.html",
    "sales_glob": "*품목별판매현황*.xlsx",
    "schedule_glob": "*행사일정*.xlsx",
    "year": None,
    "auto_push": True,
}


def log(repo_dir: Path, message: str):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}"
    print(line)
    with open(repo_dir / "scripts" / "watch_log.txt", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def latest_matching_file(folder: Path, pattern: str) -> Path | None:
    matches = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0] if matches else None


def file_signature(path: Path) -> dict:
    stat = path.stat()
    return {"path": str(path), "mtime": stat.st_mtime, "size": stat.st_size}


def load_state(state_path: Path) -> dict:
    if state_path.exists():
        return json.loads(state_path.read_text(encoding="utf-8"))
    return {}


def save_state(state_path: Path, state: dict):
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_js_syntax(html_text: str) -> str | None:
    """<script>...</script> 안쪽만 뽑아서 node --check로 문법 검증. 문제 있으면 에러 메시지 반환."""
    import re
    import tempfile

    m = re.search(r"<script>(.*)</script>", html_text, re.S)
    if not m:
        return "대시보드에서 <script> 블록을 찾지 못했습니다."
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False, encoding="utf-8") as f:
        f.write(m.group(1))
        tmp_path = f.name
    try:
        result = subprocess.run(["node", "--check", tmp_path], capture_output=True, text=True, timeout=30)
        if result.returncode != 0:
            return result.stderr.strip()
        return None
    except FileNotFoundError:
        return None  # Node.js가 없으면 문법 검증만 건너뜀(차단하지 않음) - 계산 자체는 pandas만으로 끝난 상태
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def git_commit_push(repo_dir: Path, message: str) -> str | None:
    """성공하면 None, 실패하면 에러 메시지 반환."""
    try:
        subprocess.run(["git", "add", "event_pnl_dashboard.html"], cwd=repo_dir, check=True, capture_output=True, text=True)
        commit = subprocess.run(["git", "commit", "-m", message], cwd=repo_dir, capture_output=True, text=True)
        if commit.returncode != 0 and "nothing to commit" not in commit.stdout:
            return f"git commit 실패: {commit.stdout}\n{commit.stderr}"
        push = subprocess.run(["git", "push"], cwd=repo_dir, capture_output=True, text=True)
        if push.returncode != 0:
            return f"git push 실패: {push.stdout}\n{push.stderr}"
        return None
    except subprocess.CalledProcessError as e:
        return f"git 명령 실패: {e.stdout}\n{e.stderr}"


def run_once(config_path: Path):
    config = {**DEFAULT_CONFIG, **json.loads(config_path.read_text(encoding="utf-8"))}
    repo_dir = Path(config["repo_dir"])
    watch_folder = Path(config["watch_folder"])
    dashboard_path = repo_dir / config["dashboard_filename"]
    state_path = repo_dir / "scripts" / ".watch_state.json"
    (repo_dir / "scripts").mkdir(exist_ok=True)

    sales_file = latest_matching_file(watch_folder, config["sales_glob"])
    schedule_file = latest_matching_file(watch_folder, config["schedule_glob"])
    if not sales_file or not schedule_file:
        log(repo_dir, f"대상 파일을 찾지 못함 (판매현황: {sales_file}, 행사일정: {schedule_file}) - 건너뜀")
        return

    state = load_state(state_path)
    new_sig = {"sales": file_signature(sales_file), "schedule": file_signature(schedule_file)}
    if state.get("sales") == new_sig["sales"] and state.get("schedule") == new_sig["schedule"]:
        log(repo_dir, "이전에 처리한 파일과 동일함 - 변경 없음")
        return

    log(repo_dir, f"새 파일 감지: {sales_file.name}, {schedule_file.name} - 계산 시작")
    try:
        result = compute_all(str(sales_file), str(schedule_file), str(dashboard_path), config.get("year"))
    except SystemExit as e:
        log(repo_dir, f"계산 실패(입력 파일 확인 필요): {e}")
        return
    except Exception as e:
        log(repo_dir, f"계산 중 예외 발생: {e!r}")
        return

    warnings = sanity_check(str(dashboard_path), result["totals"], result["current_year"])
    dashboard_text = dashboard_path.read_text(encoding="utf-8")
    try:
        new_text = splice_dashboard(dashboard_text, result["js_blocks"])
    except RuntimeError as e:
        log(repo_dir, f"대시보드 반영 실패: {e}")
        return
    syntax_error = validate_js_syntax(new_text)

    review_dir = repo_dir / "scripts" / "검토필요"
    if warnings or syntax_error:
        review_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        review_path = review_dir / f"event_pnl_dashboard_{stamp}.html"
        review_path.write_text(new_text, encoding="utf-8")
        reasons = warnings + ([syntax_error] if syntax_error else [])
        log(repo_dir, "자동 반영 보류 - 아래 이유로 사람 확인이 필요합니다:")
        for r in reasons:
            log(repo_dir, f"  - {r}")
        log(repo_dir, f"  계산 결과는 검토용으로 저장해뒀습니다: {review_path}")
        return  # 상태 저장도 안 함 - 같은 파일이 다음 실행 때 다시 시도되도록

    backup_dir = repo_dir / "scripts" / "백업"
    backup_dir.mkdir(exist_ok=True)
    shutil.copy2(dashboard_path, backup_dir / f"event_pnl_dashboard_{datetime.now():%Y%m%d_%H%M%S}.html")
    dashboard_path.write_text(new_text, encoding="utf-8")
    log(repo_dir, f"대시보드 반영 완료: 행사 {len(result['event_dicts'])}건, 전년동기 매칭 실패 {result['unmatched_yoy']}건")

    save_state(state_path, new_sig)

    if config.get("auto_push"):
        err = git_commit_push(repo_dir, f"자동 데이터 갱신 ({sales_file.name}, {schedule_file.name})")
        if err:
            log(repo_dir, f"git 반영 실패(로컬 파일은 갱신됨, 수동으로 push 필요): {err}")
        else:
            log(repo_dir, "GitHub에 자동 push 완료")
    else:
        log(repo_dir, "auto_push=false 설정이라 로컬 파일만 갱신함 - 직접 git push 해주세요")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=str(Path(__file__).parent / "watch_config.json"))
    args = ap.parse_args()
    run_once(Path(args.config))


if __name__ == "__main__":
    main()
