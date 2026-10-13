# -*- coding: utf-8 -*-
"""
지점 KPI 올리기 — 통계 대시보드의 전 지점 KPI 화면용.

partsbay.py 가 10분마다 만든 대시보드 원본(data)을 구역별로 Supabase 에 올린다.
  · kpi_snapshot(branch, section, data)  지점 × 구역 현재 상태 (덮어쓰기)
  · kpi_history(day, branch, section)    설정에서 고른 구역만 하루 한 번
  · kpi_config                          어떤 구역을 올릴지(upload_sections) / 기록할 구역(history_sections)
  · 구역 "_meta"                         매 주기 올리는 생존 신호 (대시보드의 "갱신 멈춤" 판단용)

로그인·설정은 재고조사 앱 모듈(stockapp.py)의 것을 그대로 쓴다 (같은 데몬 계정, daemon/config.local.json).
partsbay.py 가 매 주기 이 파일을 다시 읽으므로 고쳐도 데몬 재시작이 필요 없다.

개인정보 방어 (고정, 설정으로 풀 수 없음):
  · 고객명·차대번호 등 키(DENY_KEYS)는 값을 지우고 올린다
  · 어떤 키든 값이 차대번호(17자)·차량번호 모양이면 가린다
"""
import hashlib
import json
import re
import sys
import time
import types
from datetime import datetime, timedelta, timezone

# 설정을 못 읽을 때 쓸 기본 구역 (reasons·dday·meta 는 올리지 않음)
DEFAULT_SECTIONS = [
    "recv", "inv", "oaov", "ext", "shop", "acc", "tire", "longstock", "opart", "oavail", "oflow",
    "openro", "noshow", "sbresv", "sbcar", "nonmng", "stockcheck", "stockcheck_week", "noloc",
    "extaudit", "tires",
]

# 값을 지울 키 (소문자 비교). 설정의 deny_keys 로 더할 수는 있어도 뺄 수는 없다.
DENY_KEYS = {"cust", "custnm", "cust_nm", "customer", "vin", "vinno", "vin_no", "plate", "carno", "car_no",
             "고객", "고객명", "차량번호", "차대번호"}
MASK = "[삭제]"
VIN_RE = re.compile(r"\b[A-HJ-NPR-Z0-9]{17}\b")
PLATE_RE = re.compile(r"\d{2,3}\s?[가-힣]\s?\d{4}")

KST = timezone(timedelta(hours=9))
MAX_SECONDS = 45   # 이 시간이 넘으면 남은 구역은 다음 주기로 미룬다 (Supabase 가 느려도 DMS 갱신 주기를 붙잡지 않게)

_S = sys.modules.get("_branchsync_state")
if _S is None:
    _S = types.ModuleType("_branchsync_state")
    _S.sent = {}          # 구역 -> 마지막으로 올린 내용의 해시 (바뀐 것만 올리기)
    _S.hist_day = None    # 일별 기록을 남긴 날
    _S.hist_sections = []
    sys.modules["_branchsync_state"] = _S


def _log(msg: str) -> None:
    print(f"   [KPI 올리기] {msg}")


def scrub(obj, deny=DENY_KEYS, stats=None):
    """개인정보 키는 값을 지우고, 차대번호·차량번호 모양 문자열은 가린다. 원본은 바꾸지 않고 복사본을 돌려준다."""
    if stats is None:
        stats = {"keys": 0, "values": 0}
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            if str(k).lower() in deny:
                out[k] = "" if isinstance(v, str) else None
                if v not in (None, "", "-"):
                    stats["keys"] += 1
            else:
                out[k] = scrub(v, deny, stats)
        return out
    if isinstance(obj, list):
        return [scrub(v, deny, stats) for v in obj]
    if isinstance(obj, str) and (VIN_RE.search(obj) or PLATE_RE.search(obj)):
        stats["values"] += 1
        return VIN_RE.sub(MASK, PLATE_RE.sub(MASK, obj))
    return obj


def _hash(v) -> str:
    return hashlib.sha1(json.dumps(v, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()


def _read_config(sa) -> dict:
    """kpi_config 읽기. 실패하면 기본값 (업로드는 계속)."""
    cfg = {"upload_sections": DEFAULT_SECTIONS, "history_sections": [], "deny_keys": []}
    try:
        for r in sa._rest("GET", "kpi_config?select=key,value", prefer="return=representation") or []:
            if r["key"] in cfg and isinstance(r["value"], list):
                cfg[r["key"]] = r["value"]
    except Exception as e:
        _log(f"설정을 못 읽어 기본값 사용: {e}")
    return cfg


def run(data: dict, sa, now: datetime | None = None) -> dict:
    """data: partsbay 의 대시보드 원본(dict), sa: 불러온 stockapp 모듈. 올린 구역 수 등을 돌려준다."""
    if not sa._ensure_token():
        _log("Supabase 로그인이 안 돼 건너뜀")
        return {"ok": False}
    branch = sa._cfg()["branch"]
    now = now or datetime.now(timezone.utc)
    cfg = _read_config(sa)
    deny = DENY_KEYS | {str(k).lower() for k in cfg["deny_keys"]}
    stats = {"keys": 0, "values": 0}
    gen = now.astimezone(timezone.utc).isoformat()

    t0 = time.time()
    failed = False
    sent = changed = 0
    for sec in cfg["upload_sections"]:
        if sec not in data:
            continue
        if failed or time.time() - t0 > MAX_SECONDS:   # 한 번 실패했거나 시간이 넘으면 나머지는 다음 주기에
            break
        payload = scrub(data[sec], deny, stats)
        h = _hash(payload)
        if _S.sent.get(sec) == h:
            continue
        try:
            sa._rest("POST", "kpi_snapshot?on_conflict=branch,section",
                     {"branch": branch, "section": sec, "data": payload, "data_hash": h, "generated_at": gen},
                     "resolution=merge-duplicates,return=minimal")
            _S.sent[sec] = h
            changed += 1
        except Exception as e:
            _log(f"{sec} 올리기 실패 — 나머지는 다음 주기에: {e}")
            failed = True
        sent += 1

    # 생존 신호: 바뀐 구역이 없어도 매 주기 갱신 → 대시보드가 "마지막 갱신"을 판단
    meta = data.get("meta") or {}
    live = {"branch_name": meta.get("branch_name"), "brch_code": meta.get("brch_code"),
            "generated_at_text": meta.get("generated_at"), "sections": cfg["upload_sections"]}
    try:
        if failed:
            raise RuntimeError("앞 단계 실패로 생략")
        sa._rest("POST", "kpi_snapshot?on_conflict=branch,section",
                 {"branch": branch, "section": "_meta", "data": live, "data_hash": _hash(live), "generated_at": gen},
                 "resolution=merge-duplicates,return=minimal")
    except Exception as e:
        _log(f"생존 신호 올리기 실패: {e}")

    # 일별 기록 (설정에서 고른 구역, 하루 한 번)
    day = now.astimezone(KST).strftime("%Y-%m-%d")
    hist = [s for s in cfg["history_sections"] if s in data]
    if hist and not failed and time.time() - t0 <= MAX_SECONDS and (_S.hist_day != day or _S.hist_sections != hist):
        try:
            rows = [{"day": day, "branch": branch, "section": s, "data": scrub(data[s], deny, stats)} for s in hist]
            sa._rest("POST", "kpi_history?on_conflict=day,branch,section", rows,
                     "resolution=merge-duplicates,return=minimal")
            _S.hist_day, _S.hist_sections = day, hist
        except Exception as e:
            _log(f"일별 기록 실패: {e}")

    if stats["keys"] or stats["values"]:
        _log(f"개인정보 가림: 키 {stats['keys']}건, 값 {stats['values']}건")
    _log(f"{branch}: 구역 {len(cfg['upload_sections'])}개 중 바뀐 {changed}개 올림")
    return {"ok": True, "changed": changed, "scrubbed": stats}
