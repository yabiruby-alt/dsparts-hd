#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
동성모터스 PARTS - BMW My DMS 부품부서 대시보드 생성기

실행하면 (기본 = 상시 실행 모드):
  1. 브라우저가 최소화된 채로 한 번 뜨고 My DMS에 접속합니다. 화면엔 안 보입니다.
  2. 로그인이 안 되어 있으면 화면에 알림창이 뜹니다 - 작업표시줄에서 그 창을 열어 로그인(2FA 포함)해주세요.
  3. 로그인되면 그 브라우저를 계속 켜둔 채로 10분마다 자동으로 데이터를 뽑아 대시보드를 갱신하고
     깃허브에 push합니다. 프로그램을 끄지 않는 한 계속 이 상태로 돌아갑니다 (Ctrl+C로 종료).

`python partsbay.py --once` 로 실행하면 한 번만 갱신하고 바로 종료합니다 (테스트용).

세션(쿠키)은 authdata 폴더에 저장되어, 프로그램을 재시작해도 세션이 유지되는 동안은 재로그인이 필요 없습니다.
동시에 두 개를 띄우면 DMS가 세션을 끊어버리므로 절대 두 인스턴스를 같이 실행하지 마세요(잠금 파일로 방지됨).
"""
import json
import os
import re
import subprocess
import sys
import time
import webbrowser

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from jinja2 import Environment, FileSystemLoader
from playwright.sync_api import sync_playwright, Frame, Page

# ============================================================
# 설정 - 지점마다 이 블록만 바꿔서 배포
# ============================================================
BASE_URL = "https://bmwdms.co.kr"
DEALER_CD = "001672"
BIZ_AREA_CD = "051"
BRCH_CD = "15"
BRANCH_NAME = "AS_부산(해운중동)"

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
DOCS_DIR = BASE_DIR / "docs"  # GitHub Pages가 서빙하는 폴더
# 팀 캘린더: 이 이름의 이미지 파일을 docs 폴더에 넣어두면 "팀 캘린더" 메뉴에 그대로 표시됨.
CALENDAR_IMAGE_CANDIDATES = ["team-calendar.jpg", "team-calendar.jpeg", "team-calendar.png", "team-calendar.webp"]
GIT_REMOTE_URL = "https://github.com/yabiruby-alt/dsparts-hd.git"
PUBLISH_TO_GITHUB = True
# 크로미움이 한글 등 비-ASCII 경로의 --user-data-dir 에서 불안정하게 종료되는 문제가 있어
# 로그인 세션 저장 폴더는 항상 ASCII 전용 경로(사용자 홈 폴더 아래)에 둔다.
AUTH_DIR = Path.home() / "AppData" / "Local" / "PartsBayDMS" / "authdata"
TEMPLATE_NAME = "template.html.j2"

WEEKDAYS_KO = ["월", "화", "수", "목", "금", "토", "일"]

# ============================================================
# 사용자 알림 (창은 항상 최소화되어 있으므로, 사람 개입이 필요할 때만
# 화면에 알림창을 띄운다)
# ============================================================

def notify_user(title: str, message: str) -> None:
    safe_title = title.replace("'", "''")
    safe_msg = message.replace("'", "''")
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command",
             f"Add-Type -AssemblyName System.Windows.Forms; "
             f"[System.Windows.Forms.MessageBox]::Show('{safe_msg}', '{safe_title}', "
             f"[System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null"],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception as e:
        print(f"[알림 실패] {e}")


# ============================================================
# 브라우저 / 로그인
# ============================================================

def ensure_logged_in(page: Page, timeout_sec: int = 420) -> None:
    page.goto(f"{BASE_URL}/selectHome.do")
    deadline = time.time() + timeout_sec
    notified = False
    last_debug = 0.0
    while time.time() < deadline:
        try:
            # 좌측 메뉴(#icon-parts 등)는 iframe이 아니라 최상위 페이지 자체에 있음. 접혀 있어도
            # DOM에는 존재하므로 visible 상태를 요구하지 않고 attached 여부만 확인한다.
            page.wait_for_selector("#icon-parts", timeout=2500, state="attached")
            print(f"로그인 감지됨. 현재 URL: {page.url}")
            return
        except Exception:
            pass
        if not notified:
            print("DMS 로그인이 필요합니다. 브라우저 창에서 로그인(2FA 포함)해주세요...")
            notify_user(
                "동성모터스 PARTS — 로그인 필요",
                "DMS 세션이 만료되어 대시보드 자동 갱신이 멈춰 있습니다.\n"
                "작업표시줄에서 최소화된 브라우저 창을 열어 로그인(2FA 포함)해주세요.",
            )
            notified = True
        if time.time() - last_debug > 10:
            last_debug = time.time()
            try:
                print(f"[대기중] url={page.url}")
                debug_path = OUTPUT_DIR / "_debug_login.png"
                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(debug_path))
            except Exception as e:
                print(f"[디버그 스크린샷 실패] {e}")
        page.wait_for_timeout(1500)
    raise TimeoutError("로그인 대기 시간을 초과했습니다. 프로그램을 다시 실행해주세요.")


def click_menu(page: Page, section_id: str, label: str) -> None:
    page.evaluate(
        """({sectionId, label}) => {
            const a = [...document.querySelectorAll('#' + sectionId + ' a[data-url]')]
                .find(x => x.textContent.trim() === label);
            if (!a) throw new Error('메뉴를 찾지 못했습니다: ' + label);
            a.click();
        }""",
        {"sectionId": section_id, "label": label},
    )


def wait_for_frame(page: Page, url_substr: str, timeout_ms: int = 20000) -> Frame:
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        for f in page.frames:
            if url_substr in f.url:
                return f
        page.wait_for_timeout(300)
    raise TimeoutError(f"화면을 찾지 못했습니다: {url_substr}")


FETCH_JS = """
async ({url, body}) => {
  return await new Promise((resolve) => {
    jQuery.ajax({
      url, type: 'POST', dataType: 'json', contentType: 'application/json',
      data: JSON.stringify(body),
      success: (res) => resolve({ok: true, rows: res.data || [], total: res.total}),
      error: (xhr) => resolve({ok: false, status: xhr.status, text: (xhr.responseText || '').slice(0, 300)})
    });
  });
}
"""


def fetch_rows(frame: Frame, url: str, body: dict) -> list:
    result = frame.evaluate(FETCH_JS, {"url": url, "body": body})
    if not result.get("ok"):
        raise RuntimeError(f"요청 실패 {url}: {result}")
    return result["rows"]


# ============================================================
# 데이터 추출
# ============================================================

def extract_receiving(page: Page, today_str: str) -> list:
    click_menu(page, "icon-parts", "입고현황")
    frame = wait_for_frame(page, "selectWhStatusMain")
    body = {
        "recordCountPerPage": 5000, "pageIndex": 1,
        "sRealWhDtFrom": today_str, "sRealWhDtTo": today_str,
        "sBpNm": "", "sItemCd": "", "sWhNo": "", "sPurcOrderNo": "",
        "sAloisCd": "", "sWhTp": "", "sWhStat": "", "sGrnNoFrom": "", "sGrnNoTo": "",
    }
    return fetch_rows(frame, "/parts/whmng/selectWhStatusMain.do", body)


def extract_inventory(page: Page) -> list:
    click_menu(page, "icon-parts", "현재고리스트 조회")
    frame = wait_for_frame(page, "selectInventListMain")
    body = {
        "recordCountPerPage": 200000, "pageIndex": 1, "firstIndex": 0, "lastIndex": 200000,
        "sCorpCd": DEALER_CD, "sBizAreaCd": BIZ_AREA_CD, "sBrchCd": BRCH_CD,
        "sProdType": "", "sItemCd": "", "sItemNm": "", "sStrgeCd": "", "sCrtQtyYn": True,
    }
    return fetch_rows(frame, "/parts/inventory/selectInventoryList.do", body)


def extract_turnover(page: Page, month_start: str, today_str: str) -> list:
    click_menu(page, "icon-report", "Turn Over 리포트")
    frame = wait_for_frame(page, "selectDLRTurnOver")
    body = {
        "recordCountPerPage": 20000, "pageIndex": 1,
        "sSearchStartDt": f"{month_start}T00:00:00.000Z",
        "sSearchEndDt": f"{today_str}T23:59:59.000Z",
        "sBrands": [], "sSeriesList": [], "sCarNo": "", "sVinNo": "",
        "sCustTp": "", "sCustNo": "", "sCustNm": "",
        "sDlrCd": DEALER_CD, "sBrchCdList": [BRCH_CD], "sSaList": [], "sRoDocNo": "",
        "sCalcTpCds": ["C", "I"], "sInvcNum": "", "sInvStatCd": "", "sCalcNo": "",
        "sIctTradeYn": "", "sItemTpCd": "", "sItemCd": "", "sItemNm": "", "sProdType": "전체",
        "sAloiscd": "", "sRclCampnCdYn": "", "sCampnYn": "", "sCupnCdYn": "", "sSvcTypess": [],
    }
    return fetch_rows(frame, "/rpt/raw/selectDLRTurnOver.do", body)


# ============================================================
# 집계 (JS로 검증한 로직을 그대로 Python으로 옮김)
# ============================================================

GROUP_ORDER = ["A", "L", "O", "I", "S"]
GROUP_LABELS = {"A": "상시재고", "L": "로컬조달", "O": "특수/단종계열", "I": "비이동성", "S": "특수발주"}
PROD_LABELS = {"3": "Car Accessories", "5": "Rims and Complete Wheels", "7": "Accessories (Lifestyle)"}


def num(v) -> float:
    if v is None or v == "":
        return 0.0
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return 0.0


def group_alois(rows, val_fn, qty_fn=None):
    by_code = defaultdict(lambda: {"val": 0.0, "qty": 0.0})
    for r in rows:
        code = r.get("aloisCd") or "?"
        by_code[code]["val"] += val_fn(r)
        if qty_fn:
            by_code[code]["qty"] += qty_fn(r)
    total = sum(v["val"] for v in by_code.values()) or 1.0
    groups = []
    for g in GROUP_ORDER:
        codes = sorted(c for c in by_code if c and c[0] == g)
        if not codes:
            continue
        grows, sub_val, sub_qty = [], 0.0, 0.0
        for c in codes:
            v, q = by_code[c]["val"], by_code[c]["qty"]
            sub_val += v
            sub_qty += q
            row = {"code": c, "val": round(v), "pct": round(100 * v / total, 2)}
            if qty_fn:
                row["qty"] = round(q)
            grows.append(row)
        groups.append({
            "code": g, "label": GROUP_LABELS[g], "oseries": g == "O",
            "rows": grows, "subtotal_val": round(sub_val), "subtotal_qty": round(sub_qty),
            "subtotal_pct": round(100 * sub_val / total, 2),
        })
    return groups


def build_recv(rows, today_str):
    def val_fn(r):
        return num(r.get("purcAmt"))

    total = round(sum(val_fn(r) for r in rows))
    qty_total = round(sum(num(r.get("whQty")) for r in rows))
    groups = group_alois(rows, val_fn)

    by_grn = defaultdict(float)
    for r in rows:
        by_grn[r.get("grnNo") or "(없음)"] += val_fn(r)
    max_val = max(by_grn.values()) if by_grn else 0
    grn_rows = [
        {"grn": g, "val": round(v), "is_top": v == max_val and v > 0}
        for g, v in sorted(by_grn.items(), key=lambda kv: kv[0], reverse=True)
    ]

    top5 = sorted(rows, key=val_fn, reverse=True)[:5]
    top5_out = [{
        "item": r.get("itemCd"), "name": r.get("itemNm"), "alois": r.get("aloisCd"),
        "grn": r.get("grnNo"), "qty": int(num(r.get("whQty"))), "val": round(val_fn(r)),
        "oseries": (r.get("aloisCd") or "")[:1] == "O",
    } for r in top5]

    return {
        "total": total, "date": today_str, "grn_count": len(by_grn),
        "lines": len(rows), "qty": qty_total,
        "alois_groups": groups, "grn_rows": grn_rows, "top5": top5_out,
    }


def build_inventory(rows):
    pw = [r for r in rows if r.get("strgNm") == "부품창고"]

    def val_fn(r):
        return num(r.get("crtQty")) * num(r.get("movPrc"))

    def qty_fn(r):
        return num(r.get("crtQty"))

    total = round(sum(val_fn(r) for r in pw))
    groups = group_alois(pw, val_fn, qty_fn)
    codes = {r.get("aloisCd") for r in pw if r.get("aloisCd")}
    inv = {"total": total, "item_count": len(pw), "code_count": len(codes), "alois_groups": groups}
    return inv, pw


def build_oaov(pw_rows, inv_total):
    def val_fn(r):
        return num(r.get("crtQty")) * num(r.get("movPrc"))

    def items_of(rows):
        return sorted([{
            "item": r.get("itemCd"), "name": r.get("itemNm"), "alois": r.get("aloisCd"),
            "qty": int(num(r.get("crtQty"))), "val": round(val_fn(r)),
        } for r in rows], key=lambda i: i["val"], reverse=True)

    oa = [r for r in pw_rows if r.get("aloisCd") == "OA"]
    ov = [r for r in pw_rows if r.get("aloisCd") == "OV"]
    oa_val, ov_val = round(sum(val_fn(r) for r in oa)), round(sum(val_fn(r) for r in ov))
    oa_qty = round(sum(num(r.get("crtQty")) for r in oa))
    ov_qty = round(sum(num(r.get("crtQty")) for r in ov))
    oa_items, ov_items = items_of(oa), items_of(ov)
    top_items = sorted(oa_items + ov_items, key=lambda i: i["val"], reverse=True)[:10]
    total = oa_val + ov_val
    return {
        "total": total, "inv_total": inv_total,
        "pct": round(100 * total / inv_total, 2) if inv_total else 0,
        "oa_val": oa_val, "ov_val": ov_val, "oa_lines": len(oa), "ov_lines": len(ov),
        "oa_qty": oa_qty, "ov_qty": ov_qty, "item_count": len(oa) + len(ov),
        "oa_items": oa_items, "ov_items": ov_items, "top_items": top_items,
    }


def build_longstock(pw_rows, inv_total, now):
    """lastSaleDt(최종판매일) 기준 12개월+/24개월+ 미판매(장기재고) 집계. 판매이력 없음도 포함."""
    def val_fn(r):
        return num(r.get("crtQty")) * num(r.get("movPrc"))

    def months_since(dt_str):
        if not dt_str:
            return None
        d = datetime.strptime(dt_str[:10], "%Y-%m-%d")
        return (now.year - d.year) * 12 + (now.month - d.month) - (1 if now.day < d.day else 0)

    def items_of(rows):
        return sorted([{
            "item": r.get("itemCd"), "name": r.get("itemNm"), "alois": r.get("aloisCd"),
            "qty": int(num(r.get("crtQty"))), "val": round(val_fn(r)),
            "last_sale": (r.get("lastSaleDt") or "")[:10] or "판매이력 없음",
        } for r in rows], key=lambda i: i["val"], reverse=True)

    m12_rows, m24_rows = [], []
    for r in pw_rows:
        months = months_since(r.get("lastSaleDt"))
        if months is None or months >= 12:
            m12_rows.append(r)
        if months is None or months >= 24:
            m24_rows.append(r)

    m12_items, m24_items = items_of(m12_rows), items_of(m24_rows)
    m12_val = round(sum(i["val"] for i in m12_items))
    m24_val = round(sum(i["val"] for i in m24_items))
    return {
        "m12_items": m12_items, "m12_val": m12_val, "m12_count": len(m12_items),
        "m12_pct": round(100 * m12_val / inv_total, 2) if inv_total else 0,
        "m24_items": m24_items, "m24_val": m24_val, "m24_count": len(m24_items),
        "m24_pct": round(100 * m24_val / inv_total, 2) if inv_total else 0,
    }


def _num_of(r):
    return r.get("parInvNo") or r.get("roNo") or ""


def build_ext_shop(rows, period_label):
    c_rows = [r for r in rows if r.get("calcTpCd") == "C"]

    def group_by(detl):
        grp = defaultdict(lambda: {"invTot": 0.0, "cost": 0.0, "margin": 0.0, "num": None, "cust": None, "dt": None})
        for r in c_rows:
            if r.get("calcDetlTpNm") != detl:
                continue
            g = grp[r.get("invNo")]
            g["invTot"] += num(r.get("invTotAmt"))
            g["cost"] += num(r.get("oriSumAmt"))
            g["margin"] += num(r.get("marginAmt"))
            g["num"] = g["num"] or _num_of(r)
            g["cust"] = g["cust"] or r.get("custNm")
            g["dt"] = g["dt"] or r.get("invDt")
        return grp

    ext_grp, shop_grp = group_by("외부"), group_by("외부공업사")

    def items_of(grp):
        return [{
            "inv": k, "num": g["num"], "cust": g["cust"], "dt": (g["dt"] or "")[:16],
            "amt": round(g["invTot"]), "cost": round(g["cost"]),
        } for k, g in grp.items()]

    ext_items, shop_items = items_of(ext_grp), items_of(shop_grp)
    ext_valid = [i for i in ext_items if i["amt"] != 0]
    ext_all = sorted(ext_valid, key=lambda i: i["dt"], reverse=True)
    ext_total = sum(i["amt"] for i in ext_items)
    ext_cost = sum(i["cost"] for i in ext_items)
    ext_margin = round(sum(num(r.get("marginAmt")) for r in c_rows if r.get("calcDetlTpNm") == "외부"))

    shop_sorted = sorted(shop_items, key=lambda i: i["dt"], reverse=True)
    for i in shop_sorted:
        i["cancelled"] = i["amt"] == 0
        i["status"] = "완료" if i["amt"] != 0 else "취소"
    shop_total = sum(i["amt"] for i in shop_items)
    shop_cost = sum(i["cost"] for i in shop_items)
    if shop_items and shop_total == 0:
        shop_note = f"발행 {len(shop_items)}건 · 전액 취소되어 순매출 없음"
    elif not shop_items:
        shop_note = "이번 달 발행 없음"
    else:
        shop_note = f"유효 {len([i for i in shop_items if i['amt'] != 0])}건"

    ext = {
        "period": period_label, "total": ext_total, "cost": ext_cost, "margin": ext_margin,
        "valid_count": len(ext_valid), "cancelled_count": len(ext_items) - len(ext_valid),
        "lines": len([r for r in c_rows if r.get("calcDetlTpNm") == "외부"]), "top5": ext_all,
    }
    shop = {
        "total": shop_total, "cost": shop_cost,
        "valid_count": len([i for i in shop_items if i["amt"] != 0]),
        "lines": len([r for r in c_rows if r.get("calcDetlTpNm") == "외부공업사"]),
        "invoices": shop_sorted, "foot_note": shop_note,
    }
    return ext, shop


def _prod_code(r):
    m = re.match(r"^\[(\d+)\]", r.get("prodType") or "")
    return m.group(1) if m else None


def build_acc_tire(rows):
    ci_rows = [r for r in rows if r.get("calcTpCd") in ("C", "I")]
    acc_codes = {"3", "5", "7"}

    def group_by_invoice(codes):
        grp = defaultdict(lambda: {"invTot": 0.0, "cost": 0.0, "num": None, "cust": None, "dt": None})
        for r in ci_rows:
            if _prod_code(r) not in codes:
                continue
            k = r.get("invNo")
            if not k:
                continue
            g = grp[k]
            g["invTot"] += num(r.get("invTotAmt"))
            g["cost"] += num(r.get("oriSumAmt"))
            g["num"] = g["num"] or _num_of(r)
            g["cust"] = g["cust"] or r.get("custNm")
            g["dt"] = g["dt"] or r.get("invDt")
        return grp

    acc_grp = group_by_invoice(acc_codes)
    tire_grp = group_by_invoice({"8"})

    def sorted_items(grp):
        return sorted(
            [{"inv": k, "num": g["num"], "cust": g["cust"], "dt": (g["dt"] or "")[:16],
              "amt": round(g["invTot"]), "cost": round(g["cost"])} for k, g in grp.items()],
            key=lambda i: i["dt"], reverse=True,
        )

    acc_by_type, acc_lines, acc_total = [], 0, 0.0
    for c in sorted(acc_codes):
        crows = [r for r in ci_rows if _prod_code(r) == c]
        v = sum(num(r.get("invTotAmt")) for r in crows)
        acc_by_type.append({"code": c, "label": PROD_LABELS.get(c, c), "lines": len(crows), "val": round(v)})
        acc_lines += len(crows)
        acc_total += v
    acc_cost = round(sum(num(r.get("oriSumAmt")) for r in ci_rows if _prod_code(r) in acc_codes))

    tire_rows = [r for r in ci_rows if _prod_code(r) == "8"]
    tire_total = round(sum(num(r.get("invTotAmt")) for r in tire_rows))
    tire_cost = round(sum(num(r.get("oriSumAmt")) for r in tire_rows))

    by_sa = defaultdict(lambda: {"qty": 0.0, "val": 0.0, "margin": 0.0})
    for r in tire_rows:
        sa = r.get("emplNm") or "(미지정)"
        by_sa[sa]["qty"] += num(r.get("qty"))
        by_sa[sa]["val"] += num(r.get("invTotAmt"))
        by_sa[sa]["margin"] += num(r.get("marginAmt"))
    sa_list = sorted(
        [{"sa": k, "qty": round(v["qty"]), "val": round(v["val"]), "margin": round(v["margin"])}
         for k, v in by_sa.items()],
        key=lambda x: x["val"], reverse=True,
    )

    acc = {
        "total": round(acc_total), "cost": acc_cost, "lines": acc_lines,
        "invoices_count": len(acc_grp), "by_type": acc_by_type, "recent5": sorted_items(acc_grp)[:5],
    }
    tire = {
        "total": tire_total, "cost": tire_cost, "lines": len(tire_rows),
        "invoices_count": len(tire_grp), "by_sa": sa_list,
        "qty_total": sum(x["qty"] for x in sa_list), "margin_total": sum(x["margin"] for x in sa_list),
        "all_invoices": sorted_items(tire_grp),
    }
    return acc, tire


# ============================================================
# 렌더 & 저장
# ============================================================

def _git(*args, check=True, timeout=None):
    return subprocess.run(
        ["git", *args], cwd=str(BASE_DIR), capture_output=True, text=True, encoding="utf-8",
        check=check, timeout=timeout,
    )


def publish_to_github(commit_message: str) -> bool:
    """docs/ 안의 최신 대시보드를 깃허브에 push. 저장소가 아직 없으면 초기화."""
    if not (BASE_DIR / ".git").exists():
        print("깃 저장소 초기화...")
        _git("init")
        _git("remote", "add", "origin", GIT_REMOTE_URL)
        gitignore = BASE_DIR / ".gitignore"
        if not gitignore.exists():
            gitignore.write_text("__pycache__/\noutput/\n*.pyc\n", encoding="utf-8")

    _git("add", "docs", ".gitignore", "partsbay.py", "template.html.j2", "requirements.txt", check=False)
    status = _git("status", "--porcelain")
    if not status.stdout.strip():
        print("깃허브에 올릴 변경사항이 없습니다 (이전과 동일한 데이터).")
        return True

    _git("commit", "-m", commit_message)
    print("깃허브에 push 중... (세션 만료 시 브라우저 로그인 창이 뜰 수 있습니다)")
    push_timeout = 60  # 인증 대기 등으로 상시 실행 루프가 무한정 멈추지 않게
    try:
        result = _git("push", "-u", "origin", "main", check=False, timeout=push_timeout)
        if result.returncode != 0:
            # main이 처음이라 브랜치 이름이 다를 수 있음 (master 등) -> 강제로 main 사용
            _git("branch", "-M", "main", check=False)
            result = _git("push", "-u", "origin", "main", check=False, timeout=push_timeout)
    except subprocess.TimeoutExpired:
        print("push가 시간 내에 끝나지 않았습니다 (인증 대기 중일 수 있음). 다음 주기에 재시도합니다.")
        return False
    if result.returncode != 0:
        print("push 실패:")
        print(result.stdout)
        print(result.stderr)
        return False
    print("깃허브 push 완료.")
    return True


def render(data: dict) -> str:
    env = Environment(loader=FileSystemLoader(str(BASE_DIR)), autoescape=False)
    template = env.get_template(TEMPLATE_NAME)
    return template.render(**data)


LOCK_FILE = BASE_DIR / ".partsbay.lock"
CYCLE_INTERVAL_SEC = 600  # 상시 실행 모드에서 갱신 주기 (10분)
ONCE = "--once" in sys.argv  # 테스트용: 한 번만 돌고 종료 (기본은 계속 켜져 있는 상시 실행)


def _pid_alive(pid: int) -> bool:
    """Windows에선 os.kill(pid, 0)이 생존여부와 무관하게 항상 예외를 던져서 못 씀
    (그래서 잠금이 사실상 계속 무시되고 있었음) -> tasklist로 직접 확인."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True, text=True, timeout=5,
        )
        return str(pid) in result.stdout
    except Exception:
        return False  # 확인 불가하면 안전하게 "죽은 것"으로 간주해 새로 진행


def _acquire_lock() -> bool:
    """이미 실행 중인 인스턴스가 있으면 False. (중복 실행 -> 동시 로그인으로 세션 끊기는 사고 방지)"""
    if LOCK_FILE.exists():
        try:
            pid = int(LOCK_FILE.read_text().strip())
        except ValueError:
            pid = None
        if pid is not None and _pid_alive(pid):
            return False  # 여전히 실행 중
    LOCK_FILE.write_text(str(os.getpid()))
    return True


def run_cycle(page: Page) -> None:
    """이미 로그인된 page로 데이터 한 번 뽑아서 대시보드 생성 + 깃허브 push."""
    now = datetime.now(ZoneInfo("Asia/Seoul"))
    today_str = now.strftime("%Y-%m-%d")
    month_start = now.strftime("%Y-%m-01")
    period_label = f"{month_start} ~ {today_str}"
    generated_at = f"{now.month}월 {now.day}일 ({WEEKDAYS_KO[now.weekday()]}) · {now.strftime('%H:%M')} 기준"

    print(" - 오늘 입고현황")
    recv_rows = extract_receiving(page, today_str)
    print(" - 현재고리스트")
    inv_rows = extract_inventory(page)
    print(" - Turn Over 리포트 (당월)")
    to_rows = extract_turnover(page, month_start, today_str)

    print("집계 중...")
    recv = build_recv(recv_rows, today_str)
    inv, pw_rows = build_inventory(inv_rows)
    oaov = build_oaov(pw_rows, inv["total"])
    ext, shop = build_ext_shop(to_rows, period_label)
    acc, tire = build_acc_tire(to_rows)
    longstock = build_longstock(pw_rows, inv["total"], now)
    calendar_image = next((f for f in CALENDAR_IMAGE_CANDIDATES if (DOCS_DIR / f).exists()), None)

    data = {
        "meta": {"branch_name": BRANCH_NAME, "brch_code": f"BRCH {BRCH_CD}", "generated_at": generated_at,
                 "calendar_image": calendar_image},
        "recv": recv, "inv": inv, "oaov": oaov, "ext": ext, "shop": shop, "acc": acc, "tire": tire,
        "longstock": longstock,
    }

    html = render(data)
    stamp = now.strftime("%Y%m%d_%H%M")
    out_path = OUTPUT_DIR / f"dashboard_{stamp}.html"
    out_path.write_text(html, encoding="utf-8")
    (OUTPUT_DIR / "latest.html").write_text(html, encoding="utf-8")
    data_json = json.dumps(data, ensure_ascii=False, indent=2)
    (OUTPUT_DIR / "latest.json").write_text(data_json, encoding="utf-8")

    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")
    (DOCS_DIR / "data.json").write_text(data_json, encoding="utf-8")

    print(f"완료: {out_path}")
    if ONCE:
        webbrowser.open(out_path.resolve().as_uri())

    if PUBLISH_TO_GITHUB:
        publish_to_github(f"대시보드 갱신 {generated_at}")


def main():
    if not _acquire_lock():
        print("이미 실행 중인 인스턴스가 있습니다. (같은 프로그램을 두 번 띄우면 DMS 세션이 끊길 수 있어 종료합니다)")
        return
    AUTH_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        print("동성모터스 PARTS — DMS 접속 중...")
        with sync_playwright() as p:
            # 브라우저는 한 번만 띄워서 계속 켜둔다(로그인 세션 유지). 화면엔 안 보이게
            # 항상 최소화 상태로 띄우고, 로그인이 필요할 때만 notify_user()가 알려준다.
            ctx = p.chromium.launch_persistent_context(
                str(AUTH_DIR), headless=False, viewport={"width": 1440, "height": 900},
                args=["--start-minimized"],
            )
            page = ctx.new_page()
            ensure_logged_in(page)

            if ONCE:
                print("로그인 확인 완료. 데이터를 추출합니다...")
                run_cycle(page)
                ctx.close()
                return

            print(f"로그인 확인 완료. 상시 실행 시작 — {CYCLE_INTERVAL_SEC//60}분마다 자동 갱신합니다 (창은 계속 켜둔 채 백그라운드로 동작).")
            while True:
                try:
                    ensure_logged_in(page)  # 그 사이 세션이 끊겼으면 재확인
                    run_cycle(page)
                except Exception as e:
                    print(f"[이번 주기 실패, 다음 주기에 재시도] {type(e).__name__}: {e}")
                print(f"다음 갱신까지 {CYCLE_INTERVAL_SEC}초 대기...")
                time.sleep(CYCLE_INTERVAL_SEC)
    finally:
        LOCK_FILE.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
