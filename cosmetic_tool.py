#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
화장품 성분표 자동 작성 도구
=================================

공공데이터포털 "식품의약품안전처_화장품 원료성분정보" OpenAPI를 이용해
한글 성분명 목록을 넣으면 영문명 / CAS.NO / 기능(기원·정의)을 자동으로 채워
엑셀 성분표를 만들어 줍니다. 성분 비율은 직접 입력하시면 됩니다.

End Point : https://apis.data.go.kr/1471000/CsmtcsIngdCpntInfoService01

사용 흐름
---------
1) (최초 1회) 전체 성분 DB 내려받기:   python cosmetic_tool.py fetch
2) 엑셀(input.xlsx)에 한글 성분명/비율 입력 후:  python cosmetic_tool.py lookup
   (또는 한 번에:  python cosmetic_tool.py run )

자세한 사용법은 README.md 를 보세요.
"""

import argparse
import configparser
import json
import os
import re
import sys
import time
import unicodedata
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("requests 패키지가 필요합니다.  pip install requests")

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.exit("openpyxl 패키지가 필요합니다.  pip install openpyxl")


# ----------------------------------------------------------------------------
# 설정
# ----------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CACHE = BASE_DIR / "cache" / "ingredients.json"
DEFAULT_INPUT = BASE_DIR / "input.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output.xlsx"

ENDPOINT = "https://apis.data.go.kr/1471000/CsmtcsIngdCpntInfoService01"

# 오퍼레이션명을 모르더라도 동작하도록 후보를 순서대로 시도한다.
# (End Point 의 마지막 경로 = 서비스명. 오퍼레이션명은 보통 그 변형이다.)
OPERATION_CANDIDATES = [
    "getCsmtcsIngdCpntInfoService01",
    "getCsmtcsIngdCpntInfoList",
    "getCsmtcsIngdCpntInfo",
    "getCsmtcsIngdCpntInfo01",
    "getCsmtcsIngdCpntInfoServiceList",
    "CsmtcsIngdCpntInfoService01",
    "getList",
]

# 인증키 문제로 보이는 신호 (조기 안내용)
KEY_ERROR_HINTS = ("SERVICE_KEY", "SERVICEKEY", "인증", "CERTIF", "REGISTERED", "ACCESS_DENIED")

# 출력 엑셀 열 순서 (요청하신 항목)
OUTPUT_HEADERS = ["한글성분명", "영문명", "성분비율", "CAS.NO", "기능", "매칭상태"]


# ----------------------------------------------------------------------------
# 인증키 읽기
# ----------------------------------------------------------------------------
def load_service_key(cli_key=None):
    """우선순위: CLI 인자 > 환경변수 DATA_GO_KR_KEY > config.ini"""
    if cli_key:
        return cli_key.strip()
    env = os.environ.get("DATA_GO_KR_KEY")
    if env:
        return env.strip()
    cfg = BASE_DIR / "config.ini"
    if cfg.exists():
        parser = configparser.ConfigParser()
        parser.read(cfg, encoding="utf-8")
        key = parser.get("api", "service_key", fallback="").strip()
        if key and key != "여기에_인증키_입력":
            return key
    sys.exit(
        "인증키를 찾을 수 없습니다.\n"
        "  - config.ini 의 [api] service_key 에 입력하거나\n"
        "  - 환경변수 DATA_GO_KR_KEY 로 지정하거나\n"
        "  - --service-key 옵션으로 전달하세요."
    )


# ----------------------------------------------------------------------------
# API 호출
# ----------------------------------------------------------------------------
def _parse_response(text):
    """data.go.kr 응답(JSON 또는 XML 오류)에서 (header_code, items, total) 추출."""
    text = text.strip()
    # 1) JSON 시도
    try:
        data = json.loads(text)
        resp = data.get("response", data)
        header = resp.get("header", {}) or {}
        code = str(header.get("resultCode", header.get("code", ""))).strip()
        msg = header.get("resultMsg", header.get("message", ""))
        body = resp.get("body", {}) or {}
        items = body.get("items", [])
        # items 가 {"item": [...]} 형태인 경우
        if isinstance(items, dict):
            items = items.get("item", [])
        if isinstance(items, dict):  # 단일 item
            items = [items]
        total = body.get("totalCount", len(items) if items else 0)
        return code, msg, items, int(total or 0)
    except (json.JSONDecodeError, AttributeError):
        pass
    # 2) XML 오류 응답 처리 (인증키 오류 등은 보통 XML 로 옴)
    m_code = re.search(r"<returnReasonCode>(.*?)</returnReasonCode>", text)
    m_msg = re.search(r"<returnAuthMsg>(.*?)</returnAuthMsg>", text)
    if m_code:
        return m_code.group(1), (m_msg.group(1) if m_msg else ""), [], 0
    m_msg2 = re.search(r"<errMsg>(.*?)</errMsg>|<resultMsg>(.*?)</resultMsg>", text)
    err = m_msg2.group(0) if m_msg2 else text[:200]
    return "ERR", err, [], 0


def call_api(session, service_key, operation, page_no=1, num_rows=100, timeout=30):
    url = f"{ENDPOINT}/{operation}"
    params = {
        "serviceKey": service_key,
        "pageNo": page_no,
        "numOfRows": num_rows,
        "type": "json",
        "_type": "json",
    }
    r = session.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return _parse_response(r.text)


def detect_operation(session, service_key, debug=False):
    """동작하는 오퍼레이션명을 자동 탐색. 모든 후보를 시도하고 첫 성공을 반환."""
    key_error = None
    last = None
    for op in OPERATION_CANDIDATES:
        try:
            code, msg, items, total = call_api(session, service_key, op, 1, 1)
        except requests.RequestException as e:
            if debug:
                print(f"  [{op}] 요청 실패: {e}")
            last = ("REQERR", str(e))
            continue
        if debug:
            print(f"  [{op}] code={code} msg={msg!r} total={total} items={len(items)}")
        if code == "00" or items:
            return op
        last = (code, msg)
        # 인증키 신호는 기억만 해두고(조기 종료하지 않음) 모든 후보를 끝까지 시도
        if any(h in str(msg).upper() for h in KEY_ERROR_HINTS):
            key_error = (code, msg)
    if key_error:
        sys.exit(f"인증키 오류로 보입니다 (code={key_error[0]}, msg={key_error[1]}).\n"
                 f"공공데이터포털에서 발급한 일반 인증키(Decoding)와 '활용신청' 승인 상태를 확인하세요.\n"
                 f"신청 직후라면 적용까지 시간이 걸릴 수 있습니다.")
    if last and debug:
        print(f"  마지막 응답: code={last[0]} msg={last[1]!r}")
    return None


# ----------------------------------------------------------------------------
# 응답 필드 자동 감지
# ----------------------------------------------------------------------------
CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")
HANGUL_RE = re.compile(r"[가-힣]")


def _is_mostly_ascii(s):
    s = str(s)
    letters = [c for c in s if c.isalpha()]
    if not letters:
        return False
    ascii_letters = [c for c in letters if ord(c) < 128]
    return len(ascii_letters) / len(letters) > 0.8


def detect_fields(items):
    """샘플 item 들에서 한글명/영문명/CAS/정의 필드명을 추론."""
    sample = [it for it in items if isinstance(it, dict)][:50]
    if not sample:
        return {}
    keys = list(sample[0].keys())
    mapping = {}

    def pick_by_name(substrs):
        for k in keys:
            ku = k.upper()
            if any(s in ku for s in substrs):
                return k
        return None

    # 이름 기반 우선 추정
    mapping["cas"] = pick_by_name(["CAS"])
    mapping["eng"] = pick_by_name(["ENG"])
    mapping["kor"] = pick_by_name(["KOR", "KRN", "KORN"])
    mapping["defn"] = pick_by_name(
        ["DEFN", "DEFINITION", "ORIGIN", "ORGN", "기원", "정의", "PRPOS", "USE"]
    )

    # 이름으로 못 찾으면 값 패턴으로 추정
    def value_samples(k):
        return [str(it.get(k, "")) for it in sample if it.get(k)]

    if not mapping["cas"]:
        for k in keys:
            vals = value_samples(k)
            if vals and sum(1 for v in vals if CAS_RE.search(v)) >= max(1, len(vals) // 2):
                mapping["cas"] = k
                break

    if not mapping["eng"]:
        best, best_score = None, 0
        for k in keys:
            vals = value_samples(k)
            if not vals:
                continue
            score = sum(1 for v in vals if _is_mostly_ascii(v))
            if score > best_score:
                best, best_score = k, score
        mapping["eng"] = best

    if not mapping["kor"]:
        best, best_score = None, 0
        for k in keys:
            if k in (mapping.get("eng"), mapping.get("cas")):
                continue
            vals = value_samples(k)
            if not vals:
                continue
            score = sum(1 for v in vals if HANGUL_RE.search(v) and len(v) < 60)
            if score > best_score:
                best, best_score = k, score
        mapping["kor"] = best

    return {k: v for k, v in mapping.items() if v}


# ----------------------------------------------------------------------------
# 전체 DB 내려받기 (fetch)
# ----------------------------------------------------------------------------
def fetch_all(service_key, cache_path, num_rows=100, max_pages=None,
              operation=None, debug=False):
    session = requests.Session()
    session.headers.update({"User-Agent": "cosmetic-ingredient-tool/1.0"})

    if not operation:
        print("동작하는 오퍼레이션명을 탐색 중...")
        operation = detect_operation(session, service_key, debug=debug)
        if not operation:
            sys.exit(
                "오퍼레이션명을 자동으로 찾지 못했습니다.\n"
                "공공데이터포털 상세페이지의 '상세기능정보'에 있는 오퍼레이션명을 확인해\n"
                "--operation 옵션으로 직접 지정해 주세요.\n"
                f"시도한 후보: {OPERATION_CANDIDATES}"
            )
        print(f"  -> 오퍼레이션: {operation}")

    all_items = []
    page = 1
    field_map = {}
    while True:
        code, msg, items, total = call_api(session, service_key, operation, page, num_rows)
        if code not in ("00", "", "ERR") and not items:
            sys.exit(f"API 오류 (page {page}): code={code}, msg={msg}")
        if not items:
            break
        if not field_map:
            field_map = detect_fields(items)
            print(f"  감지된 필드 매핑: {field_map}")
            if debug:
                print(f"  샘플 item: {json.dumps(items[0], ensure_ascii=False)[:500]}")
        all_items.extend(items)
        print(f"  page {page}: 누적 {len(all_items)} / 전체 {total}")
        if total and len(all_items) >= total:
            break
        if max_pages and page >= max_pages:
            break
        page += 1
        time.sleep(0.2)  # 서버 배려

    if not all_items:
        sys.exit("내려받은 데이터가 없습니다. 인증키/오퍼레이션을 확인하세요.")

    cache_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "operation": operation,
        "field_map": field_map,
        "count": len(all_items),
        "items": all_items,
    }
    cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"완료: {len(all_items)}건을 {cache_path} 에 저장했습니다.")
    return payload


# ----------------------------------------------------------------------------
# 검색용 인덱스
# ----------------------------------------------------------------------------
def normalize_name(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.lower()
    s = re.sub(r"\s+", "", s)               # 공백 제거
    s = re.sub(r"[()\[\]{}·,./\-_'\"]", "", s)  # 흔한 구분문자 제거
    return s


def build_index(payload):
    field_map = payload.get("field_map", {})
    kor_key = field_map.get("kor")
    index = {}
    if not kor_key:
        return index, field_map
    for item in payload["items"]:
        if not isinstance(item, dict):
            continue
        name = item.get(kor_key, "")
        norm = normalize_name(name)
        if not norm:
            continue
        index.setdefault(norm, []).append(item)
    return index, field_map


def lookup_one(name, index, field_map):
    norm = normalize_name(name)
    matches = index.get(norm, [])
    status = "찾음"
    if not matches:
        # 부분 일치(포함) 보조 검색
        partial = [its[0] for k, its in index.items() if norm and (norm in k or k in norm)]
        if len(partial) == 1:
            matches, status = [partial[0]], "유사일치"
        elif len(partial) > 1:
            matches, status = [partial[0]], "다중일치(첫번째)"
        else:
            return None, "못찾음"
    elif len(matches) > 1:
        status = "다중일치(첫번째)"
    item = matches[0]
    return {
        "eng": item.get(field_map.get("eng", ""), ""),
        "cas": item.get(field_map.get("cas", ""), ""),
        "defn": item.get(field_map.get("defn", ""), ""),
    }, status


# ----------------------------------------------------------------------------
# 입력 엑셀 읽기
# ----------------------------------------------------------------------------
def read_input(input_path, name_col=None, ratio_col=None):
    wb = load_workbook(input_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("입력 엑셀이 비어 있습니다.")
    header = [str(c).strip() if c is not None else "" for c in rows[0]]

    def find_col(candidates, override):
        if override:
            if override in header:
                return header.index(override)
            try:
                return int(override)  # 0-based 인덱스 허용
            except ValueError:
                sys.exit(f"열 '{override}' 을 찾을 수 없습니다. 헤더: {header}")
        for i, h in enumerate(header):
            if any(c in h for c in candidates):
                return i
        return None

    name_idx = find_col(["한글성분명", "성분명", "성분", "한글"], name_col)
    if name_idx is None:
        # 헤더가 없다고 보고 첫 열을 성분명으로 사용
        name_idx = 0
        data_rows = rows
    else:
        data_rows = rows[1:]
    ratio_idx = find_col(["비율", "함량", "ratio"], ratio_col)

    entries = []
    for r in data_rows:
        if not r:
            continue
        name = r[name_idx] if name_idx < len(r) else None
        if name is None or str(name).strip() == "":
            continue
        ratio = ""
        if ratio_idx is not None and ratio_idx < len(r) and r[ratio_idx] is not None:
            ratio = r[ratio_idx]
        entries.append({"name": str(name).strip(), "ratio": ratio})
    return entries


# ----------------------------------------------------------------------------
# 출력 엑셀 쓰기
# ----------------------------------------------------------------------------
def write_output(entries_results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "성분표"
    ws.append(OUTPUT_HEADERS)
    for row in entries_results:
        ws.append([
            row["name"], row["eng"], row["ratio"], row["cas"], row["defn"], row["status"],
        ])
    # 열 너비 보기 좋게
    widths = [22, 32, 10, 16, 40, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def do_lookup(cache_path, input_path, output_path, name_col, ratio_col):
    if not cache_path.exists():
        sys.exit(f"캐시가 없습니다: {cache_path}\n먼저 'fetch' 를 실행하세요.")
    payload = json.loads(cache_path.read_text(encoding="utf-8"))
    index, field_map = build_index(payload)
    if not index:
        sys.exit("캐시에서 한글 성분명 필드를 찾지 못했습니다. fetch 를 --debug 로 다시 실행해 보세요.")

    entries = read_input(input_path, name_col, ratio_col)
    results = []
    found = 0
    for e in entries:
        info, status = lookup_one(e["name"], index, field_map)
        if info:
            found += 1
        results.append({
            "name": e["name"],
            "ratio": e["ratio"],
            "eng": (info or {}).get("eng", ""),
            "cas": (info or {}).get("cas", ""),
            "defn": (info or {}).get("defn", ""),
            "status": status,
        })
    write_output(results, output_path)
    print(f"완료: {len(results)}개 중 {found}개 매칭. 결과 -> {output_path}")
    not_found = [r["name"] for r in results if r["status"] == "못찾음"]
    if not_found:
        print(f"못 찾은 성분({len(not_found)}): {', '.join(not_found[:20])}"
              + (" ..." if len(not_found) > 20 else ""))


# ----------------------------------------------------------------------------
# probe: 응답 원문 확인용
# ----------------------------------------------------------------------------
def do_probe(service_key, operation):
    session = requests.Session()
    if not operation:
        operation = detect_operation(session, service_key, debug=True)
        if not operation:
            sys.exit("동작하는 오퍼레이션을 찾지 못했습니다.")
    code, msg, items, total = call_api(session, service_key, operation, 1, 3)
    print(f"operation={operation}")
    print(f"resultCode={code}  resultMsg={msg}  totalCount={total}  items={len(items)}")
    if items:
        print("필드명:", list(items[0].keys()))
        print("샘플:")
        print(json.dumps(items[0], ensure_ascii=False, indent=2))
        print("감지된 매핑:", detect_fields(items))


# ----------------------------------------------------------------------------
# 입력 템플릿 생성
# ----------------------------------------------------------------------------
def make_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "입력"
    ws.append(["한글성분명", "비율"])
    for sample in ["정제수", "글리세린", "부틸렌글라이콜"]:
        ws.append([sample, ""])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 10
    wb.save(path)
    print(f"입력 템플릿 생성: {path}")


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(
        description="화장품 성분표 자동 작성 도구 (식약처 화장품 원료성분정보 OpenAPI)")
    sub = p.add_subparsers(dest="cmd", required=True)

    common_key = argparse.ArgumentParser(add_help=False)
    common_key.add_argument("--service-key", help="공공데이터포털 인증키 (미지정시 config.ini/환경변수 사용)")

    sp = sub.add_parser("fetch", parents=[common_key], help="전체 성분 DB 내려받아 캐시")
    sp.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    sp.add_argument("--num", type=int, default=100, help="페이지당 건수 (기본 100)")
    sp.add_argument("--max-pages", type=int, default=None, help="최대 페이지 수 제한")
    sp.add_argument("--operation", default=None, help="오퍼레이션명 직접 지정")
    sp.add_argument("--debug", action="store_true")

    lp = sub.add_parser("lookup", help="입력 엑셀의 성분명을 캐시와 대조해 성분표 작성")
    lp.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    lp.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    lp.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    lp.add_argument("--name-col", default=None, help="성분명 열 이름 또는 0-based 인덱스")
    lp.add_argument("--ratio-col", default=None, help="비율 열 이름 또는 0-based 인덱스")

    rp = sub.add_parser("run", parents=[common_key],
                        help="fetch(캐시 없을 때) + lookup 한 번에 실행")
    rp.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    rp.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    rp.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    rp.add_argument("--name-col", default=None)
    rp.add_argument("--ratio-col", default=None)
    rp.add_argument("--operation", default=None)

    pp = sub.add_parser("probe", parents=[common_key], help="API 응답 원문/필드 확인")
    pp.add_argument("--operation", default=None)

    tp = sub.add_parser("template", help="입력 엑셀 템플릿 생성")
    tp.add_argument("--path", type=Path, default=DEFAULT_INPUT)

    args = p.parse_args()

    if args.cmd == "fetch":
        key = load_service_key(args.service_key)
        fetch_all(key, args.cache, args.num, args.max_pages, args.operation, args.debug)
    elif args.cmd == "lookup":
        do_lookup(args.cache, args.input, args.output, args.name_col, args.ratio_col)
    elif args.cmd == "run":
        key = load_service_key(args.service_key)
        if not args.cache.exists():
            fetch_all(key, args.cache, operation=args.operation)
        do_lookup(args.cache, args.input, args.output, args.name_col, args.ratio_col)
    elif args.cmd == "probe":
        key = load_service_key(args.service_key)
        do_probe(key, args.operation)
    elif args.cmd == "template":
        make_template(args.path)


if __name__ == "__main__":
    main()
