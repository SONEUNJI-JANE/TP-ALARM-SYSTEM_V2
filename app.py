"""
TP Alarm System V2 - 사전 경고 시뮬레이터
작업지시서 작성 전 원/부자재 조합 선택으로 품질 리스크를 실시간 예측
"""
import json
import threading
import webbrowser
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_from_directory

import pandas as pd

# 로컬 core 모듈 임포트
from core.classifier import ClassifiedFabric, ClassifiedZipper, ClassifiedPack, classify_pack
from core.classifier import FabricEntry, ZipperEntry, ArtworkEntry
from core.alarm_engine import check_alarms
from core.excel_loader import load_excel

app = Flask(__name__)
app.secret_key = "tp_alarm_v2_secret"

BASE_DIR   = Path(__file__).parent
RAW_DIR    = BASE_DIR / "raw data"
DATA_EXCEL = RAW_DIR / "라벨링 취합.xlsx"
ALARM_EXCEL = RAW_DIR / "라벨링_매장조사데이터_260424.xlsx"

# ── 캐시 ──────────────────────────────────────────────
_cache: dict = {}

def _load_data() -> dict:
    """엑셀에서 원단/지퍼/부자재/RAW 데이터를 로드하여 캐시에 저장"""
    if _cache:
        return _cache

    xl = pd.ExcelFile(str(DATA_EXCEL))
    sheet_names = xl.sheet_names  # 인코딩 무관 순서 유지

    # ── 1. RAW(Apparel) : header=3, Q열(idx 16) = Style Code ──
    raw_df = pd.read_excel(str(DATA_EXCEL), sheet_name='RAW(Apparel)', header=3)
    raw_df = raw_df.rename(columns={raw_df.columns[16]: 'Style Code'})
    raw_df = raw_df.dropna(subset=['Style Code'])
    raw_df['Style Code'] = raw_df['Style Code'].astype(str).str.strip()

    styles = []
    seen_styles = set()
    for _, row in raw_df.iterrows():
        sc = str(row.get('Style Code', '')).strip()
        if not sc or sc == 'nan' or sc in seen_styles:
            continue
        seen_styles.add(sc)
        styles.append({
            "style_code":    sc,
            "desc_eng":      str(row.get('Description (Eng)', '') or ''),
            "category":      str(row.get('Category', '') or ''),
            "item":          str(row.get('item', '') or ''),
            "fit":           str(row.get('FIT', '') or ''),
            "length":        str(row.get('Length', '') or ''),
            "fabric":        str(row.get('Fabric', '') or ''),
            "fiber_content": str(row.get('Fiber_content', '') or ''),
            "fb_weight":     str(row.get('Fb_Weight', '') or ''),
            "function":      str(row.get('Function', '') or ''),
            "designer":      str(row.get('DS', '') or ''),
            "designer_fab":  str(row.get('DS_Fabric', '') or ''),
            "td":            str(row.get('DS_Tech', '') or ''),
            "vendor":        str(row.get('Vendor', '') or ''),
            "finish":        str(row.get('Finishing', '') or ''),
            "strech":        str(row.get('Strech', '') or ''),
        })

    # ── 2. 원단 시트 ──
    fab_df = pd.read_excel(str(DATA_EXCEL), sheet_name=sheet_names[3], header=0)
    # 첫 행이 보조 헤더인 경우 skip
    fab_df.columns = [str(c).split('\n')[0].strip() for c in fab_df.columns]
    # 컬럼 리맵
    col_map_fab = {
        fab_df.columns[0]: 'fabric_name',
        fab_df.columns[1]: 'composition',
        fab_df.columns[5]: 'fabrication',
        fab_df.columns[7]: 'weight_gsm',
    }
    # finish 컬럼은 별도 처리 (시리즈 ambiguous 방지)
    finish_col = fab_df.columns[6]
    fab_df = fab_df.rename(columns=col_map_fab)
    fab_df = fab_df.dropna(subset=['fabric_name'])
    fab_df = fab_df[fab_df['fabric_name'].astype(str).str.strip().ne('')]

    def _safe(val):
        if val is None: return ''
        try:
            import math
            if isinstance(val, float) and math.isnan(val): return ''
        except Exception: pass
        return str(val).strip() if str(val).strip() != 'nan' else ''

    fabrics = []
    for _, row in fab_df.iterrows():
        nm = _safe(row['fabric_name'])
        if nm:
            fabrics.append({
                "name":        nm,
                "composition": _safe(row.get('composition', '')),
                "fabrication": _safe(row.get('fabrication', '')),
                "finish":      _safe(row[finish_col]),
                "weight_gsm":  _safe(row.get('weight_gsm', '')),
            })

    # ── 3. 지퍼 시트 ──
    zip_df = pd.read_excel(str(DATA_EXCEL), sheet_name=sheet_names[0], header=0)
    zip_df.columns = [str(c).strip() for c in zip_df.columns]
    # 주요 컬럼 찾기
    zipper_name_col = zip_df.columns[12] if len(zip_df.columns) > 12 else zip_df.columns[0]  # ZIPPER+SLIDER(ENG)
    zip_df = zip_df.dropna(subset=[zipper_name_col])

    zippers = []
    for _, row in zip_df.iterrows():
        nm = str(row.get(zipper_name_col, '')).strip()
        if nm and nm != 'nan':
            zippers.append({
                "name":    nm,
                "quality": str(row.get('QUALITY', '') or row.get(zip_df.columns[4], '') or ''),
                "size":    str(row.get('SIZE', '') or row.get(zip_df.columns[5], '') or ''),
            })

    # ── 4. 부자재 시트 ──
    acc_df = pd.read_excel(str(DATA_EXCEL), sheet_name=sheet_names[1], header=0)
    acc_df.columns = [str(c).strip() for c in acc_df.columns]
    acc_name_col = acc_df.columns[0]
    acc_df = acc_df.dropna(subset=[acc_name_col])

    accessories = []
    seen_acc = set()
    # 부자재: col[6] = TO-BE 이름, col[3] = 코드, col[0] = 분류
    to_be_col = acc_df.columns[6] if len(acc_df.columns) > 6 else acc_df.columns[0]
    code_col   = acc_df.columns[3] if len(acc_df.columns) > 3 else acc_df.columns[0]
    category_col = acc_df.columns[0]
    for _, row in acc_df.iterrows():
        nm = _safe(row[to_be_col])
        if not nm or nm in seen_acc:
            # TO-BE가 없으면 AS-IS로 폴백
            nm = _safe(row[acc_df.columns[4]] if len(acc_df.columns) > 4 else '')
        if nm and nm not in seen_acc:
            seen_acc.add(nm)
            accessories.append({
                "name":     nm,
                "code":     _safe(row[code_col]),
                "category": _safe(row[category_col]),
            })

    _cache['styles']      = styles
    _cache['fabrics']     = fabrics
    _cache['zippers']     = zippers
    _cache['accessories'] = accessories
    return _cache


# ── 알람 엔진 ────────────────────────────────
def _get_alarm_excel():
    try:
        return load_excel(str(ALARM_EXCEL))
    except Exception as e:
        print(f"[WARN] 알람 엑셀 로드 실패: {e}")
        return None


def _run_alarm_check(fabric_name: str, composition: str, finish: str,
                     weight_gsm: str, zipper_names: list,
                     artwork_codes: list, artwork_big: bool) -> list[dict]:
    """선택된 조합으로 알람 엔진을 돌려 결과 반환."""
    try:
        excel = load_excel(str(ALARM_EXCEL))
        if not excel:
            return []

        # 원단 Entry 수동 생성 (FabricEntry 실제 필드: raw_name, composition, weight_gsm, finish, is_rib, is_yoko)
        weight_num = 0
        try:
            weight_num = float(weight_gsm) if weight_gsm else 0
        except Exception:
            pass

        raw_fabric = FabricEntry(
            raw_name=fabric_name,
            composition=composition,
            weight_gsm=weight_num,
            finish=finish,
            is_rib=False,
            is_yoko=False
        )

        # 지퍼 Entry 목록 생성
        raw_zippers = []
        for zn in zipper_names:
            zn_upper = zn.upper()
            material = "NY" if "NYLON" in zn_upper else \
                       "VS" if "VISLON" in zn_upper or "PLASTIC" in zn_upper else \
                       "METAL" if "METAL" in zn_upper else "NY"
            size_str = ""
            for tok in zn.split():
                if tok.startswith('#') or tok[0].isdigit():
                    size_str = tok.replace('#','').strip()
                    break
            raw_zippers.append(ZipperEntry(
                raw_spec=zn,
                material=material,
                size=size_str
            ))

        # 아트웍 Entry 목록 생성
        raw_artworks = []
        for ac in artwork_codes:
            raw_artworks.append(ArtworkEntry(raw_spec=ac, artwork_hint=ac))
        if artwork_big and not artwork_codes:
            raw_artworks.append(ArtworkEntry(raw_spec='대형아트웍', artwork_hint='BIG'))

        pack = classify_pack(
            style_code="",
            fabrics=[raw_fabric],
            zippers=raw_zippers,
            accessories=[],
            artworks=raw_artworks,
            excel_data=excel
        )
        alarms = check_alarms(pack, excel)
        return [
            {
                "severity":    a.severity,
                "section":     a.section,
                "alarm_msg":   a.alarm_msg,
                "mechanism":   a.mechanism,
                "checklist":   a.checklist,
                "combination": a.combination,
            }
            for a in alarms
        ]
    except Exception as e:
        import traceback; traceback.print_exc()
        return [{"severity": "INFO", "section": "오류", "alarm_msg": f"알람 엔진 오류: {e}",
                 "mechanism": "", "checklist": "", "combination": ""}]


# ── Routes ──────────────────────────────────────────────
@app.route('/')
def index():
    data = _load_data()
    return render_template('index.html',
                           style_count=len(data['styles']),
                           fabric_count=len(data['fabrics']),
                           zipper_count=len(data['zippers']),
                           acc_count=len(data['accessories']))


@app.route('/api/data')
def api_data():
    data = _load_data()
    return jsonify({
        "styles":      data['styles'][:500],   # 너무 크면 짤라서 전달
        "fabrics":     data['fabrics'],
        "zippers":     data['zippers'],
        "accessories": data['accessories'],
    })


@app.route('/api/style/<style_code>')
def api_style(style_code):
    data = _load_data()
    style = next((s for s in data['styles'] if s['style_code'] == style_code), None)
    if not style:
        return jsonify({"error": "Not found"}), 404
    return jsonify(style)


@app.route('/api/simulate', methods=['POST'])
def api_simulate():
    body = request.get_json(force=True)
    fabric_name   = body.get('fabric_name', '')
    composition   = body.get('composition', '')
    finish        = body.get('finish', '')
    weight_gsm    = body.get('weight_gsm', '')
    zipper_names  = body.get('zipper_names', [])
    artwork_codes = body.get('artwork_codes', [])
    artwork_big   = bool(body.get('artwork_big', False))

    alarms = _run_alarm_check(fabric_name, composition, finish, weight_gsm,
                               zipper_names, artwork_codes, artwork_big)
    return jsonify({"alarms": alarms, "total": len(alarms)})


if __name__ == '__main__':
    # 미리 데이터 로드
    print("[V2] 마스터 데이터 로드 중...")
    _load_data()
    print("[V2] 로드 완료!")
    port = 5001
    threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(host='127.0.0.1', port=port, debug=False)
