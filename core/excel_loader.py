"""
Excel 기준 파일 로더.
경고알람기준 38개 룰과 원단/지퍼 분류 기준을 로드한다.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import openpyxl


@dataclass
class AlarmRule:
    section: str        # 지퍼 / 봉제 / 부자재 / 소재물성 / 아트웍
    condition: str      # 원문 조건 (예: "(NX/CX/ST/P) + (METAL/VS/NY)")
    risk_name: str      # 파동 주의, 박지 주의, …
    mechanism: str      # 사고 발생 원리
    alarm_msg: str      # ⚠️ [...] 시스템 알람 메시지
    checklist: str      # 체크포인트


@dataclass
class WeightLevel:
    level: str          # "Lv 1" ~ "Lv 5"
    max_gsm: float      # 해당 레벨의 최대 g/m²
    max_denier: float   # 데니어 기반 대체 기준


@dataclass
class ExcelData:
    alarm_rules: list[AlarmRule] = field(default_factory=list)
    weight_levels: list[WeightLevel] = field(default_factory=list)
    excel_path: str = ""


# 중량 레벨 기준 (원단구분 시트 기반)
_WEIGHT_LEVELS = [
    WeightLevel("Lv 1",  80,   20),
    WeightLevel("Lv 2",  160,  40),
    WeightLevel("Lv 3",  260,  75),
    WeightLevel("Lv 4",  360,  150),
    WeightLevel("Lv 5",  9999, 9999),
]

# 알람 섹션 헤더 키워드
_SECTION_KEYWORDS = {
    "지퍼":   ["지퍼 코드", "지퍼코드"],
    "봉제":   ["봉제 코드", "봉제코드"],
    "부자재": ["부자재 코드", "부자재코드"],
    "소재물성": ["소재 물성", "소재물성"],
    "아트웍": ["아트웍 코드", "아트웍코드"],
}


def load_excel(path: str) -> ExcelData:
    wb = openpyxl.load_workbook(path, data_only=True)
    data = ExcelData(excel_path=path, weight_levels=_WEIGHT_LEVELS)

    # 경고알람기준 시트 파싱
    ws = wb["경고알람기준"]
    current_section = ""
    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row):
            continue
        col_a = str(row[0] or "")

        # 섹션 헤더 감지
        new_section = _detect_section(col_a)
        if new_section:
            current_section = new_section
            continue

        # 헤더 행 스킵 (코드 분류라는 단어가 들어간 행)
        if "코드 분류" in col_a or "조합 조건" in col_a:
            continue

        # 빈 조건 행 스킵
        if not col_a.strip():
            continue

        alarm_msg = str(row[3] or "")
        checklist = str(row[4] or "")

        # ✅ 또는 ⚠️ 가 있는 행만 유효한 룰
        if not alarm_msg or ("⚠️" not in alarm_msg and "✅" not in alarm_msg):
            continue

        rule = AlarmRule(
            section=current_section,
            condition=col_a.strip(),
            risk_name=str(row[1] or ""),
            mechanism=_clean_text(str(row[2] or "")),
            alarm_msg=alarm_msg.strip(),
            checklist=checklist.strip(),
        )
        data.alarm_rules.append(rule)

    return data


def _clean_text(text: str) -> str:
    """Excel 셀 텍스트에서 개인 주석 태그([xxx님 픽] 등)를 제거한다."""
    return re.sub(r'\[[\w\s]+님\s*픽\]', '', text).strip()


def _detect_section(text: str) -> Optional[str]:
    for section, keywords in _SECTION_KEYWORDS.items():
        for kw in keywords:
            if kw in text:
                return section
    return None


def classify_weight(gsm: Optional[float], denier: Optional[float],
                    levels: list[WeightLevel]) -> str:
    """g/m² 또는 데니어로 중량 레벨 반환."""
    if gsm is not None:
        for lv in levels:
            if gsm <= lv.max_gsm:
                return lv.level
    if denier is not None:
        for lv in levels:
            if denier <= lv.max_denier:
                return lv.level
    return "Lv 3"  # 정보 없을 때 기본값


def classify_blend(composition: str) -> str:
    """혼용율 텍스트 → 혼용율 코드."""
    c = composition.upper()
    has_cotton = "COTTON" in c
    has_nylon  = "NYLON" in c
    has_poly   = "POLYESTER" in c or ("POLY" in c and "POLYURETHANE" not in c)
    has_pu     = "POLYURETHANE" in c or re.search(r'\bPU\b', c) is not None
    has_span   = "SPANDEX" in c or "SPAN" in c or "ELASTANE" in c
    has_sorona = "SORONA" in c or "T400" in c
    has_rayon  = "RAYON" in c or "MODAL" in c or "VISCOSE" in c

    if has_cotton and (has_sorona or (has_poly and "SORONA" in c)):
        return "CS"
    if has_nylon and (has_pu or has_span):
        return "NX"
    if has_nylon and not has_cotton and not has_poly and not has_rayon:
        return "N"
    if has_cotton and not has_poly and not has_nylon and not has_rayon and not has_pu and not has_span:
        return "C"
    if has_cotton:
        return "CX"
    if has_poly and not has_cotton and not has_nylon:
        return "P"
    return "S"


def classify_finish(finish: str) -> str:
    """후가공 텍스트 → 후가공 코드."""
    f = finish.upper()
    if not f or f in ("X", "NONE", "-", ""):
        return "ST"
    if any(k in f for k in ["LAYER", "LAMINATION", "TPU", "TPE"]):
        return "WP"
    # PU는 Finish 컬럼에 단독으로 있으면 라미네이션
    if re.search(r'\bPU\b', f) and "POLYURETHANE" not in f:
        return "WP"
    if any(k in f for k in ["WR", "DWR", "CIRE", "COATING", "FACE CIRE"]):
        return "WR"
    if any(k in f for k in ["BIO", "ENZYME", "PEACH", "BRUSH", "TUMBLE", "AIRO"]):
        return "AF"
    if any(k in f for k in ["UV", "QUICK DRY", "WICKING", "HYDROLYSIS"]):
        return "FN"
    return "ST"


def classify_stretch(composition: str) -> str:
    """혼용율 텍스트 → 스트레치 레벨."""
    c = composition.upper()

    # 스판/PU/엘라스테인 비율 추출
    pct_match = re.findall(r'(\d+(?:\.\d+)?)\s*%\s*(?:SPANDEX|SPAN|POLYURETHANE|PU\b|ELASTANE)', c)
    if pct_match:
        pct = max(float(p) for p in pct_match)
        if pct >= 12:
            return "Lv 5"
        if pct >= 6:
            return "Lv 4"
        return "Lv 3"

    if "SORONA" in c or "T400" in c or "JERSEY" in c or "FLEECE" in c:
        return "Lv 2"
    return "Lv 1"


def classify_fabric_type(fabric_type_str: str) -> str:
    """원단 타입 한글/영문 → W/K/S/D."""
    t = fabric_type_str.upper()
    if any(k in t for k in ["우븐", "WOVEN", "WOVEN"]):
        return "W"
    if any(k in t for k in ["스웨터", "SWEATER"]):
        return "S"
    if any(k in t for k in ["데님", "DENIM"]):
        return "D"
    # 다이마루, 니트, 저지 등은 모두 K
    return "K"
