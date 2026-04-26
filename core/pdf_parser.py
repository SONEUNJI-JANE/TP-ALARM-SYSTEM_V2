"""
PDF 작업지시서 파서.
전략 1) pdfplumber 표 추출 → F/N/Z 섹션 파싱
전략 2) 전체 텍스트 정규식 폴백 — 표 구조가 달라도 동작
두 전략을 순서대로 시도하여 최대한 많은 데이터를 추출한다.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import Optional
import pdfplumber


# ── 데이터 클래스 ──────────────────────────────────────────────
@dataclass
class FabricEntry:
    raw_name: str
    composition: str = ""
    weight_gsm: Optional[float] = None
    denier: Optional[float] = None
    finish: str = ""
    fabric_type_hint: str = ""
    is_rib: bool = False
    is_yoko: bool = False     # 요꼬(넥 전용 립) 여부 — 넥 알람에만 사용


@dataclass
class ZipperEntry:
    raw_spec: str
    brand: str = ""
    material: str = ""
    size: str = ""
    zipper_type: str = ""


@dataclass
class AccessoryEntry:
    raw_spec: str
    category_hint: str = ""


@dataclass
class ArtworkEntry:
    raw_spec: str
    artwork_hint: str = ""


@dataclass
class TechPackData:
    pdf_path: str
    style_code: str = ""
    fabrics: list[FabricEntry] = field(default_factory=list)
    zippers: list[ZipperEntry] = field(default_factory=list)
    accessories: list[AccessoryEntry] = field(default_factory=list)
    artworks: list[ArtworkEntry] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    parse_method: str = ""   # "table" / "text" / "failed"


# ── 정규식 ────────────────────────────────────────────────────
_STYLE_CODE_RE  = re.compile(r'\b3[A-Z]{4}\d{4}\b')
_SECTION_RE     = re.compile(r'^([FNZPS])\s*\(')

# 스펙 필드 정규식
_COMPOSITION_RE = re.compile(
    r'(?:Composition|Content|Material)\s*:\s*(.+?)(?=\n(?:Price|Weight|Finish|Width|Color|Placement)|$)',
    re.I | re.S)
_COMPOSITION_INLINE_RE = re.compile(
    r'(?:Composition|Content|Material)\s*:\s*([^\n]{5,150})', re.I)
_WEIGHT_RE      = re.compile(r'Weight\s*\(?G[/\\]SQM\)?\s*:\s*(\d+(?:\.\d+)?)', re.I)
_DENIER_RE      = re.compile(r'\b(\d+(?:\.\d+)?)D\b')
_FINISH_RE      = re.compile(r'(?:Finish|Finishing)\s*:\s*([^\n]{1,80})', re.I)

# 지퍼 정규식 — 브랜드 없어도 잡힘
_ZIPPER_RE = re.compile(
    r'(?:(SBS|SAB|YKK|RIRI|Talon)\s*[_\s]\s*)?'
    r'(VS\s*WR|VSWR|VS|NY\s*WR|NYWR|NY|METAL\s*LUX|METALUX|METAL)'
    r'\s*[_\s#]*#?([35])\s*'
    r'(2[\s\-]?WAY|C[\s/]?E|O[\s/]?E|CLOSE|OPEN)',
    re.I)

_ARTWORK_SIZE_RE = re.compile(r'\b(\d+(?:\.\d+)?)\s*[Cc][Mm]\s*[↑↓]?', re.I)


# ── 진입점 ────────────────────────────────────────────────────
def parse_pdf(pdf_path: str) -> TechPackData:
    result = TechPackData(pdf_path=pdf_path)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            _extract_style_code(pdf, result)

            # 전략 1: 표 기반 파싱
            found = _parse_by_tables(pdf, result)

            # 전략 2: 아무것도 추출 못했으면 전체 텍스트 폴백
            if not result.fabrics and not result.zippers and not result.artworks:
                _parse_by_text(pdf, result)

            if not found and not result.fabrics and not result.zippers and not result.artworks:
                result.parse_method = "failed"
                result.warnings.append(
                    "원부자재리스트를 찾지 못했습니다. "
                    "F·N·Z 섹션이 있는 작업지시서인지 확인해주세요."
                )
            elif not result.fabrics and not result.zippers and not result.artworks:
                result.parse_method = "failed"
                result.warnings.append(
                    "원부자재리스트 페이지를 찾았으나 소재·지퍼 데이터를 읽지 못했습니다. "
                    "PDF 내 표 구조가 예상과 다를 수 있습니다."
                )

    except Exception as e:
        result.parse_method = "failed"
        result.warnings.append(f"PDF 파싱 오류: {e}")

    return result


# ── 스타일 코드 추출 ──────────────────────────────────────────
def _extract_style_code(pdf, result: TechPackData) -> None:
    for page in pdf.pages[:3]:
        text = page.extract_text() or ""
        codes = _STYLE_CODE_RE.findall(text)
        if codes:
            result.style_code = codes[0]
            return


# ══════════════════════════════════════════════════════════════
# 전략 1: 표 기반 파싱
# ══════════════════════════════════════════════════════════════
def _parse_by_tables(pdf, result: TechPackData) -> bool:
    """F/N/Z 섹션 헤더가 있는 페이지를 표에서 파싱한다. True=페이지 발견."""
    found = False
    for page in pdf.pages[:12]:
        tables = page.extract_tables()
        if _is_material_page_table(tables, page):
            found = True
            _extract_from_tables(tables, result)
            if not result.fabrics and not result.zippers:
                # 같은 페이지 텍스트도 시도
                _extract_from_page_text(page, result)

    if result.fabrics or result.zippers or result.artworks:
        result.parse_method = "table"
    return found


def _is_material_page_table(tables: list, page) -> bool:
    for table in tables:
        for row in table or []:
            cell = str(row[0] or "") if row else ""
            if _SECTION_RE.match(cell.strip()):
                return True
    text = page.extract_text() or ""
    upper = text.upper()
    if any(k in upper for k in ["BILL OF MATERIALS", "BOM", "TRIMS", "FABRICS", "원부자재"]):
        return True
    # Composition 필드가 있으면 원부자재 페이지로 간주
    return bool(re.search(r'(?:Composition|Content|Material)\s*:', text, re.I))


def _extract_from_tables(tables: list, result: TechPackData) -> None:
    current_section = ""
    for table in tables:
        for row in table or []:
            if not row:
                continue
            cell0 = str(row[0] or "").strip()

            m = _SECTION_RE.match(cell0)
            if m:
                current_section = m.group(1)
                continue

            # 모든 컬럼에서 스펙 텍스트 찾기 (Col 위치 무관)
            spec = _best_spec_cell(row)
            if not spec or len(spec) < 5:
                continue

            if current_section == "F":
                _handle_fabric_row(cell0, spec, result)
            elif current_section == "N":
                _handle_accessory_row(cell0, spec, result)
            elif current_section == "Z":
                _handle_artwork_row(cell0, spec, result)


def _best_spec_cell(row: list) -> str:
    """행의 모든 컬럼을 순회해 Composition/Weight/지퍼 패턴이 있는 셀을 우선 반환."""
    candidates = []
    for idx, cell in enumerate(row):
        if cell is None:
            continue
        text = str(cell).strip()
        if len(text) < 4:
            continue
        # 스펙 키워드 점수
        score = 0
        if re.search(r'(?:Composition|Content|Material)\s*:', text, re.I):
            score += 10
        if re.search(r'Weight|GSM', text, re.I):
            score += 5
        if re.search(r'(?:Finish|Finishing)\s*:', text, re.I):
            score += 5
        if _ZIPPER_RE.search(text):
            score += 8
        if score > 0 or (idx in (1, 2, 3) and len(text) > 10):
            candidates.append((score, idx, text))

    if candidates:
        candidates.sort(key=lambda x: -x[0])
        return candidates[0][2]
    return ""


# ══════════════════════════════════════════════════════════════
# 전략 2: 전체 텍스트 폴백
# ══════════════════════════════════════════════════════════════
def _parse_by_text(pdf, result: TechPackData) -> None:
    """표 추출 실패 시 페이지 전체 텍스트를 줄 단위로 파싱한다."""
    for page in pdf.pages[:12]:
        _extract_from_page_text(page, result)
    if result.fabrics or result.zippers or result.artworks:
        result.parse_method = "text"


def _extract_from_page_text(page, result: TechPackData) -> None:
    text = page.extract_text() or ""
    if not text:
        return

    # Composition 블록 전체를 줄 단위로 스캔
    # 각 Composition 블록이 하나의 원단 엔트리
    blocks = re.split(r'(?=\b(?:Composition|Content|Material)\s*:)', text, flags=re.I)
    for block in blocks:
        if not re.search(r'(?:Composition|Content|Material)\s*:', block, re.I):
            continue

        entry = FabricEntry(raw_name="(자동감지)")

        # Composition
        m = _COMPOSITION_INLINE_RE.search(block)
        if m:
            raw_comp = m.group(1).strip()
            # 여러 줄에 걸친 혼용율 처리 (다음 키워드 전까지)
            extra = re.split(r'\n(?:Price|Weight|Finish|Width)\s*:', block[m.end():], 1, re.I)
            if extra:
                continuation = extra[0].strip()
                # 첫 줄이 짧고 다음 줄이 이어지는 경우
                next_lines = continuation.split('\n')
                for line in next_lines[:3]:
                    line = line.strip()
                    if line and not re.match(r'(Price|Weight|Finish|Width)\s*:', line, re.I):
                        if len(raw_comp) < 120:
                            raw_comp += " " + line
                    else:
                        break
            entry.composition = re.sub(r'\s+', ' ', raw_comp).strip()

        # Weight
        m = _WEIGHT_RE.search(block)
        if m:
            entry.weight_gsm = float(m.group(1))
        else:
            dm = _DENIER_RE.search(entry.composition)
            if dm:
                entry.denier = float(dm.group(1))

        # Finish
        m = _FINISH_RE.search(block)
        if m:
            entry.finish = m.group(1).strip()

        # RIB / YOKO 판정
        upper = (entry.composition + " " + block[:200]).upper()
        if any(k in upper for k in ["RIB", "요꼬", "YOKO", "YOKOH"]):
            entry.is_rib = True
            if any(k in upper for k in ["요꼬", "YOKO", "YOKOH", "NECK RIB", "COLLAR RIB"]):
                entry.is_yoko = True

        # 원단 타입 힌트
        for hint in ["우븐", "다이마루", "스웨터", "데님", "WOVEN", "SWEATER", "DENIM"]:
            if hint in upper:
                entry.fabric_type_hint = hint
                break

        if entry.composition:
            result.fabrics.append(entry)

    # 지퍼 — 전체 텍스트에서 패턴 검색
    for m in _ZIPPER_RE.finditer(text):
        spec = m.group(0)
        brand    = (m.group(1) or "").upper()
        mat_raw  = re.sub(r'\s+', '', m.group(2)).upper()
        size_num = m.group(3)
        z_type   = m.group(4).upper()

        material = _normalize_material(mat_raw, spec)
        zipper = ZipperEntry(
            raw_spec=spec,
            brand=brand,
            material=material,
            size=f"#{size_num}",
            zipper_type=_normalize_zipper_type(z_type),
        )
        result.zippers.append(zipper)

    # 아트웍 — EMB / 자수 / PRINT 등 키워드 검색
    upper_text = text.upper()
    if any(k in upper_text for k in ["EMBROIDERY", "EMB ", "자수", "PRINT", "전사", "DTP"]):
        _detect_artwork_from_text(text, result)


# ── 공통 핸들러 ───────────────────────────────────────────────
def _handle_fabric_row(name_cell: str, spec: str, result: TechPackData) -> None:
    entry = FabricEntry(raw_name=name_cell)

    m = _COMPOSITION_RE.search(spec)
    if not m:
        m = _COMPOSITION_INLINE_RE.search(spec)
    if m:
        entry.composition = re.sub(r'\s+', ' ', m.group(1)).strip()

    m = _WEIGHT_RE.search(spec)
    if m:
        entry.weight_gsm = float(m.group(1))
    else:
        dm = _DENIER_RE.search(spec)
        if dm:
            entry.denier = float(dm.group(1))

    m = _FINISH_RE.search(spec)
    if m:
        entry.finish = m.group(1).strip()

    upper = (spec + " " + name_cell).upper()
    if any(k in upper for k in ["RIB", "요꼬", "YOKO", "YOKOH"]):
        entry.is_rib = True
        if any(k in upper for k in ["요꼬", "YOKO", "YOKOH", "NECK RIB", "COLLAR RIB"]):
            entry.is_yoko = True

    for hint in ["우븐", "다이마루", "스웨터", "데님", "WOVEN", "SWEATER", "DENIM"]:
        if hint in upper:
            entry.fabric_type_hint = hint
            break

    if entry.composition or entry.weight_gsm:
        result.fabrics.append(entry)


def _handle_accessory_row(name_cell: str, spec: str, result: TechPackData) -> None:
    m = _ZIPPER_RE.search(spec) or _ZIPPER_RE.search(name_cell)
    if m:
        brand    = (m.group(1) or "").upper()
        mat_raw  = re.sub(r'\s+', '', m.group(2)).upper()
        size_num = m.group(3)
        z_type   = m.group(4).upper()
        material = _normalize_material(mat_raw, spec)
        zipper = ZipperEntry(
            raw_spec=spec.strip(),
            brand=brand,
            material=material,
            size=f"#{size_num}",
            zipper_type=_normalize_zipper_type(z_type),
        )
        result.zippers.append(zipper)
        return

    acc = AccessoryEntry(raw_spec=spec.strip())
    result.accessories.append(acc)


def _handle_artwork_row(name_cell: str, spec: str, result: TechPackData) -> None:
    art = ArtworkEntry(raw_spec=spec.strip())
    upper = (spec + " " + name_cell).upper()
    art.artwork_hint = _classify_artwork_hint(upper)
    result.artworks.append(art)


def _detect_artwork_from_text(text: str, result: TechPackData) -> None:
    upper = text.upper()
    hint = _classify_artwork_hint(upper)
    if hint:
        result.artworks.append(ArtworkEntry(raw_spec="(텍스트 감지)", artwork_hint=hint))


def _classify_artwork_hint(upper: str) -> str:
    # 크기 체크
    size_m = _ARTWORK_SIZE_RE.search(upper)
    big = size_m and float(size_m.group(1)) >= 39

    if any(k in upper for k in ["볼륨", "VOLUME", "패치", "PATCH", "1T EMB"]):
        hint = "EMB-HEAVY"
    elif any(k in upper for k in ["EMB", "자수", "EMBROIDERY", "EMBOSS"]):
        hint = "EMB-LIGHT"
    elif any(k in upper for k in ["전사", "TRANSFER", "REFLECTIVE", "3M REFLECTIVE", "컷팅", "CUTTING"]):
        hint = "PRINT-FILM"
    elif any(k in upper for k in ["DTP", "실리콘PRINT", "SILICONE PRINT", "SCREEN PRINT", "다이렉트", "INK PRINT"]):
        hint = "PRINT-INK"
    elif any(k in upper for k in ["PRINT", "프린트"]):
        hint = "PRINT-INK"
    else:
        hint = "EMB-LIGHT"

    return hint + ("+BIG" if big else "")


# ── 정규화 헬퍼 ───────────────────────────────────────────────
def _normalize_material(mat_raw: str, spec: str) -> str:
    mat = mat_raw.upper().replace(" ", "")
    spec_u = spec.upper()
    if "METALUX" in mat or "METALUX" in spec_u:
        return "METALUX"
    if "METAL" in mat:
        return "METAL"
    if "VSWR" in mat or ("VS" in mat and "WR" in spec_u and "VSWR" not in spec_u and "NYWR" not in spec_u):
        return "VSWR"
    if "NYWR" in mat or ("NY" in mat and "WR" in spec_u):
        return "NYWR"
    if "VS" in mat or "VISLON" in mat:
        return "VS"
    return "NY"


def _normalize_zipper_type(z: str) -> str:
    z = z.upper().replace("-", "").replace("/", "").replace(" ", "")
    if "2WAY" in z or "TWAY" in z:
        return "2WAY"
    if "CE" in z or "CLOSE" in z:
        return "CE"
    return "OE"
