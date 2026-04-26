"""
추출된 원자재 텍스트를 분류 코드로 변환한다.
복합 원단 코드: [타입]-[혼용율]-[후가공]-[스트레치]-[중량]  예) K-NX-WR-Lv 3-Lv 2
지퍼 코드: [소재]-[#사이즈]-[타입]  예) VS-#3-CE
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field

from .excel_loader import (
    ExcelData,
    classify_blend,
    classify_finish,
    classify_stretch,
    classify_weight,
    classify_fabric_type,
)
from .pdf_parser import FabricEntry, ZipperEntry, AccessoryEntry, ArtworkEntry


# ── 결과 데이터클래스 ──────────────────────────────────────────
@dataclass
class ClassifiedFabric:
    entry: FabricEntry
    fabric_type: str    # W / K / S / D
    blend_code: str     # CS / NX / N / C / CX / P / S
    finish_code: str    # WP / WR / AF / FN / ST
    stretch_lv: str     # Lv 1 ~ Lv 5
    weight_lv: str      # Lv 1 ~ Lv 5
    composite_code: str # K-NX-WR-Lv 3-Lv 2

    @property
    def display_name(self) -> str:
        name = self.entry.raw_name.strip()
        # 숫자 prefix + 깨진 한글 제거 → 영문 부분만 추출
        parts = re.findall(r'[A-Z0-9][A-Z0-9_\-\./ ]+', name.upper())
        return " ".join(parts).strip() or name[:30]


@dataclass
class ClassifiedZipper:
    entry: ZipperEntry
    zipper_code: str    # VS-#3-CE

    @property
    def display_name(self) -> str:
        return f"{self.entry.material} #{self.entry.size.replace('#','')} {self.entry.zipper_type}"


@dataclass
class ClassifiedAccessory:
    entry: AccessoryEntry
    acc_codes: list[str]  # [ZIP, SNP, TAP, ...] 복수 가능


@dataclass
class ClassifiedArtwork:
    entry: ArtworkEntry
    artwork_code: str   # EMB-HEAVY / EMB-LIGHT / PRINT-FILM / PRINT-INK
    is_big: bool = False  # 39cm 이상 대형


@dataclass
class ClassifiedPack:
    style_code: str
    fabrics: list[ClassifiedFabric] = field(default_factory=list)
    zippers: list[ClassifiedZipper] = field(default_factory=list)
    accessories: list[ClassifiedAccessory] = field(default_factory=list)
    artworks: list[ClassifiedArtwork] = field(default_factory=list)


# ── 부자재 분류 키워드 맵 ──────────────────────────────────────
_ACC_KEYWORD_MAP: list[tuple[str, list[str]]] = [
    ("SEA",  ["SEALING", "SEAL", "WELD", "심실링"]),
    ("TAP",  ["TAPE", "테이프", "WAKI", "MOBILON", "BINDING", "PIPING", "PIP"]),
    ("STR",  ["STRING", "STR", "DRAWSTRING", "CORD", "ELASTIC", "끈", "스트링", "TIP", "팁"]),
    ("EBD",  ["BAND", "밴드", "ELASTIC BAND"]),
    ("FNC",  ["STOPPER", "CORDLOCK", "코드락", "3M", "SILICONE", "실리콘", "REFLECTIVE"]),
    ("SNP",  ["SNAP", "BUTTON", "BTN", "스냅", "버튼"]),
    ("LBL",  ["LABEL", "라벨", "TAG", "태그"]),
]


# ── 진입점 ────────────────────────────────────────────────────
def classify_pack(style_code: str,
                  fabrics: list[FabricEntry],
                  zippers: list[ZipperEntry],
                  accessories: list[AccessoryEntry],
                  artworks: list[ArtworkEntry],
                  excel_data: ExcelData) -> ClassifiedPack:
    result = ClassifiedPack(style_code=style_code)

    for f in fabrics:
        cf = _classify_fabric(f, excel_data)
        result.fabrics.append(cf)

    for z in zippers:
        result.zippers.append(_classify_zipper(z))

    for acc in accessories:
        result.accessories.append(_classify_accessory(acc))

    for art in artworks:
        result.artworks.append(_classify_artwork(art))

    return result


# ── 원단 분류 ─────────────────────────────────────────────────
def _classify_fabric(f: FabricEntry, excel_data: ExcelData) -> ClassifiedFabric:
    # 타입: PDF의 fabric_type_hint가 있으면 사용, 없으면 K(기본값)
    if f.fabric_type_hint:
        ftype = classify_fabric_type(f.fabric_type_hint)
    elif f.is_rib:
        ftype = "K"
    else:
        ftype = "K"

    blend   = classify_blend(f.composition)
    finish  = classify_finish(f.finish) if f.finish else "ST"
    stretch = _classify_stretch_smart(f.composition, blend)
    weight  = classify_weight(f.weight_gsm, f.denier, excel_data.weight_levels)

    code = f"{ftype}-{blend}-{finish}-{stretch}-{weight}"
    return ClassifiedFabric(
        entry=f,
        fabric_type=ftype,
        blend_code=blend,
        finish_code=finish,
        stretch_lv=stretch,
        weight_lv=weight,
        composite_code=code,
    )


def _classify_stretch_smart(composition: str, blend_code: str) -> str:
    """Excel 데이터 패턴 기반: NX 계열은 Lv 3 고정 (실측 데이터 기준)."""
    if blend_code == "NX":
        return "Lv 3"
    if blend_code in ("CS", "CX"):
        # 스판 함량이 있으면 Lv 3
        if re.search(r'\b\d+\s*%\s*(SPANDEX|SPAN|ELASTANE)', composition, re.I):
            return "Lv 3"
        return "Lv 2"
    return classify_stretch(composition)


# ── 지퍼 분류 ─────────────────────────────────────────────────
def _classify_zipper(z: ZipperEntry) -> ClassifiedZipper:
    code = f"{z.material}-{z.size}-{z.zipper_type}"
    return ClassifiedZipper(entry=z, zipper_code=code)


# ── 부자재 분류 ───────────────────────────────────────────────
def _classify_accessory(acc: AccessoryEntry) -> ClassifiedAccessory:
    upper = acc.raw_spec.upper() + " " + acc.category_hint.upper()
    codes: list[str] = []
    for code, keywords in _ACC_KEYWORD_MAP:
        if any(k in upper for k in keywords):
            codes.append(code)
    if not codes:
        codes = ["ACC"]
    return ClassifiedAccessory(entry=acc, acc_codes=codes)


# ── 아트웍 분류 ───────────────────────────────────────────────
def _classify_artwork(art: ArtworkEntry) -> ClassifiedArtwork:
    hint = art.artwork_hint or "EMB-LIGHT"
    is_big = "+BIG" in hint
    code = hint.replace("+BIG", "")
    return ClassifiedArtwork(entry=art, artwork_code=code, is_big=is_big)
