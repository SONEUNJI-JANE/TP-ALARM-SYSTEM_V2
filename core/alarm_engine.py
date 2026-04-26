"""
알람 엔진.
분류된 코드 조합을 경고알람기준 38개 룰과 대조하여 AlarmResult 리스트를 반환한다.
"""
from __future__ import annotations
import re
from dataclasses import dataclass

from .excel_loader import AlarmRule, ExcelData
from .classifier import ClassifiedPack, ClassifiedFabric, ClassifiedZipper


# ── 결과 타입 ─────────────────────────────────────────────────
@dataclass
class AlarmResult:
    section: str        # 지퍼 / 봉제 / 부자재 / 소재물성 / 아트웍
    risk_name: str
    alarm_msg: str      # ⚠️ [...] or ✅ [...]
    checklist: str
    mechanism: str
    combination: str    # 조합 설명 (표시용)
    severity: str       # HIGH / MED / LOW / INFO
    trigger_keyword: str = "" # 실제 문서 내 하이라이트할 원본 텍스트

    @property
    def is_warning(self) -> bool:
        return "⚠️" in self.alarm_msg

    @property
    def is_info(self) -> bool:
        return "✅" in self.alarm_msg


# 심각도 분류
_HIGH_RISKS  = {"박지 주의", "이염 주의", "원단 데미지", "카라 고위험", "기능성 매칭", "열변색 주의", "원단 미어짐", "이염"}
_MED_RISKS   = {"파동 주의", "자수 퍼커링", "코팅-금속 오염", "화학 주의", "구조 주의", "안착 불량",
                "넥 변형", "외관불량", "중량 격차", "프린트 크랙", "봉제 부하", "지퍼 꿀렁임", "오염 및 벗겨짐", "고신축-접착 꿀렁임", "폴리 승화 이염"}
_LOW_RISKS   = {"축률 주의", "물성 주의", "공정 주의", "우븐-테이프 축률", "심실링 변형",
                "넥/에리 안착 불량", "외관 불량", "초경량 하중 경고", "하중 경고",
                "신축 차이 파동", "복합 파동", "강도 불일치", "신축 단차", "물성 상성 주의",
                "핸드필 저하", "자수 파묻힘"}


def _severity(risk_name: str) -> str:
    if risk_name in _HIGH_RISKS:
        return "HIGH"
    if risk_name in _MED_RISKS:
        return "MED"
    if risk_name in _LOW_RISKS:
        return "LOW"
    return "INFO"


# ── 진입점 ────────────────────────────────────────────────────
def check_alarms(pack: ClassifiedPack, excel_data: ExcelData) -> list[AlarmResult]:
    results: list[AlarmResult] = []
    rules_by_section = _group_rules(excel_data.alarm_rules)

    for fabric in pack.fabrics:
        if fabric.entry.is_rib:
            # RIB 원단: 외관 불량 체크는 항상, 넥 알람은 요꼬(is_yoko)일 때만
            _check_rib_alarms(fabric, pack, rules_by_section.get("부자재", []), results)
            continue

        # 지퍼 알람
        for zipper in pack.zippers:
            _check_zipper(fabric, zipper, rules_by_section.get("지퍼", []), results)

        # 봉제 + 부자재 알람
        _check_seam_and_accessories(fabric, pack, rules_by_section, results)

        # 아트웍 알람
        for artwork in pack.artworks:
            _check_artwork(fabric, artwork, rules_by_section.get("아트웍", []), results)

    # 소재물성 알람: 원단끼리 조합
    if len(pack.fabrics) >= 2:
        _check_material_props(pack.fabrics, rules_by_section.get("소재물성", []), results)

    # 중복 제거: 동일 alarm_msg는 한 번만 표시
    seen: set[str] = set()
    unique: list[AlarmResult] = []
    for r in results:
        if r.alarm_msg not in seen:
            seen.add(r.alarm_msg)
            unique.append(r)

    unique.sort(key=lambda r: {"HIGH": 0, "MED": 1, "LOW": 2, "INFO": 3}[r.severity])
    return unique


# ── 지퍼 알람 ─────────────────────────────────────────────────
def _check_zipper(fabric: ClassifiedFabric, zipper: ClassifiedZipper,
                  rules: list[AlarmRule], results: list[AlarmResult]) -> None:
    fc = fabric.composite_code
    zc = zipper.zipper_code
    blend = fabric.blend_code
    finish = fabric.finish_code
    weight = fabric.weight_lv
    combo = f"{fc} + {zc}"

    for rule in rules:
        cond = rule.condition
        fired = False

        # (NX/CX/ST/P) + (METAL/VS/NY) — 파동 주의
        if "NX/CX/ST/P" in cond or "파동" in rule.risk_name or "꿀렁임" in rule.risk_name:
            # ST here = standard finish fabric (비신축), 원단구분상 P = 폴리
            # 실제로는 NX, CX 또는 P(폴리), ST(스웨터 타입)
            if blend in ("NX", "CX", "P") or fabric.fabric_type == "S":
                if any(m in zc for m in ("METAL", "VS", "NY")):
                    fired = True

        # Lv 1 + (METAL/VS/NY) — 박지 주의
        elif "Lv 1" in cond and ("박지" in rule.risk_name or "미어짐" in rule.risk_name):
            if weight == "Lv 1" and any(m in zc for m in ("METAL", "VS", "NY")):
                fired = True

        # (D/Indigo) + VS — 이염 주의
        elif "D/Indigo" in cond or "이염" in rule.risk_name:
            if fabric.fabric_type == "D" or "INDIGO" in fabric.entry.composition.upper():
                if "VS" in zc:
                    fired = True

        # WP + Non-WR — 기능성 매칭
        elif "WP" in cond and "Non-WR" in cond:
            if finish == "WP" and "WR" not in zc:
                fired = True

        if fired:
            results.append(AlarmResult(
                section=rule.section,
                risk_name=rule.risk_name,
                alarm_msg=rule.alarm_msg,
                checklist=rule.checklist,
                mechanism=rule.mechanism,
                combination=combo,
                severity=_severity(rule.risk_name),
                trigger_keyword=zipper.entry.raw_spec[:15] if zipper.entry.raw_spec else "Zipper"
            ))


# ── 봉제 + 부자재 알람 ────────────────────────────────────────
def _check_seam_and_accessories(fabric: ClassifiedFabric, pack: ClassifiedPack,
                                 rules_by_section: dict, results: list[AlarmResult]) -> None:
    fc      = fabric.composite_code
    blend   = fabric.blend_code
    finish  = fabric.finish_code
    ftype   = fabric.fabric_type
    weight  = fabric.weight_lv
    stretch = fabric.stretch_lv

    # 부자재 코드 풀셋
    all_acc_codes: set[str] = set()
    for acc in pack.accessories:
        all_acc_codes.update(acc.acc_codes)
    # 지퍼도 부자재로 포함
    for z in pack.zippers:
        all_acc_codes.add("ZIP")

    has_tape = bool(all_acc_codes & {"TAP"})

    for section in ("봉제", "부자재"):
        for rule in rules_by_section.get(section, []):
            cond = rule.condition
            fired = False
            combo = fc

            # ── 봉제 섹션 ─────────────────────────────────────
            # WP / ZIP,SNP — 코팅-금속 오염
            if ("WP" in cond and ("ZIP" in cond or "SNP" in cond)) and "코팅" in rule.risk_name:
                if finish == "WP" and bool(all_acc_codes & {"ZIP", "SNP"}):
                    fired = True
                    combo += f" + 금속 부자재"

            # Lv 1 / ZIP,SNP,STR — 초경량 하중 (Lv1 = 중량 기준)
            elif "Lv 1" in cond and ("ZIP" in cond or "SNP" in cond or "STR" in cond) and "하중" in rule.risk_name:
                if weight == "Lv 1" and bool(all_acc_codes & {"ZIP", "SNP", "STR"}):
                    fired = True
                    combo += " + 금속/스트링 부자재"

            # W / TAP,PIP — 우븐-테이프 축률
            elif re.match(r'^W\s*/', cond) and ("TAP" in cond or "PIP" in cond):
                if ftype == "W" and has_tape:
                    fired = True
                    combo += " + TAPE/PIPING"

            # NX / FNC — 고신축-접착 (봉제)
            elif re.match(r'^NX\s*/', cond) and "FNC" in cond:
                if blend == "NX" and "FNC" in all_acc_codes:
                    fired = True
                    combo += " + 기능부속(3M/실리콘)"

            # P / FNC,STR — 폴리 승화 이염 (봉제)
            elif re.match(r'^P\s*/', cond) and ("FNC" in cond or "STR" in cond):
                if blend == "P" and bool(all_acc_codes & {"FNC", "STR"}):
                    fired = True
                    combo += " + 팁/라벨"

            # Lv 4-5 / TAP,PIP — 신축 차이 파동 (stretch level 4-5 기준)
            elif "Lv 4-5" in cond and ("TAP" in cond or "PIP" in cond):
                if stretch in ("Lv 4", "Lv 5") and has_tape:
                    fired = True
                    combo += " + TAPE/PIPING"

            # WP / SEA — 심실링 변형
            elif re.match(r'^WP\s*/', cond) and "SEA" in cond:
                if finish == "WP" and "SEA" in all_acc_codes:
                    fired = True
                    combo += " + 심실링"

            # RIB,YOKO / NK — 넥/에리 안착 불량 (요꼬 넥 전용 — 일반 RIB 밑단 제외)
            elif "RIB" in cond and "YOKO" in cond and "NK" in cond:
                has_yoko = any(f.entry.is_yoko for f in pack.fabrics)
                if has_yoko and section == "봉제":
                    fired = True
                    combo = "RIB/요꼬 + 몸판 넥"

            # ── 부자재 섹션 ────────────────────────────────────
            # WP + ZIP/SNP — 화학 주의
            elif re.match(r'^WP\s*\+', cond) and ("ZIP" in cond or "SNP" in cond):
                if finish == "WP" and bool(all_acc_codes & {"ZIP", "SNP"}):
                    fired = True
                    combo += " + 지퍼/스냅"

            # Lv 1 + ZIP/SNP — 하중 경고 (중량 기준)
            elif "Lv 1" in cond and ("ZIP" in cond or "SNP" in cond) and "하중" in rule.risk_name:
                if weight == "Lv 1" and bool(all_acc_codes & {"ZIP", "SNP"}):
                    fired = True
                    combo += " + 지퍼/스냅"

            # W + TAP/PIP — 축률 주의
            elif re.match(r'^W\s*\+', cond) and ("TAP" in cond or "PIP" in cond):
                if ftype == "W" and has_tape:
                    fired = True
                    combo += " + TAPE/PIPING"

            # Lv 4/5 + TAP/PIP — 물성 주의 (stretch level 4-5 기준)
            elif "Lv 4/5" in cond and ("TAP" in cond or "PIP" in cond):
                if stretch in ("Lv 4", "Lv 5") and has_tape:
                    fired = True
                    combo += " + TAPE/PIPING"

            # WP + SEA — 공정 주의
            elif re.match(r'^WP\s*\+', cond) and "SEA" in cond:
                if finish == "WP" and "SEA" in all_acc_codes:
                    fired = True
                    combo += " + 심실링"

            # EBD + PIP/BND — 복합 파동
            elif "EBD" in cond and ("PIP" in cond or "BND" in cond):
                if bool(all_acc_codes & {"EBD"}) and has_tape:
                    fired = True
                    combo += " + 밴드+파이핑"

            # YOKO + NK-SELF — 안착 불량 (요꼬 넥 전용 — 일반 RIB 밑단은 해당 없음)
            elif "YOKO" in cond and "NK-SELF" in cond:
                has_yoko = any(f.entry.is_yoko for f in pack.fabrics)
                if has_yoko:
                    fired = True
                    combo = "요꼬 + 자가원단 넥"

            # NK + Lv 4/5 — 외관불량 (요꼬 넥 전용)
            elif re.match(r'^NK\s*\+', cond) and ("Lv 4" in cond or "Lv 5" in cond):
                has_yoko = any(f.entry.is_yoko for f in pack.fabrics)
                if has_yoko and stretch in ("Lv 4", "Lv 5"):
                    fired = True

            # RIB / YOKO — 외관 불량 (요꼬 넥 전용 — 밑단/소맷단 일반 RIB 제외)
            elif re.match(r'^RIB\s*/', cond) and "YOKO" in cond:
                has_yoko = any(f.entry.is_yoko for f in pack.fabrics)
                if has_yoko and section == "부자재":
                    fired = True
                    combo = "요꼬/넥 립 몸판 연결부"

            # Lv 1 + STR — 중량 격차 (중량 기준)
            elif "Lv 1" in cond and "STR" in cond and "중량" in rule.risk_name:
                if weight == "Lv 1" and "STR" in all_acc_codes:
                    fired = True
                    combo += " + 스트링/팁"

            # ST + Lv 5 — 봉제 부하 (중량 Lv5 + 신축 없는 원단)
            elif "ST" in cond and "Lv 5" in cond and "봉제" in rule.risk_name:
                if weight == "Lv 5" and stretch == "Lv 1":
                    fired = True

            # P + FNC/STR — 이염 주의 (부자재)
            elif re.match(r'^P\s*\+', cond) and ("FNC" in cond or "STR" in cond):
                if blend == "P" and bool(all_acc_codes & {"FNC", "STR"}):
                    fired = True
                    combo += " + 팁/부자재"

            # NX + FNC — 구조 주의 (부자재)
            elif re.match(r'^NX\s*\+', cond) and "FNC" in cond:
                if blend == "NX" and "FNC" in all_acc_codes:
                    fired = True
                    combo += " + 기능부속"

            if fired:
                tk = fabric.entry.composition[:15] if fabric.entry.composition else fabric.display_name
                results.append(AlarmResult(
                    section=rule.section,
                    risk_name=rule.risk_name,
                    alarm_msg=rule.alarm_msg,
                    checklist=rule.checklist,
                    mechanism=rule.mechanism,
                    combination=combo,
                    severity=_severity(rule.risk_name),
                    trigger_keyword=tk
                ))


# ── RIB 전용 알람 ─────────────────────────────────────────────
def _check_rib_alarms(rib_fabric: ClassifiedFabric, _pack: ClassifiedPack,
                      rules: list[AlarmRule], results: list[AlarmResult]) -> None:
    """RIB/요꼬 원단 알람.
    - 일반 RIB(밑단/소맷단): 외관 불량(RIB/YOKO) 체크만
    - 요꼬(is_yoko, 넥 전용): 넥 관련 안착 불량까지 추가
    """
    is_yoko = rib_fabric.entry.is_yoko
    for rule in rules:
        cond = rule.condition
        fired = False

        if ("RIB" in cond and "YOKO" in cond) or "NK-SELF" in cond:
            # 외관 불량 / 안착 불량 모두 요꼬(넥 전용 립) 일 때만
            # 일반 RIB(밑단/소맷단)는 이 알람 기준에 해당하지 않음
            fired = is_yoko

        if fired:
            loc = "요꼬/넥" if is_yoko else "RIB 밑단/소맷단"
            tk = rib_fabric.entry.composition[:15] if rib_fabric.entry.composition else rib_fabric.display_name
            results.append(AlarmResult(
                section=rule.section,
                risk_name=rule.risk_name,
                alarm_msg=rule.alarm_msg,
                checklist=rule.checklist,
                mechanism=rule.mechanism,
                combination=f"{loc} ({rib_fabric.composite_code})",
                severity=_severity(rule.risk_name),
                trigger_keyword=tk
            ))


# ── 소재물성 알람 ─────────────────────────────────────────────
def _check_material_props(fabrics: list[ClassifiedFabric],
                          rules: list[AlarmRule],
                          results: list[AlarmResult]) -> None:
    """원단끼리의 조합 (T=메인, W=배색/부원단)."""
    main_fabrics = [f for f in fabrics if not f.entry.is_rib]
    sub_fabrics  = [f for f in fabrics if f.entry.is_rib]
    if not main_fabrics:
        return

    main = main_fabrics[0]
    mt = main.fabric_type
    mb = main.blend_code
    mw = main.weight_lv

    for rule in rules:
        cond = rule.condition
        fired = False
        combo = ""

        # NYLON(T) + LEATHER(W)
        if "NYLON(T)" in cond and "LEATHER(W)" in cond:
            if mb == "N" or mb == "NX":
                # 레더 배색은 판정 어려움 — 조건 미발동 (없으면 스킵)
                pass

        # NYLON(T) + POLY(W) — ✅ 정상 가이드
        elif "NYLON(T)" in cond and "POLY(W)" in cond:
            for sf in sub_fabrics:
                if sf.blend_code == "P" and (mb in ("N", "NX")):
                    fired = True
                    combo = f"{main.composite_code} + {sf.composite_code} (배색)"

        # W/P(T) + COTTON(W)
        elif "W/P(T)" in cond and "COTTON(W)" in cond:
            if mt in ("W",) or mb == "P":
                for sf in sub_fabrics:
                    if sf.blend_code in ("C", "CS", "CX"):
                        fired = True
                        combo = f"{main.composite_code} + {sf.composite_code}"

        # WOVEN(T) + RIB/KNIT(W)
        elif "WOVEN(T)" in cond and ("RIB" in cond or "KNIT" in cond):
            if mt == "W" and sub_fabrics:
                fired = True
                combo = f"{main.composite_code} + RIB/요꼬"

        # Lv 1(T) + HEAVY(W)
        elif "Lv 1(T)" in cond and "HEAVY(W)" in cond:
            if mw == "Lv 1":
                for sf in sub_fabrics:
                    if sf.weight_lv in ("Lv 4", "Lv 5"):
                        fired = True
                        combo = f"{main.composite_code} + HEAVY 배색"

        # CX/C + ETC
        elif ("CX" in cond or "C +" in cond) and "ETC" in cond:
            if mb in ("C", "CX"):
                fired = True
                combo = f"{main.composite_code} + ETC 배색"

        # D + ETC
        elif cond.startswith("D +") and "ETC" in cond:
            if mt == "D":
                fired = True
                combo = f"{main.composite_code} + ETC 배색"

        if fired:
            tk = main.entry.composition[:15] if main.entry.composition else main.display_name
            results.append(AlarmResult(
                section=rule.section,
                risk_name=rule.risk_name,
                alarm_msg=rule.alarm_msg,
                checklist=rule.checklist,
                mechanism=rule.mechanism,
                combination=combo,
                severity=_severity(rule.risk_name),
                trigger_keyword=tk
            ))


# ── 아트웍 알람 ───────────────────────────────────────────────
def _check_artwork(fabric: ClassifiedFabric, artwork,
                   rules: list[AlarmRule], results: list[AlarmResult]) -> None:
    fc   = fabric.composite_code
    ac   = artwork.artwork_code
    is_big = artwork.is_big
    blend  = fabric.blend_code
    weight = fabric.weight_lv

    for rule in rules:
        cond = rule.condition
        fired = False
        combo = f"{fc} + {ac}"

        # SIZE 39CM↑ / BIG — 핸드필 저하
        if "39CM" in cond or "BIG" in cond:
            if is_big:
                fired = True

        # ST / NX + EMB — 자수 퍼커링  (ST = 스판 stretch, NX = 나일론 스판)
        elif "NX" in cond and "EMB" in cond:
            has_stretch = (blend == "NX") or (blend in ("CX", "CS") and fabric.stretch_lv != "Lv 1")
            if has_stretch and "EMB" in ac:
                fired = True

        # PRINT-FILM + N — 열변색 주의
        elif "PRINT-FILM" in cond and "+ N" in cond:
            if ac == "PRINT-FILM" and blend in ("N", "NX"):
                fired = True

        # PRINT-INK + Lv 3↑ — 프린트 크랙
        elif "PRINT-INK" in cond and "Lv 3" in cond:
            if ac == "PRINT-INK" and fabric.stretch_lv in ("Lv 3", "Lv 4", "Lv 5"):
                fired = True

        # EMB-HEAVY + Lv 1 — 원단 데미지
        elif "EMB-HEAVY" in cond and "Lv 1" in cond:
            if ac == "EMB-HEAVY" and weight == "Lv 1":
                fired = True

        # EMB-LIGHT + Lv 5 — 자수 파묻힘
        elif "EMB-LIGHT" in cond and "Lv 5" in cond:
            if ac == "EMB-LIGHT" and weight == "Lv 5":
                fired = True

        if fired:
            tk = artwork.artwork_code.replace("+BIG", "")
            results.append(AlarmResult(
                section=rule.section,
                risk_name=rule.risk_name,
                alarm_msg=rule.alarm_msg,
                checklist=rule.checklist,
                mechanism=rule.mechanism,
                combination=combo,
                severity=_severity(rule.risk_name),
                trigger_keyword=tk
            ))


# ── 헬퍼 ──────────────────────────────────────────────────────
def _group_rules(rules: list[AlarmRule]) -> dict[str, list[AlarmRule]]:
    groups: dict[str, list[AlarmRule]] = {}
    for r in rules:
        groups.setdefault(r.section, []).append(r)
    return groups
