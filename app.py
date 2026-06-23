"""
선택과목 이동반 편성 시스템 — Streamlit 웹앱
실행: streamlit run app.py
"""
import io
import random
from collections import defaultdict
from itertools import permutations

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import pulp
    HAS_PULP = True
except ImportError:
    HAS_PULP = False

# ──────────────────────────────────────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────────────────────────────────────

NON_SUBJECT_COLS = {
    '순번', '번호', '신학번', '학번', '이름', '성명', '학년', '반', '번', 'no', 'id',
    '성별', '남', '여', 'gender', 'sex',
}


def get_id_col(df: pd.DataFrame) -> str:
    for c in ['신학번', '학번', '번호', 'ID', 'id']:
        if c in df.columns:
            return c
    return df.columns[0]


def detect_gender_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        if c.lower() in ['성별', 'gender', 'sex', '남녀', '성별']:
            return c
    return None


def detect_subjects(df: pd.DataFrame) -> list[str]:
    result = []
    for col in df.columns:
        if col.lower() in NON_SUBJECT_COLS:
            continue
        vals = set(df[col].dropna().astype(str).str.strip().unique())
        vals.discard('')
        if vals <= {'0', '1', '0.0', '1.0', '0.', '1.'}:
            result.append(col)
    return result


def extract_homeroom(sid: str) -> str:
    s = str(sid).strip()
    if len(s) == 5 and s.isdigit():
        return str(int(s[1:3])) + '반'
    return ''


def get_time_labels(n: int) -> list[str]:
    return [chr(65 + i) for i in range(n)]


def abbr(name: str, length: int = 5) -> str:
    return name[:length]


# ──────────────────────────────────────────────────────────────────────────────
# 과목 그룹 유틸
# ──────────────────────────────────────────────────────────────────────────────

def validate_groups(groups_df: pd.DataFrame, subjects_auto: list[str]) -> tuple[list[dict], list[str]]:
    """
    groups_df: 그룹명 / 과목목록(쉼표구분) / 선택수 컬럼
    반환: (그룹 목록, 경고 메시지 목록)
    """
    groups  = []
    warns   = []
    all_sub = set(subjects_auto)

    for _, row in groups_df.dropna(subset=['그룹명']).iterrows():
        gname   = str(row['그룹명']).strip()
        raw_sub = str(row.get('과목목록 (쉼표로 구분)', '')).strip()
        try:
            n_pick = int(row.get('선택수', 1))
        except (ValueError, TypeError):
            n_pick = 1

        subs = [s.strip() for s in raw_sub.split(',') if s.strip()]
        unknown = [s for s in subs if s not in all_sub]
        if unknown:
            warns.append(f"그룹 '{gname}': CSV에 없는 과목 → {unknown}")
        subs = [s for s in subs if s in all_sub]

        if not subs:
            warns.append(f"그룹 '{gname}': 유효한 과목이 없어 건너뜁니다.")
            continue
        if n_pick > len(subs):
            warns.append(f"그룹 '{gname}': 선택수({n_pick}) > 과목수({len(subs)}) → 선택수를 {len(subs)}로 조정")
            n_pick = len(subs)

        groups.append({'name': gname, 'subjects': subs, 'n_pick': n_pick})

    return groups, warns


def groups_to_subject_map(groups: list[dict]) -> dict[str, str]:
    """과목명 → 그룹명 매핑"""
    return {s: g['name'] for g in groups for s in g['subjects']}


def total_picks_from_groups(groups: list[dict]) -> int:
    return sum(g['n_pick'] for g in groups)


def validate_student_picks(row: dict, groups: list[dict], subjects: list[str]) -> bool:
    """
    그룹 구조가 있을 때: 각 그룹에서 정확히 n_pick개를 선택했는지 확인
    그룹 구조가 없을 때: 단순히 len(times)개 선택 확인
    """
    for g in groups:
        picked = sum(int(row.get(s, 0)) for s in g['subjects'] if s in subjects)
        if picked != g['n_pick']:
            return False
    return True


# ──────────────────────────────────────────────────────────────────────────────
# 교우관계 분리 조건
# ──────────────────────────────────────────────────────────────────────────────

def parse_separation_pairs(df_sep: pd.DataFrame) -> set[frozenset]:
    pairs = set()
    for _, row in df_sep.iterrows():
        a = str(row.iloc[0]).strip()
        b = str(row.iloc[1]).strip()
        if a and b and a != b and a != 'nan' and b != 'nan':
            pairs.add(frozenset([a, b]))
    return pairs


def check_separation_violations(
    assignments_dict: dict, sep_pairs: set[frozenset], label: str = ""
) -> list[dict]:
    violations = []
    for pair in sep_pairs:
        pl = list(pair)
        if len(pl) < 2:
            continue
        a, b = pl[0], pl[1]
        ga, gb = assignments_dict.get(a), assignments_dict.get(b)
        if ga is not None and gb is not None and ga == gb:
            violations.append({'학번A': a, '학번B': b, '배정위치': str(ga), '구분': label})
    return violations


# ──────────────────────────────────────────────────────────────────────────────
# ILP — 그룹 내 과목 분산 제약 포함
# ──────────────────────────────────────────────────────────────────────────────

def run_ilp(subjects, n_sections, max_per_time, times, groups=None):
    """
    groups: [{'name':..., 'subjects':[...], 'n_pick':...}]
    그룹이 있으면: 같은 그룹 과목들이 같은 타임에 몰리지 않도록 제약 추가
      → 그룹 내 과목별 타임당 배정 분반 수 합계 ≤ floor(그룹총분반/타임수)+1
    """
    groups = groups or []
    total  = sum(n_sections[s] for s in subjects)
    nt     = len(times)
    lo, hi = total // nt, (total + nt - 1) // nt

    prob = pulp.LpProblem("section_time", pulp.LpMinimize)
    z = {
        s: {t: pulp.LpVariable(f"z_{i}_{t}", 0, n_sections[s], cat='Integer')
            for t in times}
        for i, s in enumerate(subjects)
    }

    # 목적함수: 그룹 내 타임별 집중도 최소화 (분산 유도)
    group_spread_cost = []
    for g in groups:
        gsubs = [s for s in g['subjects'] if s in subjects]
        for t in times:
            # 같은 타임에 같은 그룹 과목이 많을수록 페널티
            group_spread_cost.append(
                pulp.lpSum(z[s][t] for s in gsubs)
            )
    if group_spread_cost:
        prob += pulp.lpSum(group_spread_cost)
    else:
        prob += 0

    # 과목별 총 분반 수
    for s in subjects:
        prob += pulp.lpSum(z[s][t] for t in times) == n_sections[s]
        for t in times:
            prob += z[s][t] <= max_per_time[s]

    # 타임별 총 분반 균형
    for t in times:
        col_sum = pulp.lpSum(z[s][t] for s in subjects)
        prob += col_sum >= lo
        prob += col_sum <= hi

    # 그룹 분산 제약: 같은 타임에 그룹 내 과목 분반 합 ≤ ceil
    for g in groups:
        gsubs      = [s for s in g['subjects'] if s in subjects]
        g_total    = sum(n_sections[s] for s in gsubs)
        g_lo, g_hi = g_total // nt, (g_total + nt - 1) // nt
        for t in times:
            prob += pulp.lpSum(z[s][t] for s in gsubs) <= g_hi + 1  # 약간의 여유

    status = prob.solve(pulp.PULP_CBC_CMD(msg=0))
    if pulp.LpStatus[prob.status] not in ('Optimal', 'Not Solved'):
        return None, pulp.LpStatus[prob.status]
    if pulp.LpStatus[prob.status] == 'Not Solved':
        # 제약 완화 후 재시도
        prob2 = pulp.LpProblem("section_time_relaxed", pulp.LpMinimize)
        z2 = {
            s: {t: pulp.LpVariable(f"z2_{i}_{t}", 0, n_sections[s], cat='Integer')
                for t in times}
            for i, s in enumerate(subjects)
        }
        prob2 += 0
        for s in subjects:
            prob2 += pulp.lpSum(z2[s][t] for t in times) == n_sections[s]
            for t in times:
                prob2 += z2[s][t] <= max_per_time[s]
        for t in times:
            col_sum2 = pulp.lpSum(z2[s][t] for s in subjects)
            prob2 += col_sum2 >= lo
            prob2 += col_sum2 <= hi
        status2 = prob2.solve(pulp.PULP_CBC_CMD(msg=0))
        if pulp.LpStatus[prob2.status] != 'Optimal':
            return None, pulp.LpStatus[prob2.status]
        result = {s: {t: int(round(pulp.value(z2[s][t]))) for t in times} for s in subjects}
        return result, 'Optimal(relaxed)'

    result = {s: {t: int(round(pulp.value(z[s][t]))) for t in times} for s in subjects}
    return result, 'Optimal'


def build_sections(subjects, assignment, times):
    sections = {}
    for s in subjects:
        sections[s] = {}
        num = 1
        for t in times:
            cnt = assignment[s][t]
            sections[s][t] = list(range(num, num + cnt))
            num += cnt
    return sections


# ──────────────────────────────────────────────────────────────────────────────
# 학생 배정 — 그룹 구조 + 우선과목 + 분리 조건 반영
# ──────────────────────────────────────────────────────────────────────────────

def assign_students(
    df, id_col, subjects, assignment, sections,
    times, max_per_section,
    homeroom=None, sep_pairs=None,
    groups=None,           # 그룹 구조
    priority_map=None,     # {과목: 우선순위점수} — 높을수록 분반 수 우선 배정
    seed=42,
):
    """
    그룹 구조가 있을 때:
      - 학생이 선택한 과목을 그룹별로 분류 → 그룹의 타임 슬롯에 배정
      - priority_map: 분반 수 배정 시 우선과목을 먼저 배정 (greedy)
    분리 조건: sep_pairs 쌍은 같은 분반 금지 (페널티 방식)
    """
    homeroom     = homeroom or {}
    sep_pairs    = sep_pairs or set()
    groups       = groups or []
    priority_map = priority_map or {}
    random.seed(seed)

    use_groups = len(groups) > 0
    n_sel      = len(times)  # 그룹 없을 때: 타임 수 = 선택 과목 수

    sep_neighbors: dict[str, set[str]] = defaultdict(set)
    for pair in sep_pairs:
        pl = list(pair)
        sep_neighbors[pl[0]].add(pl[1])
        sep_neighbors[pl[1]].add(pl[0])

    sec_load   = {(s, k): 0 for s in subjects for t in times for k in sections[s][t]}
    sec_roster = defaultdict(list)
    records    = []
    failed     = []

    rows = df.to_dict('records')
    random.shuffle(rows)

    # 분리 쌍 학생 먼저
    sep_sids   = {sid for pair in sep_pairs for sid in pair}
    rows_sep   = [r for r in rows if str(r.get(id_col, '')) in sep_sids]
    rows_other = [r for r in rows if str(r.get(id_col, '')) not in sep_sids]
    rows       = rows_sep + rows_other

    sid_sec_map: dict[str, dict[str, int]] = {}

    def _sep_penalty(sid, s, k):
        for nb in sep_neighbors.get(sid, set()):
            if sid_sec_map.get(nb, {}).get(s) == k:
                return 1000
        return 0

    for row in rows:
        sid = str(row.get(id_col, ''))

        # 선택 과목 추출
        chosen = [s for s in subjects if int(row.get(s, 0)) == 1]

        # 유효성 검사
        if use_groups:
            if not validate_student_picks(row, groups, subjects):
                failed.append(sid)
                continue
        else:
            if len(chosen) != n_sel:
                failed.append(sid)
                continue

        # 타임 배정: 그룹별로 슬롯 할당
        if use_groups:
            # 그룹별 선택 과목 목록
            group_chosen: dict[str, list[str]] = {}
            for g in groups:
                gpicks = [s for s in g['subjects'] if int(row.get(s, 0)) == 1]
                group_chosen[g['name']] = gpicks

            # 그룹 순서대로 타임 슬롯 할당 (우선순위 높은 그룹부터)
            # 각 그룹의 n_pick개 과목에 서로 다른 타임 슬롯 배정
            group_order = sorted(
                groups,
                key=lambda g: -sum(priority_map.get(s, 0) for s in g['subjects'])
            )

            # 타임 슬롯 풀
            available_times = list(times)

            best_full_map: dict[str, str] | None = None  # 과목 → 타임
            best_full_cost = float('inf')

            def _try_assign_groups(g_idx, time_pool, current_map):
                nonlocal best_full_map, best_full_cost
                if g_idx == len(group_order):
                    # 전체 배정 완료 — 비용 계산
                    sep_c = sum(
                        sum(_sep_penalty(sid, s, k) for k in sections[s][current_map[s]])
                        for s in current_map if assignment[s][current_map[s]] > 0
                    )
                    load_c = sum(
                        min(sec_load[(s, k)] for k in sections[s][current_map[s]])
                        for s in current_map if assignment[s][current_map[s]] > 0
                    )
                    cost = sep_c * 10 + load_c
                    if cost < best_full_cost:
                        best_full_cost  = cost
                        best_full_map   = dict(current_map)
                    return

                g     = group_order[g_idx]
                gpicks = group_chosen.get(g['name'], [])
                if not gpicks:
                    _try_assign_groups(g_idx + 1, time_pool, current_map)
                    return

                # 이 그룹의 과목들에 타임 슬롯 배정 (순열)
                needed = len(gpicks)
                if needed > len(time_pool):
                    return  # 슬롯 부족

                for perm in permutations(time_pool, needed):
                    tm = dict(zip(gpicks, perm))
                    if not all(assignment[s][t] > 0 for s, t in tm.items()):
                        continue
                    new_pool = [t for t in time_pool if t not in perm]
                    new_map  = {**current_map, **tm}
                    _try_assign_groups(g_idx + 1, new_pool, new_map)

            _try_assign_groups(0, available_times, {})
            final_map = best_full_map

        else:
            # 기존 로직: permutations
            best_map, best_cost     = None, float('inf')
            overflow_map, overflow_cost = None, float('inf')

            for perm in permutations(times):
                tm = dict(zip(chosen, perm))
                if not all(assignment[s][t] > 0 for s, t in tm.items()):
                    continue
                sep_c  = sum(sum(_sep_penalty(sid, s, k) for k in sections[s][tm[s]]) for s in chosen)
                load_c = sum(min(sec_load[(s, k)] for k in sections[s][tm[s]]) for s in chosen)
                cost   = sep_c * 10 + load_c
                has_room = all(
                    any(sec_load[(s, k)] < max_per_section for k in sections[s][tm[s]])
                    for s in chosen
                )
                if has_room and cost < best_cost:
                    best_cost, best_map = cost, tm
                if cost < overflow_cost:
                    overflow_cost, overflow_map = cost, tm

            final_map = best_map if best_map is not None else overflow_map

        if final_map is None:
            failed.append(sid)
            continue

        rec = {id_col: sid, '본반': homeroom.get(sid, '')}
        for t in times:
            rec[f'{t}타임_과목'] = ''
            rec[f'{t}타임_분반'] = ''

        sid_sec_map[sid] = {}
        for s, t in final_map.items():
            if assignment[s][t] == 0:
                continue

            def pool_score(k, _s=s, _sid=sid):
                sep_v = sum(1 for nb in sep_neighbors.get(_sid, set())
                            if sid_sec_map.get(nb, {}).get(_s) == k)
                return (sep_v * 1000, sec_load[(_s, k)])

            with_room = [k for k in sections[s][t] if sec_load[(s, k)] < max_per_section]
            pool      = with_room if with_room else sections[s][t]
            k         = min(pool, key=pool_score)

            sec_load[(s, k)] += 1
            sec_roster[(s, k)].append(sid)
            sid_sec_map[sid][s] = k
            rec[f'{t}타임_과목'] = s
            rec[f'{t}타임_분반'] = f"{abbr(s)}-{k}반"

        records.append(rec)

    rec_df    = pd.DataFrame(records) if records else pd.DataFrame(columns=[id_col])
    base      = df[[id_col]].copy()
    base[id_col] = base[id_col].astype(str)
    result_df = base.merge(rec_df, on=id_col, how='left') if not rec_df.empty else base

    return result_df, sec_roster, sec_load, failed


# ──────────────────────────────────────────────────────────────────────────────
# 본반 편성
# ──────────────────────────────────────────────────────────────────────────────

def get_subject_pattern(row, subjects):
    return tuple(s for s in subjects if int(row.get(s, 0)) == 1)


def assign_homeroom_classes(
    df, id_col, subjects, gender_col,
    n_classes, max_per_class,
    gender_ratio_weight=0.5,
    sep_pairs=None,
    seed=42,
):
    sep_pairs = sep_pairs or set()
    random.seed(seed)
    rows = df.to_dict('records')

    def norm_gender(v):
        v = str(v).strip()
        if v in ('남', 'M', 'm', '1', 'male', 'Male'): return '남'
        if v in ('여', 'F', 'f', '2', 'female', 'Female'): return '여'
        return '미상'

    sep_neighbors: dict[str, set[str]] = defaultdict(set)
    for pair in sep_pairs:
        pl = list(pair)
        sep_neighbors[pl[0]].add(pl[1])
        sep_neighbors[pl[1]].add(pl[0])

    pattern_groups = defaultdict(list)
    for row in rows:
        sid    = str(row.get(id_col, ''))
        pat    = get_subject_pattern(row, subjects)
        gender = norm_gender(row.get(gender_col, '미상')) if gender_col else '미상'
        pattern_groups[pat].append({'sid': sid, 'gender': gender, 'pattern': pat})

    total_m        = sum(1 for r in rows if norm_gender(r.get(gender_col, '')) == '남') if gender_col else 0
    total          = len(rows)
    target_m_ratio = total_m / total if total else 0.5

    class_load     = [0] * n_classes
    class_males    = [0] * n_classes
    class_patterns = [defaultdict(int) for _ in range(n_classes)]
    class_sids     = [set() for _ in range(n_classes)]
    assignments    = {}

    for pat, students in sorted(pattern_groups.items(), key=lambda x: -len(x[1])):
        random.shuffle(students)
        for stu in students:
            sid, gender = stu['sid'], stu['gender']
            candidates = [i for i in range(n_classes) if class_load[i] < max_per_class] or list(range(n_classes))

            def score(i):
                sep_pen = sum(1 for nb in sep_neighbors.get(sid, set()) if nb in class_sids[i]) * 10000
                pat_bon = -class_patterns[i][pat] * 10
                load    = class_load[i]
                if load > 0:
                    nr = (class_males[i] + (1 if gender == '남' else 0)) / (load + 1)
                    gender_s = abs(nr - target_m_ratio) * gender_ratio_weight * 100
                else:
                    gender_s = 0
                return sep_pen + pat_bon + gender_s + class_load[i]

            best = min(candidates, key=score)
            assignments[sid] = best + 1
            class_load[best] += 1
            class_patterns[best][pat] += 1
            class_sids[best].add(sid)
            if gender == '남':
                class_males[best] += 1

    records = []
    for row in rows:
        sid    = str(row.get(id_col, ''))
        gender = norm_gender(row.get(gender_col, '')) if gender_col else '미상'
        pat    = get_subject_pattern(row, subjects)
        cls    = assignments.get(sid)
        records.append({
            id_col: sid, '배정_본반': f"{cls}반" if cls else '',
            '성별': gender, '과목패턴': ' / '.join(pat) if pat else '',
        })

    result_df = pd.DataFrame(records)
    stats = []
    for i in range(n_classes):
        n = class_load[i]; m = class_males[i]
        top = ', '.join(
            f"{'·'.join(p)}({c}명)"
            for p, c in sorted(class_patterns[i].items(), key=lambda x: -x[1])[:3]
        )
        stats.append({'반': f"{i+1}반", '총원': n, '남': m, '여': n-m,
                      '남비율': f"{100*m/n:.1f}%" if n else '0%', '주요 과목패턴': top})

    stats_df = pd.DataFrame(stats)
    total_students = len(rows)
    same_pat = sum(c*(c-1)//2 for pd_ in class_patterns for c in pd_.values())
    total_pairs = total_students * (total_students - 1) // 2
    cohesion = same_pat / total_pairs if total_pairs else 0

    return result_df, stats_df, cohesion, assignments


# ──────────────────────────────────────────────────────────────────────────────
# Excel 출력
# ──────────────────────────────────────────────────────────────────────────────

_H_FILL  = PatternFill("solid", fgColor="2E4699")
_H_FONT  = Font(bold=True, color="FFFFFF", size=10)
_H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
_S_FILL  = PatternFill("solid", fgColor="5B9BD5")
_W_FILL  = PatternFill("solid", fgColor="FFD7D7")
_G_FILLS = [
    PatternFill("solid", fgColor="E8F4FD"),
    PatternFill("solid", fgColor="FEF9E7"),
    PatternFill("solid", fgColor="EAFAF1"),
    PatternFill("solid", fgColor="F9EBEA"),
    PatternFill("solid", fgColor="F4ECF7"),
]


def _hdr(ws, r, c, v):
    cell = ws.cell(r, c, v)
    cell.font, cell.fill, cell.alignment = _H_FONT, _H_FILL, _H_ALIGN
    return cell


def _subhdr(ws, r, c, v):
    cell = ws.cell(r, c, v)
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.fill = _S_FILL
    cell.alignment = _H_ALIGN
    return cell


def _auto_width(ws, mn=8, mx=42):
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, mn), mx)


def build_excel(
    subjects, times, id_col, assignment, sections,
    sec_load, sec_roster, result_df, homeroom,
    n_total, failed, max_per_section,
    sep_violations=None, groups=None,
):
    sep_violations = sep_violations or []
    groups         = groups or []
    wb = Workbook()
    wb.remove(wb.active)

    # 개요
    ws = wb.create_sheet("개요")
    ws.cell(1, 1, f"이동반 편성 결과 — 총 {n_total}명 / 배정 {n_total-len(failed)}명 / 미배정 {len(failed)}명")
    ws.cell(1, 1).font = Font(bold=True, size=12)
    ws.merge_cells('A1:G1')

    # 그룹 정보 표시
    if groups:
        ws.cell(2, 1, "과목 그룹: " + " | ".join(
            f"{g['name']}({','.join(g['subjects'])}) {g['n_pick']}선택" for g in groups
        ))
        ws.cell(2, 1).font = Font(italic=True, color="555555", size=9)
        ws.merge_cells('A2:G2')

    r = 4
    for ci, h in enumerate(['타임', '과목', '그룹', '분반명', '배정인원', '최대인원', '상태'], 1):
        _hdr(ws, r, ci, h)
    r += 1

    sub_to_group = groups_to_subject_map(groups)
    group_names  = [g['name'] for g in groups]
    g_fill_map   = {g['name']: _G_FILLS[i % len(_G_FILLS)] for i, g in enumerate(groups)}

    for t in times:
        _subhdr(ws, r, 1, f"▶ {t}타임")
        ws.merge_cells(f'A{r}:G{r}')
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt   = sec_load.get((s, k), 0)
                state = "★초과" if cnt > max_per_section else "정상"
                gname = sub_to_group.get(s, '')
                fill  = g_fill_map.get(gname)
                for ci, val in enumerate([t, s, gname, f"{abbr(s)}-{k}반", cnt, max_per_section, state], 1):
                    cell = ws.cell(r, ci, val)
                    if fill:
                        cell.fill = fill
                if state != "정상":
                    ws.cell(r, 7).font = Font(bold=True, color="FF0000")
                r += 1
        r += 1
    _auto_width(ws)

    # 학생 시간표
    ws2   = wb.create_sheet("학생시간표")
    cols  = [id_col, '본반'] + [f'{t}타임_과목' for t in times] + [f'{t}타임_분반' for t in times]
    for ci, h in enumerate(cols, 1):
        _hdr(ws2, 1, ci, h)
    ws2.freeze_panes = 'A2'
    for _, row in result_df.iterrows():
        ws2.append([str(row.get(c, '') or '') for c in cols])
    _auto_width(ws2)

    # 분반별 명단
    for s in subjects:
        for t in times:
            for k in sections[s][t]:
                ws_s = wb.create_sheet(f"{abbr(s)}-{k}반")
                gname = sub_to_group.get(s, '')
                ws_s.cell(1, 1, f"[{s}] {k}분반 ({t}타임)" + (f" — {gname}" if gname else ''))
                ws_s.cell(1, 1).font = Font(bold=True, size=11)
                ws_s.merge_cells('A1:C1')
                for ci, h in enumerate([id_col, '본반', '비고'], 1):
                    _hdr(ws_s, 2, ci, h)
                for sid in sorted(sec_roster.get((s, k), [])):
                    ws_s.append([sid, homeroom.get(sid, ''), ''])
                _auto_width(ws_s)

    # 교실배치
    ws_cr = wb.create_sheet("교실배치추천")
    ws_cr.cell(1, 1, "교실 배치 추천").font = Font(bold=True, size=12)
    ws_cr.merge_cells('A1:F1')
    r = 3
    for t in times:
        _subhdr(ws_cr, r, 1, f"▶ {t}타임")
        ws_cr.merge_cells(f'A{r}:F{r}')
        r += 1
        for ci, h in enumerate(['과목', '그룹', '분반명', '배정인원', '본반 구성', '권장교실'], 1):
            _hdr(ws_cr, r, ci, h)
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt    = sec_load.get((s, k), 0)
                hr_cnt = defaultdict(int)
                for sid in sec_roster.get((s, k), []):
                    hr = homeroom.get(sid, '')
                    if hr:
                        hr_cnt[hr] += 1
                top      = sorted(hr_cnt.items(), key=lambda x: (-x[1], x[0]))
                priority = '  >  '.join(f"{hr} {c}명" for hr, c in top) if top else ''
                ws_cr.cell(r, 1, s)
                ws_cr.cell(r, 2, sub_to_group.get(s, ''))
                ws_cr.cell(r, 3, f"{abbr(s)}-{k}반")
                ws_cr.cell(r, 4, cnt)
                ws_cr.cell(r, 5, priority).alignment = Alignment(horizontal='left')
                r += 1
        r += 1
    ws_cr.column_dimensions['A'].width = 22
    ws_cr.column_dimensions['B'].width = 12
    ws_cr.column_dimensions['C'].width = 14
    ws_cr.column_dimensions['D'].width = 10
    ws_cr.column_dimensions['E'].width = 60

    # 분리 위반
    if sep_violations:
        ws_v = wb.create_sheet("⚠️분리위반")
        ws_v.cell(1, 1, f"분리 조건 위반 — {len(sep_violations)}건").font = Font(bold=True, size=12, color="CC0000")
        ws_v.merge_cells('A1:D1')
        for ci, h in enumerate(['학번A', '학번B', '배정위치', '구분'], 1):
            _hdr(ws_v, 2, ci, h)
        for viol in sep_violations:
            ws_v.append([viol.get('학번A',''), viol.get('학번B',''),
                         viol.get('배정위치',''), viol.get('구분','')])
            for ci in range(1, 5):
                ws_v.cell(ws_v.max_row, ci).fill = _W_FILL
        _auto_width(ws_v)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_homeroom_excel(hr_result_df, hr_stats_df, id_col, n_classes,
                          assignments, df, gender_col, subjects, sep_violations=None):
    sep_violations = sep_violations or []
    wb = Workbook()
    wb.remove(wb.active)

    def hdr(ws, r, c, v):
        cell = ws.cell(r, c, v)
        cell.font  = Font(bold=True, color="FFFFFF", size=10)
        cell.fill  = PatternFill("solid", fgColor="2E4699")
        cell.alignment = Alignment(horizontal='center', vertical='center')

    def aw(ws):
        for col in ws.columns:
            w = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w+2,8),42)

    ws = wb.create_sheet("본반편성_개요")
    ws.cell(1, 1, "본반 편성 결과 요약").font = Font(bold=True, size=13)
    ws.merge_cells('A1:F1')
    r = 3
    for ci, h in enumerate(['반', '총원', '남', '여', '남비율', '주요 과목패턴'], 1):
        hdr(ws, r, ci, h)
    r += 1
    for _, row in hr_stats_df.iterrows():
        ws.append([row['반'], row['총원'], row['남'], row['여'], row['남비율'], row['주요 과목패턴']])
    aw(ws)

    ws2 = wb.create_sheet("본반편성_학생목록")
    for ci, h in enumerate([id_col, '배정_본반', '성별', '과목패턴'], 1):
        hdr(ws2, 1, ci, h)
    ws2.freeze_panes = 'A2'
    for _, row in hr_result_df.iterrows():
        ws2.append([str(row.get(c, '') or '') for c in [id_col, '배정_본반', '성별', '과목패턴']])
    aw(ws2)

    def ng(v):
        v = str(v).strip()
        if v in ('남', 'M', 'm', '1', 'male'): return '남'
        if v in ('여', 'F', 'f', '2', 'female'): return '여'
        return '미상'

    for cls_num in range(1, n_classes + 1):
        ws_c = wb.create_sheet(f"{cls_num}반")
        ws_c.cell(1, 1, f"{cls_num}반 학생 명단").font = Font(bold=True, size=12)
        ws_c.merge_cells('A1:D1')
        for ci, h in enumerate([id_col, '성별', '과목패턴', '비고'], 1):
            hdr(ws_c, 2, ci, h)
        class_stu = hr_result_df[hr_result_df['배정_본반'] == f"{cls_num}반"]
        for _, row in class_stu.sort_values(id_col).iterrows():
            ws_c.append([str(row[id_col]), row['성별'], row['과목패턴'], ''])
        aw(ws_c)

    if sep_violations:
        ws_v = wb.create_sheet("⚠️분리위반")
        ws_v.cell(1, 1, f"분리 조건 위반 — {len(sep_violations)}건").font = Font(bold=True, size=12, color="CC0000")
        ws_v.merge_cells('A1:D1')
        for ci, h in enumerate(['학번A', '학번B', '배정위치', '구분'], 1):
            hdr(ws_v, 2, ci, h)
        for viol in sep_violations:
            ws_v.append([viol.get('학번A',''), viol.get('학번B',''),
                         viol.get('배정위치',''), viol.get('구분','')])
            for ci in range(1, 5):
                ws_v.cell(ws_v.max_row, ci).fill = PatternFill("solid", fgColor="FFD7D7")
        aw(ws_v)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ──────────────────────────────────────────────────────────────────────────────
# 예시 파일 생성 함수
# ──────────────────────────────────────────────────────────────────────────────

def make_example_student_csv() -> bytes:
    """학생 선택과목 CSV 예시"""
    df = pd.DataFrame({
        '신학번': ['10101', '10102', '10103', '10104', '10105'],
        '성별':   ['남',    '여',    '남',    '여',    '남'],
        '수학Ⅰ':  [1, 0, 1, 0, 1],
        '수학Ⅱ':  [0, 1, 0, 1, 0],
        '영어Ⅰ':  [1, 1, 0, 0, 1],
        '영어Ⅱ':  [0, 0, 1, 1, 0],
        '음악':   [1, 0, 1, 0, 0],
        '미술':   [0, 1, 0, 1, 1],
        '윤리':   [1, 1, 1, 1, 1],
        '철학':   [0, 0, 0, 0, 0],
    })
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


def make_example_student_excel() -> bytes:
    """학생 선택과목 Excel 예시"""
    wb = Workbook()
    ws = wb.active
    ws.title = "학생선택과목"

    H_FILL  = PatternFill("solid", fgColor="2E4699")
    H_FONT  = Font(bold=True, color="FFFFFF", size=10)
    H_ALIGN = Alignment(horizontal='center', vertical='center')
    NOTE_FONT = Font(color="888888", italic=True, size=9)

    # 안내 행
    ws.cell(1, 1, "※ 신학번·성별 컬럼 + 과목명 컬럼(0/1). 첫 행은 헤더.")
    ws.cell(1, 1).font = NOTE_FONT
    ws.merge_cells('A1:J1')

    headers = ['신학번', '성별', '수학Ⅰ', '수학Ⅱ', '영어Ⅰ', '영어Ⅱ', '음악', '미술', '윤리', '철학']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(2, ci, h)
        cell.font, cell.fill, cell.alignment = H_FONT, H_FILL, H_ALIGN

    data = [
        ['10101','남',1,0,1,0,1,0,1,0],
        ['10102','여',0,1,1,0,0,1,1,0],
        ['10103','남',1,0,0,1,1,0,1,0],
        ['10104','여',0,1,0,1,0,1,1,0],
        ['10105','남',1,0,1,0,0,1,1,0],
    ]
    for row in data:
        ws.append(row)

    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w+2, 8), 20)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


def make_example_settings_csv(subjects: list[str] | None = None) -> bytes:
    """과목 설정 CSV 예시"""
    subs = subjects or ['수학Ⅰ', '수학Ⅱ', '영어Ⅰ', '영어Ⅱ', '음악', '미술', '윤리']
    df = pd.DataFrame({
        '과목명':            subs,
        '학급수(총 분반 수)': [3] * len(subs),
        '교사수(타임당 최대)': [1] * len(subs),
        '우선순위':          list(range(len(subs), 0, -1)),
    })
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


def make_example_settings_excel(subjects: list[str] | None = None) -> bytes:
    """과목 설정 Excel 예시"""
    subs = subjects or ['수학Ⅰ', '수학Ⅱ', '영어Ⅰ', '영어Ⅱ', '음악', '미술', '윤리']
    wb = Workbook()
    ws = wb.active
    ws.title = "과목설정"

    H_FILL  = PatternFill("solid", fgColor="2E4699")
    H_FONT  = Font(bold=True, color="FFFFFF", size=10)
    H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    NOTE_FONT = Font(color="888888", italic=True, size=9)

    ws.cell(1, 1, "※ 과목명·학급수·교사수·우선순위 컬럼. 과목명은 학생CSV와 정확히 일치해야 합니다.")
    ws.cell(1, 1).font = NOTE_FONT
    ws.merge_cells('A1:D1')

    headers = ['과목명', '학급수(총 분반 수)', '교사수(타임당 최대)', '우선순위']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(2, ci, h)
        cell.font, cell.fill, cell.alignment = H_FONT, H_FILL, H_ALIGN
    ws.row_dimensions[2].height = 30

    for i, s in enumerate(subs):
        ws.append([s, 3, 1, len(subs)-i])

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 10

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


def make_example_groups_csv() -> bytes:
    """과목 그룹 CSV 예시"""
    df = pd.DataFrame({
        '그룹명':               ['교양', '예술', '교과'],
        '과목목록 (쉼표로 구분)': ['윤리, 철학', '음악, 미술', '수학Ⅰ, 수학Ⅱ, 영어Ⅰ, 영어Ⅱ'],
        '선택수':               [1, 1, 3],
    })
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


def make_example_groups_excel() -> bytes:
    """과목 그룹 Excel 예시"""
    wb = Workbook()
    ws = wb.active
    ws.title = "과목그룹"

    H_FILL  = PatternFill("solid", fgColor="2E4699")
    H_FONT  = Font(bold=True, color="FFFFFF", size=10)
    H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    NOTE_FONT = Font(color="888888", italic=True, size=9)

    ws.cell(1, 1, "※ 그룹명 / 과목목록(쉼표구분) / 선택수. 선택수 합계 = 타임 수.")
    ws.cell(1, 1).font = NOTE_FONT
    ws.merge_cells('A1:C1')

    for ci, h in enumerate(['그룹명', '과목목록 (쉼표로 구분)', '선택수'], 1):
        cell = ws.cell(2, ci, h)
        cell.font, cell.fill, cell.alignment = H_FONT, H_FILL, H_ALIGN
    ws.row_dimensions[2].height = 28

    data = [
        ['교양', '윤리, 철학', 1],
        ['예술', '음악, 미술', 1],
        ['교과', '수학Ⅰ, 수학Ⅱ, 영어Ⅰ, 영어Ⅱ', 3],
    ]
    for row in data:
        ws.append(row)

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 10

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


def make_example_separation_csv() -> bytes:
    """분리 조건 CSV 예시"""
    df = pd.DataFrame({'학번A': ['10101', '10203'], '학번B': ['10205', '10418']})
    return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')


def make_example_separation_excel() -> bytes:
    """분리 조건 Excel 예시"""
    wb = Workbook()
    ws = wb.active
    ws.title = "분리조건"

    H_FILL  = PatternFill("solid", fgColor="C0392B")
    H_FONT  = Font(bold=True, color="FFFFFF", size=10)
    H_ALIGN = Alignment(horizontal='center', vertical='center')
    NOTE_FONT = Font(color="888888", italic=True, size=9)

    ws.cell(1, 1, "※ 학번A / 학번B 두 열. 이 두 학생은 본반·이동반 모두에서 분리됩니다.")
    ws.cell(1, 1).font = NOTE_FONT
    ws.merge_cells('A1:B1')

    for ci, h in enumerate(['학번A', '학번B'], 1):
        cell = ws.cell(2, ci, h)
        cell.font, cell.fill, cell.alignment = H_FONT, H_FILL, H_ALIGN

    for row in [['10101', '10205'], ['10203', '10418']]:
        ws.append(row)

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 14

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


def read_uploaded_file(f, dtype=str) -> pd.DataFrame:
    """CSV / Excel 모두 읽기"""
    name = f.name.lower()
    if name.endswith('.xlsx') or name.endswith('.xls'):
        return pd.read_excel(f, dtype=dtype)
    else:
        for enc in ['utf-8-sig', 'utf-8', 'cp949', 'euc-kr']:
            try:
                f.seek(0)
                return pd.read_csv(f, encoding=enc, dtype=dtype)
            except UnicodeDecodeError:
                continue
        f.seek(0)
        return pd.read_csv(f, dtype=dtype)


# ──────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="이동반 편성 시스템", layout="wide", page_icon="🏫")
st.title("🏫 선택과목 이동반 편성 시스템")
st.caption("CSV 업로드 → 과목 그룹 설정 → 분리 조건 → 편성 → 결과")

if not HAS_PULP:
    st.error("⚠️ PuLP 라이브러리 없음. `pip install pulp` 후 재시작하세요.")
    st.stop()

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 기본 설정")
    n_times = st.number_input("타임 수", min_value=1, max_value=10, value=5,
                               help="과목 그룹을 사용할 때는 '그룹별 선택수 합계'와 일치해야 합니다")
    times   = get_time_labels(n_times)
    st.info(f"타임: **{' / '.join(times)}**")
    max_per_section = st.number_input("분반당 최대 인원", min_value=10, max_value=100, value=41)
    seed = int(st.number_input("무작위 시드", min_value=0, max_value=9999, value=42))
    st.divider()
    st.caption("※ 설정 변경 후 편성 시작을 다시 눌러주세요.")

# ── 탭 ───────────────────────────────────────────────────────────────────────
tab1, tab_grp, tab_sep, tab2, tab3, tab4 = st.tabs([
    "① CSV 업로드",
    "② 과목 그룹",
    "③ 분리 조건",
    "④ 과목 설정 및 편성",
    "⑤ 편성 결과",
    "⑥ 본반 편성",
])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1: CSV 업로드
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("학생 선택과목 파일 업로드")
    st.markdown("""
    **파일 형식:** `신학번`(또는 `학번`) + `성별` + 과목명 컬럼들  
    과목 선택 여부: `1`(선택) / `0`(미선택) | 성별: `남`/`여` (또는 M/F, 1/2)  
    **지원 형식:** CSV, Excel(.xlsx)
    """)

    with st.expander("📥 예시 파일 다운로드"):
        st.caption("아래 예시 파일을 참고해 업로드 파일을 준비하세요.")
        ex1, ex2 = st.columns(2)
        with ex1:
            st.download_button("📄 예시 CSV", data=make_example_student_csv(),
                               file_name="학생선택과목_예시.csv", mime="text/csv",
                               use_container_width=True)
        with ex2:
            st.download_button("📊 예시 Excel", data=make_example_student_excel(),
                               file_name="학생선택과목_예시.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

    uploaded = st.file_uploader("파일 선택 (CSV 또는 Excel)", type=['csv','xlsx','xls'],
                                label_visibility='collapsed')

    if uploaded:
        try:
            raw_df = read_uploaded_file(uploaded, dtype=str)
            raw_df.columns = raw_df.columns.str.strip()
            for col in raw_df.columns:
                raw_df[col] = raw_df[col].str.strip()

            subjects_auto = detect_subjects(raw_df)
            id_col        = get_id_col(raw_df)
            gender_col    = detect_gender_col(raw_df)

            for s in subjects_auto:
                raw_df[s] = pd.to_numeric(raw_df[s], errors='coerce').fillna(0).astype(int)

            if not subjects_auto:
                st.error("과목 컬럼을 자동 감지하지 못했습니다.")
                st.stop()

            st.session_state.update({
                'raw_df': raw_df, 'subjects_auto': subjects_auto,
                'id_col': id_col, 'gender_col': gender_col,
                'settings_init': False, 'groups_init': False,
            })

            n_sel   = raw_df[subjects_auto].sum(axis=1)
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("총 학생 수", f"{len(raw_df)}명")
            c2.metric("감지된 과목 수", f"{len(subjects_auto)}개")
            c3.metric("학생당 선택 과목 (평균)", f"{n_sel.mean():.1f}개")
            bad_cnt = int((n_sel != n_times).sum())
            c4.metric(f"{n_times}과목 미준수", f"{bad_cnt}명",
                      delta_color="inverse" if bad_cnt else "off")

            if gender_col:
                st.success(f"✅ 성별 컬럼 감지: **{gender_col}**")
            else:
                st.warning("⚠️ 성별 컬럼을 찾지 못했습니다.")

            enroll_data = [{'과목': s, '신청 인원': int(raw_df[s].sum()),
                            '비율': f"{100*raw_df[s].mean():.1f}%"}
                           for s in subjects_auto]
            st.subheader("과목별 신청 인원")
            st.dataframe(pd.DataFrame(enroll_data), use_container_width=True, hide_index=True)
            st.bar_chart({r['과목']: r['신청 인원'] for r in enroll_data})
            st.subheader("원본 데이터 미리보기 (상위 10행)")
            st.dataframe(raw_df.head(10), use_container_width=True)

        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")
            st.exception(e)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2: 과목 그룹 설정
# ════════════════════════════════════════════════════════════════════════════════
with tab_grp:
    st.subheader("📚 과목 그룹 설정")
    st.markdown("""
    과목을 그룹으로 묶고 그룹별 선택 수를 지정합니다.

    **예시:**
    | 그룹명 | 과목목록 | 선택수 |
    |--------|----------|--------|
    | 교양 | 윤리와사상, 생활과윤리, 철학 | 1 |
    | 예술 | 음악, 미술, 체육 | 1 |
    | 교과 | 수학Ⅰ, 수학Ⅱ, 영어Ⅰ, ... | 3 |

    → 교양 1타임 + 예술 1타임 + 교과 3타임 = **5타임** (사이드바 타임 수와 일치)

    **그룹 효과:**
    - 🔵 **인원 많은 과목 우선 분반 배정**
    - 🔴 **같은 그룹 과목끼리 다른 타임에 분산** (ILP 제약)
    """)

    use_groups = st.toggle("과목 그룹 사용", value=False,
                           help="끄면 기존 방식(모든 과목 동등)으로 편성됩니다")
    st.session_state['use_groups'] = use_groups

    if use_groups:
        if 'raw_df' not in st.session_state:
            st.info("먼저 [① CSV 업로드] 탭에서 파일을 업로드해주세요.")
        else:
            subjects_auto = st.session_state['subjects_auto']
            raw_df        = st.session_state['raw_df']

            # 예시 다운로드
            with st.expander("📥 예시 파일 다운로드"):
                ex1, ex2 = st.columns(2)
                with ex1:
                    st.download_button("📄 예시 CSV", data=make_example_groups_csv(),
                                       file_name="과목그룹_예시.csv", mime="text/csv",
                                       use_container_width=True)
                with ex2:
                    st.download_button("📊 예시 Excel", data=make_example_groups_excel(),
                                       file_name="과목그룹_예시.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)

            # 파일 업로드
            st.markdown("#### 방법 ① — 파일 업로드 (CSV / Excel)")
            grp_uploaded = st.file_uploader("과목 그룹 파일 선택", type=['csv','xlsx','xls'],
                                            key='grp_uploader', label_visibility='collapsed')
            if grp_uploaded:
                try:
                    grp_file_df = read_uploaded_file(grp_uploaded, dtype=str)
                    grp_file_df.columns = grp_file_df.columns.str.strip()
                    # 필요한 컬럼만 매핑
                    col_map = {}
                    for c in grp_file_df.columns:
                        cl = c.strip().lower()
                        if '그룹' in cl or 'group' in cl:
                            col_map['그룹명'] = c
                        elif '과목' in cl or 'subject' in cl:
                            col_map['과목목록 (쉼표로 구분)'] = c
                        elif '선택' in cl or 'pick' in cl or 'count' in cl or '수' in cl:
                            col_map['선택수'] = c
                    if len(col_map) >= 2:
                        grp_file_df = grp_file_df.rename(columns={v: k for k, v in col_map.items()})
                        st.session_state['groups_df'] = grp_file_df[list(col_map.keys())].copy()
                        st.session_state['groups_init'] = True
                        st.success(f"✅ {len(grp_file_df)}개 그룹 로드됨")
                    else:
                        st.error("그룹명·과목목록·선택수 컬럼을 찾지 못했습니다. 예시 파일 형식을 확인해주세요.")
                except Exception as e:
                    st.error(f"파일 읽기 오류: {e}")

            st.markdown("#### 방법 ② — 직접 입력")
            # 초기 그룹 테이블
            if not st.session_state.get('groups_init'):
                st.session_state['groups_df'] = pd.DataFrame({
                    '그룹명':               [''],
                    '과목목록 (쉼표로 구분)': [', '.join(subjects_auto)],
                    '선택수':               [n_times],
                })
                st.session_state['groups_init'] = True

            st.caption("행 추가(+) / 삭제(행 선택 후 Delete) 가능. 과목명은 CSV와 정확히 일치해야 합니다.")

            edited_grp = st.data_editor(
                st.session_state['groups_df'],
                num_rows='dynamic',
                use_container_width=True,
                height=min(80 + 55 * max(len(st.session_state['groups_df']), 3), 500),
                column_config={
                    '그룹명': st.column_config.TextColumn('그룹명', width='small'),
                    '과목목록 (쉼표로 구분)': st.column_config.TextColumn('과목목록 (쉼표로 구분)', width='large'),
                    '선택수': st.column_config.NumberColumn('선택수', min_value=1, max_value=10, step=1, width='small'),
                },
                key='groups_editor',
            )
            st.session_state['groups_df'] = edited_grp

            # 검증
            parsed_groups, warns = validate_groups(edited_grp, subjects_auto)
            for w in warns:
                st.warning(f"⚠️ {w}")

            if parsed_groups:
                total_picks = total_picks_from_groups(parsed_groups)
                col_g1, col_g2, col_g3 = st.columns(3)
                col_g1.metric("그룹 수", f"{len(parsed_groups)}개")
                col_g2.metric("선택수 합계", f"{total_picks}타임")

                if total_picks == n_times:
                    col_g3.metric("✅ 타임 수 일치", f"{n_times}타임")
                else:
                    col_g3.metric("⚠️ 타임 수 불일치",
                                  f"그룹합계 {total_picks} ≠ 설정 {n_times}",
                                  delta_color="inverse")
                    st.warning(f"사이드바 타임 수를 **{total_picks}**으로 변경하거나, 그룹 선택수 합계를 **{n_times}**으로 맞춰주세요.")

                # 그룹별 인원 현황
                st.subheader("그룹별 과목 신청 현황")
                for g in parsed_groups:
                    with st.expander(f"**{g['name']}** ({g['n_pick']}선택 / {len(g['subjects'])}과목)"):
                        enroll = [(s, int(raw_df[s].sum())) for s in g['subjects']]
                        enroll.sort(key=lambda x: -x[1])
                        g_df = pd.DataFrame(enroll, columns=['과목', '신청인원'])
                        g_df['우선순위'] = range(1, len(g_df)+1)
                        st.dataframe(g_df, use_container_width=True, hide_index=True)
                        st.caption(f"↑ 인원 많은 순 = 분반 배정 우선순위")

                st.session_state['parsed_groups'] = parsed_groups
            else:
                st.session_state['parsed_groups'] = []
    else:
        st.session_state['parsed_groups'] = []
        st.info("그룹을 사용하지 않으면 모든 과목이 동등하게 처리됩니다.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3: 분리 조건
# ════════════════════════════════════════════════════════════════════════════════
with tab_sep:
    st.subheader("🚫 교우관계 분리 조건 설정")
    st.markdown("""
    지정한 두 학생은 **본반·이동반 모두에서 같은 반/분반에 배정되지 않도록** 처리됩니다.
    > 구조적으로 분리 불가능한 경우 위반 내역이 별도로 표시됩니다.
    """)

    st.markdown("#### 방법 ① — 파일 업로드 (CSV / Excel)")

    with st.expander("📥 예시 파일 다운로드"):
        ex1, ex2 = st.columns(2)
        with ex1:
            st.download_button("📄 예시 CSV", data=make_example_separation_csv(),
                               file_name="분리조건_예시.csv", mime="text/csv",
                               use_container_width=True)
        with ex2:
            st.download_button("📊 예시 Excel", data=make_example_separation_excel(),
                               file_name="분리조건_예시.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

    st.markdown("`학번A, 학번B` 두 열 (헤더 유무 자동 감지)")
    sep_uploaded = st.file_uploader("분리 조건 파일 (CSV / Excel)", type=['csv','xlsx','xls'],
                                    key='sep_uploader', label_visibility='collapsed')

    uploaded_pairs: set[frozenset] = set()
    if sep_uploaded:
        try:
            sep_raw = read_uploaded_file(sep_uploaded, dtype=str)
            sep_raw.columns = [str(c).strip() for c in sep_raw.columns]
            # 헤더가 학번A/B 형태면 그대로, 아니면 첫 행 검사
            has_header = any(
                str(c).strip().lower() in ['학번a','학번b','a','b','학번','id']
                for c in sep_raw.columns
            )
            if not has_header:
                # 헤더 없는 경우 첫 행이 헤더인지 확인
                first_row = sep_raw.iloc[0].tolist()
                if any(str(v).strip().lower() in ['학번a','학번b','a','b','학번','id'] for v in first_row):
                    sep_raw = sep_raw.iloc[1:].reset_index(drop=True)
            uploaded_pairs = parse_separation_pairs(sep_raw)
            st.success(f"✅ {len(uploaded_pairs)}쌍 로드됨")
            st.dataframe(
                pd.DataFrame([{'학번A': list(p)[0], '학번B': list(p)[1]} for p in uploaded_pairs]),
                use_container_width=True, hide_index=True, height=180,
            )
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")

    st.divider()
    st.markdown("#### 방법 ② — 직접 입력")

    if 'sep_manual_df' not in st.session_state:
        st.session_state['sep_manual_df'] = pd.DataFrame({'학번A': [''], '학번B': ['']})

    edited_sep = st.data_editor(
        st.session_state['sep_manual_df'], num_rows='dynamic',
        use_container_width=True, height=220,
        column_config={
            '학번A': st.column_config.TextColumn('학번A', width='medium'),
            '학번B': st.column_config.TextColumn('학번B', width='medium'),
        },
        key='sep_editor',
    )
    st.session_state['sep_manual_df'] = edited_sep
    manual_pairs = parse_separation_pairs(edited_sep)

    all_sep_pairs = uploaded_pairs | manual_pairs
    st.session_state['sep_pairs'] = all_sep_pairs

    st.divider()
    col_p1, col_p2 = st.columns(2)
    col_p1.metric("CSV 업로드", f"{len(uploaded_pairs)}쌍")
    col_p2.metric("직접 입력", f"{len(manual_pairs)}쌍")

    if all_sep_pairs:
        st.info(f"✅ 총 **{len(all_sep_pairs)}쌍** 적용 예정")
        with st.expander(f"전체 분리 목록 ({len(all_sep_pairs)}쌍)"):
            st.dataframe(
                pd.DataFrame([{'학번A': sorted(list(p))[0], '학번B': sorted(list(p))[1]}
                              for p in all_sep_pairs]).sort_values(['학번A','학번B']),
                use_container_width=True, hide_index=True,
            )
    else:
        st.info("분리 조건 없음 — 조건 없이 편성됩니다.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 4: 과목 설정 및 편성
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    if 'raw_df' not in st.session_state:
        st.info("먼저 [① CSV 업로드] 탭에서 파일을 업로드해주세요.")
    else:
        raw_df        = st.session_state['raw_df']
        subjects_auto = st.session_state['subjects_auto']
        parsed_groups = st.session_state.get('parsed_groups', [])
        use_groups    = st.session_state.get('use_groups', False)

        st.subheader("과목별 분반·교사 수 설정")

        # 예시 + 파일 업로드
        with st.expander("📥 예시 파일 다운로드 / 파일로 불러오기"):
            st.caption("예시 파일을 받아 수정한 뒤 업로드하면 표가 자동으로 채워집니다.")
            cur_subs = st.session_state.get('subjects_auto', None)
            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button("📄 예시 CSV", data=make_example_settings_csv(cur_subs),
                                   file_name="과목설정_예시.csv", mime="text/csv",
                                   use_container_width=True)
            with dl2:
                st.download_button("📊 예시 Excel", data=make_example_settings_excel(cur_subs),
                                   file_name="과목설정_예시.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)

            settings_file = st.file_uploader(
                "과목 설정 파일 업로드 (CSV / Excel)",
                type=['csv','xlsx','xls'], key='settings_uploader',
                label_visibility='collapsed',
            )
            if settings_file:
                try:
                    sf = read_uploaded_file(settings_file, dtype=str)
                    sf.columns = sf.columns.str.strip()
                    # 컬럼 매핑
                    rename = {}
                    for c in sf.columns:
                        cl = c.lower()
                        if '과목' in cl and '명' in cl: rename[c] = '과목명'
                        elif '학급' in cl or '분반' in cl: rename[c] = '학급수(총 분반 수)'
                        elif '교사' in cl: rename[c] = '교사수(타임당 최대)'
                        elif '우선' in cl or 'priority' in cl: rename[c] = '우선순위'
                    sf = sf.rename(columns=rename)
                    needed = ['과목명', '학급수(총 분반 수)', '교사수(타임당 최대)']
                    missing = [c for c in needed if c not in sf.columns]
                    if missing:
                        st.error(f"필수 컬럼 없음: {missing}. 예시 파일 형식을 확인해주세요.")
                    else:
                        for col in ['학급수(총 분반 수)', '교사수(타임당 최대)', '우선순위']:
                            if col in sf.columns:
                                sf[col] = pd.to_numeric(sf[col], errors='coerce').fillna(
                                    1 if '우선' not in col else 0
                                ).astype(int)
                        if '우선순위' not in sf.columns:
                            sf['우선순위'] = 0
                        if '그룹' not in sf.columns:
                            sub_to_g = groups_to_subject_map(parsed_groups)
                            sf['그룹'] = sf['과목명'].apply(
                                lambda s: sub_to_g.get(str(s).strip(), '(미지정)')
                            )
                        st.session_state['settings_df'] = sf[['과목명','학급수(총 분반 수)','교사수(타임당 최대)','그룹','우선순위']]
                        st.session_state['settings_init'] = True
                        st.success(f"✅ {len(sf)}개 과목 설정 로드됨")
                except Exception as e:
                    st.error(f"파일 읽기 오류: {e}")

        # 우선순위 자동 계산 (그룹 사용 시)
        priority_map: dict[str, int] = {}
        if use_groups and parsed_groups:
            st.info("💡 그룹 내 신청 인원 순으로 우선순위가 자동 계산됩니다. 아래에서 조정 가능합니다.")
            for g in parsed_groups:
                enroll = sorted(g['subjects'], key=lambda s: -raw_df[s].sum())
                for rank, s in enumerate(enroll):
                    priority_map[s] = len(enroll) - rank  # 1위가 가장 높음

        st.markdown("""
        - **학급수**: 해당 과목의 전체 분반 수 (타임 합산)
        - **교사수**: 같은 타임 최대 분반 수
        - **우선순위**: 높을수록 분반 배정 우선 (그룹 내에서만 의미)
        """)

        if not st.session_state.get('settings_init'):
            init_data = {
                '과목명':            subjects_auto,
                '학급수(총 분반 수)': [3] * len(subjects_auto),
                '교사수(타임당 최대)': [1] * len(subjects_auto),
                '그룹':              [groups_to_subject_map(parsed_groups).get(s, '(미지정)') for s in subjects_auto],
                '우선순위':          [priority_map.get(s, 0) for s in subjects_auto],
            }
            st.session_state['settings_df'] = pd.DataFrame(init_data)
            st.session_state['settings_init'] = True
        else:
            # 그룹·우선순위 컬럼 갱신
            cur = st.session_state['settings_df']
            sub_to_g = groups_to_subject_map(parsed_groups)
            cur['그룹']    = cur['과목명'].apply(lambda s: sub_to_g.get(str(s).strip(), '(미지정)'))
            cur['우선순위'] = cur['과목명'].apply(lambda s: priority_map.get(str(s).strip(), 0))
            st.session_state['settings_df'] = cur

        edited = st.data_editor(
            st.session_state['settings_df'],
            num_rows='dynamic',
            use_container_width=True,
            height=min(80 + 35 * len(st.session_state['settings_df']), 600),
            column_config={
                '과목명':            st.column_config.TextColumn('과목명', width='large'),
                '학급수(총 분반 수)': st.column_config.NumberColumn('학급수', min_value=1, max_value=30, step=1),
                '교사수(타임당 최대)': st.column_config.NumberColumn('교사수', min_value=1, max_value=6, step=1),
                '그룹':              st.column_config.TextColumn('그룹', width='medium', disabled=True),
                '우선순위':          st.column_config.NumberColumn('우선순위', min_value=0, max_value=99, step=1,
                                                                   help="같은 그룹 안에서 높을수록 분반 배정 우선"),
            },
            key='settings_editor',
        )
        st.session_state['settings_df'] = edited

        valid          = edited.dropna(subset=['과목명'])
        total_sections = int(valid['학급수(총 분반 수)'].fillna(0).sum())
        lo = total_sections // n_times
        hi = (total_sections + n_times - 1) // n_times

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("총 분반 수", total_sections)
        col_b.metric("타임 수", n_times)
        col_c.metric("타임당 분반", f"{lo}~{hi}개")

        sep_pairs = st.session_state.get('sep_pairs', set())
        hints = []
        if sep_pairs:
            hints.append(f"🚫 분리 조건 {len(sep_pairs)}쌍 적용 예정")
        if use_groups and parsed_groups:
            hints.append(f"📚 과목 그룹 {len(parsed_groups)}개 적용 예정")
        for h in hints:
            st.info(h)

        if total_sections == 0:
            st.error("학급수 합계가 0입니다.")
        elif total_sections % n_times != 0:
            st.warning(f"⚠️ {total_sections}개 분반을 {n_times}타임에 균등 배분 불가 — ILP가 {lo}~{hi}개 범위로 맞춥니다.")
        else:
            st.success(f"✅ 타임당 정확히 {total_sections // n_times}개 분반 배정 예정")

        st.divider()
        run_btn = st.button("🚀 편성 시작", type="primary",
                            use_container_width=True, disabled=(total_sections == 0))

        if run_btn:
            subjects    = valid['과목명'].str.strip().tolist()
            n_sections  = dict(zip(valid['과목명'].str.strip(), valid['학급수(총 분반 수)'].astype(int)))
            mpt         = dict(zip(valid['과목명'].str.strip(), valid['교사수(타임당 최대)'].astype(int)))
            prio_map    = dict(zip(valid['과목명'].str.strip(), valid['우선순위'].fillna(0).astype(int)))
            id_col      = st.session_state['id_col']
            sep_pairs   = st.session_state.get('sep_pairs', set())
            use_g       = st.session_state.get('use_groups', False)
            p_groups    = st.session_state.get('parsed_groups', []) if use_g else []

            with st.spinner("① ILP로 분반 타임 배정 중 (그룹 분산 제약 포함)..."):
                assignment, ilp_status = run_ilp(subjects, n_sections, mpt, times, groups=p_groups)

            if assignment is None:
                st.error(f"❌ ILP 실패 (상태: {ilp_status}). 설정을 확인해주세요.")
            else:
                if 'relaxed' in ilp_status:
                    st.warning("⚠️ 그룹 분산 제약을 완화해 풀었습니다.")

                sections = build_sections(subjects, assignment, times)
                homeroom = {str(r[id_col]): extract_homeroom(str(r[id_col]))
                            for _, r in raw_df.iterrows()}

                with st.spinner("② 학생 분반 배정 중 (그룹·우선순위·분리 조건 반영)..."):
                    result_df, sec_roster, sec_load, failed = assign_students(
                        raw_df, id_col, subjects, assignment, sections,
                        times, int(max_per_section), homeroom,
                        sep_pairs=sep_pairs, groups=p_groups,
                        priority_map=prio_map, seed=seed,
                    )

                # 이동반 분리 위반 확인
                moving_violations = []
                for s in subjects:
                    for t in times:
                        for k in sections[s][t]:
                            roster = set(sec_roster.get((s, k), []))
                            for pair in sep_pairs:
                                pl = list(pair)
                                if pl[0] in roster and pl[1] in roster:
                                    moving_violations.append({
                                        '학번A': pl[0], '학번B': pl[1],
                                        '배정위치': f"{abbr(s)}-{k}반 ({t}타임)",
                                        '구분': '이동반',
                                    })

                st.session_state.update({
                    'assignment': assignment, 'sections': sections,
                    'sec_load': sec_load,     'sec_roster': sec_roster,
                    'result_df': result_df,   'failed': failed,
                    'subjects': subjects,     'homeroom': homeroom,
                    'n_total': len(raw_df),   'id_col': id_col,
                    'has_result': True,
                    'moving_violations': moving_violations,
                    'active_groups': p_groups,
                    'has_homeroom_result': False,
                })

                n_ok = len(raw_df) - len(failed)
                n_ov = sum(1 for v in sec_load.values() if v > max_per_section)
                msgs = []
                if moving_violations: msgs.append(f"분리 위반 {len(moving_violations)}건")
                if failed:            msgs.append(f"미배정 {len(failed)}명")
                if n_ov:              msgs.append(f"초과 분반 {n_ov}개")

                if msgs:
                    st.warning(f"⚠️ 편성 완료 — {' / '.join(msgs)} → [⑤ 편성 결과] 탭 확인")
                else:
                    st.success(f"✅ 편성 완료! {n_ok}/{len(raw_df)}명 전원 배정, 모든 조건 충족")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 5: 편성 결과
# ════════════════════════════════════════════════════════════════════════════════
with tab3:
    if not st.session_state.get('has_result'):
        st.info("편성을 실행하면 여기에 결과가 표시됩니다.")
        st.stop()

    subjects      = st.session_state['subjects']
    assignment    = st.session_state['assignment']
    sections      = st.session_state['sections']
    sec_load      = st.session_state['sec_load']
    sec_roster    = st.session_state['sec_roster']
    result_df     = st.session_state['result_df']
    failed        = st.session_state['failed']
    homeroom      = st.session_state['homeroom']
    n_total       = st.session_state['n_total']
    id_col        = st.session_state['id_col']
    mv_viol       = st.session_state.get('moving_violations', [])
    active_groups = st.session_state.get('active_groups', [])
    n_ok          = n_total - len(failed)
    n_ov          = sum(1 for v in sec_load.values() if v > max_per_section)
    sub_to_group  = groups_to_subject_map(active_groups)

    st.subheader("📊 편성 결과 요약")
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("총 학생", f"{n_total}명")
    m2.metric("배정 완료", f"{n_ok}명", f"{100*n_ok/n_total:.1f}%")
    m3.metric("미배정", f"{len(failed)}명", delta_color="inverse" if failed else "off")
    m4.metric("정원 초과 분반", f"{n_ov}개",   delta_color="inverse" if n_ov else "off")
    m5.metric("🚫 분리 위반",  f"{len(mv_viol)}건", delta_color="inverse" if mv_viol else "off")

    if mv_viol:
        with st.expander(f"⚠️ 이동반 분리 위반 {len(mv_viol)}건"):
            st.dataframe(pd.DataFrame(mv_viol), use_container_width=True, hide_index=True)
    elif st.session_state.get('sep_pairs'):
        st.success(f"✅ 분리 조건 {len(st.session_state['sep_pairs'])}쌍 모두 충족")
    if failed:
        with st.expander(f"⚠️ 미배정 {len(failed)}명"):
            st.write(failed)

    st.divider()
    rt1, rt2, rt3 = st.tabs(["📋 타임 배정표", "👥 분반별 인원", "📄 학생 시간표"])

    with rt1:
        st.subheader("과목별 타임 배정")
        if active_groups:
            st.caption("그룹 색상 구분: " + " | ".join(f"**{g['name']}**" for g in active_groups))
        rows = []
        for s in subjects:
            row = {'과목': s, '그룹': sub_to_group.get(s, ''), '총 분반': sum(assignment[s][t] for t in times)}
            for t in times:
                cnt  = assignment[s][t]
                secs = ', '.join(f"{abbr(s)}-{k}반" for k in sections[s][t])
                row[f'{t}타임'] = f"{cnt}분반 ({secs})" if cnt else "—"
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # 그룹별 타임 분포 확인
        if active_groups:
            st.subheader("그룹별 타임 분산 현황")
            for g in active_groups:
                gsubs = [s for s in g['subjects'] if s in subjects]
                g_data = {t: sum(assignment[s][t] for s in gsubs) for t in times}
                st.caption(f"**{g['name']}** ({g['n_pick']}선택) — 타임별 분반 수")
                st.bar_chart(g_data)

    with rt2:
        rows = []
        for t in times:
            for s in subjects:
                for k in sections[s][t]:
                    n     = sec_load.get((s, k), 0)
                    state = "★ 초과" if n > max_per_section else "✓"
                    rows.append({'타임': t, '그룹': sub_to_group.get(s,''), '과목': s,
                                 '분반명': f"{abbr(s)}-{k}반", '배정인원': n, '상태': state})
        load_df = pd.DataFrame(rows)
        def _rc(row):
            c = '#FFE0E0' if row['상태'] == '★ 초과' else ''
            return [f'background-color: {c}'] * len(row)
        st.dataframe(load_df.style.apply(_rc, axis=1), use_container_width=True, hide_index=True)

    with rt3:
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            search = st.text_input("🔍 신학번 검색", placeholder="신학번 일부 입력...")
        with col_f2:
            filter_time = st.selectbox("타임 필터", ["전체"] + times)
        disp = result_df.copy().fillna('')
        if search:
            disp = disp[disp[id_col].astype(str).str.contains(search, na=False)]
        if filter_time != "전체":
            col_s = f"{filter_time}타임_과목"
            if col_s in disp.columns:
                disp = disp[disp[col_s].ne('')]
        st.caption(f"{len(disp)}명 표시 중")
        st.dataframe(disp, use_container_width=True, hide_index=True, height=450)

    st.divider()
    st.subheader("📥 결과 다운로드")
    with st.spinner("Excel 생성 중..."):
        excel_buf = build_excel(
            subjects, times, id_col, assignment, sections,
            sec_load, sec_roster, result_df, homeroom,
            n_total, failed, int(max_per_section),
            sep_violations=mv_viol, groups=active_groups,
        )
    csv_bytes = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button("📊 Excel 다운로드", data=excel_buf,
                           file_name="이동반편성결과.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with dl2:
        st.download_button("📄 CSV 다운로드", data=csv_bytes,
                           file_name="학생시간표.csv", mime="text/csv",
                           use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 6: 본반 편성
# ════════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("🏠 본반 편성")
    st.markdown("""
    **편성 기준:** 🔀 이동 최소화 · ⚖️ 남녀 균형 · 🚫 분리 조건
    """)

    if 'raw_df' not in st.session_state:
        st.info("먼저 [① CSV 업로드]를 완료해주세요.")
        st.stop()
    if not st.session_state.get('has_result'):
        st.info("먼저 [④ 과목 설정 및 편성]에서 이동반 편성을 완료해주세요.")
        st.stop()

    raw_df     = st.session_state['raw_df']
    id_col     = st.session_state['id_col']
    subjects   = st.session_state['subjects']
    gender_col = st.session_state.get('gender_col')
    sep_pairs  = st.session_state.get('sep_pairs', set())

    if sep_pairs:
        st.info(f"🚫 분리 조건 **{len(sep_pairs)}쌍** 적용됨")

    st.divider()
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        n_hr_classes = st.number_input("편성할 본반 수", min_value=1, max_value=30, value=6)
    with col_s2:
        max_per_hr   = st.number_input("본반당 최대 인원", min_value=10, max_value=60, value=30)
    with col_s3:
        if gender_col:
            gender_weight = st.slider("남녀 비율 균형 가중치", 0.0, 1.0, 0.5, 0.1)
        else:
            gender_weight = 0.0
            st.info("성별 정보 없음")

    if not gender_col:
        mc = st.selectbox("성별 컬럼 직접 선택", ["(없음)"] + [c for c in raw_df.columns if c != id_col])
        if mc != "(없음)":
            gender_col = mc
            st.session_state['gender_col'] = mc

    with st.expander("📊 과목 선택 패턴 분포"):
        pc = defaultdict(int)
        for _, row in raw_df.iterrows():
            pc[get_subject_pattern(row.to_dict(), subjects)] += 1
        st.dataframe(pd.DataFrame([
            {'과목 조합': ' / '.join(p) if p else '(없음)', '학생 수': c,
             '비율': f"{100*c/len(raw_df):.1f}%"}
            for p, c in sorted(pc.items(), key=lambda x: -x[1])
        ]), use_container_width=True, hide_index=True)

    n_students = len(raw_df)
    expected   = n_students / n_hr_classes
    col_v1, col_v2, col_v3 = st.columns(3)
    col_v1.metric("전체 학생 수", f"{n_students}명")
    col_v2.metric("반당 예상 인원", f"{expected:.1f}명")
    col_v3.metric("✅ 인원 여유" if expected <= max_per_hr else "⚠️ 초과 위험",
                  f"반당 {max_per_hr-expected:.1f}명" if expected <= max_per_hr else f"{expected:.0f} > {max_per_hr}명")

    st.divider()
    if st.button("🏠 본반 편성 시작", type="primary", use_container_width=True):
        with st.spinner("본반 편성 중..."):
            hr_result_df, hr_stats_df, cohesion, hr_assign = assign_homeroom_classes(
                df=raw_df, id_col=id_col, subjects=subjects, gender_col=gender_col,
                n_classes=n_hr_classes, max_per_class=max_per_hr,
                gender_ratio_weight=gender_weight, sep_pairs=sep_pairs, seed=seed,
            )
        hr_viol = check_separation_violations(hr_assign, sep_pairs, label='본반')
        st.session_state.update({
            'hr_result_df': hr_result_df, 'hr_stats_df': hr_stats_df,
            'hr_assignments': hr_assign,  'hr_cohesion': cohesion,
            'hr_n_classes': n_hr_classes, 'hr_violations': hr_viol,
            'has_homeroom_result': True,
        })
        if hr_viol:
            st.warning(f"⚠️ 본반 편성 완료 — 분리 위반 {len(hr_viol)}건")
        else:
            st.success("✅ 본반 편성 완료! 분리 조건 모두 충족")

    if st.session_state.get('has_homeroom_result'):
        hr_result_df = st.session_state['hr_result_df']
        hr_stats_df  = st.session_state['hr_stats_df']
        cohesion     = st.session_state['hr_cohesion']
        n_classes    = st.session_state['hr_n_classes']
        hr_assign    = st.session_state['hr_assignments']
        hr_viol      = st.session_state.get('hr_violations', [])

        st.divider()
        m1,m2,m3,m4 = st.columns(4)
        m1.metric("편성된 반 수", f"{n_classes}반")
        m2.metric("과목 패턴 응집도", f"{cohesion*100:.1f}%")
        m3.metric("🚫 분리 위반", f"{len(hr_viol)}건", delta_color="inverse" if hr_viol else "off")
        if gender_col:
            ratios = [r['남']/r['총원'] for _, r in hr_stats_df.iterrows() if r['총원'] > 0]
            m4.metric("남비율 표준편차", f"{pd.Series(ratios).std()*100:.1f}%p" if ratios else "—")
        else:
            m4.metric("성별 정보", "없음")

        if hr_viol:
            with st.expander(f"⚠️ 본반 분리 위반 {len(hr_viol)}건"):
                st.dataframe(pd.DataFrame(hr_viol), use_container_width=True, hide_index=True)
        elif sep_pairs:
            st.success(f"✅ 분리 조건 {len(sep_pairs)}쌍 모두 충족")

        def color_gender(val):
            if isinstance(val, str) and '%' in val:
                try:
                    v = float(val.replace('%',''))
                    if v > 65 or v < 35: return 'color: #CC0000; font-weight: bold'
                except: pass
            return ''

        st.subheader("반별 현황")
        st.dataframe(hr_stats_df.style.applymap(color_gender, subset=['남비율']),
                     use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2)
        with c1:
            st.caption("반별 총원")
            st.bar_chart({r['반']: r['총원'] for _, r in hr_stats_df.iterrows()})
        with c2:
            if gender_col:
                st.caption("반별 남녀")
                st.bar_chart(pd.DataFrame({'남': hr_stats_df.set_index('반')['남'],
                                           '여': hr_stats_df.set_index('반')['여']}))

        search_hr = st.text_input("🔍 학생 검색", placeholder="학번 일부 입력...", key="hr_search")
        disp_hr   = hr_result_df.copy()
        if search_hr:
            disp_hr = disp_hr[disp_hr[id_col].astype(str).str.contains(search_hr, na=False)]
        st.caption(f"{len(disp_hr)}명 표시 중")
        st.dataframe(disp_hr, use_container_width=True, hide_index=True, height=400)

        st.divider()
        with st.spinner("Excel 생성 중..."):
            hr_excel = build_homeroom_excel(
                hr_result_df, hr_stats_df, id_col, n_classes, hr_assign,
                raw_df, gender_col, subjects, sep_violations=hr_viol,
            )
        hr_csv = hr_result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button("📊 본반 편성 Excel", data=hr_excel,
                               file_name="본반편성결과.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary")
        with dl2:
            st.download_button("📄 본반 편성 CSV", data=hr_csv,
                               file_name="본반편성결과.csv", mime="text/csv",
                               use_container_width=True)
