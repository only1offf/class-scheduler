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

NON_SUBJECT_COLS = {'순번', '번호', '신학번', '학번', '이름', '성명', '학년', '반', '번', 'no', 'id',
                    '성별', '남', '여', 'gender', 'sex'}


def get_id_col(df: pd.DataFrame) -> str:
    for c in ['신학번', '학번', '번호', 'ID', 'id']:
        if c in df.columns:
            return c
    return df.columns[0]


def detect_gender_col(df: pd.DataFrame) -> str | None:
    candidates = ['성별', 'gender', 'sex', '남녀', '性別']
    for c in df.columns:
        if c.lower() in [x.lower() for x in candidates]:
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
# 교우관계 분리 조건 파싱
# ──────────────────────────────────────────────────────────────────────────────

def parse_separation_pairs(df_sep: pd.DataFrame) -> set[frozenset]:
    """
    DataFrame (학번A, 학번B 컬럼) → frozenset 쌍 집합
    중복·자기 자신 쌍 제거
    """
    pairs = set()
    for _, row in df_sep.iterrows():
        a = str(row.iloc[0]).strip()
        b = str(row.iloc[1]).strip()
        if a and b and a != b and a != 'nan' and b != 'nan':
            pairs.add(frozenset([a, b]))
    return pairs


def check_separation_violations(assignments_dict: dict, sep_pairs: set[frozenset],
                                  label: str = "") -> list[dict]:
    """
    assignments_dict: {sid: group_key}  (group_key = 반번호 or (과목,분반) 등)
    sep_pairs: 분리되어야 할 쌍
    반환: 위반 목록
    """
    violations = []
    for pair in sep_pairs:
        pair_list = list(pair)
        if len(pair_list) < 2:
            continue
        a, b = pair_list[0], pair_list[1]
        ga = assignments_dict.get(a)
        gb = assignments_dict.get(b)
        if ga is not None and gb is not None and ga == gb:
            violations.append({'학번A': a, '학번B': b, '배정위치': str(ga), '구분': label})
    return violations


# ──────────────────────────────────────────────────────────────────────────────
# 편성 로직
# ──────────────────────────────────────────────────────────────────────────────

def run_ilp(subjects, n_sections, max_per_time, times):
    total = sum(n_sections[s] for s in subjects)
    nt    = len(times)
    lo, hi = total // nt, (total + nt - 1) // nt

    prob = pulp.LpProblem("section_time", pulp.LpMinimize)
    z = {
        s: {t: pulp.LpVariable(f"z_{i}_{t}", 0, n_sections[s], cat='Integer')
            for t in times}
        for i, s in enumerate(subjects)
    }

    prob += 0

    for s in subjects:
        prob += pulp.lpSum(z[s][t] for t in times) == n_sections[s]
        for t in times:
            prob += z[s][t] <= max_per_time[s]

    for t in times:
        col_sum = pulp.lpSum(z[s][t] for s in subjects)
        prob += col_sum >= lo
        prob += col_sum <= hi

    status = prob.solve(pulp.PULP_CBC_CMD(msg=0))
    if pulp.LpStatus[prob.status] != 'Optimal':
        return None, pulp.LpStatus[prob.status]

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


def assign_students(df, id_col, subjects, assignment, sections,
                    times, max_per_section, homeroom=None, sep_pairs=None, seed=42):
    """
    이동반 배정 — sep_pairs 쌍은 같은 분반에 배정되지 않도록 우선 처리
    (hard constraint: 가능하면 분리, 불가능 시 위반 기록 후 진행)
    """
    homeroom  = homeroom or {}
    sep_pairs = sep_pairs or set()
    random.seed(seed)
    n_sel = len(times)

    # 분리 쌍에 속한 학번 → 연결된 학번들 인덱스
    sep_neighbors: dict[str, set[str]] = defaultdict(set)
    for pair in sep_pairs:
        pl = list(pair)
        sep_neighbors[pl[0]].add(pl[1])
        sep_neighbors[pl[1]].add(pl[0])

    sec_load   = {(s, k): 0 for s in subjects for t in times for k in sections[s][t]}
    sec_roster = defaultdict(list)   # (subject, section_num) → [sid, ...]
    records    = []
    failed     = []

    rows = df.to_dict('records')
    random.shuffle(rows)

    # 분리 쌍 우선 처리: 쌍에 속한 학생 먼저 배정
    sep_sids = {sid for pair in sep_pairs for sid in pair}
    rows_sep   = [r for r in rows if str(r.get(id_col, '')) in sep_sids]
    rows_other = [r for r in rows if str(r.get(id_col, '')) not in sep_sids]
    rows = rows_sep + rows_other

    # 현재 배정 상태: sid → {subject: section_num}
    sid_sec_map: dict[str, dict[str, int]] = {}

    def _sep_penalty(sid, s, k):
        """분리 위반 시 큰 페널티 (0 = 무위반, 1000 = 위반)"""
        neighbors = sep_neighbors.get(sid, set())
        for nb in neighbors:
            nb_map = sid_sec_map.get(nb, {})
            if nb_map.get(s) == k:
                return 1000
        return 0

    for row in rows:
        sid    = str(row.get(id_col, ''))
        chosen = [s for s in subjects if int(row.get(s, 0)) == 1]

        if len(chosen) != n_sel:
            failed.append(sid)
            continue

        best_map, best_cost     = None, float('inf')
        overflow_map, overflow_cost = None, float('inf')

        for perm in permutations(times):
            tm = dict(zip(chosen, perm))
            if not all(assignment[s][t] > 0 for s, t in tm.items()):
                continue

            sep_cost = sum(
                sum(_sep_penalty(sid, s, k) for k in sections[s][tm[s]])
                for s in chosen
            )
            load_cost = sum(
                min(sec_load[(s, k)] for k in sections[s][tm[s]])
                for s in chosen
            )
            cost = sep_cost * 10 + load_cost

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
        for s in chosen:
            t = final_map[s]
            # 분리 조건 우선, 그 다음 여석, 그 다음 최소 부하
            neighbors = sep_neighbors.get(sid, set())

            def pool_score(k):
                sep_v = sum(1 for nb in neighbors if sid_sec_map.get(nb, {}).get(s) == k)
                return (sep_v * 1000, sec_load[(s, k)])

            with_room = [k for k in sections[s][t] if sec_load[(s, k)] < max_per_section]
            pool = with_room if with_room else sections[s][t]
            k = min(pool, key=pool_score)

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
# 본반 편성 로직
# ──────────────────────────────────────────────────────────────────────────────

def get_subject_pattern(row, subjects):
    return tuple(s for s in subjects if int(row.get(s, 0)) == 1)


def assign_homeroom_classes(
    df, id_col, subjects, gender_col,
    n_classes, max_per_class,
    gender_ratio_weight=0.5,
    sep_pairs=None,
    seed=42
):
    """
    본반 편성:
    - 같은 과목 패턴 학생을 같은 반으로 최대한 묶기 (이동 최소화)
    - 남녀 비율 균형 유지
    - sep_pairs: 같은 본반에 배정되지 않아야 할 쌍 (hard constraint 우선)
    """
    sep_pairs = sep_pairs or set()
    random.seed(seed)

    rows = df.to_dict('records')

    def norm_gender(v):
        v = str(v).strip()
        if v in ('남', 'M', 'm', '1', 'male', 'Male'):
            return '남'
        if v in ('여', 'F', 'f', '2', 'female', 'Female'):
            return '여'
        return '미상'

    # 분리 이웃 인덱스
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

    total_m = sum(1 for r in rows if norm_gender(r.get(gender_col, '')) == '남') if gender_col else 0
    total   = len(rows)
    target_m_ratio = total_m / total if total else 0.5

    class_load     = [0] * n_classes
    class_males    = [0] * n_classes
    class_patterns = [defaultdict(int) for _ in range(n_classes)]
    class_sids     = [set() for _ in range(n_classes)]   # 분리 확인용

    assignments = {}   # sid -> 1-based 반 번호

    sorted_patterns = sorted(pattern_groups.items(), key=lambda x: -len(x[1]))

    for pat, students in sorted_patterns:
        random.shuffle(students)

        for stu in students:
            sid    = stu['sid']
            gender = stu['gender']

            candidates = [i for i in range(n_classes) if class_load[i] < max_per_class]
            if not candidates:
                candidates = list(range(n_classes))

            def score(i):
                # 분리 위반: 이 반에 이미 이웃이 있으면 큰 페널티
                sep_penalty = sum(1 for nb in sep_neighbors.get(sid, set())
                                  if nb in class_sids[i]) * 10000

                pattern_bonus = -class_patterns[i][pat] * 10

                load = class_load[i]
                if load > 0:
                    if gender == '남':
                        new_ratio = (class_males[i] + 1) / (load + 1)
                    else:
                        new_ratio = class_males[i] / (load + 1)
                    gender_score = abs(new_ratio - target_m_ratio) * gender_ratio_weight * 100
                else:
                    gender_score = 0

                balance_score = class_load[i]
                return sep_penalty + pattern_bonus + gender_score + balance_score

            best_class = min(candidates, key=score)
            assignments[sid] = best_class + 1

            class_load[best_class] += 1
            class_patterns[best_class][pat] += 1
            class_sids[best_class].add(sid)
            if gender == '남':
                class_males[best_class] += 1

    records = []
    for row in rows:
        sid    = str(row.get(id_col, ''))
        gender = norm_gender(row.get(gender_col, '')) if gender_col else '미상'
        pat    = get_subject_pattern(row, subjects)
        cls    = assignments.get(sid)
        records.append({
            id_col:       sid,
            '배정_본반':  f"{cls}반" if cls else '',
            '성별':       gender,
            '과목패턴':   ' / '.join(pat) if pat else '',
        })

    result_df = pd.DataFrame(records)

    stats = []
    for i in range(n_classes):
        cls_num = i + 1
        n = class_load[i]
        m = class_males[i]
        f = n - m
        pat_dist = sorted(class_patterns[i].items(), key=lambda x: -x[1])
        top_pats = ', '.join(f"{'·'.join(p)}({c}명)" for p, c in pat_dist[:3]) if pat_dist else ''
        stats.append({
            '반': f"{cls_num}반",
            '총원': n, '남': m, '여': f,
            '남비율': f"{100*m/n:.1f}%" if n else '0%',
            '주요 과목패턴': top_pats,
        })

    stats_df = pd.DataFrame(stats)

    total_students = len(rows)
    same_pattern_count = sum(
        cnt * (cnt - 1) // 2
        for pat_dict in class_patterns
        for cnt in pat_dict.values()
    )
    total_pairs = total_students * (total_students - 1) // 2
    cohesion_rate = same_pattern_count / total_pairs if total_pairs else 0

    return result_df, stats_df, cohesion_rate, assignments


# ──────────────────────────────────────────────────────────────────────────────
# Excel 출력
# ──────────────────────────────────────────────────────────────────────────────

_H_FILL  = PatternFill("solid", fgColor="2E4699")
_H_FONT  = Font(bold=True, color="FFFFFF", size=10)
_H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
_S_FILL  = PatternFill("solid", fgColor="5B9BD5")
_W_FILL  = PatternFill("solid", fgColor="FFD7D7")  # 위반 강조


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


def build_excel(subjects, times, id_col, assignment, sections,
                sec_load, sec_roster, result_df, homeroom,
                n_total, failed, max_per_section, sep_violations=None):
    sep_violations = sep_violations or []
    wb = Workbook()
    wb.remove(wb.active)

    # ── 개요 ──────────────────────────────────────────────────────────────
    ws = wb.create_sheet("개요")
    ws.cell(1, 1, f"이동반 편성 결과 — 총 {n_total}명 / 충족 {n_total-len(failed)}명 / 미배정 {len(failed)}명")
    ws.cell(1, 1).font = Font(bold=True, size=12)
    ws.merge_cells('A1:G1')
    ws.row_dimensions[1].height = 22

    r = 3
    for ci, h in enumerate(['타임', '과목', '분반명', '배정인원', '최대인원', '상태'], 1):
        _hdr(ws, r, ci, h)
    r += 1

    for t in times:
        _subhdr(ws, r, 1, f"▶ {t}타임")
        ws.merge_cells(f'A{r}:G{r}')
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt   = sec_load.get((s, k), 0)
                state = "★초과" if cnt > max_per_section else "정상"
                ws.cell(r, 1, t); ws.cell(r, 2, s)
                ws.cell(r, 3, f"{abbr(s)}-{k}반"); ws.cell(r, 4, cnt)
                ws.cell(r, 5, max_per_section)
                c6 = ws.cell(r, 6, state)
                if state != "정상":
                    c6.font = Font(bold=True, color="FF0000")
                r += 1
        r += 1
    _auto_width(ws)

    # ── 학생 시간표 ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("학생시간표")
    cols = [id_col, '본반'] + [f'{t}타임_과목' for t in times] + [f'{t}타임_분반' for t in times]
    for ci, h in enumerate(cols, 1):
        _hdr(ws2, 1, ci, h)
    ws2.freeze_panes = 'A2'
    for _, row in result_df.iterrows():
        ws2.append([str(row.get(c, '') or '') for c in cols])
    _auto_width(ws2)

    # ── 분반별 명단 ────────────────────────────────────────────────────────
    for s in subjects:
        for t in times:
            for k in sections[s][t]:
                ws_s = wb.create_sheet(f"{abbr(s)}-{k}반")
                ws_s.cell(1, 1, f"[{s}]  {k}분반  ({t}타임)").font = Font(bold=True, size=11)
                ws_s.merge_cells('A1:C1')
                for ci, h in enumerate([id_col, '본반', '비고'], 1):
                    _hdr(ws_s, 2, ci, h)
                for sid in sorted(sec_roster.get((s, k), [])):
                    ws_s.append([sid, homeroom.get(sid, ''), ''])
                _auto_width(ws_s)

    # ── 교실배치 ──────────────────────────────────────────────────────────
    ws_cr = wb.create_sheet("교실배치추천")
    ws_cr.cell(1, 1, "교실 배치 추천").font = Font(bold=True, size=12)
    ws_cr.merge_cells('A1:F1')
    r = 3
    for t in times:
        _subhdr(ws_cr, r, 1, f"▶ {t}타임")
        ws_cr.merge_cells(f'A{r}:F{r}')
        r += 1
        for ci, h in enumerate(['과목', '분반명', '배정인원', '본반 구성 (우선순위)', '권장교실', '비고'], 1):
            _hdr(ws_cr, r, ci, h)
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt = sec_load.get((s, k), 0)
                hr_cnt = defaultdict(int)
                for sid in sec_roster.get((s, k), []):
                    hr = homeroom.get(sid, '')
                    if hr:
                        hr_cnt[hr] += 1
                top      = sorted(hr_cnt.items(), key=lambda x: (-x[1], x[0]))
                priority = '  >  '.join(f"{hr} {c}명" for hr, c in top) if top else ''
                ws_cr.cell(r, 1, s); ws_cr.cell(r, 2, f"{abbr(s)}-{k}반")
                ws_cr.cell(r, 3, cnt)
                ws_cr.cell(r, 4, priority).alignment = Alignment(horizontal='left')
                r += 1
        r += 1
    ws_cr.column_dimensions['A'].width = 22
    ws_cr.column_dimensions['B'].width = 14
    ws_cr.column_dimensions['C'].width = 10
    ws_cr.column_dimensions['D'].width = 65
    ws_cr.column_dimensions['E'].width = 14

    # ── 분리 위반 시트 ────────────────────────────────────────────────────
    if sep_violations:
        ws_v = wb.create_sheet("⚠️분리위반")
        ws_v.cell(1, 1, f"분리 조건 위반 목록 — {len(sep_violations)}건").font = Font(bold=True, size=12, color="CC0000")
        ws_v.merge_cells('A1:D1')
        for ci, h in enumerate(['학번A', '학번B', '배정위치', '구분'], 1):
            _hdr(ws_v, 2, ci, h)
        for viol in sep_violations:
            row_vals = [viol.get('학번A',''), viol.get('학번B',''),
                        viol.get('배정위치',''), viol.get('구분','')]
            ws_v.append(row_vals)
            # 위반 행 배경색
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

    _H_FILL  = PatternFill("solid", fgColor="2E4699")
    _H_FONT  = Font(bold=True, color="FFFFFF", size=10)
    _H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    _W_FILL2 = PatternFill("solid", fgColor="FFD7D7")

    def hdr(ws, r, c, v):
        cell = ws.cell(r, c, v)
        cell.font, cell.fill, cell.alignment = _H_FONT, _H_FILL, _H_ALIGN

    def aw(ws, mn=8, mx=42):
        for col in ws.columns:
            w = max((len(str(c.value or '')) for c in col), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w+2,mn),mx)

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
    cols = [id_col, '배정_본반', '성별', '과목패턴']
    for ci, h in enumerate(cols, 1):
        hdr(ws2, 1, ci, h)
    ws2.freeze_panes = 'A2'
    for _, row in hr_result_df.iterrows():
        ws2.append([str(row.get(c, '') or '') for c in cols])
    aw(ws2)

    def norm_gender(v):
        v = str(v).strip()
        if v in ('남', 'M', 'm', '1', 'male', 'Male'): return '남'
        if v in ('여', 'F', 'f', '2', 'female', 'Female'): return '여'
        return '미상'

    for cls_num in range(1, n_classes + 1):
        ws_c = wb.create_sheet(f"{cls_num}반")
        ws_c.cell(1, 1, f"{cls_num}반 학생 명단").font = Font(bold=True, size=12)
        ws_c.merge_cells('A1:D1')
        for ci, h in enumerate([id_col, '성별', '과목패턴', '비고'], 1):
            hdr(ws_c, 2, ci, h)
        class_students = hr_result_df[hr_result_df['배정_본반'] == f"{cls_num}반"]
        for _, row in class_students.sort_values(id_col).iterrows():
            ws_c.append([str(row[id_col]), row['성별'], row['과목패턴'], ''])
        aw(ws_c)

    if sep_violations:
        ws_v = wb.create_sheet("⚠️분리위반")
        ws_v.cell(1, 1, f"분리 조건 위반 목록 — {len(sep_violations)}건").font = Font(bold=True, size=12, color="CC0000")
        ws_v.merge_cells('A1:D1')
        for ci, h in enumerate(['학번A', '학번B', '배정위치', '구분'], 1):
            hdr(ws_v, 2, ci, h)
        for viol in sep_violations:
            ws_v.append([viol.get('학번A',''), viol.get('학번B',''),
                         viol.get('배정위치',''), viol.get('구분','')])
            for ci in range(1, 5):
                ws_v.cell(ws_v.max_row, ci).fill = _W_FILL2
        aw(ws_v)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ──────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="이동반 편성 시스템", layout="wide", page_icon="🏫")
st.title("🏫 선택과목 이동반 편성 시스템")
st.caption("CSV 업로드 → 과목 설정 → 편성 시작 → 결과 확인 및 다운로드")

if not HAS_PULP:
    st.error("⚠️ PuLP 라이브러리가 없습니다. `pip install pulp` 후 재시작하세요.")
    st.stop()

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 기본 설정")
    n_times = st.number_input("타임 수", min_value=2, max_value=6, value=3)
    times   = get_time_labels(n_times)
    st.info(f"타임: **{' / '.join(times)}**")
    max_per_section = st.number_input("분반당 최대 인원", min_value=10, max_value=100, value=41)
    seed = int(st.number_input("무작위 시드", min_value=0, max_value=9999, value=42))
    st.divider()
    st.caption("※ 설정 변경 후 편성 시작을 다시 눌러주세요.")

# ── 탭 ───────────────────────────────────────────────────────────────────────
tab1, tab_sep, tab2, tab3, tab4 = st.tabs([
    "① CSV 업로드",
    "② 분리 조건",
    "③ 과목 설정 및 편성",
    "④ 편성 결과",
    "⑤ 본반 편성",
])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1: 파일 업로드
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("학생 선택과목 CSV 업로드")
    st.markdown("""
    **파일 형식:**
    - 컬럼: `신학번`(또는 `학번`) + `성별` + 과목명들
    - 과목 선택 여부: `1`(선택) / `0`(미선택)
    - 성별: `남` / `여` (또는 M/F, 1/2)
    """)

    uploaded = st.file_uploader("CSV 파일 선택", type=['csv'], label_visibility='collapsed')

    if uploaded:
        try:
            raw_df = pd.read_csv(uploaded, encoding='utf-8-sig', dtype=str)
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

            st.session_state['raw_df']        = raw_df
            st.session_state['subjects_auto'] = subjects_auto
            st.session_state['id_col']        = id_col
            st.session_state['gender_col']    = gender_col
            st.session_state['settings_init'] = False

            n_sel = raw_df[subjects_auto].sum(axis=1)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("총 학생 수", f"{len(raw_df)}명")
            c2.metric("감지된 과목 수", f"{len(subjects_auto)}개")
            c3.metric("학생당 선택 과목", f"{n_sel.mean():.1f}개 (평균)")
            bad_cnt = int((n_sel != n_times).sum())
            c4.metric(f"{n_times}과목 미준수", f"{bad_cnt}명",
                      delta_color="inverse" if bad_cnt else "off")

            if gender_col:
                st.success(f"✅ 성별 컬럼 감지: **{gender_col}**")
            else:
                st.warning("⚠️ 성별 컬럼을 찾지 못했습니다.")

            if bad_cnt:
                st.warning(f"⚠️ {n_times}과목 미선택 학생 {bad_cnt}명 → 미배정 처리됩니다.")
            else:
                st.success(f"✅ 전원 {n_times}과목 선택 확인")

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
# TAB 2: 분리 조건 (교우관계)
# ════════════════════════════════════════════════════════════════════════════════
with tab_sep:
    st.subheader("🚫 교우관계 분리 조건 설정")
    st.markdown("""
    지정한 두 학생은 **본반·이동반 모두에서 같은 반/분반에 배정되지 않도록** 우선 처리됩니다.

    > ⚠️ 과목 조합·인원 등 구조적 제약으로 분리가 불가능한 경우, 편성은 진행되며 위반 내역이 따로 표시됩니다.
    """)

    # ── 방법 1: CSV 업로드 ────────────────────────────────────────────────
    st.markdown("#### 방법 ① — CSV 파일 업로드")
    st.markdown("""
    **형식:** 헤더 없이 (또는 `학번A`, `학번B` 헤더 포함) 두 열로 구성
    ```
    10101,10205
    10302,10418
    ```
    """)

    sep_uploaded = st.file_uploader(
        "분리 조건 CSV 업로드",
        type=['csv'],
        key='sep_uploader',
        label_visibility='collapsed',
    )

    uploaded_pairs: set[frozenset] = set()
    if sep_uploaded:
        try:
            sep_raw = pd.read_csv(sep_uploaded, encoding='utf-8-sig', dtype=str, header=None)
            # 첫 행이 헤더처럼 보이면 제거
            first_row = sep_raw.iloc[0].tolist()
            if any(str(v).strip().lower() in ['학번a','학번b','a','b','학번','id'] for v in first_row):
                sep_raw = sep_raw.iloc[1:].reset_index(drop=True)
            uploaded_pairs = parse_separation_pairs(sep_raw)
            st.success(f"✅ {len(uploaded_pairs)}쌍 로드됨")
            preview = pd.DataFrame([{'학번A': list(p)[0], '학번B': list(p)[1]} for p in uploaded_pairs])
            st.dataframe(preview, use_container_width=True, hide_index=True, height=180)
        except Exception as e:
            st.error(f"CSV 읽기 오류: {e}")

    st.divider()

    # ── 방법 2: 직접 입력 ─────────────────────────────────────────────────
    st.markdown("#### 방법 ② — 앱 안에서 직접 입력")
    st.caption("행을 추가(+)하거나 삭제(행 선택 후 Delete)할 수 있습니다.")

    if 'sep_manual_df' not in st.session_state:
        st.session_state['sep_manual_df'] = pd.DataFrame({'학번A': [''], '학번B': ['']})

    edited_sep = st.data_editor(
        st.session_state['sep_manual_df'],
        num_rows='dynamic',
        use_container_width=True,
        height=220,
        column_config={
            '학번A': st.column_config.TextColumn('학번A', width='medium'),
            '학번B': st.column_config.TextColumn('학번B', width='medium'),
        },
        key='sep_editor',
    )
    st.session_state['sep_manual_df'] = edited_sep

    manual_pairs = parse_separation_pairs(edited_sep)

    # ── 전체 합산 ─────────────────────────────────────────────────────────
    all_sep_pairs = uploaded_pairs | manual_pairs
    st.session_state['sep_pairs'] = all_sep_pairs

    st.divider()
    col_p1, col_p2 = st.columns(2)
    col_p1.metric("CSV 업로드 분리 쌍", f"{len(uploaded_pairs)}쌍")
    col_p2.metric("직접 입력 분리 쌍", f"{len(manual_pairs)}쌍")

    if all_sep_pairs:
        st.info(f"✅ 총 **{len(all_sep_pairs)}쌍**의 분리 조건이 적용됩니다 → 이동반·본반 편성 시 자동 반영")
        with st.expander(f"전체 분리 조건 목록 ({len(all_sep_pairs)}쌍)"):
            all_pairs_df = pd.DataFrame([
                {'학번A': sorted(list(p))[0], '학번B': sorted(list(p))[1]}
                for p in all_sep_pairs
            ]).sort_values(['학번A', '학번B']).reset_index(drop=True)
            st.dataframe(all_pairs_df, use_container_width=True, hide_index=True)
    else:
        st.info("분리 조건이 없습니다. 조건 없이 편성됩니다.")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3: 과목 설정 및 편성
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    if 'raw_df' not in st.session_state:
        st.info("먼저 [① CSV 업로드] 탭에서 파일을 업로드해주세요.")
    else:
        subjects_auto = st.session_state['subjects_auto']

        st.subheader("과목별 분반·교사 수 설정")
        st.markdown("""
        - **학급수**: 해당 과목의 전체 분반 수 (타임 합산)
        - **교사수**: 같은 타임에 최대 몇 분반까지 동시 운영 가능한지
        """)

        if not st.session_state.get('settings_init'):
            st.session_state['settings_df'] = pd.DataFrame({
                '과목명':            subjects_auto,
                '학급수(총 분반 수)': [3] * len(subjects_auto),
                '교사수(타임당 최대)': [1] * len(subjects_auto),
            })
            st.session_state['settings_init'] = True

        edited = st.data_editor(
            st.session_state['settings_df'],
            num_rows='dynamic',
            use_container_width=True,
            height=min(80 + 35 * len(st.session_state['settings_df']), 600),
            column_config={
                '과목명': st.column_config.TextColumn('과목명', width='large'),
                '학급수(총 분반 수)': st.column_config.NumberColumn(
                    '학급수 (총 분반)', min_value=1, max_value=30, step=1, width='medium'),
                '교사수(타임당 최대)': st.column_config.NumberColumn(
                    '교사수 (타임당 최대 분반)', min_value=1, max_value=6, step=1, width='medium'),
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

        # 분리 조건 현황 표시
        sep_pairs = st.session_state.get('sep_pairs', set())
        if sep_pairs:
            st.info(f"🚫 분리 조건 {len(sep_pairs)}쌍 적용 예정")

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
            subjects   = valid['과목명'].str.strip().tolist()
            n_sections = dict(zip(valid['과목명'].str.strip(), valid['학급수(총 분반 수)'].astype(int)))
            mpt        = dict(zip(valid['과목명'].str.strip(), valid['교사수(타임당 최대)'].astype(int)))
            raw_df     = st.session_state['raw_df']
            id_col     = st.session_state['id_col']
            sep_pairs  = st.session_state.get('sep_pairs', set())

            with st.spinner("① ILP로 분반 타임 배정 중..."):
                assignment, ilp_status = run_ilp(subjects, n_sections, mpt, times)

            if assignment is None:
                st.error(f"❌ ILP 해를 찾지 못했습니다 (상태: {ilp_status}).")
            else:
                sections = build_sections(subjects, assignment, times)
                homeroom = {str(r[id_col]): extract_homeroom(str(r[id_col]))
                            for _, r in raw_df.iterrows()}

                with st.spinner("② 학생 분반 배정 중 (분리 조건 반영)..."):
                    result_df, sec_roster, sec_load, failed = assign_students(
                        raw_df, id_col, subjects, assignment, sections,
                        times, int(max_per_section), homeroom,
                        sep_pairs=sep_pairs, seed=seed,
                    )

                # 분리 위반 확인 (이동반)
                moving_violations = []
                for s in subjects:
                    for t in times:
                        for k in sections[s][t]:
                            roster = sec_roster.get((s, k), [])
                            sid_to_sec = {sid: (s, k) for sid in roster}
                            for pair in sep_pairs:
                                pl = list(pair)
                                if pl[0] in sid_to_sec and pl[1] in sid_to_sec:
                                    moving_violations.append({
                                        '학번A': pl[0], '학번B': pl[1],
                                        '배정위치': f"{abbr(s)}-{k}반 ({t}타임)",
                                        '구분': '이동반',
                                    })

                st.session_state.update({
                    'assignment':  assignment, 'sections': sections,
                    'sec_load':    sec_load,   'sec_roster': sec_roster,
                    'result_df':   result_df,  'failed': failed,
                    'subjects':    subjects,   'homeroom': homeroom,
                    'n_total':     len(raw_df),'id_col': id_col,
                    'has_result':  True,
                    'moving_violations': moving_violations,
                    'has_homeroom_result': False,
                })

                n_ok = len(raw_df) - len(failed)
                n_ov = sum(1 for v in sec_load.values() if v > max_per_section)

                if moving_violations:
                    st.warning(f"⚠️ 편성 완료 — 이동반 분리 위반 {len(moving_violations)}건 발생 → [④ 편성 결과] 탭 확인")
                elif n_ov == 0 and not failed:
                    st.success(f"✅ 편성 완료! {n_ok}/{len(raw_df)}명 전원 배정, 분리 조건 모두 충족")
                else:
                    st.warning(f"⚠️ 편성 완료 — 미배정 {len(failed)}명 / 초과 분반 {n_ov}개 → [④ 편성 결과] 탭 확인")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 4: 편성 결과
# ════════════════════════════════════════════════════════════════════════════════
with tab3:
    if not st.session_state.get('has_result'):
        st.info("편성을 실행하면 여기에 결과가 표시됩니다.")
        st.stop()

    subjects   = st.session_state['subjects']
    assignment = st.session_state['assignment']
    sections   = st.session_state['sections']
    sec_load   = st.session_state['sec_load']
    sec_roster = st.session_state['sec_roster']
    result_df  = st.session_state['result_df']
    failed     = st.session_state['failed']
    homeroom   = st.session_state['homeroom']
    n_total    = st.session_state['n_total']
    id_col     = st.session_state['id_col']
    mv_viol    = st.session_state.get('moving_violations', [])
    n_ok       = n_total - len(failed)
    n_ov       = sum(1 for v in sec_load.values() if v > max_per_section)

    st.subheader("📊 편성 결과 요약")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("총 학생", f"{n_total}명")
    m2.metric("배정 완료", f"{n_ok}명", f"{100*n_ok/n_total:.1f}%")
    m3.metric("미배정", f"{len(failed)}명",
              delta=f"-{len(failed)}" if failed else "없음",
              delta_color="inverse" if failed else "off")
    m4.metric(f"정원 초과 분반", f"{n_ov}개",
              delta_color="inverse" if n_ov else "off")
    m5.metric("🚫 분리 위반", f"{len(mv_viol)}건",
              delta_color="inverse" if mv_viol else "off")

    if mv_viol:
        with st.expander(f"⚠️ 이동반 분리 조건 위반 {len(mv_viol)}건 — 클릭하여 확인"):
            st.dataframe(pd.DataFrame(mv_viol), use_container_width=True, hide_index=True)
            st.caption("구조적으로 분리가 불가능한 경우 발생합니다. 과목 분반 수를 늘리거나 조건을 재검토해주세요.")
    elif st.session_state.get('sep_pairs'):
        st.success(f"✅ 분리 조건 {len(st.session_state['sep_pairs'])}쌍 모두 충족")

    if failed:
        with st.expander(f"⚠️ 미배정 학생 {len(failed)}명"):
            st.write(failed)

    st.divider()

    rt1, rt2, rt3 = st.tabs(["📋 타임 배정표", "👥 분반별 인원", "📄 학생 시간표"])

    with rt1:
        rows = []
        for s in subjects:
            row = {'과목': s, '총 분반': sum(assignment[s][t] for t in times)}
            for t in times:
                cnt  = assignment[s][t]
                secs = ', '.join(f"{abbr(s)}-{k}반" for k in sections[s][t])
                row[f'{t}타임'] = f"{cnt}분반 ({secs})" if cnt else "—"
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        totals = {t: sum(assignment[s][t] for s in subjects) for t in times}
        st.bar_chart(totals)

    with rt2:
        rows = []
        for t in times:
            for s in subjects:
                for k in sections[s][t]:
                    n     = sec_load.get((s, k), 0)
                    state = "★ 초과" if n > max_per_section else "✓"
                    rows.append({'타임': t, '과목': s,
                                 '분반명': f"{abbr(s)}-{k}반",
                                 '배정인원': n, '상태': state})
        load_df = pd.DataFrame(rows)
        def _rc(row):
            c = '#FFE0E0' if row['상태'] == '★ 초과' else ''
            return [f'background-color: {c}'] * len(row)
        st.dataframe(load_df.style.apply(_rc, axis=1),
                     use_container_width=True, hide_index=True)
        chart_data = {}
        for s in subjects:
            for t in times:
                for k in sections[s][t]:
                    chart_data[f"{abbr(s)}-{k}반({t})"] = sec_load.get((s, k), 0)
        st.bar_chart(chart_data)

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
            sep_violations=mv_viol,
        )
    csv_bytes = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button("📊 Excel 다운로드 (전체)", data=excel_buf,
                           file_name="이동반편성결과.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with dl2:
        st.download_button("📄 CSV 다운로드 (학생 시간표)", data=csv_bytes,
                           file_name="학생시간표.csv", mime="text/csv",
                           use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 5: 본반 편성
# ════════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("🏠 본반 편성")
    st.markdown("""
    **편성 기준:**
    - 🔀 **이동 최소화**: 같은 과목 조합 학생끼리 같은 반으로 묶기
    - ⚖️ **남녀 균형**: 전체 비율에 맞도록 배분
    - 🚫 **분리 조건**: [② 분리 조건] 탭에서 입력한 쌍은 다른 반으로 분리
    """)

    if 'raw_df' not in st.session_state:
        st.info("먼저 [① CSV 업로드] 탭에서 파일을 업로드해주세요.")
        st.stop()
    if not st.session_state.get('has_result'):
        st.info("먼저 [③ 과목 설정 및 편성] 탭에서 이동반 편성을 완료해주세요.")
        st.stop()

    raw_df     = st.session_state['raw_df']
    id_col     = st.session_state['id_col']
    subjects   = st.session_state['subjects']
    gender_col = st.session_state.get('gender_col')
    sep_pairs  = st.session_state.get('sep_pairs', set())

    if sep_pairs:
        st.info(f"🚫 분리 조건 **{len(sep_pairs)}쌍** 적용됨")

    st.divider()
    st.subheader("⚙️ 본반 편성 설정")

    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        n_hr_classes = st.number_input("편성할 본반 수", min_value=1, max_value=30, value=6)
    with col_s2:
        max_per_hr = st.number_input("본반당 최대 인원", min_value=10, max_value=60, value=30)
    with col_s3:
        if gender_col:
            gender_weight = st.slider("남녀 비율 균형 가중치", 0.0, 1.0, 0.5, 0.1,
                                      help="0 = 이동 최소화 우선 / 1 = 남녀 균형 우선")
        else:
            gender_weight = 0.0
            st.info("성별 정보 없음 — 이동 최소화만 적용")

    if not gender_col:
        all_cols = [c for c in raw_df.columns if c != id_col]
        manual_gender = st.selectbox("성별 컬럼 직접 선택 (선택사항)", ["(없음)"] + all_cols)
        if manual_gender != "(없음)":
            gender_col = manual_gender
            st.session_state['gender_col'] = gender_col

    with st.expander("📊 과목 선택 패턴 분포 미리보기"):
        pattern_counts = defaultdict(int)
        for _, row in raw_df.iterrows():
            pat = get_subject_pattern(row.to_dict(), subjects)
            pattern_counts[pat] += 1
        pat_df = pd.DataFrame([
            {'과목 조합': ' / '.join(p) if p else '(없음)', '학생 수': c,
             '비율': f"{100*c/len(raw_df):.1f}%"}
            for p, c in sorted(pattern_counts.items(), key=lambda x: -x[1])
        ])
        st.dataframe(pat_df, use_container_width=True, hide_index=True)

    n_students = len(raw_df)
    expected   = n_students / n_hr_classes
    col_v1, col_v2, col_v3 = st.columns(3)
    col_v1.metric("전체 학생 수", f"{n_students}명")
    col_v2.metric("반당 예상 인원", f"{expected:.1f}명")
    if expected > max_per_hr:
        col_v3.metric("⚠️ 최대 인원 초과 위험", f"{expected:.0f} > {max_per_hr}명")
    else:
        col_v3.metric("✅ 인원 여유", f"반당 {max_per_hr - expected:.1f}명")

    st.divider()
    hr_run_btn = st.button("🏠 본반 편성 시작", type="primary", use_container_width=True)

    if hr_run_btn:
        with st.spinner("본반 편성 중... (분리 조건 + 이동 최소화 + 남녀 균형)"):
            hr_result_df, hr_stats_df, cohesion_rate, hr_assignments = assign_homeroom_classes(
                df=raw_df, id_col=id_col, subjects=subjects,
                gender_col=gender_col, n_classes=n_hr_classes,
                max_per_class=max_per_hr, gender_ratio_weight=gender_weight,
                sep_pairs=sep_pairs, seed=seed,
            )

        # 본반 분리 위반 확인
        hr_assign_map = {sid: cls for sid, cls in hr_assignments.items()}
        hr_violations = check_separation_violations(hr_assign_map, sep_pairs, label='본반')

        st.session_state.update({
            'hr_result_df':   hr_result_df,
            'hr_stats_df':    hr_stats_df,
            'hr_assignments': hr_assignments,
            'hr_cohesion':    cohesion_rate,
            'hr_n_classes':   n_hr_classes,
            'hr_violations':  hr_violations,
            'has_homeroom_result': True,
        })
        if hr_violations:
            st.warning(f"⚠️ 본반 편성 완료 — 분리 위반 {len(hr_violations)}건 발생")
        else:
            st.success("✅ 본반 편성 완료! 분리 조건 모두 충족")

    # ── 결과 표시 ─────────────────────────────────────────────────────────
    if st.session_state.get('has_homeroom_result'):
        hr_result_df = st.session_state['hr_result_df']
        hr_stats_df  = st.session_state['hr_stats_df']
        cohesion     = st.session_state['hr_cohesion']
        n_classes    = st.session_state['hr_n_classes']
        hr_assign    = st.session_state['hr_assignments']
        hr_viol      = st.session_state.get('hr_violations', [])

        st.divider()
        st.subheader("📊 본반 편성 결과")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("편성된 반 수", f"{n_classes}반")
        m2.metric("과목 패턴 응집도", f"{cohesion*100:.1f}%")
        m3.metric("🚫 분리 위반", f"{len(hr_viol)}건",
                  delta_color="inverse" if hr_viol else "off")
        if gender_col:
            ratios = []
            for _, row in hr_stats_df.iterrows():
                if row['총원'] > 0:
                    ratios.append(row['남'] / row['총원'])
            if ratios:
                m4.metric("반간 남비율 표준편차", f"{pd.Series(ratios).std()*100:.1f}%p")
        else:
            m4.metric("성별 정보", "없음")

        if hr_viol:
            with st.expander(f"⚠️ 본반 분리 조건 위반 {len(hr_viol)}건"):
                st.dataframe(pd.DataFrame(hr_viol), use_container_width=True, hide_index=True)
                st.caption("반 수를 늘리거나 분리 조건을 재검토해주세요.")
        elif sep_pairs:
            st.success(f"✅ 분리 조건 {len(sep_pairs)}쌍 모두 충족")

        st.subheader("반별 현황")

        def color_gender(val):
            if isinstance(val, str) and '%' in val:
                try:
                    v = float(val.replace('%', ''))
                    if v > 65 or v < 35:
                        return 'color: #CC0000; font-weight: bold'
                except:
                    pass
            return ''

        st.dataframe(
            hr_stats_df.style.applymap(color_gender, subset=['남비율']),
            use_container_width=True, hide_index=True,
        )
        st.caption("남비율 35%~65% 범위 벗어나면 빨간색 표시")

        chart_col1, chart_col2 = st.columns(2)
        with chart_col1:
            st.caption("반별 총원")
            st.bar_chart({row['반']: row['총원'] for _, row in hr_stats_df.iterrows()})
        with chart_col2:
            if gender_col:
                st.caption("반별 남녀 인원")
                st.bar_chart(pd.DataFrame({
                    '남': hr_stats_df.set_index('반')['남'],
                    '여': hr_stats_df.set_index('반')['여'],
                }))

        st.subheader("학생별 본반 배정 결과")
        search_hr = st.text_input("🔍 학생 검색", placeholder="학번 일부 입력...", key="hr_search")
        disp_hr   = hr_result_df.copy()
        if search_hr:
            disp_hr = disp_hr[disp_hr[id_col].astype(str).str.contains(search_hr, na=False)]
        st.caption(f"{len(disp_hr)}명 표시 중")
        st.dataframe(disp_hr, use_container_width=True, hide_index=True, height=400)

        st.divider()
        st.subheader("📥 본반 편성 결과 다운로드")
        with st.spinner("Excel 생성 중..."):
            hr_excel = build_homeroom_excel(
                hr_result_df, hr_stats_df, id_col, n_classes, hr_assign,
                raw_df, gender_col, subjects,
                sep_violations=hr_viol,
            )
        hr_csv = hr_result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📊 본반 편성 Excel 다운로드", data=hr_excel,
                file_name="본반편성결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )
        with dl2:
            st.download_button(
                "📄 본반 편성 CSV 다운로드", data=hr_csv,
                file_name="본반편성결과.csv", mime="text/csv",
                use_container_width=True,
            )
