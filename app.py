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

NON_SUBJECT_COLS = {'순번', '번호', '신학번', '학번', '이름', '성명', '학년', '반', '번', 'no', 'id'}


def get_id_col(df: pd.DataFrame) -> str:
    for c in ['신학번', '학번', '번호', 'ID', 'id']:
        if c in df.columns:
            return c
    return df.columns[0]


def detect_subjects(df: pd.DataFrame) -> list[str]:
    """값이 0/1인 컬럼을 과목으로 간주"""
    result = []
    for col in df.columns:
        if col.lower() in NON_SUBJECT_COLS:
            continue
        vals = set(df[col].dropna().astype(str).str.strip().unique())
        vals.discard('')  # 빈 셀은 0으로 간주
        if vals <= {'0', '1', '0.0', '1.0', '0.', '1.'}:
            result.append(col)
    return result


def extract_homeroom(sid: str) -> str:
    """신학번 5자리 → 본반: 20103 → '1반'"""
    s = str(sid).strip()
    if len(s) == 5 and s.isdigit():
        return str(int(s[1:3])) + '반'
    return ''


def get_time_labels(n: int) -> list[str]:
    return [chr(65 + i) for i in range(n)]


def abbr(name: str, length: int = 5) -> str:
    return name[:length]


# ──────────────────────────────────────────────────────────────────────────────
# 편성 로직 (매개변수화)
# ──────────────────────────────────────────────────────────────────────────────

def run_ilp(subjects, n_sections, max_per_time, times):
    """
    ILP: 각 과목의 타임별 분반 수 결정
    타임당 분반 수는 [floor, ceil] 범위 내에서 균형 배분
    """
    total = sum(n_sections[s] for s in subjects)
    nt    = len(times)
    lo, hi = total // nt, (total + nt - 1) // nt

    prob = pulp.LpProblem("section_time", pulp.LpMinimize)
    z = {
        s: {t: pulp.LpVariable(f"z_{i}_{t}", 0, n_sections[s], cat='Integer')
            for t in times}
        for i, s in enumerate(subjects)
    }

    prob += 0  # 실현 가능성만 확인

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
    """분반 목록: {과목: {타임: [분반번호, ...]}}"""
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
                    times, max_per_section, homeroom=None, seed=42):
    """
    학생 → 분반 배정 (capacity-aware greedy)
    1순위: 모든 분반에 여석이 있는 타임 순열 중 최소 부하
    2순위: 여석 없으면 부하 최소 (overflow 최소화 fallback)
    """
    homeroom = homeroom or {}
    random.seed(seed)
    n_sel = len(times)  # 학생당 선택해야 할 과목 수

    sec_load   = {(s, k): 0 for s in subjects for t in times for k in sections[s][t]}
    sec_roster = defaultdict(list)
    records    = []
    failed     = []

    rows = df.to_dict('records')
    random.shuffle(rows)

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
            cost = sum(
                min(sec_load[(s, k)] for k in sections[s][tm[s]])
                for s in chosen
            )
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

        for s in chosen:
            t = final_map[s]
            with_room = [k for k in sections[s][t] if sec_load[(s, k)] < max_per_section]
            pool = with_room if with_room else sections[s][t]
            k = min(pool, key=lambda x: sec_load[(s, x)])
            sec_load[(s, k)] += 1
            sec_roster[(s, k)].append(sid)
            rec[f'{t}타임_과목'] = s
            rec[f'{t}타임_분반'] = f"{abbr(s)}-{k}반"

        records.append(rec)

    rec_df    = pd.DataFrame(records) if records else pd.DataFrame(columns=[id_col])
    base      = df[[id_col]].copy()
    base[id_col] = base[id_col].astype(str)
    result_df = base.merge(rec_df, on=id_col, how='left') if not rec_df.empty else base

    return result_df, sec_roster, sec_load, failed


# ──────────────────────────────────────────────────────────────────────────────
# Excel 출력
# ──────────────────────────────────────────────────────────────────────────────

_H_FILL  = PatternFill("solid", fgColor="2E4699")
_H_FONT  = Font(bold=True, color="FFFFFF", size=10)
_H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
_S_FILL  = PatternFill("solid", fgColor="5B9BD5")


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
                n_total, failed, max_per_section):
    wb = Workbook()
    wb.remove(wb.active)

    # ── 개요 ──────────────────────────────────────────────────────────────
    ws = wb.create_sheet("개요")
    ws.cell(1, 1, f"이동반 편성 결과 — 총 {n_total}명 / 충족 {n_total-len(failed)}명 / 미배정 {len(failed)}명")
    ws.cell(1, 1).font = Font(bold=True, size=12)
    ws.merge_cells(f'A1:G1')
    ws.row_dimensions[1].height = 22

    r = 3
    for ci, h in enumerate(['타임', '과목', '분반명', '배정인원', '최대인원', '상태'], 1):
        _hdr(ws, r, ci, h)
    ws.row_dimensions[r].height = 18
    r += 1

    for t in times:
        _subhdr(ws, r, 1, f"▶ {t}타임")
        ws.merge_cells(f'A{r}:G{r}')
        ws.row_dimensions[r].height = 16
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt   = sec_load.get((s, k), 0)
                state = "★초과" if cnt > max_per_section else "정상"
                ws.cell(r, 1, t)
                ws.cell(r, 2, s)
                ws.cell(r, 3, f"{abbr(s)}-{k}반")
                ws.cell(r, 4, cnt)
                ws.cell(r, 5, max_per_section)
                c6 = ws.cell(r, 6, state)
                if state != "정상":
                    c6.font = Font(bold=True, color="FF0000")
                r += 1
        r += 1
    _auto_width(ws)

    # ── 학생 시간표 ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("학생시간표")
    t_cols = [f'{t}타임_과목' for t in times] + [f'{t}타임_분반' for t in times]
    cols   = [id_col, '본반'] + [f'{t}타임_과목' for t in times] + [f'{t}타임_분반' for t in times]
    for ci, h in enumerate(cols, 1):
        _hdr(ws2, 1, ci, h)
    ws2.row_dimensions[1].height = 18
    ws2.freeze_panes = 'A2'
    for _, row in result_df.iterrows():
        ws2.append([str(row.get(c, '') or '') for c in cols])
    _auto_width(ws2)

    # ── 분반별 명단 ────────────────────────────────────────────────────────
    for s in subjects:
        for t in times:
            for k in sections[s][t]:
                sname = f"{abbr(s)}-{k}반"
                ws_s  = wb.create_sheet(sname)
                ws_s.cell(1, 1, f"[{s}]  {k}분반  ({t}타임)").font = Font(bold=True, size=11)
                ws_s.merge_cells('A1:C1')
                ws_s.row_dimensions[1].height = 22
                for ci, h in enumerate([id_col, '본반', '비고'], 1):
                    _hdr(ws_s, 2, ci, h)
                ws_s.row_dimensions[2].height = 18
                for sid in sorted(sec_roster.get((s, k), [])):
                    ws_s.append([sid, homeroom.get(sid, ''), ''])
                _auto_width(ws_s)

    # ── 교실 배치 추천 ─────────────────────────────────────────────────────
    ws_cr = wb.create_sheet("교실배치추천")
    ws_cr.cell(1, 1, "교실 배치 추천 — 본반 우선순위 기준").font = Font(bold=True, size=12)
    ws_cr.merge_cells('A1:F1')
    ws_cr.row_dimensions[1].height = 22

    r = 3
    for t in times:
        _subhdr(ws_cr, r, 1, f"▶ {t}타임")
        ws_cr.merge_cells(f'A{r}:F{r}')
        ws_cr.row_dimensions[r].height = 16
        r += 1
        for ci, h in enumerate(['과목', '분반명', '배정인원', '본반 구성 (우선순위)', '권장교실', '비고'], 1):
            _hdr(ws_cr, r, ci, h)
        ws_cr.row_dimensions[r].height = 18
        r += 1
        for s in subjects:
            for k in sections[s][t]:
                cnt    = sec_load.get((s, k), 0)
                hr_cnt = defaultdict(int)
                for sid in sec_roster.get((s, k), []):
                    hr = homeroom.get(sid, '')
                    if hr:
                        hr_cnt[hr] += 1
                top = sorted(hr_cnt.items(), key=lambda x: (-x[1], x[0]))
                priority = '  >  '.join(f"{hr} {c}명" for hr, c in top) if top else ''
                ws_cr.cell(r, 1, s)
                ws_cr.cell(r, 2, f"{abbr(s)}-{k}반")
                ws_cr.cell(r, 3, cnt)
                ws_cr.cell(r, 4, priority).alignment = Alignment(horizontal='left')
                r += 1
        r += 1

    ws_cr.column_dimensions['A'].width = 22
    ws_cr.column_dimensions['B'].width = 14
    ws_cr.column_dimensions['C'].width = 10
    ws_cr.column_dimensions['D'].width = 65
    ws_cr.column_dimensions['E'].width = 14

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
    st.error("⚠️ PuLP 라이브러리가 없습니다. 터미널에서 `pip install pulp` 후 재시작해주세요.")
    st.stop()

# ── 사이드바 ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 기본 설정")
    n_times = st.number_input("타임 수", min_value=2, max_value=6, value=3,
                               help="예: 3이면 A·B·C 타임으로 자동 생성")
    times = get_time_labels(n_times)
    st.info(f"타임: **{' / '.join(times)}**")

    max_per_section = st.number_input("분반당 최대 인원", min_value=10, max_value=100, value=41)
    seed = int(st.number_input("무작위 시드", min_value=0, max_value=9999, value=42,
                                help="같은 시드는 같은 결과를 재현합니다"))

    st.divider()
    st.caption("※ 설정 변경 후 편성 시작을 다시 눌러주세요.")

# ── 메인 탭 ───────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["① CSV 업로드", "② 과목 설정 및 편성", "③ 편성 결과"])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1: 파일 업로드
# ════════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("학생 선택과목 CSV 업로드")
    st.markdown("""
    **파일 형식:**
    - 컬럼: `신학번`(또는 `학번`) + 과목명들
    - 과목 선택 여부: `1`(선택) / `0`(미선택)
    - 학생마다 정확히 **타임 수**만큼 과목을 선택해야 합니다
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

            for s in subjects_auto:
                raw_df[s] = pd.to_numeric(raw_df[s], errors='coerce').fillna(0).astype(int)

            if not subjects_auto:
                st.error("과목 컬럼을 자동 감지하지 못했습니다. CSV 구조를 확인해주세요.")
                st.stop()

            # 세션 저장
            st.session_state['raw_df']          = raw_df
            st.session_state['subjects_auto']   = subjects_auto
            st.session_state['id_col']          = id_col
            st.session_state['settings_init']   = False  # 재업로드 시 설정 초기화

            # 통계 표시
            n_sel = raw_df[subjects_auto].sum(axis=1)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("총 학생 수", f"{len(raw_df)}명")
            c2.metric("감지된 과목 수", f"{len(subjects_auto)}개")
            c3.metric("학생당 선택 과목", f"{n_sel.mean():.1f}개 (평균)")
            bad_cnt = int((n_sel != n_times).sum())
            c4.metric(f"{n_times}과목 미준수", f"{bad_cnt}명",
                      delta_color="inverse" if bad_cnt else "off")

            if bad_cnt:
                st.warning(f"⚠️ {n_times}과목을 선택하지 않은 학생이 {bad_cnt}명입니다. 편성 시 미배정 처리됩니다.")
            else:
                st.success(f"✅ 전원 {n_times}과목 선택 확인")

            # 과목별 신청 인원
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
# TAB 2: 과목 설정 및 편성 시작
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    if 'raw_df' not in st.session_state:
        st.info("먼저 [① CSV 업로드] 탭에서 파일을 업로드해주세요.")
    else:
        subjects_auto = st.session_state['subjects_auto']

        st.subheader("과목별 분반·교사 수 설정")
        st.markdown("""
        - **학급수**: 해당 과목의 전체 분반 수 (A·B·C 타임 합산)
        - **교사수**: 같은 타임에 최대 몇 분반까지 동시 운영할 수 있는지 (= 교사 수)
        - 행을 **추가·삭제**할 수 있습니다 (표 하단 + 버튼 / 행 선택 후 Delete)
        """)

        # 초기 설정 테이블 (업로드 직후 한 번만 초기화)
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

        # 검증 정보
        valid = edited.dropna(subset=['과목명'])
        total_sections = int(valid['학급수(총 분반 수)'].fillna(0).sum())
        lo = total_sections // n_times
        hi = (total_sections + n_times - 1) // n_times

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("총 분반 수", total_sections)
        col_b.metric("타임 수", n_times)
        col_c.metric("타임당 분반", f"{lo}~{hi}개")

        if total_sections == 0:
            st.error("학급수 합계가 0입니다. 설정을 확인해주세요.")
        elif total_sections % n_times != 0:
            st.warning(f"⚠️ {total_sections}개 분반을 {n_times}타임에 균등 배분할 수 없습니다. "
                       f"ILP가 {lo}~{hi}개 범위에서 최대한 균형을 맞춥니다.")
        else:
            st.success(f"✅ 타임당 정확히 {total_sections // n_times}개 분반 배정 예정")

        st.divider()

        run_btn = st.button("🚀 편성 시작", type="primary",
                            use_container_width=True,
                            disabled=(total_sections == 0))

        if run_btn:
            subjects   = valid['과목명'].str.strip().tolist()
            n_sections = dict(zip(valid['과목명'].str.strip(),
                                  valid['학급수(총 분반 수)'].astype(int)))
            mpt        = dict(zip(valid['과목명'].str.strip(),
                                  valid['교사수(타임당 최대)'].astype(int)))
            raw_df     = st.session_state['raw_df']
            id_col     = st.session_state['id_col']

            with st.spinner("① ILP로 분반 타임 배정 중..."):
                assignment, ilp_status = run_ilp(subjects, n_sections, mpt, times)

            if assignment is None:
                st.error(f"❌ ILP 해를 찾지 못했습니다 (상태: {ilp_status}). "
                         "학급수·교사수·타임 수 설정을 확인해주세요.")
            else:
                sections = build_sections(subjects, assignment, times)

                # 본반 추출
                homeroom = {}
                if id_col in raw_df.columns:
                    homeroom = {str(r[id_col]): extract_homeroom(str(r[id_col]))
                                for _, r in raw_df.iterrows()}

                with st.spinner("② 학생 분반 배정 중..."):
                    result_df, sec_roster, sec_load, failed = assign_students(
                        raw_df, id_col, subjects, assignment, sections,
                        times, int(max_per_section), homeroom, seed=seed,
                    )

                st.session_state.update({
                    'assignment':  assignment,
                    'sections':    sections,
                    'sec_load':    sec_load,
                    'sec_roster':  sec_roster,
                    'result_df':   result_df,
                    'failed':      failed,
                    'subjects':    subjects,
                    'homeroom':    homeroom,
                    'n_total':     len(raw_df),
                    'id_col':      id_col,
                    'has_result':  True,
                })

                n_ok  = len(raw_df) - len(failed)
                n_ov  = sum(1 for v in sec_load.values() if v > max_per_section)
                if n_ov == 0 and len(failed) == 0:
                    st.success(f"✅ 편성 완료! {n_ok}/{len(raw_df)}명 전원 배정, 정원 초과 없음 → [③ 편성 결과] 탭 확인")
                elif n_ov == 0:
                    st.warning(f"⚠️ 편성 완료 — 미배정 {len(failed)}명 발생 → [③ 편성 결과] 탭 확인")
                else:
                    st.warning(f"⚠️ 편성 완료 — 정원 초과 분반 {n_ov}개 발생 → [③ 편성 결과] 탭 확인")

# ════════════════════════════════════════════════════════════════════════════════
# TAB 3: 결과
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
    n_ok       = n_total - len(failed)
    n_ov       = sum(1 for v in sec_load.values() if v > max_per_section)

    # ── 요약 지표 ────────────────────────────────────────────────────────────
    st.subheader("📊 편성 결과 요약")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("총 학생", f"{n_total}명")
    m2.metric("배정 완료", f"{n_ok}명", f"{100*n_ok/n_total:.1f}%")
    m3.metric("미배정", f"{len(failed)}명",
              delta=f"-{len(failed)}" if failed else "없음",
              delta_color="inverse" if failed else "off")
    m4.metric(f"정원({int(max_per_section)}명) 초과 분반", f"{n_ov}개",
              delta_color="inverse" if n_ov else "off")

    if failed:
        with st.expander(f"⚠️ 미배정 학생 {len(failed)}명 목록"):
            st.write(failed)
    if n_ov == 0 and not failed:
        st.success("✅ 전원 배정 완료, 모든 분반 정원 이내")

    st.divider()

    # ── 결과 탭 ─────────────────────────────────────────────────────────────
    rt1, rt2, rt3 = st.tabs(["📋 타임 배정표", "👥 분반별 인원", "📄 학생 시간표"])

    with rt1:
        st.subheader("과목별 타임 배정")
        rows = []
        for s in subjects:
            row = {'과목': s, '총 분반': sum(assignment[s][t] for t in times)}
            for t in times:
                cnt  = assignment[s][t]
                secs = ', '.join(f"{abbr(s)}-{k}반" for k in sections[s][t])
                row[f'{t}타임'] = f"{cnt}분반 ({secs})" if cnt else "—"
            rows.append(row)

        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # 타임별 분반 수 합산
        st.caption("타임별 총 분반 수")
        totals = {t: sum(assignment[s][t] for s in subjects) for t in times}
        st.bar_chart(totals)

    with rt2:
        st.subheader("분반별 배정 인원")
        rows = []
        for t in times:
            for s in subjects:
                for k in sections[s][t]:
                    n     = sec_load.get((s, k), 0)
                    state = "★ 초과" if n > max_per_section else "✓"
                    rows.append({'타임': t, '과목': s,
                                 '분반명': f"{abbr(s)}-{k}반",
                                 '배정인원': n,
                                 '상태': state})

        load_df = pd.DataFrame(rows)

        def _row_color(row):
            c = '#FFE0E0' if row['상태'] == '★ 초과' else ''
            return [f'background-color: {c}'] * len(row)

        st.dataframe(
            load_df.style.apply(_row_color, axis=1),
            use_container_width=True, hide_index=True,
        )

        # 과목별 분반 인원 분포 차트
        st.caption("과목별 분반 인원 분포")
        chart_data = {}
        for s in subjects:
            for t in times:
                for k in sections[s][t]:
                    chart_data[f"{abbr(s)}-{k}반({t})"] = sec_load.get((s, k), 0)
        st.bar_chart(chart_data)

    with rt3:
        st.subheader("학생별 시간표")

        # 검색 필터
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            search = st.text_input("🔍 신학번 검색", placeholder="신학번 일부 입력...")
        with col_f2:
            filter_time = st.selectbox("타임 필터 (과목 포함)", ["전체"] + times)

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

    # ── 다운로드 ─────────────────────────────────────────────────────────────
    st.subheader("📥 결과 다운로드")

    with st.spinner("Excel 파일 생성 중..."):
        excel_buf = build_excel(
            subjects, times, id_col, assignment, sections,
            sec_load, sec_roster, result_df, homeroom,
            n_total, failed, int(max_per_section),
        )

    csv_bytes = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="📊 Excel 다운로드 (전체 — 개요·시간표·분반명단·교실배치)",
            data=excel_buf,
            file_name="이동반편성결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
    with dl2:
        st.download_button(
            label="📄 CSV 다운로드 (학생 시간표만)",
            data=csv_bytes,
            file_name="학생시간표.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.caption("Excel 파일 구성: 개요 / 학생시간표 / 분반별 명단(30개 시트) / 교실배치추천")
