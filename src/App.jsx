import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import IconAlertTriangle from "@tabler/icons-react/dist/esm/icons/IconAlertTriangle.mjs";
import IconArmchair from "@tabler/icons-react/dist/esm/icons/IconArmchair.mjs";
import IconBook from "@tabler/icons-react/dist/esm/icons/IconBook.mjs";
import IconBuildingBank from "@tabler/icons-react/dist/esm/icons/IconBuildingBank.mjs";
import IconCalendar from "@tabler/icons-react/dist/esm/icons/IconCalendar.mjs";
import IconCheck from "@tabler/icons-react/dist/esm/icons/IconCheck.mjs";
import IconChevronRight from "@tabler/icons-react/dist/esm/icons/IconChevronRight.mjs";
import IconClipboardCheck from "@tabler/icons-react/dist/esm/icons/IconClipboardCheck.mjs";
import IconClock from "@tabler/icons-react/dist/esm/icons/IconClock.mjs";
import IconDeviceFloppy from "@tabler/icons-react/dist/esm/icons/IconDeviceFloppy.mjs";
import IconDownload from "@tabler/icons-react/dist/esm/icons/IconDownload.mjs";
import IconFileSpreadsheet from "@tabler/icons-react/dist/esm/icons/IconFileSpreadsheet.mjs";
import IconLayoutGrid from "@tabler/icons-react/dist/esm/icons/IconLayoutGrid.mjs";
import IconListCheck from "@tabler/icons-react/dist/esm/icons/IconListCheck.mjs";
import IconLock from "@tabler/icons-react/dist/esm/icons/IconLock.mjs";
import IconLogout from "@tabler/icons-react/dist/esm/icons/IconLogout.mjs";
import IconPlus from "@tabler/icons-react/dist/esm/icons/IconPlus.mjs";
import IconPrinter from "@tabler/icons-react/dist/esm/icons/IconPrinter.mjs";
import IconRefresh from "@tabler/icons-react/dist/esm/icons/IconRefresh.mjs";
import IconSettings from "@tabler/icons-react/dist/esm/icons/IconSettings.mjs";
import IconShieldLock from "@tabler/icons-react/dist/esm/icons/IconShieldLock.mjs";
import IconTrash from "@tabler/icons-react/dist/esm/icons/IconTrash.mjs";
import IconUpload from "@tabler/icons-react/dist/esm/icons/IconUpload.mjs";
import IconUserCheck from "@tabler/icons-react/dist/esm/icons/IconUserCheck.mjs";
import IconUsers from "@tabler/icons-react/dist/esm/icons/IconUsers.mjs";
import IconX from "@tabler/icons-react/dist/esm/icons/IconX.mjs";
import {
  ABSENCE_REASONS,
  aggregateAbsenceReport,
  analyzeEnrollmentRows,
  analyzeManualSeatRows,
  analyzeStudentRows,
  buildPersonalTimetable,
  buildSeatSlots,
  classLabel,
  formatKoreanDate,
  formatKoreanWeekday,
  formatPresenceLabel,
  generateSeatAssignment,
  makeId,
  resolveStudentRoom,
  seatChartKey,
  sortStudents,
  summarizeAbsences,
  toCsvSafeFileName,
} from "./core.js";
import {
  apiRequest,
  appMode,
  clearSessions,
  hasAdminSession,
  hasSchoolSession,
  saveAdminSession,
  saveSchoolSession,
} from "./api.js";

const PERIODS = [1, 2, 3, 4, 5, 6, 7];
const EMPTY_BOOTSTRAP = {
  settings: {},
  classes: [],
  exam_dates: [],
  students: [],
  timetable: [],
  enrollments: [],
  seat_charts: [],
};

function messageFrom(error) {
  return error instanceof Error ? error.message : String(error || "오류가 발생했습니다.");
}

function importErrorMessage(prefix, errors = []) {
  const labels = {
    INVALID_HEADERS: "열 이름 오류",
    INVALID_STUDENT_ROW: "학생 값 오류",
    DUPLICATE_STUDENT_NUMBER: "학급 내 번호 중복",
    INVALID_ENROLLMENT_ROW: "선택과목 값 오류",
    DUPLICATE_ENROLLMENT: "선택과목 중복",
    INVALID_SEAT_ROW: "수동 명단 값 오류",
    DUPLICATE_SEAT_STUDENT: "수동 명단 학생 중복",
    STUDENT_NOT_FOUND: "등록 학생과 불일치",
    CLASS_NOT_FOUND: "학급 불일치",
    STUDENT_NAME_MISMATCH: "학생 이름 불일치",
    SUBJECT_NOT_FOUND: "선택과목 없음",
    SUBJECT_CLASS_MISMATCH: "과목 적용 학급 아님",
  };
  const details = errors
    .slice(0, 10)
    .map((item) => `${item.row_number}행 ${labels[item.code] || item.code}`)
    .join(", ");
  return `${prefix}: ${details}${errors.length > 10 ? ` 외 ${errors.length - 10}건` : ""}`;
}

function downloadWorkbook(fileName, sheets) {
  if (!window.XLSX) throw new Error("엑셀 모듈을 불러오지 못했습니다.");
  const workbook = window.XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const worksheet = window.XLSX.utils.aoa_to_sheet(rows);
    window.XLSX.utils.book_append_sheet(workbook, worksheet, name.slice(0, 31));
  });
  window.XLSX.writeFile(workbook, `${toCsvSafeFileName(fileName)}.xlsx`);
}

async function readWorkbookRows(file) {
  if (!window.XLSX) throw new Error("엑셀 모듈을 불러오지 못했습니다.");
  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return window.XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "" });
}

function selectFirst(items, key = "id") {
  return items?.[0]?.[key] || "";
}

function EmptyState({ icon: Icon = IconClipboardCheck, title, description }) {
  return (
    <div className="empty-state">
      <span className="empty-icon"><Icon size={28} stroke={1.6} /></span>
      <strong>{title}</strong>
      <p>{description}</p>
    </div>
  );
}

function Notice({ tone = "info", children }) {
  if (!children) return null;
  return (
    <div className={`notice notice-${tone}`} role={tone === "error" ? "alert" : "status"}>
      {tone === "error" ? <IconAlertTriangle size={18} /> : <IconCheck size={18} />}
      <span>{children}</span>
    </div>
  );
}

function LoadingButton({ loading, children, ...props }) {
  return (
    <button {...props} disabled={loading || props.disabled}>
      {loading ? <span className="spinner" aria-hidden="true" /> : null}
      {children}
    </button>
  );
}

function LoginScreen({ onLogin }) {
  const [schoolCode, setSchoolCode] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setLoading(true);
    setError("");
    try {
      const result = await apiRequest("school_login", { school_code: schoolCode.trim() });
      saveSchoolSession(result.session);
      onLogin(result.bootstrap);
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login-page">
      <div className="login-ornament ornament-one" />
      <div className="login-ornament ornament-two" />
      <section className="login-card">
        <div className="brand-seal"><IconBook size={34} stroke={1.5} /></div>
        <p className="eyebrow">SCHOOL EXAM OPERATIONS</p>
        <h1>정기고사<br />관리 시스템</h1>
        <p className="login-copy">결시 입력부터 고사본부 현황, 자리배치와 출력까지 한 곳에서 관리합니다.</p>
        <form onSubmit={submit}>
          <label htmlFor="school-code">학교코드</label>
          <input
            id="school-code"
            value={schoolCode}
            onChange={(event) => setSchoolCode(event.target.value)}
            placeholder="학교에서 안내한 코드를 입력하세요"
            autoComplete="off"
            autoCapitalize="characters"
            required
          />
          <LoadingButton className="button button-primary button-wide" loading={loading} type="submit">
            접속하기 <IconChevronRight size={18} />
          </LoadingButton>
        </form>
        <Notice tone="error">{error}</Notice>
        {appMode.demo ? (
          <div className="demo-hint">
            <strong>데모 모드</strong>
            <span>학교코드는 아무 값이나, 관리자 암호는 <code>demo-admin</code>을 사용하세요.</span>
          </div>
        ) : (
          <a className="demo-link" href="?demo=1">실제 데이터가 없는 데모로 둘러보기</a>
        )}
        <p className="privacy-note"><IconShieldLock size={15} /> 입력 정보는 학교 소유 시트에만 저장됩니다.</p>
      </section>
    </main>
  );
}

function AdminDialog({ open, onClose, onUnlock }) {
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const passwordRef = useRef(null);
  const dialogRef = useRef(null);
  const returnFocusRef = useRef(null);

  useEffect(() => {
    if (!open) return undefined;
    returnFocusRef.current = document.activeElement;
    setPassword("");
    setError("");
    const focusTimer = setTimeout(() => passwordRef.current?.focus(), 50);
    const handleKeyDown = (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
        return;
      }
      if (event.key !== "Tab") return;
      const focusable = [...dialogRef.current?.querySelectorAll(
        'button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), a[href], [tabindex]:not([tabindex="-1"])',
      ) || []];
      if (!focusable.length) return;
      const first = focusable[0];
      const last = focusable.at(-1);
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      clearTimeout(focusTimer);
      document.removeEventListener("keydown", handleKeyDown);
      setTimeout(() => returnFocusRef.current?.focus(), 0);
    };
  }, [open]);

  if (!open) return null;

  async function submit(event) {
    event.preventDefault();
    setLoading(true);
    setError("");
    try {
      const result = await apiRequest("admin_login", { password });
      saveAdminSession(result.admin_session);
      onUnlock(result.bootstrap);
      onClose();
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={(event) => {
      if (event.target === event.currentTarget) onClose();
    }}>
      <section ref={dialogRef} className="dialog" role="dialog" aria-modal="true" aria-labelledby="admin-title">
        <button className="icon-button dialog-close" onClick={onClose} aria-label="닫기">
          <IconX size={20} />
        </button>
        <span className="dialog-icon"><IconLock size={26} /></span>
        <p className="eyebrow">ADMIN ACCESS</p>
        <h2 id="admin-title">관리자 확인</h2>
        <p>고사본부·자리배치·설정은 관리자 암호가 필요합니다.</p>
        <form onSubmit={submit}>
          <label htmlFor="admin-password">관리자 암호</label>
          <input
            ref={passwordRef}
            id="admin-password"
            type="password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            autoComplete="current-password"
            required
          />
          <LoadingButton className="button button-primary button-wide" loading={loading} type="submit">
            관리자 화면 열기
          </LoadingButton>
        </form>
        <Notice tone="error">{error}</Notice>
      </section>
    </div>
  );
}

function SectionHeader({ eyebrow, title, description, actions }) {
  return (
    <header className="section-header">
      <div>
        <p className="eyebrow">{eyebrow}</p>
        <h2>{title}</h2>
        {description ? <p>{description}</p> : null}
      </div>
      {actions ? <div className="section-actions no-print">{actions}</div> : null}
    </header>
  );
}

function AbsencePanel({ bootstrap }) {
  const [examDate, setExamDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [classId, setClassId] = useState(selectFirst(bootstrap.classes));
  const [period, setPeriod] = useState(1);
  const [context, setContext] = useState({
    students: [],
    absences: [],
    completed: false,
    submitted: false,
    revision: 0,
  });
  const [loaded, setLoaded] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [confirmSubmit, setConfirmSubmit] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    if (!examDate && bootstrap.exam_dates.length) setExamDate(bootstrap.exam_dates[0].exam_date);
    if (!classId && bootstrap.classes.length) setClassId(bootstrap.classes[0].id);
  }, [bootstrap.exam_dates, bootstrap.classes, examDate, classId]);

  function resetLoaded(next) {
    next();
    setLoaded(false);
    setDirty(false);
    setConfirmSubmit(false);
    setMessage("");
    setError("");
    setContext({ students: [], absences: [], completed: false, submitted: false, revision: 0 });
  }

  const load = useCallback(async () => {
    if (!examDate || !classId || !period) {
      setError("고사일, 학급, 교시를 모두 선택하세요.");
      return;
    }
    setLoading(true);
    setMessage("");
    setError("");
    try {
      const result = await apiRequest("get_absence_context", {
        exam_date: examDate,
        class_id: classId,
        period,
      });
      setContext({
        students: result.students || [],
        absences: result.absences || [],
        completed: Boolean(result.completed ?? result.submitted),
        submitted: Boolean(result.submitted ?? result.completed),
        revision: Number(result.revision || 0),
      });
      setLoaded(true);
      setDirty(false);
      setConfirmSubmit(false);
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }, [examDate, classId, period]);

  const absenceMap = useMemo(
    () => new Map(context.absences.map((item) => [String(item.student_id), item])),
    [context.absences],
  );

  function updateStudent(student, patch) {
    const current = absenceMap.get(String(student.id));
    const nextReason = patch.reason ?? current?.reason ?? "미인정";
    const next = {
      absent: patch.absent ?? Boolean(current),
      reason: nextReason,
      reason_detail: nextReason === "기타"
        ? patch.reason_detail ?? current?.reason_detail ?? ""
        : "",
    };
    setMessage("");
    setError("");
    setDirty(true);
    setConfirmSubmit(false);
    setContext((value) => ({
      ...value,
      completed: false,
      submitted: false,
      absences: value.absences
        .filter((item) => String(item.student_id) !== String(student.id))
        .concat(next.absent ? [{
          ...current,
          student_id: student.id,
          class_id: classId,
          exam_date: examDate,
          period,
          reason: next.reason,
          reason_detail: next.reason_detail,
        }] : []),
    }));
  }

  async function submitAll() {
    const missingDetail = context.absences.find(
      (item) => item.reason === "기타" && !String(item.reason_detail || "").trim(),
    );
    if (missingDetail) {
      const student = context.students.find((item) => String(item.id) === String(missingDetail.student_id));
      setError(`${student?.number || ""}번 ${student?.name || "학생"}의 기타 상세 사유를 입력하세요.`);
      return;
    }
    if (!confirmSubmit) {
      setConfirmSubmit(true);
      setMessage(`제출 전 확인: 결시 ${context.absences.length}명 명단을 고사본부에 제출합니다.`);
      setError("");
      return;
    }
    setLoading(true);
    setMessage("");
    setError("");
    try {
      const result = await apiRequest("submit_absence_input", {
        exam_date: examDate,
        class_id: classId,
        period,
        expected_revision: Number(context.revision || 0),
        absences: context.absences.map((item) => ({
          student_id: item.student_id,
          reason: item.reason,
          reason_detail: item.reason === "기타" ? String(item.reason_detail || "").trim() : "",
        })),
      });
      setContext((value) => ({
        ...value,
        absences: result.absences || value.absences,
        completed: true,
        submitted: true,
        revision: Number(result.revision ?? value.revision + 1),
      }));
      setDirty(false);
      setConfirmSubmit(false);
      setMessage("결시 명단을 일괄 제출했습니다. 다시 수정해 제출하면 최신 내용으로 교체됩니다.");
    } catch (nextError) {
      if (nextError?.code === "REVISION_CONFLICT") {
        await load();
        setError("다른 사용자가 같은 명단을 먼저 수정했습니다. 최신 상태를 다시 불러왔으니 확인 후 제출하세요.");
      } else {
        setError(messageFrom(nextError));
      }
      setConfirmSubmit(false);
    } finally {
      setLoading(false);
    }
  }

  const selectedClass = bootstrap.classes.find((item) => String(item.id) === String(classId));

  return (
    <section className="page-section print-section">
      <SectionHeader
        eyebrow="TEACHER DESK"
        title="결시 학생 입력"
        description="고사일·학급·교시를 선택해 명단을 불러온 뒤, 결시자와 사유를 한 번에 제출하세요."
        actions={<button className="button button-light" onClick={() => window.print()} disabled={!loaded}><IconPrinter size={17} /> 응시현황표 인쇄</button>}
      />
      <div className="filter-bar no-print">
        <label>고사일
          <select value={examDate} onChange={(event) => resetLoaded(() => setExamDate(event.target.value))}>
            {bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {formatKoreanDate(item.exam_date)}</option>)}
          </select>
        </label>
        <label>학급
          <select value={classId} onChange={(event) => resetLoaded(() => setClassId(event.target.value))}>
            {bootstrap.classes.map((item) => <option key={item.id} value={item.id}>{classLabel(item)}</option>)}
          </select>
        </label>
        <label>교시
          <select value={period} onChange={(event) => resetLoaded(() => setPeriod(Number(event.target.value)))}>
            {PERIODS.map((item) => <option key={item} value={item}>{item}교시</option>)}
          </select>
        </label>
        <div className="filter-action">
          <LoadingButton className="button button-primary" loading={loading} onClick={load}>
            <IconRefresh size={17} /> 명단 불러오기
          </LoadingButton>
        </div>
      </div>

      <div className="print-heading">
        <h3>{formatKoreanDate(examDate)} {period}교시 응시현황표</h3>
        <p>{selectedClass ? classLabel(selectedClass) : ""} · 재적 {context.students.length}명 · 결시 {context.absences.length}명</p>
      </div>
      <Notice>{message}</Notice>
      <Notice tone="error">{error}</Notice>

      {loading && !loaded ? (
        <div className="loading-block"><span className="spinner spinner-dark" /> 명단을 불러오는 중입니다.</div>
      ) : loaded && context.students.length ? (
        <div className="student-list">
          {context.students.map((student) => {
            const absence = absenceMap.get(String(student.id));
            return (
              <article className={`student-row ${absence ? "is-absent" : ""}`} key={student.id}>
                <button
                  className="absence-toggle no-print"
                  onClick={() => updateStudent(student, { absent: !absence })}
                  aria-pressed={Boolean(absence)}
                  aria-label={`${student.number}번 ${student.name} ${absence ? "결시 해제" : "결시 선택"}`}
                >
                  {absence ? <IconCheck size={17} /> : null}
                </button>
                <div className="student-identity">
                  <span>{student.number}번</span>
                  <strong>{student.name}</strong>
                </div>
                <span className={`status-pill ${absence ? "status-absent" : "status-present"}`}>
                  {absence ? "결시" : "응시"}
                </span>
                {absence ? (
                  <div className="reason-fields no-print">
                    <select
                      value={absence.reason}
                      onChange={(event) => updateStudent(student, { reason: event.target.value })}
                      aria-label={`${student.name} 결시 사유`}
                    >
                      {ABSENCE_REASONS.map((reason) => <option key={reason}>{reason}</option>)}
                    </select>
                    {absence.reason === "기타" ? (
                      <input
                        value={absence.reason_detail || ""}
                        placeholder="상세 사유(필수)"
                        aria-label={`${student.name} 상세 사유`}
                        onChange={(event) => updateStudent(student, { reason_detail: event.target.value })}
                      />
                    ) : null}
                  </div>
                ) : null}
                <span className="print-only">{absence ? `${absence.reason}${absence.reason_detail ? `(${absence.reason_detail})` : ""}` : "응시"}</span>
              </article>
            );
          })}
        </div>
      ) : loaded ? (
        <EmptyState icon={IconUsers} title="등록된 학생이 없습니다" description="관리자 설정에서 학급과 학생 명단을 먼저 등록하세요." />
      ) : (
        <EmptyState icon={IconListCheck} title="결시 명단을 불러오세요" description="고사일, 학급, 교시를 확인한 뒤 ‘명단 불러오기’를 누르세요." />
      )}

      {loaded ? <div className="completion-bar no-print">
        <div>
          <span className={context.submitted && !dirty ? "completion-dot done" : "completion-dot"} />
          <div>
            <strong>{context.submitted && !dirty ? "제출 완료" : dirty ? "수정 중" : "입력 확인 전"}</strong>
            <p>현재 명단 버전 {context.revision} · 결시 {context.absences.length}명</p>
          </div>
        </div>
        <div className="completion-actions">
          {confirmSubmit ? (
            <button
              className="button button-light"
              type="button"
              onClick={() => {
                setConfirmSubmit(false);
                setMessage("");
              }}
            >
              취소
            </button>
          ) : null}
          <LoadingButton className="button button-gold" loading={loading} onClick={submitAll} disabled={!context.students.length}>
            <IconClipboardCheck size={18} /> {confirmSubmit ? "확인하고 제출" : context.revision ? "수정 내용 다시 제출" : "결시 명단 일괄 제출"}
          </LoadingButton>
        </div>
      </div> : null}
    </section>
  );
}

function HqPanel({ bootstrap }) {
  const [examDate, setExamDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [period, setPeriod] = useState(1);
  const [grade, setGrade] = useState("");
  const [classId, setClassId] = useState("");
  const [status, setStatus] = useState({ absences: [], completions: [], summary: null });
  const [report, setReport] = useState({
    absences: [],
    reason_summary: [],
    class_summary: [],
    date_summary: [],
    student_summary: [],
  });
  const [startDate, setStartDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [endDate, setEndDate] = useState(bootstrap.exam_dates.at(-1)?.exam_date || "");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const grades = [...new Set(bootstrap.classes.map((item) => Number(item.grade)))];

  useEffect(() => {
    if (!examDate && bootstrap.exam_dates.length) {
      setExamDate(bootstrap.exam_dates[0].exam_date);
      setStartDate(bootstrap.exam_dates[0].exam_date);
      setEndDate(bootstrap.exam_dates.at(-1).exam_date);
    }
  }, [bootstrap.exam_dates, examDate]);

  const refresh = useCallback(async (quiet = false) => {
    if (!examDate) return;
    if (!quiet) setLoading(true);
    try {
      const result = await apiRequest("get_hq_status", {
        exam_date: examDate,
        period,
        grade: grade || "",
        class_id: classId || "",
      });
      setStatus(result);
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      if (!quiet) setLoading(false);
    }
  }, [examDate, period, grade, classId]);

  useEffect(() => {
    refresh();
    const tick = () => {
      if (document.visibilityState === "visible") refresh(true);
    };
    const timer = setInterval(tick, 30_000);
    const onVisibility = () => {
      if (document.visibilityState === "visible") refresh(true);
    };
    document.addEventListener("visibilitychange", onVisibility);
    return () => {
      clearInterval(timer);
      document.removeEventListener("visibilitychange", onVisibility);
    };
  }, [refresh]);

  const filteredClasses = bootstrap.classes.filter(
    (item) => (!grade || Number(item.grade) === Number(grade)) && (!classId || String(item.id) === String(classId)),
  );
  const completionKeys = new Set(
    status.completions
      .filter(
        (item) =>
          item.exam_date === examDate &&
          Number(item.period) === Number(period) &&
          Boolean(item.submitted ?? item.completed_at),
      )
      .map((item) => String(item.class_id)),
  );
  const filteredClassIds = new Set(filteredClasses.map((item) => String(item.id)));
  const enrolled = Number(
    status.summary?.enrolled ??
      bootstrap.students.filter((item) => filteredClassIds.has(String(item.class_id))).length,
  );
  const absent = Number(status.summary?.absent ?? status.absences.length);
  const present = Number(status.summary?.present ?? Math.max(0, enrolled - absent));
  const reasonSummary = status.summary?.reasons || summarizeAbsences(status.absences);

  async function loadReport() {
    setLoading(true);
    try {
      const result = await apiRequest("get_period_report", {
        start_date: startDate,
        end_date: endDate,
        grade: grade || "",
        class_id: classId || "",
      });
      const fallback = aggregateAbsenceReport(result.absences || []);
      setReport({
        absences: result.absences || [],
        reason_summary: result.reason_summary || fallback.reason_summary,
        class_summary: result.class_summary || fallback.class_summary,
        date_summary: result.date_summary || fallback.date_summary,
        student_summary: result.student_summary || fallback.student_summary,
      });
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }

  function exportCurrentStatus() {
    downloadWorkbook("교시별 결시·제출 현황", {
      결시현황: [
        ["고사일", "교시", "학년", "반", "번호", "이름", "사유", "세부사유", "수정시각"],
        ...status.absences.map((item) => [
          item.exam_date, item.period, item.grade, item.class_num, item.student_number,
          item.student_name, item.reason, item.reason_detail || "", item.updated_at || "",
        ]),
      ],
      제출현황: [
        ["고사일", "교시", "학급", "제출상태", "제출시각", "버전"],
        ...filteredClasses.map((classInfo) => {
          const completion = status.completions.find(
            (item) =>
              String(item.class_id) === String(classInfo.id) &&
              Boolean(item.submitted ?? item.completed_at),
          );
          return [
            examDate,
            period,
            classLabel(classInfo),
            completion ? "완료" : "대기",
            completion?.completed_at || "",
            completion?.revision || "",
          ];
        }),
      ],
      교시합계: [
        ["구분", "인원"],
        ["재적", enrolled],
        ["응시", present],
        ["결시", absent],
        ...ABSENCE_REASONS.map((reason) => [reason, reasonSummary[reason] || 0]),
      ],
    });
  }

  function exportPeriodReport() {
    downloadWorkbook("기간 결시 통계", {
      결시상세: [
        ["고사일", "교시", "학년", "반", "번호", "이름", "사유", "세부사유"],
        ...report.absences.map((item) => [
          item.exam_date, item.period, item.grade, item.class_num,
          item.student_number, item.student_name, item.reason, item.reason_detail || "",
        ]),
      ],
      사유별: [["사유", "건수"], ...report.reason_summary.map((item) => [item.reason, item.count])],
      학급별: [["학급", "건수"], ...report.class_summary.map((item) => [item.class_label, item.count])],
      일자별: [["고사일", "건수"], ...report.date_summary.map((item) => [item.exam_date, item.count])],
      학생별: [
        ["학급", "번호", "이름", "건수"],
        ...report.student_summary.map((item) => [
          item.class_label, item.student_number, item.student_name, item.count,
        ]),
      ],
    });
  }

  return (
    <section className="page-section print-section">
      <SectionHeader
        eyebrow="EXAM HEADQUARTERS"
        title="고사본부 현황"
        description="교시별 결시 현황과 학급 제출 여부를 확인합니다. 이 화면에서만 30초마다 자동 갱신됩니다."
        actions={
          <>
            <button className="button button-light" onClick={() => refresh()}><IconRefresh size={17} /> 새로고침</button>
            <button className="button button-light" onClick={exportCurrentStatus}><IconDownload size={17} /> 교시 현황 엑셀</button>
          </>
        }
      />
      <div className="filter-bar no-print">
        <label>고사일
          <select value={examDate} onChange={(event) => setExamDate(event.target.value)}>
            {bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {formatKoreanDate(item.exam_date)}</option>)}
          </select>
        </label>
        <label>교시
          <select value={period} onChange={(event) => setPeriod(Number(event.target.value))}>
            {PERIODS.map((item) => <option key={item} value={item}>{item}교시</option>)}
          </select>
        </label>
        <label>학년
          <select value={grade} onChange={(event) => { setGrade(event.target.value); setClassId(""); }}>
            <option value="">전체 학년</option>
            {grades.map((item) => <option key={item} value={item}>{item}학년</option>)}
          </select>
        </label>
        <label>학급
          <select value={classId} onChange={(event) => setClassId(event.target.value)}>
            <option value="">전체 학급</option>
            {bootstrap.classes.filter((item) => !grade || Number(item.grade) === Number(grade)).map((item) => (
              <option key={item.id} value={item.id}>{classLabel(item)}</option>
            ))}
          </select>
        </label>
      </div>
      <Notice tone="error">{error}</Notice>

      <div className="metric-grid">
        <article><span>재적</span><strong>{enrolled}<small>명</small></strong></article>
        <article><span>응시</span><strong>{present}<small>명</small></strong></article>
        <article><span>결시</span><strong>{absent}<small>명</small></strong></article>
      </div>
      <div className="reason-summary" aria-label="사유별 결시 인원">
        {ABSENCE_REASONS.map((reason) => (
          <span key={reason}><strong>{reason}</strong> {reasonSummary[reason] || 0}명</span>
        ))}
      </div>

      <div className="split-grid">
        <article className="panel-card">
          <div className="card-heading"><div><IconListCheck size={20} /><h3>결시 학생</h3></div><span>{loading ? "갱신 중" : "최신 상태"}</span></div>
          {status.absences.length ? (
            <div className="table-wrap">
              <table>
                <thead><tr><th>학급</th><th>번호</th><th>이름</th><th>사유</th><th>상세</th></tr></thead>
                <tbody>{status.absences.map((item) => (
                  <tr key={item.id || `${item.student_id}-${item.period}`}>
                    <td>{item.class_label}</td><td>{item.student_number}</td><td>{item.student_name}</td><td><span className="status-pill status-absent">{item.reason}</span></td><td>{item.reason_detail || "—"}</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          ) : <EmptyState title="현재 결시 학생이 없습니다" description="교사가 입력하면 이곳에 바로 표시됩니다." />}
        </article>
        <article className="panel-card">
          <div className="card-heading"><div><IconUserCheck size={20} /><h3>학급 제출 여부</h3></div><span>{completionKeys.size}/{filteredClasses.length}</span></div>
          <div className="completion-grid">
            {filteredClasses.map((item) => {
              const done = completionKeys.has(String(item.id));
              return <div className={done ? "class-status done" : "class-status"} key={item.id}><span>{classLabel(item)}</span><strong>{done ? "완료" : "대기"}</strong></div>;
            })}
          </div>
        </article>
      </div>

      <article className="panel-card report-card no-print">
        <div className="card-heading"><div><IconCalendar size={20} /><h3>기간 합산 통계</h3></div></div>
        <div className="inline-form">
          <label>시작일<input type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} /></label>
          <label>종료일<input type="date" value={endDate} onChange={(event) => setEndDate(event.target.value)} /></label>
          <button className="button button-secondary" onClick={loadReport}>통계 조회</button>
          <button className="button button-light" disabled={!report.absences.length} onClick={exportPeriodReport}><IconFileSpreadsheet size={17} /> 기간 통계 엑셀</button>
        </div>
        {report.absences.length ? (
          <div className="report-results">
            <p className="report-summary">선택 기간 결시 기록 <strong>{report.absences.length}건</strong></p>
            <div className="report-summary-grid">
              <div><h4>사유별</h4>{report.reason_summary.map((item) => <span key={item.reason}>{item.reason}<strong>{item.count}건</strong></span>)}</div>
              <div><h4>학급별</h4>{report.class_summary.map((item) => <span key={item.class_id || item.class_label}>{item.class_label}<strong>{item.count}건</strong></span>)}</div>
              <div><h4>일자별</h4>{report.date_summary.map((item) => <span key={item.exam_date}>{formatKoreanDate(item.exam_date)}<strong>{item.count}건</strong></span>)}</div>
              <div><h4>학생별</h4>{report.student_summary.map((item) => <span key={item.student_id}>{item.class_label} {item.student_number}번 {item.student_name}<strong>{item.count}건</strong></span>)}</div>
            </div>
          </div>
        ) : null}
      </article>
    </section>
  );
}

function SeatPanel({ bootstrap, onBootstrap }) {
  const initialClass = bootstrap.classes[0];
  const [mode, setMode] = useState("separate");
  const [examDate, setExamDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [grade, setGrade] = useState(Number(initialClass?.grade || 1));
  const [classId, setClassId] = useState(selectFirst(bootstrap.classes));
  const [period, setPeriod] = useState(1);
  const [subjectName, setSubjectName] = useState("");
  const [rows, setRows] = useState(6);
  const [cols, setCols] = useState(5);
  const [startSide, setStartSide] = useState("window");
  const [seatOrder, setSeatOrder] = useState("row");
  const [roomName, setRoomName] = useState("");
  const [disabledSeats, setDisabledSeats] = useState([]);
  const [selectedIds, setSelectedIds] = useState([]);
  const [manualPool, setManualPool] = useState([]);
  const [absentIds, setAbsentIds] = useState([]);
  const [result, setResult] = useState({ assignments: [], ok: true, error: "" });
  const [loadedChartId, setLoadedChartId] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const manualFileRef = useRef(null);
  const seatEditorRef = useRef(null);

  const classMap = useMemo(
    () => new Map(bootstrap.classes.map((item) => [String(item.id), item])),
    [bootstrap.classes],
  );
  const gradeClasses = useMemo(
    () => bootstrap.classes.filter((item) => Number(item.grade) === Number(grade)),
    [bootstrap.classes, grade],
  );
  const selectedClass = classMap.get(String(classId));
  const students = useMemo(
    () => sortStudents(bootstrap.students.filter((item) =>
      mode === "separate"
        ? Number(classMap.get(String(item.class_id))?.grade) === Number(grade)
        : String(item.class_id) === String(classId)
    )),
    [bootstrap.students, classMap, mode, grade, classId],
  );
  const timetableItems = useMemo(
    () => bootstrap.timetable.filter((item) =>
      item.exam_date === examDate &&
      Number(item.period) === Number(period) &&
      Number(item.grade) === Number(grade)
    ),
    [bootstrap.timetable, examDate, period, grade],
  );
  const selectedTimetable = timetableItems.find((item) => item.subject_name === subjectName);
  const roomOptions = useMemo(() => {
    if (mode === "own") {
      return selectedClass ? [classLabel(selectedClass)] : [];
    }
    const options = new Set();
    if (selectedTimetable?.room_name) options.add(selectedTimetable.room_name);
    bootstrap.enrollments
      .filter((item) => Number(item.grade) === Number(grade) && item.subject_name === subjectName)
      .forEach((item) => {
        if (item.room_name) options.add(item.room_name);
      });
    return [...options];
  }, [bootstrap.enrollments, grade, subjectName, selectedTimetable, mode, selectedClass]);
  const seatSequenceMap = useMemo(
    () => new Map(
      buildSeatSlots(rows, cols, startSide, disabledSeats, seatOrder)
        .map((slot, index) => [slot.key, index + 1]),
    ),
    [rows, cols, startSide, disabledSeats, seatOrder],
  );

  const defaultExamineeIds = useMemo(() => {
    if (!selectedTimetable) return mode === "own" ? students.map((item) => item.id) : [];
    const applicable = new Set(
      String(selectedTimetable.class_ids || "")
        .split(/[|,]/)
        .map((item) => item.trim())
        .filter(Boolean),
    );
    return students.filter((student) => {
      if (selectedTimetable.subject_type === "common") {
        return !applicable.size || applicable.has(String(student.class_id));
      }
      const enrollment = bootstrap.enrollments.find(
        (item) =>
          String(item.student_id) === String(student.id) &&
          item.subject_name === subjectName,
      );
      if (!enrollment) return false;
      if (mode === "own") return true;
      if (!roomName) return true;
      return resolveStudentRoom({
        studentId: student.id,
        subjectName,
        timetableItem: selectedTimetable,
        enrollments: bootstrap.enrollments,
        classInfo: classMap.get(String(student.class_id)),
      }) === roomName;
    }).map((item) => item.id);
  }, [
    selectedTimetable,
    mode,
    students,
    bootstrap.enrollments,
    subjectName,
    roomName,
    classMap,
  ]);

  useEffect(() => {
    if (!examDate && bootstrap.exam_dates.length) setExamDate(bootstrap.exam_dates[0].exam_date);
    if (!classId && bootstrap.classes.length) setClassId(bootstrap.classes[0].id);
  }, [bootstrap.exam_dates, bootstrap.classes, examDate, classId]);

  useEffect(() => {
    if (loadedChartId) return;
    if (!gradeClasses.some((item) => String(item.id) === String(classId))) {
      setClassId(gradeClasses[0]?.id || "");
    }
  }, [gradeClasses, classId, loadedChartId]);

  useEffect(() => {
    if (loadedChartId) return;
    if (!timetableItems.some((item) => item.subject_name === subjectName)) {
      setSubjectName(timetableItems[0]?.subject_name || "");
    }
  }, [timetableItems, subjectName, loadedChartId]);

  useEffect(() => {
    if (loadedChartId) return;
    if (!roomName || (roomOptions.length && !roomOptions.includes(roomName))) {
      setRoomName(roomOptions[0] || selectedTimetable?.room_name || (mode === "own" && selectedClass ? classLabel(selectedClass) : ""));
    }
  }, [roomOptions, selectedTimetable, roomName, mode, selectedClass, loadedChartId]);

  useEffect(() => {
    if (loadedChartId || manualPool.length) return;
    setSelectedIds(defaultExamineeIds);
    setResult({ assignments: [], ok: true, error: "" });
  }, [defaultExamineeIds, loadedChartId, manualPool.length]);

  useEffect(() => {
    if (!examDate) return;
    const request = mode === "separate"
      ? apiRequest("get_hq_status", { exam_date: examDate, period, grade, class_id: "" })
      : apiRequest("get_absence_context", { exam_date: examDate, class_id: classId, period });
    request
      .then((value) => setAbsentIds((value.absences || []).map((item) => item.student_id)))
      .catch(() => setAbsentIds([]));
  }, [mode, examDate, grade, classId, period]);

  function changeSeatContext(change) {
    setLoadedChartId("");
    setManualPool([]);
    setResult({ assignments: [], ok: true, error: "" });
    change();
  }

  function invalidateSeatResult() {
    setLoadedChartId("");
    setResult({ assignments: [], ok: true, error: "" });
  }

  async function createChart() {
    if (mode === "separate" && (!subjectName.trim() || !roomName.trim())) {
      setError("별실 자리배치는 시간표 과목과 고사실을 모두 선택하세요.");
      return;
    }
    let latestAbsentIds = absentIds;
    try {
      const current = mode === "separate"
        ? await apiRequest("get_hq_status", { exam_date: examDate, period, grade, class_id: "" })
        : await apiRequest("get_absence_context", { exam_date: examDate, class_id: classId, period });
      latestAbsentIds = (current.absences || []).map((item) => item.student_id);
      setAbsentIds(latestAbsentIds);
    } catch (nextError) {
      setError(`최신 결시 명단을 확인하지 못했습니다. ${messageFrom(nextError)}`);
      return;
    }
    const sourceStudents = manualPool.length ? manualPool : students;
    const next = generateSeatAssignment({
      students: sourceStudents,
      rows,
      cols,
      startSide,
      seatOrder,
      disabledSeats,
      mode,
      selectedIds,
      absentIds: latestAbsentIds,
    });
    setResult(next);
    setError(next.error);
    if (next.ok) {
      setMessage(
        mode === "own"
          ? "응시자와 결시자를 앞에, 미응시자를 뒤에 배치했습니다."
          : "학년 전체 명단에서 선택한 과목·호실 응시자를 배치했습니다.",
      );
    }
  }

  function toggleSeat(key) {
    invalidateSeatResult();
    setDisabledSeats((items) => items.includes(key) ? items.filter((item) => item !== key) : [...items, key]);
  }

  function moveSeatFocus(event, index) {
    const offsets = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -cols, ArrowDown: cols };
    if (!(event.key in offsets)) return;
    event.preventDefault();
    const next = Math.min(rows * cols - 1, Math.max(0, index + offsets[event.key]));
    seatEditorRef.current?.querySelector(`[data-seat-index="${next}"]`)?.focus();
  }

  function chartPayload() {
    const classValue = mode === "own" ? classId : "";
    const base = {
      id: loadedChartId || undefined,
      mode,
      exam_date: examDate,
      period,
      grade,
      class_id: classValue,
      subject_name: subjectName,
      room_name: roomName,
      rows,
      cols,
      start_side: startSide,
      seat_order: seatOrder,
      disabled_seats: disabledSeats,
      examinee_ids: selectedIds,
      assignment: result.assignments.map((item) => ({
        key: item.key,
        sequence: item.sequence,
        student_id: item.student?.id || "",
        absent: item.absent,
      })),
      updated_at: new Date().toISOString(),
    };
    if (!base.id) {
      base.id = bootstrap.seat_charts.find((item) => seatChartKey(item) === seatChartKey(base))?.id;
    }
    return base;
  }

  async function saveChart() {
    if (!result.assignments.length) return setError("자리배치를 먼저 생성하세요.");
    if (mode === "separate" && (!subjectName.trim() || !roomName.trim())) {
      setError("별실 자리배치는 시간표 과목과 고사실을 모두 선택하세요.");
      return;
    }
    try {
      const response = await apiRequest("save_seat_chart", { chart: chartPayload() });
      onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
      setLoadedChartId(response.chart?.id || loadedChartId);
      setMessage(response.updated_count
        ? "같은 고사·과목·호실 자리배치를 최신 내용으로 갱신했습니다."
        : "자리배치를 새로 저장했습니다.");
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    }
  }

  function loadChart(chart) {
    setLoadedChartId(chart.id);
    setMode(chart.mode);
    setExamDate(chart.exam_date);
    setPeriod(Number(chart.period));
    setGrade(Number(chart.grade));
    setClassId(chart.class_id || bootstrap.classes.find((item) => Number(item.grade) === Number(chart.grade))?.id || "");
    setSubjectName(chart.subject_name || "");
    setRoomName(chart.room_name || "");
    setRows(Number(chart.rows || 6));
    setCols(Number(chart.cols || 5));
    setStartSide(chart.start_side || "window");
    setSeatOrder(chart.seat_order === "column" ? "column" : "row");
    setDisabledSeats(chart.disabled_seats || []);
    setSelectedIds(chart.examinee_ids || []);
    const studentMap = new Map(bootstrap.students.map((item) => [String(item.id), item]));
    setResult({
      ok: true,
      error: "",
      assignments: (chart.assignment || []).map((item) => ({
        ...item,
        student: item.student_id ? studentMap.get(String(item.student_id)) : null,
      })),
    });
    setMessage("저장된 자리배치를 불러왔습니다.");
  }

  async function deleteChart(chart) {
    if (!window.confirm(`'${chart.subject_name || "과목 미지정"} · ${chart.room_name}' 자리배치를 삭제할까요?`)) return;
    try {
      const response = await apiRequest("delete_seat_chart", { id: chart.id });
      onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
      if (loadedChartId === chart.id) setLoadedChartId("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    }
  }

  async function importManualList(file) {
    if (!file) return;
    try {
      const analyzed = analyzeManualSeatRows(await readWorkbookRows(file));
      if (analyzed.errors.length) {
        throw new Error(importErrorMessage("수동 명단을 가져오지 않았습니다", analyzed.errors));
      }
      const rowsFromFile = analyzed.rows;
      const classByNumber = new Map(
        gradeClasses.map((item) => [Number(item.class_num), item]),
      );
      const unmatched = [];
      const matched = rowsFromFile.map((row) => {
        const classInfo = classByNumber.get(Number(row.class_num));
        const student = bootstrap.students.find((candidate) =>
          classInfo &&
          String(candidate.class_id) === String(classInfo.id) &&
          Number(candidate.number) === Number(row.number) &&
          candidate.name === row.name
        );
        if (!student) {
          unmatched.push({ row_number: row.row_number, code: "STUDENT_NOT_FOUND" });
        }
        return student;
      }).filter(Boolean);
      if (unmatched.length) {
        throw new Error(importErrorMessage("수동 명단을 가져오지 않았습니다", unmatched));
      }
      if (!matched.length) throw new Error("학급·번호·이름이 등록 명단과 일치하는 학생이 없습니다.");
      setLoadedChartId("");
      setMode("separate");
      setManualPool(matched);
      setSelectedIds([...new Set(matched.map((item) => item.id))]);
      setMessage(`수동 명단에서 ${matched.length}명을 선택했습니다.`);
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      if (manualFileRef.current) manualFileRef.current.value = "";
    }
  }

  async function createElectiveBatch() {
    let batchAbsentIds = absentIds;
    try {
      const hq = await apiRequest("get_hq_status", { exam_date: examDate, period, grade, class_id: "" });
      batchAbsentIds = (hq.absences || []).map((item) => item.student_id);
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const subjectRows = bootstrap.timetable.filter((item) =>
      item.exam_date === examDate &&
      Number(item.period) === Number(period) &&
      Number(item.grade) === Number(grade) &&
      item.subject_type === "elective"
    );
    const subjectMap = new Map(subjectRows.map((item) => [item.subject_name, item]));
    const studentMap = new Map(bootstrap.students.map((item) => [String(item.id), item]));
    const groups = new Map();
    bootstrap.enrollments
      .filter((item) => Number(item.grade) === Number(grade) && subjectMap.has(item.subject_name))
      .forEach((item) => {
        const timetableItem = subjectMap.get(item.subject_name);
        const student = studentMap.get(String(item.student_id));
        if (!student) return;
        const effectiveRoom = resolveStudentRoom({
          studentId: student.id,
          subjectName: item.subject_name,
          timetableItem,
          enrollments: bootstrap.enrollments,
          classInfo: classMap.get(String(student.class_id)),
        }) || "별실";
        const key = `${item.subject_name}|${effectiveRoom}`;
        if (!groups.has(key)) groups.set(key, { subject_name: item.subject_name, room_name: effectiveRoom, students: [] });
        groups.get(key).students.push(student);
      });
    if (!groups.size) {
      setError("선택한 고사일·교시·학년에 연결된 선택과목 응시자 명단이 없습니다.");
      return;
    }
    const charts = [];
    for (const group of groups.values()) {
      const generated = generateSeatAssignment({
        students: group.students,
        rows,
        cols,
        startSide,
        seatOrder,
        disabledSeats,
        mode: "separate",
        selectedIds: group.students.map((item) => item.id),
        absentIds: batchAbsentIds,
      });
      if (!generated.ok) {
        setError(`${group.subject_name} · ${group.room_name}: ${generated.error}`);
        return;
      }
      charts.push({
        mode: "separate",
        exam_date: examDate,
        period,
        grade,
        class_id: "",
        subject_name: group.subject_name,
        room_name: group.room_name,
        rows,
        cols,
        start_side: startSide,
        seat_order: seatOrder,
        disabled_seats: disabledSeats,
        examinee_ids: group.students.map((item) => item.id),
        assignment: generated.assignments.map((item) => ({
          key: item.key,
          sequence: item.sequence,
          student_id: item.student?.id || "",
          absent: item.absent,
        })),
        updated_at: new Date().toISOString(),
      });
    }
    try {
      const response = await apiRequest("save_seat_charts_batch", { charts });
      onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
      setMessage(
        `선택과목·호실 기준 자리배치: 신규 ${response.created_count || 0}개 · 갱신 ${response.updated_count || 0}개`,
      );
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    }
  }

  const displayStudents = manualPool.length ? manualPool : students;
  const scopeLabel = mode === "separate"
    ? `${grade}학년 전체`
    : selectedClass ? classLabel(selectedClass) : "";

  return (
    <section className="page-section">
      <SectionHeader
        eyebrow="SEATING PLAN"
        title="자리배치"
        description="별실은 학년 전체, 각자 교실은 학급 전체를 기준으로 6행×5열 좌석을 만듭니다."
        actions={<button className="button button-light" onClick={() => window.print()} disabled={!result.assignments.length}><IconPrinter size={17} /> PDF·인쇄</button>}
      />
      <div className="segmented no-print">
        <button className={mode === "separate" ? "active" : ""} onClick={() => changeSeatContext(() => setMode("separate"))}><IconArmchair size={17} /> 별실 배치</button>
        <button className={mode === "own" ? "active" : ""} onClick={() => changeSeatContext(() => setMode("own"))}><IconLayoutGrid size={17} /> 각자 교실</button>
      </div>
      <div className="seat-batch-actions no-print">
        <button className="button button-light" onClick={() => downloadWorkbook("자리배치 수동 명단 양식", { 수동명단: [["반", "번호", "이름"], [1, 1, "홍길동"]] })}><IconDownload size={17} /> 수동 명단 양식</button>
        <button className="button button-light" onClick={() => manualFileRef.current?.click()}><IconUpload size={17} /> 수동 명단 가져오기</button>
        <input ref={manualFileRef} className="sr-only" type="file" accept=".xlsx,.xls" onChange={(event) => importManualList(event.target.files?.[0])} />
        <button className="button button-secondary" onClick={createElectiveBatch}><IconLayoutGrid size={17} /> 시간표 기반 일괄 생성</button>
      </div>
      <Notice>{message}</Notice><Notice tone="error">{error}</Notice>
      <div className="seat-layout">
        <aside className="seat-controls no-print">
          <div className="control-group">
            <h3>1. 고사·과목·호실</h3>
            <label>고사일<select value={examDate} onChange={(event) => changeSeatContext(() => setExamDate(event.target.value))}>{bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {formatKoreanDate(item.exam_date)}</option>)}</select></label>
            <div className="two-fields">
              {mode === "separate" ? (
                <label>학년<select value={grade} onChange={(event) => changeSeatContext(() => setGrade(Number(event.target.value)))}>{[...new Set(bootstrap.classes.map((item) => Number(item.grade)))].map((item) => <option key={item}>{item}학년</option>)}</select></label>
              ) : (
                <label>학급<select value={classId} onChange={(event) => { const id = event.target.value; changeSeatContext(() => { setClassId(id); setGrade(Number(classMap.get(String(id))?.grade || grade)); }); }}>{bootstrap.classes.map((item) => <option key={item.id} value={item.id}>{classLabel(item)}</option>)}</select></label>
              )}
              <label>교시<select value={period} onChange={(event) => changeSeatContext(() => setPeriod(Number(event.target.value)))}>{PERIODS.map((item) => <option key={item} value={item}>{item}교시</option>)}</select></label>
            </div>
            <label>과목<select value={subjectName} onChange={(event) => changeSeatContext(() => setSubjectName(event.target.value))}><option value="">과목 미지정</option>{timetableItems.map((item) => <option key={item.id || `${item.subject_name}-${item.subject_type}`} value={item.subject_name}>{item.subject_name} · {item.subject_type === "common" ? "공통" : "선택"}</option>)}</select></label>
            <label>고사실<input list="seat-room-options" value={roomName} onChange={(event) => changeSeatContext(() => setRoomName(event.target.value))} placeholder="호실을 입력하세요" /></label>
            <datalist id="seat-room-options">{roomOptions.map((item) => <option key={item} value={item} />)}</datalist>
          </div>
          <div className="control-group">
            <h3>2. 좌석 구성</h3>
            <div className="three-fields">
              <label>행<input type="number" min="1" max="12" value={rows} onChange={(event) => { invalidateSeatResult(); setRows(Number(event.target.value)); }} /></label>
              <label>열<input type="number" min="1" max="12" value={cols} onChange={(event) => { invalidateSeatResult(); setCols(Number(event.target.value)); }} /></label>
              <label>시작<select value={startSide} onChange={(event) => { invalidateSeatResult(); setStartSide(event.target.value); }}><option value="window">창가</option><option value="aisle">복도</option></select></label>
            </div>
            <label>배치 순서<select value={seatOrder} onChange={(event) => { invalidateSeatResult(); setSeatOrder(event.target.value); }}><option value="row">가로 순번 · 행 우선 (→)</option><option value="column">세로 순번 · 열 우선 (↓)</option></select></label>
            <p className="field-help">{seatOrder === "column" ? "한 열을 칠판 쪽부터 채운 뒤 다음 열로 이동합니다." : "한 행을 창가·복도 방향으로 채운 뒤 다음 행으로 이동합니다."} Enter·Space로 좌석을 제외하고 방향키로 이동할 수 있습니다.</p>
            <div ref={seatEditorRef} className="mini-seat-grid" style={{ "--seat-cols": cols }}>
              {Array.from({ length: rows * cols }, (_, index) => {
                const row = Math.floor(index / cols);
                const col = index % cols;
                const key = `${row}-${col}`;
                const sequence = seatSequenceMap.get(key);
                return <button type="button" data-seat-index={index} key={key} className={disabledSeats.includes(key) ? "disabled" : ""} onKeyDown={(event) => moveSeatFocus(event, index)} onClick={() => toggleSeat(key)} aria-pressed={disabledSeats.includes(key)} aria-label={`${row + 1}행 ${col + 1}열 ${disabledSeats.includes(key) ? "복원" : `배치 ${sequence}번 · 제외`}`}>{disabledSeats.includes(key) ? <IconX size={13} /> : sequence}</button>;
              })}
            </div>
          </div>
          <div className="control-group">
            <h3>3. 응시자 선택 <span>{selectedIds.length}명</span></h3>
            <div className="selection-actions"><button onClick={() => { invalidateSeatResult(); setSelectedIds(displayStudents.map((item) => item.id)); }}>전체</button><button onClick={() => { invalidateSeatResult(); setSelectedIds([]); }}>해제</button></div>
            {manualPool.length ? <p className="manual-pool-note">수동 명단 {manualPool.length}명 선택됨</p> : null}
            <div className="check-list">
              {displayStudents.map((student) => <label key={student.id}><input type="checkbox" checked={selectedIds.map(String).includes(String(student.id))} onChange={() => { invalidateSeatResult(); setSelectedIds((items) => items.map(String).includes(String(student.id)) ? items.filter((id) => String(id) !== String(student.id)) : [...items, student.id]); }} /><span>{mode === "separate" ? `${classLabel(classMap.get(String(student.class_id)))} ` : ""}{student.number}번 {student.name}</span>{absentIds.map(String).includes(String(student.id)) ? <em>결시</em> : null}</label>)}
            </div>
          </div>
          <button className="button button-primary button-wide" onClick={createChart}><IconArmchair size={18} /> 자리배치 생성</button>
        </aside>

        <div className="seat-preview print-section">
          <div className="seat-title">
            <p>{formatKoreanDate(examDate)} · {period}교시</p>
            <h3>{subjectName || "과목 미지정"} · {roomName || "호실 미지정"}</h3>
            <span>{scopeLabel} · 응시 {selectedIds.length}명 · {rows}행×{cols}열 · {seatOrder === "column" ? "세로 순번" : "가로 순번"}</span>
          </div>
          <div className="room-orientation"><span className={startSide === "window" ? "active" : ""}>창가</span><strong>칠판 · 교탁</strong><span className={startSide === "aisle" ? "active" : ""}>복도</span></div>
          {result.assignments.length ? (
            <div className="seat-grid" style={{ "--seat-cols": cols }}>
              {Array.from({ length: rows * cols }, (_, index) => {
                const row = Math.floor(index / cols);
                const col = index % cols;
                const key = `${row}-${col}`;
                if (disabledSeats.includes(key)) return <div className="seat-hole" key={key} />;
                const item = result.assignments.find((slot) => slot.key === key);
                const studentClass = item?.student ? classMap.get(String(item.student.class_id)) : null;
                return (
                  <div className={`seat ${item?.absent ? "seat-absent" : ""} ${!item?.student ? "seat-empty" : ""}`} key={key}>
                    <span>{item?.sequence || seatSequenceMap.get(key)}번 · {row + 1}행 {col + 1}열</span>
                    {item?.student ? <><strong>{mode === "separate" && studentClass ? `${studentClass.class_num}반 ` : ""}{item.student.number}번 {item.student.name}</strong>{item.absent ? <em>결시 · 자리 유지</em> : !item.selected && mode === "own" ? <small>미응시</small> : null}</> : <small>빈 좌석</small>}
                  </div>
                );
              })}
            </div>
          ) : <EmptyState icon={IconArmchair} title="아직 생성된 자리배치가 없습니다" description="과목·호실·응시자와 좌석 구성을 확인한 뒤 자리배치를 만드세요." />}
          {result.assignments.length ? <div className="seat-preview-actions no-print"><button className="button button-gold" onClick={saveChart}><IconDeviceFloppy size={17} /> 저장</button><button className="button button-light" onClick={() => window.print()}><IconPrinter size={17} /> PDF·인쇄</button></div> : null}
        </div>
      </div>

      <article className="panel-card saved-charts no-print">
        <div className="card-heading"><div><IconDeviceFloppy size={20} /><h3>저장된 자리배치</h3></div><span>{bootstrap.seat_charts.length}개</span></div>
        {bootstrap.seat_charts.length ? <div className="saved-list">{bootstrap.seat_charts.map((chart) => (
          <div key={chart.id}><div><strong>{chart.subject_name || "과목 미지정"} · {chart.room_name}</strong><span>{formatKoreanDate(chart.exam_date)} · {chart.period}교시 · {chart.class_id ? classLabel(bootstrap.classes.find((item) => item.id === chart.class_id)) : `${chart.grade}학년`} · {chart.seat_order === "column" ? "세로 순번" : "가로 순번"}</span></div><div><button className="button button-light" onClick={() => loadChart(chart)}>불러오기</button><button className="icon-button danger" onClick={() => deleteChart(chart)} aria-label={`${chart.room_name} 자리배치 삭제`}><IconTrash size={17} /></button></div></div>
        ))}</div> : <EmptyState icon={IconDeviceFloppy} title="저장된 배치가 없습니다" description="같은 고사·과목·호실 조합은 중복 없이 최신 배치로 저장됩니다." />}
      </article>
    </section>
  );
}

function SetupPanel({ bootstrap, onBootstrap }) {
  const [active, setActive] = useState("dates");
  const [newDate, setNewDate] = useState("");
  const [newDateLabel, setNewDateLabel] = useState("");
  const [grade, setGrade] = useState(1);
  const [classNum, setClassNum] = useState(1);
  const [selectedClassId, setSelectedClassId] = useState(selectFirst(bootstrap.classes));
  const [studentText, setStudentText] = useState("");
  const [studentEdits, setStudentEdits] = useState({});
  const [newStudent, setNewStudent] = useState({ number: "", name: "" });
  const [timetableDraft, setTimetableDraft] = useState({
    exam_date: selectFirst(bootstrap.exam_dates, "exam_date"),
    grade: 1,
    period: 1,
    start_time: "09:00",
    end_time: "09:50",
    subject_name: "",
    subject_type: "common",
    room_name: "",
    class_ids: [],
  });
  const [personalStudentId, setPersonalStudentId] = useState(selectFirst(bootstrap.students));
  const [printScope, setPrintScope] = useState("student");
  const [printGrade, setPrintGrade] = useState(Number(bootstrap.classes[0]?.grade || 1));
  const [printClassId, setPrintClassId] = useState(selectFirst(bootstrap.classes));
  const [cleanupDate, setCleanupDate] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const studentFileRef = useRef(null);
  const enrollmentFileRef = useRef(null);

  useEffect(() => {
    if (!selectedClassId && bootstrap.classes.length) setSelectedClassId(bootstrap.classes[0].id);
    if (!personalStudentId && bootstrap.students.length) setPersonalStudentId(bootstrap.students[0].id);
    if (!timetableDraft.exam_date && bootstrap.exam_dates.length) {
      setTimetableDraft((value) => ({ ...value, exam_date: bootstrap.exam_dates[0].exam_date }));
    }
  }, [bootstrap.classes, bootstrap.students, bootstrap.exam_dates, selectedClassId, personalStudentId, timetableDraft.exam_date]);

  useEffect(() => {
    const rows = bootstrap.students
      .filter((item) => String(item.class_id) === String(selectedClassId))
      .sort((a, b) => Number(a.number) - Number(b.number));
    setStudentText(rows.map((item) => `${item.number}\t${item.name}`).join("\n"));
    setStudentEdits(Object.fromEntries(rows.map((item) => [
      String(item.id),
      { number: item.number, name: item.name },
    ])));
  }, [bootstrap.students, selectedClassId]);

  useEffect(() => {
    const available = bootstrap.classes
      .filter((item) => Number(item.grade) === Number(timetableDraft.grade))
      .map((item) => item.id);
    setTimetableDraft((value) => ({
      ...value,
      class_ids: value.class_ids.filter((id) => available.includes(id)).length
        ? value.class_ids.filter((id) => available.includes(id))
        : available,
    }));
  }, [bootstrap.classes, timetableDraft.grade]);

  async function commit(action, payload, success) {
    setMessage("");
    setError("");
    try {
      const result = await apiRequest(action, payload);
      onBootstrap({ ...bootstrap, ...result });
      setMessage(success);
      return result;
    } catch (nextError) {
      setError(messageFrom(nextError));
      return null;
    }
  }

  async function addDate() {
    if (!newDate) return setError("고사일을 선택하세요.");
    const existing = bootstrap.exam_dates.find((item) => item.exam_date === newDate);
    const next = [...bootstrap.exam_dates.filter((item) => item.exam_date !== newDate), {
      id: existing?.id || makeId("date"),
      exam_date: newDate,
      label: newDateLabel.trim() || existing?.label || `${bootstrap.exam_dates.length + 1}일차`,
      active: true,
    }].sort((a, b) => a.exam_date.localeCompare(b.exam_date));
    if (await commit("save_exam_dates", { exam_dates: next }, "고사일을 저장했습니다.")) {
      setNewDate(""); setNewDateLabel("");
    }
  }

  async function deleteExamDate(item) {
    setMessage("");
    setError("");
    let preview;
    try {
      preview = await apiRequest("delete_exam_date", {
        exam_date: item.exam_date,
        preview_only: true,
      });
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const impact = preview?.impact || {};
    if (!window.confirm(
      `${item.label || item.exam_date} 고사일을 삭제할까요?\n` +
      `시간표 ${impact.timetable || 0}건, 결시 ${impact.absences || 0}건, 제출 ${impact.completions || 0}건, 자리배치 ${impact.seat_charts || 0}건이 함께 정리됩니다.`,
    )) return;
    await commit(
      "delete_exam_date",
      { exam_date: item.exam_date },
      "고사일과 참조 데이터를 정리했습니다.",
    );
  }

  async function addClass() {
    const result = await commit("save_class", {
      class_item: { grade: Number(grade), class_num: Number(classNum), active: true },
    }, "학급을 저장했습니다.");
    if (result?.saved_class_id) setSelectedClassId(result.saved_class_id);
  }

  async function saveStudents() {
    const students = [];
    const errors = [];
    const seen = new Set();
    studentText.split(/\r?\n/).forEach((raw, index) => {
      const line = raw.trim();
      if (!line) return;
      const [numberText, ...nameParts] = line.split(/[\t, ]+/);
      const number = Number(numberText);
      const name = nameParts.join(" ").trim();
      if (!Number.isInteger(number) || number < 1 || !name) {
        errors.push({ row_number: index + 1, code: "INVALID_STUDENT_ROW" });
        return;
      }
      if (seen.has(number)) {
        errors.push({ row_number: index + 1, code: "DUPLICATE_STUDENT_NUMBER" });
        return;
      }
      seen.add(number);
      students.push({ number, name });
    });
    if (errors.length) {
      setError(importErrorMessage("명단을 저장하지 않았습니다", errors));
      return;
    }
    if (!students.length) {
      setError("학생 명단을 한 명 이상 입력하세요.");
      return;
    }
    await commit(
      "replace_students",
      { class_id: selectedClassId, students },
      `${students.length}명의 학생 명단을 저장했습니다.`,
    );
  }

  async function saveStudent(studentId = "") {
    const draft = studentId ? studentEdits[String(studentId)] : newStudent;
    const number = Number(draft?.number);
    const name = String(draft?.name || "").trim();
    if (!selectedClassId || !Number.isInteger(number) || number < 1 || !name) {
      setError("학생 번호와 이름을 확인하세요.");
      return;
    }
    const result = await commit("save_student", {
      student: {
        id: studentId || undefined,
        class_id: selectedClassId,
        number,
        name,
      },
    }, studentId ? "학생 정보를 수정했습니다." : "학생을 추가했습니다.");
    if (result && !studentId) setNewStudent({ number: "", name: "" });
  }

  async function deleteStudent(student) {
    setMessage("");
    setError("");
    let preview;
    try {
      preview = await apiRequest("delete_student", {
        student_id: student.id,
        preview_only: true,
      });
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const impact = preview?.impact || {};
    if (!window.confirm(
      `${student.number}번 ${student.name} 학생을 비활성화할까요?\n` +
      `선택과목 ${impact.enrollments || 0}건, 자리배치 ${impact.seat_charts || 0}건이 정리됩니다. ` +
      `과거 결시 ${impact.absences || 0}건의 이름과 연결은 보존됩니다.`,
    )) return;
    await commit("delete_student", { student_id: student.id }, "학생을 비활성화했습니다.");
  }

  async function deleteClass() {
    setMessage("");
    setError("");
    let preview;
    try {
      preview = await apiRequest("delete_class", {
        class_id: selectedClassId,
        preview_only: true,
      });
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const impact = preview?.impact || {};
    if (!window.confirm(
      `이 학급을 삭제할까요?\n` +
      `학생 ${impact.students || 0}명, 시간표 ${impact.timetable || 0}건, 선택과목 ${impact.enrollments || 0}건, ` +
      `결시 ${impact.absences || 0}건, 제출 ${impact.completions || 0}건, 자리배치 ${impact.seat_charts || 0}건에 영향을 줍니다.\n` +
      "학생과 학급은 비활성화해 과거 결시 기록 연결을 보존하고 운영 참조 자료만 정리합니다.",
    )) return;
    const result = await commit("delete_class", { class_id: selectedClassId }, "학급과 참조 데이터를 정리했습니다.");
    if (result?.classes?.length) setSelectedClassId(result.classes[0].id);
  }

  async function importStudentFile(file) {
    if (!file) return;
    try {
      const analyzed = analyzeStudentRows(await readWorkbookRows(file));
      if (analyzed.errors.length) {
        throw new Error(importErrorMessage("학생 명단을 가져오지 않았습니다", analyzed.errors));
      }
      const rows = analyzed.rows;
      if (!rows.length) throw new Error("가져올 학생이 없습니다. 열 이름과 값 형식을 확인하세요.");
      setMessage("");
      setError("");
      const result = await apiRequest("import_students", { rows });
      onBootstrap({ ...bootstrap, ...result });
      setMessage(`${rows.length}명의 학생을 가져왔습니다.`);
    } catch (nextError) {
      const backendErrors = nextError?.details?.errors || [];
      setError(backendErrors.length
        ? importErrorMessage("학생 명단을 저장하지 않았습니다", backendErrors)
        : messageFrom(nextError));
    } finally {
      studentFileRef.current.value = "";
    }
  }

  async function addTimetable() {
    if (!timetableDraft.subject_name.trim()) return setError("과목명을 입력하세요.");
    if (!timetableDraft.class_ids.length) return setError("적용 학급을 한 곳 이상 선택하세요.");
    if (bootstrap.timetable.some((item) =>
      item.exam_date === timetableDraft.exam_date &&
      Number(item.grade) === Number(timetableDraft.grade) &&
      Number(item.period) === Number(timetableDraft.period) &&
      item.subject_name === timetableDraft.subject_name.trim()
    )) return setError("같은 교시에 같은 과목명이 이미 등록되어 있습니다.");
    const next = [...bootstrap.timetable, {
      ...timetableDraft,
      id: makeId("timetable"),
      grade: Number(timetableDraft.grade),
      period: Number(timetableDraft.period),
      class_ids: timetableDraft.class_ids.join("|"),
    }];
    if (await commit("save_timetable", { timetable: next }, "시간표 항목을 저장했습니다.")) {
      setTimetableDraft((value) => ({ ...value, subject_name: "", room_name: "" }));
    }
  }

  async function removeTimetable(id) {
    const item = bootstrap.timetable.find((row) => String(row.id) === String(id));
    if (!item) return;
    setMessage("");
    setError("");
    let preview;
    try {
      preview = await apiRequest("delete_timetable", { id, preview_only: true });
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const impact = preview?.impact || {};
    const enrollmentText = impact.remove_subject_enrollments
      ? `이 과목의 선택과목 등록 ${impact.enrollments || 0}건`
      : "다른 시간표에서 같은 과목을 사용하므로 선택과목 등록은 유지";
    if (!window.confirm(
      `${item.exam_date} ${item.period}교시 '${item.subject_name}' 시간표를 삭제할까요?\n` +
      `${enrollmentText}, 자리배치 ${impact.seat_charts || 0}건이 함께 정리됩니다.`,
    )) return;
    await commit("delete_timetable", { id }, "시간표와 참조 데이터를 정리했습니다.");
  }

  async function importEnrollmentFile(file) {
    if (!file) return;
    try {
      const analyzed = analyzeEnrollmentRows(await readWorkbookRows(file));
      if (analyzed.errors.length) {
        throw new Error(importErrorMessage("선택과목 명단을 가져오지 않았습니다", analyzed.errors));
      }
      const raw = analyzed.rows;
      const targetGrade = Number(timetableDraft.grade);
      if (!raw.length) throw new Error("가져올 선택과목 행이 없습니다. 열 이름과 값 형식을 확인하세요.");
      setMessage("");
      setError("");
      const result = await apiRequest("save_enrollments", {
        grade: targetGrade,
        enrollments: raw,
      });
      onBootstrap({ ...bootstrap, ...result });
      const importResult = result?.enrollment_result;
      if (importResult) {
        setMessage(`선택과목 ${importResult.matched_count}건을 저장했습니다.`);
      }
    } catch (nextError) {
      const backendErrors = nextError?.details?.errors || [];
      setError(backendErrors.length
        ? importErrorMessage("선택과목 명단을 저장하지 않았습니다", backendErrors)
        : messageFrom(nextError));
    } finally {
      enrollmentFileRef.current.value = "";
    }
  }

  const printStudents = sortStudents(bootstrap.students.filter((student) => {
    if (printScope === "student") return String(student.id) === String(personalStudentId);
    if (printScope === "class") return String(student.class_id) === String(printClassId);
    const classInfo = bootstrap.classes.find((item) => String(item.id) === String(student.class_id));
    return Number(classInfo?.grade) === Number(printGrade);
  }));
  const personalSheets = printStudents.map((student) => {
    const classInfo = bootstrap.classes.find((item) => String(item.id) === String(student.class_id));
    return {
      student,
      classInfo,
      rows: buildPersonalTimetable({
        student,
        classInfo,
        timetable: bootstrap.timetable,
        enrollments: bootstrap.enrollments,
      }),
    };
  });

  const enrollmentRows = bootstrap.enrollments.filter(
    (item) => Number(item.grade) === Number(timetableDraft.grade),
  );
  const enrollmentSummary = [...enrollmentRows.reduce((map, item) => {
    const room = item.room_name || bootstrap.timetable.find(
      (row) =>
        Number(row.grade) === Number(item.grade) &&
        row.subject_name === item.subject_name,
    )?.room_name || "미지정";
    const key = `${item.subject_name}|${room}`;
    if (!map.has(key)) map.set(key, { subject_name: item.subject_name, room_name: room, count: 0 });
    map.get(key).count += 1;
    return map;
  }, new Map()).values()];

  async function cleanup() {
    if (!cleanupDate) return setError("정리 기준일을 선택하세요.");
    if (!window.confirm(`${cleanupDate} 이전의 고사·결시·제출·자리배치 기록을 정리할까요? 이 작업은 되돌릴 수 없습니다.`)) return;
    await commit("cleanup", { cutoff_date: cleanupDate }, "오래된 기록을 정리했습니다.");
  }

  const menu = [
    ["dates", IconCalendar, "고사일"],
    ["classes", IconUsers, "학급·학생"],
    ["timetable", IconClock, "시간표"],
    ["electives", IconBook, "선택과목"],
    ["print", IconPrinter, "개인 시간표"],
    ["cleanup", IconTrash, "기록 정리"],
  ];

  return (
    <section className="page-section">
      <SectionHeader eyebrow="ADMIN SETTINGS" title="관리자 설정" description="고사 운영에 필요한 기준 정보를 순서대로 등록하세요." />
      <Notice>{message}</Notice><Notice tone="error">{error}</Notice>
      <div className="settings-layout">
        <nav className="settings-nav no-print" aria-label="설정 메뉴">
          {menu.map(([id, Icon, label]) => <button key={id} className={active === id ? "active" : ""} onClick={() => setActive(id)}><Icon size={18} /><span>{label}</span><IconChevronRight size={16} /></button>)}
        </nav>
        <div className="settings-content">
          {active === "dates" ? (
            <article className="panel-card">
              <div className="card-heading"><div><IconCalendar size={20} /><h3>고사일 관리</h3></div></div>
              <div className="inline-form">
                <label>날짜<span className="date-input-row"><input type="date" value={newDate} onChange={(event) => setNewDate(event.target.value)} aria-describedby="new-exam-date-weekday" /><output id="new-exam-date-weekday" className="weekday-output" aria-live="polite">{formatKoreanWeekday(newDate)}</output></span></label>
                <label>표시 이름<input value={newDateLabel} onChange={(event) => setNewDateLabel(event.target.value)} placeholder="예: 1일차" /></label>
                <button className="button button-secondary" onClick={addDate}><IconPlus size={17} /> 추가</button>
              </div>
              <div className="item-list">{bootstrap.exam_dates.map((item) => <div key={item.id}><div><strong>{item.label}</strong><span>{formatKoreanDate(item.exam_date)}</span></div><button className="icon-button danger" onClick={() => deleteExamDate(item)} aria-label={`${item.label} 고사일 삭제`}><IconTrash size={17} /></button></div>)}</div>
            </article>
          ) : null}

          {active === "classes" ? (
            <article className="panel-card">
              <div className="card-heading"><div><IconUsers size={20} /><h3>학급과 학생 명단</h3></div><button className="button button-light" onClick={() => downloadWorkbook("학생 명단 등록 양식", { 학생명단: [["학년", "반", "번호", "이름"], [1, 1, 1, "홍길동"]] })}><IconDownload size={17} /> 양식</button></div>
              <div className="inline-form">
                <label>학년<input type="number" min="1" max="9" value={grade} onChange={(event) => setGrade(event.target.value)} /></label>
                <label>반<input type="number" min="1" value={classNum} onChange={(event) => setClassNum(event.target.value)} /></label>
                <button className="button button-secondary" onClick={addClass}><IconPlus size={17} /> 학급 추가</button>
                <button className="button button-light" onClick={() => studentFileRef.current?.click()}><IconUpload size={17} /> 엑셀 가져오기</button>
                <input ref={studentFileRef} className="sr-only" type="file" accept=".xlsx,.xls" onChange={(event) => importStudentFile(event.target.files?.[0])} />
              </div>
              <div className="class-chips">{bootstrap.classes.map((item) => <button key={item.id} className={selectedClassId === item.id ? "active" : ""} onClick={() => setSelectedClassId(item.id)}>{classLabel(item)}</button>)}</div>
              {selectedClassId ? <>
                <label className="stacked-label">학생 명단 <span>한 줄에 ‘번호 이름’ 형식</span>
                  <textarea rows="10" value={studentText} onChange={(event) => setStudentText(event.target.value)} placeholder={"1 홍길동\n2 김한양"} />
                </label>
                <div className="button-row">
                  <button className="button button-primary" onClick={saveStudents}><IconDeviceFloppy size={17} /> 명단 저장</button>
                  <button className="button button-danger" onClick={deleteClass}><IconTrash size={17} /> 영향 확인 후 학급 삭제</button>
                </div>
                <div className="student-admin">
                  <div className="card-heading"><div><IconUserCheck size={19} /><h3>학생 개별 수정</h3></div><span>명단 교체 시 같은 번호의 ID 유지</span></div>
                  <div className="student-admin-new">
                    <label>번호<input type="number" min="1" value={newStudent.number} onChange={(event) => setNewStudent({ ...newStudent, number: event.target.value })} /></label>
                    <label>이름<input value={newStudent.name} onChange={(event) => setNewStudent({ ...newStudent, name: event.target.value })} /></label>
                    <button className="button button-secondary" onClick={() => saveStudent()}><IconPlus size={17} /> 학생 추가</button>
                  </div>
                  <div className="student-edit-list">
                    {bootstrap.students.filter((item) => String(item.class_id) === String(selectedClassId)).map((student) => {
                      const draft = studentEdits[String(student.id)] || student;
                      return (
                        <div key={student.id}>
                          <input aria-label={`${student.name} 번호`} type="number" min="1" value={draft.number} onChange={(event) => setStudentEdits((value) => ({ ...value, [student.id]: { ...draft, number: event.target.value } }))} />
                          <input aria-label={`${student.number}번 이름`} value={draft.name} onChange={(event) => setStudentEdits((value) => ({ ...value, [student.id]: { ...draft, name: event.target.value } }))} />
                          <button className="button button-light" onClick={() => saveStudent(student.id)}><IconDeviceFloppy size={16} /> 수정</button>
                          <button className="icon-button danger" onClick={() => deleteStudent(student)} aria-label={`${student.name} 삭제`}><IconTrash size={16} /></button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </> : null}
            </article>
          ) : null}

          {active === "timetable" ? (
            <article className="panel-card">
              <div className="card-heading"><div><IconClock size={20} /><h3>공통·선택과목 시간표</h3></div></div>
              <div className="form-grid">
                <label>고사일<select value={timetableDraft.exam_date} onChange={(event) => setTimetableDraft({ ...timetableDraft, exam_date: event.target.value })}>{bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {formatKoreanDate(item.exam_date)}</option>)}</select></label>
                <label>학년<input type="number" min="1" value={timetableDraft.grade} onChange={(event) => setTimetableDraft({ ...timetableDraft, grade: event.target.value })} /></label>
                <label>교시<select value={timetableDraft.period} onChange={(event) => setTimetableDraft({ ...timetableDraft, period: event.target.value })}>{PERIODS.map((item) => <option key={item}>{item}교시</option>)}</select></label>
                <label>시작<input type="time" value={timetableDraft.start_time} onChange={(event) => setTimetableDraft({ ...timetableDraft, start_time: event.target.value })} /></label>
                <label>종료<input type="time" value={timetableDraft.end_time} onChange={(event) => setTimetableDraft({ ...timetableDraft, end_time: event.target.value })} /></label>
                <label>구분<select value={timetableDraft.subject_type} onChange={(event) => setTimetableDraft({ ...timetableDraft, subject_type: event.target.value })}><option value="common">공통과목</option><option value="elective">선택과목</option></select></label>
                <label className="span-two">과목명<input value={timetableDraft.subject_name} onChange={(event) => setTimetableDraft({ ...timetableDraft, subject_name: event.target.value })} /></label>
                <label>기본 호실<input value={timetableDraft.room_name} onChange={(event) => setTimetableDraft({ ...timetableDraft, room_name: event.target.value })} /></label>
              </div>
              <fieldset className="class-scope">
                <legend>적용 학급</legend>
                {bootstrap.classes.filter((item) => Number(item.grade) === Number(timetableDraft.grade)).map((item) => (
                  <label key={item.id}>
                    <input
                      type="checkbox"
                      checked={timetableDraft.class_ids.includes(item.id)}
                      onChange={() => setTimetableDraft((value) => ({
                        ...value,
                        class_ids: value.class_ids.includes(item.id)
                          ? value.class_ids.filter((id) => id !== item.id)
                          : [...value.class_ids, item.id],
                      }))}
                    />
                    {classLabel(item)}
                  </label>
                ))}
              </fieldset>
              <button className="button button-secondary" onClick={addTimetable}><IconPlus size={17} /> 시간표 추가</button>
              <div className="table-wrap"><table><thead><tr><th>고사일</th><th>학년</th><th>교시</th><th>시간</th><th>구분</th><th>과목</th><th>적용 학급</th><th></th></tr></thead><tbody>{bootstrap.timetable.map((item) => <tr key={item.id}><td>{formatKoreanDate(item.exam_date)}</td><td>{item.grade}</td><td>{item.period}</td><td>{item.start_time}~{item.end_time}</td><td>{item.subject_type === "common" ? "공통" : "선택"}</td><td>{item.subject_name}</td><td>{String(item.class_ids || "").split(/[|,]/).filter(Boolean).map((id) => classLabel(bootstrap.classes.find((row) => String(row.id) === String(id)))).join(", ") || "학년 전체"}</td><td><button className="icon-button danger" onClick={() => removeTimetable(item.id)} aria-label={`${item.subject_name} 삭제`}><IconTrash size={16} /></button></td></tr>)}</tbody></table></div>
            </article>
          ) : null}

          {active === "electives" ? (
            <article className="panel-card">
              <div className="card-heading"><div><IconBook size={20} /><h3>선택과목 응시자·호실</h3></div><button className="button button-light" onClick={() => downloadWorkbook("선택과목 등록 양식", { 선택과목: [["과목명", "반", "번호", "이름", "호실"], ["경제", 1, 1, "홍길동", "별실 1"]] })}><IconDownload size={17} /> 양식</button></div>
              <p className="body-copy">시간표의 과목명과 엑셀의 과목명이 정확히 같아야 개인 시간표와 별실 배치에 연결됩니다.</p>
              <div className="inline-form">
                <label>대상 학년<select value={timetableDraft.grade} onChange={(event) => setTimetableDraft({ ...timetableDraft, grade: event.target.value })}>{[...new Set(bootstrap.classes.map((item) => item.grade))].map((item) => <option key={item}>{item}학년</option>)}</select></label>
                <button className="button button-primary" onClick={() => enrollmentFileRef.current?.click()}><IconUpload size={17} /> 엑셀 가져오기</button>
                <input ref={enrollmentFileRef} className="sr-only" type="file" accept=".xlsx,.xls" onChange={(event) => importEnrollmentFile(event.target.files?.[0])} />
              </div>
              <div className="metric-grid compact"><article><span>등록 건수</span><strong>{bootstrap.enrollments.filter((item) => Number(item.grade) === Number(timetableDraft.grade)).length}<small>건</small></strong></article><article><span>선택과목 수</span><strong>{new Set(bootstrap.enrollments.filter((item) => Number(item.grade) === Number(timetableDraft.grade)).map((item) => item.subject_name)).size}<small>개</small></strong></article></div>
              {enrollmentSummary.length ? (
                <>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>과목</th><th>호실</th><th>응시자</th></tr></thead>
                      <tbody>{enrollmentSummary.map((item) => <tr key={`${item.subject_name}-${item.room_name}`}><td>{item.subject_name}</td><td>{item.room_name}</td><td>{item.count}명</td></tr>)}</tbody>
                    </table>
                  </div>
                  <details className="enrollment-details">
                    <summary>등록 결과 상세 보기</summary>
                    <div className="table-wrap">
                      <table>
                        <thead><tr><th>과목</th><th>학급</th><th>번호</th><th>이름</th><th>호실</th></tr></thead>
                        <tbody>{enrollmentRows.map((item) => {
                          const student = bootstrap.students.find((row) => String(row.id) === String(item.student_id));
                          const classInfo = bootstrap.classes.find((row) => String(row.id) === String(item.class_id));
                          const defaultRoom = bootstrap.timetable.find((row) => Number(row.grade) === Number(item.grade) && row.subject_name === item.subject_name)?.room_name;
                          return <tr key={item.id || `${item.student_id}-${item.subject_name}`}><td>{item.subject_name}</td><td>{classLabel(classInfo)}</td><td>{student?.number}</td><td>{student?.name}</td><td>{item.room_name || defaultRoom || "미지정"}</td></tr>;
                        })}</tbody>
                      </table>
                    </div>
                  </details>
                </>
              ) : <EmptyState icon={IconBook} title="등록된 선택과목 응시자가 없습니다" description="엑셀 양식에 과목·반·번호·이름·호실을 작성해 가져오세요." />}
            </article>
          ) : null}

          {active === "print" ? (
            <article className="panel-card print-section">
              <div className="card-heading no-print"><div><IconPrinter size={20} /><h3>학생별 개인 시간표</h3></div><button className="button button-light" onClick={() => window.print()} disabled={!personalSheets.length}><IconPrinter size={17} /> A4 일괄 인쇄</button></div>
              <div className="print-scope no-print">
                <label>출력 범위<select value={printScope} onChange={(event) => setPrintScope(event.target.value)}><option value="student">학생 1명</option><option value="class">특정 학급 전체</option><option value="grade">학년 전체</option></select></label>
                {printScope === "student" ? <label>학생<select value={personalStudentId} onChange={(event) => setPersonalStudentId(event.target.value)}>{bootstrap.classes.map((classInfo) => <optgroup key={classInfo.id} label={classLabel(classInfo)}>{bootstrap.students.filter((item) => item.class_id === classInfo.id).map((student) => <option key={student.id} value={student.id}>{student.number}번 {student.name}</option>)}</optgroup>)}</select></label> : null}
                {printScope === "class" ? <label>학급<select value={printClassId} onChange={(event) => setPrintClassId(event.target.value)}>{bootstrap.classes.map((item) => <option key={item.id} value={item.id}>{classLabel(item)}</option>)}</select></label> : null}
                {printScope === "grade" ? <label>학년<select value={printGrade} onChange={(event) => setPrintGrade(Number(event.target.value))}>{[...new Set(bootstrap.classes.map((item) => Number(item.grade)))].map((item) => <option key={item}>{item}학년</option>)}</select></label> : null}
                <p>{personalSheets.length}명의 시간표를 각각 A4 한 장으로 출력합니다.</p>
              </div>
              {personalSheets.map(({ student, classInfo, rows: scheduleRows }) => (
                <div className="personal-sheet" key={student.id}>
                  <p>PERSONAL EXAM SCHEDULE</p><h3>개인 고사 시간표</h3>
                  <div className="student-meta"><span>{classLabel(classInfo)}</span><strong>{student.number}번 {student.name}</strong></div>
                  <table><thead><tr><th>고사일</th><th>교시</th><th>시간</th><th>과목</th><th>고사실</th></tr></thead><tbody>{scheduleRows.map((row) => <tr key={`${row.exam_date}-${row.period}`}><td>{formatKoreanDate(row.exam_date)}</td><td>{row.period}</td><td>{row.time}</td><td>{row.subject_name}</td><td>{row.room_name}</td></tr>)}</tbody></table>
                </div>
              ))}
            </article>
          ) : null}

          {active === "cleanup" ? (
            <article className="panel-card danger-card">
              <div className="card-heading"><div><IconTrash size={20} /><h3>오래된 기록 정리</h3></div></div>
              <p>기준일 이전의 고사일, 시간표, 결시, 입력완료, 자리배치 기록을 정리합니다. 학급·학생·선택과목 명단은 유지됩니다.</p>
              <div className="inline-form"><label>정리 기준일<input type="date" value={cleanupDate} onChange={(event) => setCleanupDate(event.target.value)} /></label><button className="button button-danger" onClick={cleanup}><IconTrash size={17} /> 이전 기록 정리</button></div>
            </article>
          ) : null}
        </div>
      </div>
    </section>
  );
}

export function App() {
  const [loggedIn, setLoggedIn] = useState(hasSchoolSession());
  const [admin, setAdmin] = useState(hasAdminSession());
  const [bootstrap, setBootstrap] = useState(EMPTY_BOOTSTRAP);
  const [activeTab, setActiveTab] = useState("absence");
  const [adminDialog, setAdminDialog] = useState(false);
  const [pendingTab, setPendingTab] = useState("");
  const [startupError, setStartupError] = useState("");
  const [presence, setPresence] = useState({ status: "connecting", onlineCount: 0 });
  const lastActivityAtRef = useRef(Date.now());

  useEffect(() => {
    if (!loggedIn) return;
    apiRequest(admin ? "get_admin_bootstrap" : "get_bootstrap")
      .then((value) => setBootstrap((current) => ({ ...current, ...value })))
      .catch((error) => {
        setStartupError(messageFrom(error));
        if (/세션|로그인|만료/.test(messageFrom(error))) {
          clearSessions(); setLoggedIn(false); setAdmin(false);
        }
      });
  }, [loggedIn, admin]);

  useEffect(() => {
    const handleExpired = (event) => {
      if (event.detail?.code === "ADMIN_SESSION_EXPIRED") {
        setAdmin(false);
        setActiveTab("absence");
        setBootstrap((current) => ({
          ...EMPTY_BOOTSTRAP,
          settings: current.settings,
          classes: current.classes,
          exam_dates: current.exam_dates,
        }));
        setStartupError("관리자 세션이 만료되었습니다. 관리자 화면을 열 때 암호를 다시 입력하세요.");
        return;
      }
      setLoggedIn(false);
      setAdmin(false);
      setActiveTab("absence");
      setBootstrap(EMPTY_BOOTSTRAP);
      setStartupError("세션이 만료되었습니다. 학교코드로 다시 접속하세요.");
    };
    window.addEventListener("exam-session-expired", handleExpired);
    return () => window.removeEventListener("exam-session-expired", handleExpired);
  }, []);

  useEffect(() => {
    if (!loggedIn) {
      setPresence({ status: "connecting", onlineCount: 0 });
      return undefined;
    }

    let cancelled = false;
    let inFlight = false;
    let timer;
    const ping = async () => {
      if (cancelled || inFlight || document.visibilityState === "hidden") return;
      if (navigator.onLine === false) {
        setPresence((current) => ({ ...current, status: "disconnected" }));
        return;
      }
      inFlight = true;
      setPresence((current) => ({
        ...current,
        status: current.status === "connected" ? "connected" : "connecting",
      }));
      try {
        const result = await apiRequest("presence_ping", {
          user_active: Date.now() - lastActivityAtRef.current < 35_000,
        });
        if (!cancelled) {
          setPresence({
            status: result.connected === false ? "disconnected" : "connected",
            onlineCount: Math.max(0, Number(result.online_count) || 0),
          });
        }
      } catch (error) {
        if (!cancelled && hasSchoolSession()) {
          setPresence((current) => ({ ...current, status: "disconnected" }));
        }
      } finally {
        inFlight = false;
      }
    };
    const schedule = () => {
      clearTimeout(timer);
      if (cancelled || document.visibilityState === "hidden" || navigator.onLine === false) return;
      const jitter = Math.floor(Math.random() * 8_001) - 4_000;
      timer = setTimeout(tick, 30_000 + jitter);
    };
    const tick = async () => {
      await ping();
      schedule();
    };
    const handleVisibility = () => {
      if (document.visibilityState === "visible") tick();
      else clearTimeout(timer);
    };
    const handleOnline = () => tick();
    const handleOffline = () => {
      clearTimeout(timer);
      setPresence((current) => ({ ...current, status: "disconnected" }));
    };

    tick();
    document.addEventListener("visibilitychange", handleVisibility);
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      cancelled = true;
      clearTimeout(timer);
      document.removeEventListener("visibilitychange", handleVisibility);
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, [loggedIn]);

  useEffect(() => {
    if (!loggedIn) return undefined;
    let timer;
    const reset = () => {
      lastActivityAtRef.current = Date.now();
      clearTimeout(timer);
      timer = setTimeout(() => {
        apiRequest("logout").catch(() => {});
        clearSessions();
        setLoggedIn(false);
        setAdmin(false);
        setActiveTab("absence");
      }, 5 * 60 * 1000);
    };
    const events = ["pointerdown", "keydown", "scroll", "touchstart"];
    events.forEach((event) => window.addEventListener(event, reset, { passive: true }));
    reset();
    return () => {
      clearTimeout(timer);
      events.forEach((event) => window.removeEventListener(event, reset));
    };
  }, [loggedIn]);

  function chooseTab(id, locked) {
    if (locked && !admin) {
      setPendingTab(id);
      setAdminDialog(true);
      return;
    }
    setActiveTab(id);
  }

  function logout() {
    apiRequest("logout").catch(() => {});
    clearSessions();
    setLoggedIn(false);
    setAdmin(false);
    setBootstrap(EMPTY_BOOTSTRAP);
  }

  if (!loggedIn) {
    return <LoginScreen onLogin={(value) => { setBootstrap({ ...EMPTY_BOOTSTRAP, ...value }); setLoggedIn(true); }} />;
  }

  const tabs = [
    ["absence", IconClipboardCheck, "결시 입력", false],
    ["hq", IconBuildingBank, "고사본부", true],
    ["seating", IconArmchair, "자리배치", true],
    ["settings", IconSettings, "설정", true],
  ];

  return (
    <div className="app-shell">
      <header className="app-header no-print">
        <div className="header-inner">
          <div className="header-brand"><span><IconBook size={22} /></span><div><strong>{bootstrap.settings.app_name || "정기고사 관리 시스템"}</strong><small>{bootstrap.settings.school_name || "학교 고사 운영"}</small></div></div>
          <div className="header-actions">
            <span
              className={`connection-badge ${presence.status}`}
              role="status"
              aria-live="polite"
              title="현재 화면을 제외한 최근 90초 이내 다른 활성 화면 수"
            >
              <span className="connection-dot" aria-hidden="true" />
              {formatPresenceLabel(presence.status, presence.onlineCount)}
            </span>
            <span className={`mode-badge ${appMode.demo ? "demo" : ""}`}>{appMode.name}</span>
            {admin ? <span className="admin-badge"><IconShieldLock size={14} /> 관리자</span> : null}
            <button className="icon-button header-logout" onClick={logout} aria-label="로그아웃"><IconLogout size={19} /></button>
          </div>
        </div>
        <nav className="main-tabs">
          {tabs.map(([id, Icon, label, locked]) => (
            <button key={id} className={activeTab === id ? "active" : ""} onClick={() => chooseTab(id, locked)}>
              <Icon size={19} /><span>{label}</span>{locked && !admin ? <IconLock className="tab-lock" size={12} /> : null}
            </button>
          ))}
        </nav>
      </header>

      <main className="app-main">
        <Notice tone="error">{startupError}</Notice>
        {activeTab === "absence" ? <AbsencePanel bootstrap={bootstrap} /> : null}
        {activeTab === "hq" && admin ? <HqPanel bootstrap={bootstrap} /> : null}
        {activeTab === "seating" && admin ? <SeatPanel bootstrap={bootstrap} onBootstrap={setBootstrap} /> : null}
        {activeTab === "settings" && admin ? <SetupPanel bootstrap={bootstrap} onBootstrap={setBootstrap} /> : null}
      </main>

      <footer className="app-footer no-print">
        <span>학교 소유 Google Sheets · Apps Script</span>
        <span>5분 미사용 시 자동 로그아웃</span>
      </footer>

      <AdminDialog
        open={adminDialog}
        onClose={() => { setAdminDialog(false); setPendingTab(""); }}
        onUnlock={(value) => {
          setBootstrap((current) => ({ ...current, ...value }));
          setAdmin(true);
          setActiveTab(pendingTab || "hq");
        }}
      />
    </div>
  );
}
