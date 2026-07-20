import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  IconAlertTriangle,
  IconArmchair,
  IconBook,
  IconBuildingBank,
  IconCalendar,
  IconCheck,
  IconChevronRight,
  IconClipboardCheck,
  IconClock,
  IconDeviceFloppy,
  IconDownload,
  IconFileSpreadsheet,
  IconLayoutGrid,
  IconListCheck,
  IconLock,
  IconLogout,
  IconPlus,
  IconPrinter,
  IconRefresh,
  IconSettings,
  IconShieldLock,
  IconTrash,
  IconUpload,
  IconUserCheck,
  IconUsers,
  IconX,
} from "@tabler/icons-react";
import {
  ABSENCE_REASONS,
  buildPersonalTimetable,
  classLabel,
  formatKoreanDate,
  generateSeatAssignment,
  makeId,
  normalizeEnrollmentRows,
  normalizeManualSeatRows,
  normalizeStudentRows,
  sortStudents,
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

  useEffect(() => {
    if (open) {
      setPassword("");
      setError("");
      setTimeout(() => passwordRef.current?.focus(), 50);
    }
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
      <section className="dialog" role="dialog" aria-modal="true" aria-labelledby="admin-title">
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
  const [context, setContext] = useState({ students: [], absences: [], completed: false });
  const [loading, setLoading] = useState(false);
  const [savingId, setSavingId] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const manualFileRef = useRef(null);

  useEffect(() => {
    if (!examDate && bootstrap.exam_dates.length) setExamDate(bootstrap.exam_dates[0].exam_date);
    if (!classId && bootstrap.classes.length) setClassId(bootstrap.classes[0].id);
  }, [bootstrap.exam_dates, bootstrap.classes, examDate, classId]);

  const load = useCallback(async () => {
    if (!examDate || !classId || !period) return;
    setLoading(true);
    setError("");
    try {
      setContext(await apiRequest("get_absence_context", {
        exam_date: examDate,
        class_id: classId,
        period,
      }));
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }, [examDate, classId, period]);

  useEffect(() => { load(); }, [load]);

  const absenceMap = useMemo(
    () => new Map(context.absences.map((item) => [String(item.student_id), item])),
    [context.absences],
  );

  async function updateStudent(student, patch) {
    const current = absenceMap.get(String(student.id));
    const next = {
      absent: patch.absent ?? Boolean(current),
      reason: patch.reason ?? current?.reason ?? "질병",
      reason_detail: patch.reason_detail ?? current?.reason_detail ?? "",
    };
    setSavingId(student.id);
    setMessage("");
    setError("");
    setContext((value) => ({
      ...value,
      completed: false,
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
    try {
      await apiRequest("set_absence", {
        exam_date: examDate,
        class_id: classId,
        period,
        student_id: student.id,
        ...next,
      });
      setMessage(`${student.number}번 ${student.name} 학생 정보를 저장했습니다.`);
    } catch (nextError) {
      setError(messageFrom(nextError));
      await load();
    } finally {
      setSavingId("");
    }
  }

  async function complete() {
    setLoading(true);
    setMessage("");
    try {
      await apiRequest("complete_input", { exam_date: examDate, class_id: classId, period });
      setContext((value) => ({ ...value, completed: true }));
      setMessage("고사본부에 입력 완료로 표시했습니다.");
    } catch (nextError) {
      setError(messageFrom(nextError));
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
        description="날짜와 학급, 교시를 선택한 뒤 결시 학생과 사유를 저장하세요."
        actions={<button className="button button-light" onClick={() => window.print()}><IconPrinter size={17} /> 응시현황표 인쇄</button>}
      />
      <div className="filter-bar no-print">
        <label>고사일
          <select value={examDate} onChange={(event) => setExamDate(event.target.value)}>
            {bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {formatKoreanDate(item.exam_date)}</option>)}
          </select>
        </label>
        <label>학급
          <select value={classId} onChange={(event) => setClassId(event.target.value)}>
            {bootstrap.classes.map((item) => <option key={item.id} value={item.id}>{classLabel(item)}</option>)}
          </select>
        </label>
        <label>교시
          <select value={period} onChange={(event) => setPeriod(Number(event.target.value))}>
            {PERIODS.map((item) => <option key={item} value={item}>{item}교시</option>)}
          </select>
        </label>
      </div>

      <div className="print-heading">
        <h3>{formatKoreanDate(examDate)} {period}교시 응시현황표</h3>
        <p>{selectedClass ? classLabel(selectedClass) : ""} · 재적 {context.students.length}명 · 결시 {context.absences.length}명</p>
      </div>
      <Notice>{message}</Notice>
      <Notice tone="error">{error}</Notice>

      {loading && !context.students.length ? (
        <div className="loading-block"><span className="spinner spinner-dark" /> 명단을 불러오는 중입니다.</div>
      ) : context.students.length ? (
        <div className="student-list">
          {context.students.map((student) => {
            const absence = absenceMap.get(String(student.id));
            return (
              <article className={`student-row ${absence ? "is-absent" : ""}`} key={student.id}>
                <button
                  className="absence-toggle no-print"
                  onClick={() => updateStudent(student, { absent: !absence })}
                  aria-pressed={Boolean(absence)}
                  disabled={savingId === student.id}
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
                    >
                      {ABSENCE_REASONS.map((reason) => <option key={reason}>{reason}</option>)}
                    </select>
                    {absence.reason === "기타" ? (
                      <input
                        value={absence.reason_detail || ""}
                        placeholder="기타 사유"
                        onBlur={(event) => updateStudent(student, { reason_detail: event.target.value })}
                        onChange={(event) => setContext((value) => ({
                          ...value,
                          absences: value.absences.map((item) =>
                            String(item.student_id) === String(student.id)
                              ? { ...item, reason_detail: event.target.value }
                              : item,
                          ),
                        }))}
                      />
                    ) : null}
                  </div>
                ) : null}
                <span className="print-only">{absence ? `${absence.reason}${absence.reason_detail ? `(${absence.reason_detail})` : ""}` : "응시"}</span>
              </article>
            );
          })}
        </div>
      ) : (
        <EmptyState icon={IconUsers} title="등록된 학생이 없습니다" description="관리자 설정에서 학급과 학생 명단을 먼저 등록하세요." />
      )}

      <div className="completion-bar no-print">
        <div>
          <span className={context.completed ? "completion-dot done" : "completion-dot"} />
          <div>
            <strong>{context.completed ? "입력 완료" : "입력 확인 전"}</strong>
            <p>결시 입력을 모두 확인한 뒤 고사본부에 완료 상태를 전송합니다.</p>
          </div>
        </div>
        <LoadingButton className="button button-gold" loading={loading} onClick={complete} disabled={!context.students.length}>
          <IconClipboardCheck size={18} /> 입력 완료
        </LoadingButton>
      </div>
    </section>
  );
}

function HqPanel({ bootstrap }) {
  const [examDate, setExamDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [period, setPeriod] = useState(1);
  const [grade, setGrade] = useState("");
  const [classId, setClassId] = useState("");
  const [status, setStatus] = useState({ absences: [], completions: [] });
  const [report, setReport] = useState([]);
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
    const timer = setInterval(() => refresh(true), 30_000);
    return () => clearInterval(timer);
  }, [refresh]);

  const filteredClasses = bootstrap.classes.filter(
    (item) => (!grade || Number(item.grade) === Number(grade)) && (!classId || String(item.id) === String(classId)),
  );
  const completionKeys = new Set(
    status.completions
      .filter((item) => item.exam_date === examDate && Number(item.period) === Number(period))
      .map((item) => String(item.class_id)),
  );

  async function loadReport() {
    setLoading(true);
    try {
      const result = await apiRequest("get_period_report", {
        start_date: startDate,
        end_date: endDate,
        grade: grade || "",
        class_id: classId || "",
      });
      setReport(result.absences || []);
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      setLoading(false);
    }
  }

  function exportRows(rows, title) {
    downloadWorkbook(title, {
      결시현황: [
        ["고사일", "교시", "학년", "반", "번호", "이름", "사유", "세부사유", "수정시각"],
        ...rows.map((item) => [
          item.exam_date, item.period, item.grade, item.class_num, item.student_number,
          item.student_name, item.reason, item.reason_detail || "", item.updated_at || "",
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
            <button className="button button-light" onClick={() => exportRows(status.absences, "교시별 결시 현황")}><IconDownload size={17} /> 엑셀</button>
          </>
        }
      />
      <div className="filter-bar no-print">
        <label>고사일
          <select value={examDate} onChange={(event) => setExamDate(event.target.value)}>
            {bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {item.exam_date}</option>)}
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
        <article><span>결시 학생</span><strong>{status.absences.length}<small>명</small></strong></article>
        <article><span>입력 완료 학급</span><strong>{filteredClasses.filter((item) => completionKeys.has(String(item.id))).length}<small>개</small></strong></article>
        <article><span>입력 대기 학급</span><strong>{filteredClasses.filter((item) => !completionKeys.has(String(item.id))).length}<small>개</small></strong></article>
      </div>

      <div className="split-grid">
        <article className="panel-card">
          <div className="card-heading"><div><IconListCheck size={20} /><h3>결시 학생</h3></div><span>{loading ? "갱신 중" : "최신 상태"}</span></div>
          {status.absences.length ? (
            <div className="table-wrap">
              <table>
                <thead><tr><th>학급</th><th>번호</th><th>이름</th><th>사유</th></tr></thead>
                <tbody>{status.absences.map((item) => (
                  <tr key={item.id || `${item.student_id}-${item.period}`}>
                    <td>{item.class_label}</td><td>{item.student_number}</td><td>{item.student_name}</td><td><span className="status-pill status-absent">{item.reason}</span></td>
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
          <button className="button button-light" disabled={!report.length} onClick={() => exportRows(report, "기간 결시 통계")}><IconFileSpreadsheet size={17} /> 통계 엑셀</button>
        </div>
        {report.length ? <p className="report-summary">선택 기간 결시 기록 <strong>{report.length}건</strong>을 확인했습니다.</p> : null}
      </article>
    </section>
  );
}

function SeatPanel({ bootstrap, onBootstrap }) {
  const [mode, setMode] = useState("separate");
  const [examDate, setExamDate] = useState(selectFirst(bootstrap.exam_dates, "exam_date"));
  const [classId, setClassId] = useState(selectFirst(bootstrap.classes));
  const [period, setPeriod] = useState(1);
  const [rows, setRows] = useState(5);
  const [cols, setCols] = useState(6);
  const [startSide, setStartSide] = useState("window");
  const [roomName, setRoomName] = useState("별실 1");
  const [disabledSeats, setDisabledSeats] = useState([]);
  const [selectedIds, setSelectedIds] = useState([]);
  const [manualPool, setManualPool] = useState([]);
  const [absentIds, setAbsentIds] = useState([]);
  const [result, setResult] = useState({ assignments: [], ok: true, error: "" });
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const students = useMemo(
    () => sortStudents(bootstrap.students.filter((item) => String(item.class_id) === String(classId))),
    [bootstrap.students, classId],
  );
  const selectedClass = bootstrap.classes.find((item) => String(item.id) === String(classId));

  useEffect(() => {
    if (!examDate && bootstrap.exam_dates.length) setExamDate(bootstrap.exam_dates[0].exam_date);
    if (!classId && bootstrap.classes.length) setClassId(bootstrap.classes[0].id);
  }, [bootstrap.exam_dates, bootstrap.classes, examDate, classId]);

  useEffect(() => {
    setManualPool([]);
    setSelectedIds(students.map((item) => item.id));
  }, [students]);
  useEffect(() => {
    if (!examDate || !classId) return;
    apiRequest("get_absence_context", { exam_date: examDate, class_id: classId, period })
      .then((value) => setAbsentIds(value.absences.map((item) => item.student_id)))
      .catch(() => setAbsentIds([]));
  }, [examDate, classId, period]);

  function createChart() {
    const next = generateSeatAssignment({
      students: manualPool.length ? manualPool : students,
      rows,
      cols,
      startSide,
      disabledSeats,
      mode,
      selectedIds,
      absentIds,
    });
    setResult(next);
    setError(next.error);
    if (next.ok) setMessage("자리배치를 생성했습니다. 결시자 좌석은 유지됩니다.");
  }

  function toggleSeat(key) {
    setDisabledSeats((items) => items.includes(key) ? items.filter((item) => item !== key) : [...items, key]);
  }

  async function saveChart() {
    if (!result.assignments.length) return setError("자리배치를 먼저 생성하세요.");
    try {
      const response = await apiRequest("save_seat_chart", {
        chart: {
          mode,
          exam_date: examDate,
          period,
          grade: selectedClass?.grade || "",
          class_id: classId,
          subject_name: "",
          room_name: roomName,
          rows,
          cols,
          start_side: startSide,
          disabled_seats: disabledSeats,
          examinee_ids: selectedIds,
          assignment: result.assignments.map((item) => ({
            key: item.key,
            student_id: item.student?.id || "",
            absent: item.absent,
          })),
          updated_at: new Date().toISOString(),
        },
      });
      onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
      setMessage("자리배치를 저장했습니다.");
    } catch (nextError) {
      setError(messageFrom(nextError));
    }
  }

  function loadChart(chart) {
    setMode(chart.mode);
    setExamDate(chart.exam_date);
    setPeriod(Number(chart.period));
    setClassId(chart.class_id);
    setRoomName(chart.room_name);
    setRows(Number(chart.rows));
    setCols(Number(chart.cols));
    setStartSide(chart.start_side);
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
  }

  async function deleteChart(chart) {
    if (!window.confirm(`'${chart.room_name}' 자리배치를 삭제할까요?`)) return;
    const response = await apiRequest("delete_seat_chart", { id: chart.id });
    onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
  }

  async function importManualList(file) {
    if (!file) return;
    try {
      const rowsFromFile = normalizeManualSeatRows(await readWorkbookRows(file));
      const classByNumber = new Map(
        bootstrap.classes
          .filter((item) => Number(item.grade) === Number(selectedClass?.grade))
          .map((item) => [Number(item.class_num), item]),
      );
      const matched = rowsFromFile.map((row) => {
        const classInfo = classByNumber.get(Number(row.class_num));
        return bootstrap.students.find((student) =>
          classInfo &&
          String(student.class_id) === String(classInfo.id) &&
          Number(student.number) === Number(row.number) &&
          student.name === row.name
        );
      }).filter(Boolean);
      if (!matched.length) throw new Error("학급·번호·이름이 등록 명단과 일치하는 학생이 없습니다.");
      setManualPool(matched);
      setSelectedIds([...new Set(matched.map((item) => item.id))]);
      setMessage(`수동 명단에서 ${matched.length}명을 선택했습니다.`);
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      manualFileRef.current.value = "";
    }
  }

  async function createElectiveBatch() {
    const grade = Number(selectedClass?.grade);
    let batchAbsentIds = absentIds;
    try {
      const hq = await apiRequest("get_hq_status", { exam_date: examDate, period, grade, class_id: "" });
      batchAbsentIds = (hq.absences || []).map((item) => item.student_id);
    } catch (nextError) {
      setError(messageFrom(nextError));
      return;
    }
    const subjects = new Set(
      bootstrap.timetable
        .filter((item) =>
          item.exam_date === examDate &&
          Number(item.period) === Number(period) &&
          Number(item.grade) === grade &&
          item.subject_type === "elective"
        )
        .map((item) => item.subject_name),
    );
    const studentMap = new Map(bootstrap.students.map((item) => [String(item.id), item]));
    const groups = new Map();
    bootstrap.enrollments
      .filter((item) => Number(item.grade) === grade && subjects.has(item.subject_name))
      .forEach((item) => {
        const key = `${item.subject_name}|${item.room_name || "별실"}`;
        if (!groups.has(key)) groups.set(key, { subject_name: item.subject_name, room_name: item.room_name || "별실", students: [] });
        const student = studentMap.get(String(item.student_id));
        if (student) groups.get(key).students.push(student);
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
        disabled_seats: disabledSeats,
        examinee_ids: group.students.map((item) => item.id),
        assignment: generated.assignments.map((item) => ({
          key: item.key,
          student_id: item.student?.id || "",
          absent: item.absent,
        })),
        updated_at: new Date().toISOString(),
      });
    }
    try {
      const response = await apiRequest("save_seat_charts_batch", { charts });
      onBootstrap({ ...bootstrap, seat_charts: response.seat_charts });
      setMessage(`선택과목·호실 기준 자리배치 ${charts.length}개를 일괄 생성해 저장했습니다.`);
      setError("");
    } catch (nextError) {
      setError(messageFrom(nextError));
    }
  }

  return (
    <section className="page-section">
      <SectionHeader
        eyebrow="SEATING PLAN"
        title="자리배치"
        description="별실 또는 각자 교실 기준으로 좌석을 만들고 저장·불러오기·인쇄할 수 있습니다."
        actions={<button className="button button-light" onClick={() => window.print()}><IconPrinter size={17} /> PDF·인쇄</button>}
      />
      <div className="segmented no-print">
        <button className={mode === "separate" ? "active" : ""} onClick={() => setMode("separate")}><IconArmchair size={17} /> 별실 배치</button>
        <button className={mode === "own" ? "active" : ""} onClick={() => setMode("own")}><IconLayoutGrid size={17} /> 각자 교실</button>
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
            <h3>1. 기본 정보</h3>
            <label>고사일<select value={examDate} onChange={(event) => setExamDate(event.target.value)}>{bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label} · {item.exam_date}</option>)}</select></label>
            <div className="two-fields">
              <label>학급<select value={classId} onChange={(event) => setClassId(event.target.value)}>{bootstrap.classes.map((item) => <option key={item.id} value={item.id}>{classLabel(item)}</option>)}</select></label>
              <label>교시<select value={period} onChange={(event) => setPeriod(Number(event.target.value))}>{PERIODS.map((item) => <option key={item}>{item}교시</option>)}</select></label>
            </div>
            <label>고사실<input value={roomName} onChange={(event) => setRoomName(event.target.value)} /></label>
          </div>
          <div className="control-group">
            <h3>2. 좌석 구성</h3>
            <div className="three-fields">
              <label>행<input type="number" min="1" max="12" value={rows} onChange={(event) => setRows(Number(event.target.value))} /></label>
              <label>열<input type="number" min="1" max="12" value={cols} onChange={(event) => setCols(Number(event.target.value))} /></label>
              <label>시작<select value={startSide} onChange={(event) => setStartSide(event.target.value)}><option value="window">창가</option><option value="aisle">복도</option></select></label>
            </div>
            <p className="field-help">아래 격자에서 사용하지 않을 좌석을 눌러 제외하세요.</p>
            <div className="mini-seat-grid" style={{ "--seat-cols": cols }}>
              {Array.from({ length: rows * cols }, (_, index) => {
                const row = Math.floor(index / cols);
                const col = index % cols;
                const key = `${row}-${col}`;
                return <button key={key} className={disabledSeats.includes(key) ? "disabled" : ""} onClick={() => toggleSeat(key)} aria-label={`${row + 1}행 ${col + 1}열 ${disabledSeats.includes(key) ? "복원" : "제외"}`}>{disabledSeats.includes(key) ? <IconX size={13} /> : index + 1}</button>;
              })}
            </div>
          </div>
          <div className="control-group">
            <h3>3. 응시자 선택 <span>{selectedIds.length}명</span></h3>
            <div className="selection-actions"><button onClick={() => setSelectedIds((manualPool.length ? manualPool : students).map((item) => item.id))}>전체</button><button onClick={() => setSelectedIds([])}>해제</button></div>
            {manualPool.length ? <p className="manual-pool-note">수동 명단 {manualPool.length}명 선택됨</p> : null}
            <div className="check-list">
              {students.map((student) => <label key={student.id}><input type="checkbox" checked={selectedIds.includes(student.id)} onChange={() => setSelectedIds((items) => items.includes(student.id) ? items.filter((id) => id !== student.id) : [...items, student.id])} /><span>{student.number}번 {student.name}</span>{absentIds.includes(student.id) ? <em>결시</em> : null}</label>)}
            </div>
          </div>
          <button className="button button-primary button-wide" onClick={createChart}><IconArmchair size={18} /> 자리배치 생성</button>
        </aside>

        <div className="seat-preview print-section">
          <div className="seat-title">
            <p>{formatKoreanDate(examDate)} · {period}교시</p>
            <h3>{roomName} 자리배치표</h3>
            <span>{selectedClass ? classLabel(selectedClass) : ""} · 응시 {selectedIds.length}명</span>
          </div>
          <div className="teacher-desk">교탁</div>
          {result.assignments.length ? (
            <div className="seat-grid" style={{ "--seat-cols": cols }}>
              {Array.from({ length: rows * cols }, (_, index) => {
                const row = Math.floor(index / cols);
                const col = index % cols;
                const key = `${row}-${col}`;
                if (disabledSeats.includes(key)) return <div className="seat-hole" key={key} />;
                const item = result.assignments.find((slot) => slot.key === key);
                return (
                  <div className={`seat ${item?.absent ? "seat-absent" : ""} ${!item?.student ? "seat-empty" : ""}`} key={key}>
                    <span>{row + 1}-{col + 1}</span>
                    {item?.student ? <><strong>{item.student.number}번 {item.student.name}</strong>{item.absent ? <em>결시 · 자리 유지</em> : null}</> : <small>빈 좌석</small>}
                  </div>
                );
              })}
            </div>
          ) : <EmptyState icon={IconArmchair} title="아직 생성된 자리배치가 없습니다" description="응시자와 좌석 구성을 확인한 뒤 자리배치 생성을 누르세요." />}
          {result.assignments.length ? <div className="seat-preview-actions no-print"><button className="button button-gold" onClick={saveChart}><IconDeviceFloppy size={17} /> 저장</button><button className="button button-light" onClick={() => window.print()}><IconPrinter size={17} /> 인쇄</button></div> : null}
        </div>
      </div>

      <article className="panel-card saved-charts no-print">
        <div className="card-heading"><div><IconDeviceFloppy size={20} /><h3>저장된 자리배치</h3></div><span>{bootstrap.seat_charts.length}개</span></div>
        {bootstrap.seat_charts.length ? <div className="saved-list">{bootstrap.seat_charts.map((chart) => (
          <div key={chart.id}><div><strong>{chart.room_name}</strong><span>{chart.exam_date} · {chart.period}교시 · {chart.class_id ? classLabel(bootstrap.classes.find((item) => item.id === chart.class_id)) : `${chart.grade}학년`}</span></div><div><button className="button button-light" onClick={() => loadChart(chart)}>불러오기</button><button className="icon-button danger" onClick={() => deleteChart(chart)} aria-label="삭제"><IconTrash size={17} /></button></div></div>
        ))}</div> : <EmptyState icon={IconDeviceFloppy} title="저장된 배치가 없습니다" description="생성한 자리배치를 저장하면 이곳에서 다시 불러올 수 있습니다." />}
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
  const [timetableDraft, setTimetableDraft] = useState({
    exam_date: selectFirst(bootstrap.exam_dates, "exam_date"),
    grade: 1,
    period: 1,
    start_time: "09:00",
    end_time: "09:50",
    subject_name: "",
    subject_type: "common",
    room_name: "",
  });
  const [personalStudentId, setPersonalStudentId] = useState(selectFirst(bootstrap.students));
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
  }, [bootstrap.students, selectedClassId]);

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
    const next = [...bootstrap.exam_dates.filter((item) => item.exam_date !== newDate), {
      id: makeId("date"), exam_date: newDate, label: newDateLabel.trim() || `${bootstrap.exam_dates.length + 1}일차`, active: true,
    }].sort((a, b) => a.exam_date.localeCompare(b.exam_date));
    if (await commit("save_exam_dates", { exam_dates: next }, "고사일을 저장했습니다.")) {
      setNewDate(""); setNewDateLabel("");
    }
  }

  async function addClass() {
    const result = await commit("save_class", {
      class_item: { grade: Number(grade), class_num: Number(classNum), active: true },
    }, "학급을 저장했습니다.");
    if (result?.classes?.length) setSelectedClassId(result.classes.at(-1).id);
  }

  async function saveStudents() {
    const students = studentText.split(/\r?\n/).map((line) => {
      const [number, ...nameParts] = line.trim().split(/[\t, ]+/);
      return { number: Number(number), name: nameParts.join(" ").trim() };
    }).filter((item) => Number.isInteger(item.number) && item.number > 0 && item.name);
    const unique = new Map(students.map((item) => [item.number, item]));
    await commit("replace_students", { class_id: selectedClassId, students: [...unique.values()] }, `${unique.size}명의 학생 명단을 저장했습니다.`);
  }

  async function importStudentFile(file) {
    if (!file) return;
    try {
      const rows = normalizeStudentRows(await readWorkbookRows(file));
      if (!rows.length) throw new Error("가져올 학생이 없습니다. 열 이름과 값 형식을 확인하세요.");
      await commit("import_students", { rows }, `${rows.length}명의 학생을 가져왔습니다.`);
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      studentFileRef.current.value = "";
    }
  }

  async function addTimetable() {
    if (!timetableDraft.subject_name.trim()) return setError("과목명을 입력하세요.");
    const next = [...bootstrap.timetable, {
      ...timetableDraft,
      id: makeId("timetable"),
      grade: Number(timetableDraft.grade),
      period: Number(timetableDraft.period),
      class_ids: bootstrap.classes.filter((item) => Number(item.grade) === Number(timetableDraft.grade)).map((item) => item.id).join("|"),
    }];
    if (await commit("save_timetable", { timetable: next }, "시간표 항목을 저장했습니다.")) {
      setTimetableDraft((value) => ({ ...value, subject_name: "", room_name: "" }));
    }
  }

  async function removeTimetable(id) {
    await commit("save_timetable", { timetable: bootstrap.timetable.filter((item) => item.id !== id) }, "시간표 항목을 삭제했습니다.");
  }

  async function importEnrollmentFile(file) {
    if (!file) return;
    try {
      const raw = normalizeEnrollmentRows(await readWorkbookRows(file));
      const targetGrade = Number(timetableDraft.grade);
      const classMap = new Map(bootstrap.classes.filter((item) => Number(item.grade) === targetGrade).map((item) => [Number(item.class_num), item]));
      const studentMap = new Map(bootstrap.students.map((item) => [`${item.class_id}|${Number(item.number)}`, item]));
      const enrollments = raw.map((row) => {
        const classInfo = classMap.get(Number(row.class_num));
        const student = classInfo ? studentMap.get(`${classInfo.id}|${Number(row.number)}`) : null;
        if (!classInfo || !student || student.name !== row.name) return null;
        return { id: makeId("enrollment"), grade: targetGrade, class_id: classInfo.id, student_id: student.id, subject_name: row.subject_name, room_name: row.room_name };
      }).filter(Boolean);
      if (!enrollments.length) throw new Error("학급·번호·이름이 학생 명단과 일치하는 선택과목 행이 없습니다.");
      await commit("save_enrollments", { grade: targetGrade, enrollments }, `${enrollments.length}건의 선택과목을 저장했습니다.`);
    } catch (nextError) {
      setError(messageFrom(nextError));
    } finally {
      enrollmentFileRef.current.value = "";
    }
  }

  const personalStudent = bootstrap.students.find((item) => String(item.id) === String(personalStudentId));
  const personalClass = bootstrap.classes.find((item) => String(item.id) === String(personalStudent?.class_id));
  const personalRows = buildPersonalTimetable({
    student: personalStudent,
    classInfo: personalClass,
    timetable: bootstrap.timetable,
    enrollments: bootstrap.enrollments,
  });

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
                <label>날짜<input type="date" value={newDate} onChange={(event) => setNewDate(event.target.value)} /></label>
                <label>표시 이름<input value={newDateLabel} onChange={(event) => setNewDateLabel(event.target.value)} placeholder="예: 1일차" /></label>
                <button className="button button-secondary" onClick={addDate}><IconPlus size={17} /> 추가</button>
              </div>
              <div className="item-list">{bootstrap.exam_dates.map((item) => <div key={item.id}><div><strong>{item.label}</strong><span>{formatKoreanDate(item.exam_date)}</span></div><button className="icon-button danger" onClick={() => commit("save_exam_dates", { exam_dates: bootstrap.exam_dates.filter((date) => date.id !== item.id) }, "고사일을 삭제했습니다.")}><IconTrash size={17} /></button></div>)}</div>
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
                  <button className="button button-danger" onClick={() => window.confirm("학급과 소속 학생을 모두 삭제할까요?") && commit("delete_class", { class_id: selectedClassId }, "학급을 삭제했습니다.")}><IconTrash size={17} /> 학급 삭제</button>
                </div>
              </> : null}
            </article>
          ) : null}

          {active === "timetable" ? (
            <article className="panel-card">
              <div className="card-heading"><div><IconClock size={20} /><h3>공통·선택과목 시간표</h3></div></div>
              <div className="form-grid">
                <label>고사일<select value={timetableDraft.exam_date} onChange={(event) => setTimetableDraft({ ...timetableDraft, exam_date: event.target.value })}>{bootstrap.exam_dates.map((item) => <option key={item.id} value={item.exam_date}>{item.label}</option>)}</select></label>
                <label>학년<input type="number" min="1" value={timetableDraft.grade} onChange={(event) => setTimetableDraft({ ...timetableDraft, grade: event.target.value })} /></label>
                <label>교시<select value={timetableDraft.period} onChange={(event) => setTimetableDraft({ ...timetableDraft, period: event.target.value })}>{PERIODS.map((item) => <option key={item}>{item}교시</option>)}</select></label>
                <label>시작<input type="time" value={timetableDraft.start_time} onChange={(event) => setTimetableDraft({ ...timetableDraft, start_time: event.target.value })} /></label>
                <label>종료<input type="time" value={timetableDraft.end_time} onChange={(event) => setTimetableDraft({ ...timetableDraft, end_time: event.target.value })} /></label>
                <label>구분<select value={timetableDraft.subject_type} onChange={(event) => setTimetableDraft({ ...timetableDraft, subject_type: event.target.value })}><option value="common">공통과목</option><option value="elective">선택과목</option></select></label>
                <label className="span-two">과목명<input value={timetableDraft.subject_name} onChange={(event) => setTimetableDraft({ ...timetableDraft, subject_name: event.target.value })} /></label>
                <label>기본 호실<input value={timetableDraft.room_name} onChange={(event) => setTimetableDraft({ ...timetableDraft, room_name: event.target.value })} /></label>
              </div>
              <button className="button button-secondary" onClick={addTimetable}><IconPlus size={17} /> 시간표 추가</button>
              <div className="table-wrap"><table><thead><tr><th>고사일</th><th>학년</th><th>교시</th><th>시간</th><th>구분</th><th>과목</th><th></th></tr></thead><tbody>{bootstrap.timetable.map((item) => <tr key={item.id}><td>{item.exam_date}</td><td>{item.grade}</td><td>{item.period}</td><td>{item.start_time}~{item.end_time}</td><td>{item.subject_type === "common" ? "공통" : "선택"}</td><td>{item.subject_name}</td><td><button className="icon-button danger" onClick={() => removeTimetable(item.id)}><IconTrash size={16} /></button></td></tr>)}</tbody></table></div>
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
            </article>
          ) : null}

          {active === "print" ? (
            <article className="panel-card print-section">
              <div className="card-heading no-print"><div><IconPrinter size={20} /><h3>학생별 개인 시간표</h3></div><button className="button button-light" onClick={() => window.print()}><IconPrinter size={17} /> 인쇄</button></div>
              <label className="no-print">학생<select value={personalStudentId} onChange={(event) => setPersonalStudentId(event.target.value)}>{bootstrap.classes.map((classInfo) => <optgroup key={classInfo.id} label={classLabel(classInfo)}>{bootstrap.students.filter((item) => item.class_id === classInfo.id).map((student) => <option key={student.id} value={student.id}>{student.number}번 {student.name}</option>)}</optgroup>)}</select></label>
              <div className="personal-sheet">
                <p>PERSONAL EXAM SCHEDULE</p><h3>개인 고사 시간표</h3>
                <div className="student-meta"><span>{personalClass ? classLabel(personalClass) : ""}</span><strong>{personalStudent?.number}번 {personalStudent?.name}</strong></div>
                <table><thead><tr><th>고사일</th><th>교시</th><th>시간</th><th>과목</th><th>고사실</th></tr></thead><tbody>{personalRows.map((row) => <tr key={`${row.exam_date}-${row.period}`}><td>{formatKoreanDate(row.exam_date)}</td><td>{row.period}</td><td>{row.time}</td><td>{row.subject_name}</td><td>{row.room_name}</td></tr>)}</tbody></table>
              </div>
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
    if (!loggedIn) return undefined;
    let timer;
    const reset = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
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
