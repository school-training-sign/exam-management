import {
  absenceKey,
  classLabel,
  completionKey,
  filterAbsences,
  makeId,
  sortClasses,
  sortStudents,
} from "./core.js";
import { createDemoState } from "./demo-data.js";

const SCHOOL_SESSION_KEY = "exam-management:school-session";
const ADMIN_SESSION_KEY = "exam-management:admin-session";

function readConfig() {
  return window.EXAM_MANAGEMENT_CONFIG || {};
}

function decorateAbsences(state, rows) {
  const studentMap = new Map(state.students.map((item) => [String(item.id), item]));
  const classMap = new Map(state.classes.map((item) => [String(item.id), item]));
  return rows.map((item) => {
    const student = studentMap.get(String(item.student_id));
    const classInfo = classMap.get(String(item.class_id));
    return {
      ...item,
      grade: classInfo?.grade,
      class_num: classInfo?.class_num,
      class_label: classLabel(classInfo),
      student_number: student?.number,
      student_name: student?.name,
    };
  });
}

class DemoRepository {
  constructor() {
    this.state = createDemoState();
  }

  async request(action, payload = {}) {
    await new Promise((resolve) => setTimeout(resolve, 70));
    const state = this.state;

    switch (action) {
      case "school_login":
        if (!String(payload.school_code || "").trim()) {
          throw new Error("학교코드를 입력하세요.");
        }
        return { session: "demo-teacher-session", bootstrap: this.teacherBootstrap() };

      case "admin_login":
        if (payload.password !== "demo-admin") {
          throw new Error("데모 관리자 암호가 올바르지 않습니다.");
        }
        return { admin_session: "demo-admin-session", bootstrap: this.adminBootstrap() };

      case "get_bootstrap":
        return this.teacherBootstrap();

      case "get_admin_bootstrap":
        return this.adminBootstrap();

      case "get_absence_context": {
        const students = sortStudents(
          state.students.filter((item) => String(item.class_id) === String(payload.class_id)),
        );
        const absences = state.absences.filter(
          (item) =>
            item.exam_date === payload.exam_date &&
            Number(item.period) === Number(payload.period) &&
            String(item.class_id) === String(payload.class_id),
        );
        const completion = state.completions.find(
          (item) =>
            item.exam_date === payload.exam_date &&
            Number(item.period) === Number(payload.period) &&
            String(item.class_id) === String(payload.class_id),
        );
        return { students, absences, completed: Boolean(completion) };
      }

      case "set_absence": {
        const key = absenceKey(payload);
        state.absences = state.absences.filter((item) => item.key !== key);
        if (payload.absent) {
          state.absences.push({
            key,
            id: makeId("absence"),
            exam_date: payload.exam_date,
            period: Number(payload.period),
            class_id: payload.class_id,
            student_id: payload.student_id,
            reason: payload.reason || "미인정",
            reason_detail: payload.reason === "기타" ? payload.reason_detail || "" : "",
            updated_at: new Date().toISOString(),
          });
        }
        return { saved: true };
      }

      case "complete_input": {
        const key = completionKey(payload);
        state.completions = state.completions.filter((item) => item.key !== key);
        state.completions.push({
          key,
          id: makeId("completion"),
          exam_date: payload.exam_date,
          period: Number(payload.period),
          class_id: payload.class_id,
          completed_at: new Date().toISOString(),
        });
        return { completed: true };
      }

      case "get_hq_status": {
        const absences = decorateAbsences(
          state,
          filterAbsences(state.absences, {
            examDate: payload.exam_date,
            period: payload.period,
            grade: payload.grade,
            classId: payload.class_id,
          }),
        );
        return { absences, completions: structuredClone(state.completions) };
      }

      case "get_period_report": {
        const absences = decorateAbsences(
          state,
          filterAbsences(state.absences, {
            startDate: payload.start_date,
            endDate: payload.end_date,
            grade: payload.grade,
            classId: payload.class_id,
          }),
        );
        return { absences };
      }

      case "save_exam_dates":
        state.examDates = payload.exam_dates;
        return this.adminBootstrap();

      case "save_class": {
        const existing = state.classes.find((item) => item.id === payload.class_item.id);
        if (existing) Object.assign(existing, payload.class_item);
        else state.classes.push({ ...payload.class_item, id: makeId("class") });
        state.classes = sortClasses(state.classes);
        return this.adminBootstrap();
      }

      case "delete_class":
        state.classes = state.classes.filter((item) => item.id !== payload.class_id);
        state.students = state.students.filter((item) => item.class_id !== payload.class_id);
        return this.adminBootstrap();

      case "replace_students":
        state.students = state.students.filter((item) => item.class_id !== payload.class_id);
        state.students.push(
          ...payload.students.map((item) => ({
            ...item,
            id: item.id || makeId("student"),
            class_id: payload.class_id,
          })),
        );
        return this.adminBootstrap();

      case "import_students": {
        payload.rows.forEach((row) => {
          let classInfo = state.classes.find(
            (item) =>
              Number(item.grade) === Number(row.grade) &&
              Number(item.class_num) === Number(row.class_num),
          );
          if (!classInfo) {
            classInfo = {
              id: makeId("class"),
              grade: Number(row.grade),
              class_num: Number(row.class_num),
              active: true,
            };
            state.classes.push(classInfo);
          }
          state.students = state.students.filter(
            (item) =>
              !(
                item.class_id === classInfo.id &&
                Number(item.number) === Number(row.number)
              ),
          );
          state.students.push({
            id: makeId("student"),
            class_id: classInfo.id,
            number: Number(row.number),
            name: row.name,
          });
        });
        state.classes = sortClasses(state.classes);
        return this.adminBootstrap();
      }

      case "save_timetable":
        state.timetable = payload.timetable;
        return this.adminBootstrap();

      case "save_enrollments":
        state.enrollments = state.enrollments
          .filter((item) => Number(item.grade) !== Number(payload.grade))
          .concat(payload.enrollments);
        return this.adminBootstrap();

      case "save_seat_chart": {
        const item = { ...payload.chart, id: payload.chart.id || makeId("seat") };
        state.seatCharts = state.seatCharts.filter((chart) => chart.id !== item.id);
        state.seatCharts.unshift(item);
        return { chart: item, seat_charts: state.seatCharts };
      }

      case "save_seat_charts_batch":
        payload.charts.forEach((chart) => {
          state.seatCharts.unshift({ ...chart, id: chart.id || makeId("seat") });
        });
        return { seat_charts: state.seatCharts };

      case "delete_seat_chart":
        state.seatCharts = state.seatCharts.filter((item) => item.id !== payload.id);
        return { seat_charts: state.seatCharts };

      case "cleanup":
        state.absences = state.absences.filter((item) => item.exam_date >= payload.cutoff_date);
        state.completions = state.completions.filter((item) => item.exam_date >= payload.cutoff_date);
        state.examDates = state.examDates.filter((item) => item.exam_date >= payload.cutoff_date);
        state.timetable = state.timetable.filter((item) => item.exam_date >= payload.cutoff_date);
        state.seatCharts = state.seatCharts.filter((item) => item.exam_date >= payload.cutoff_date);
        return this.adminBootstrap();

      case "logout":
        return { logged_out: true };

      default:
        throw new Error(`지원하지 않는 데모 작업입니다: ${action}`);
    }
  }

  teacherBootstrap() {
    return {
      settings: this.state.settings,
      classes: sortClasses(this.state.classes),
      exam_dates: structuredClone(this.state.examDates),
    };
  }

  adminBootstrap() {
    return {
      ...this.teacherBootstrap(),
      students: sortStudents(this.state.students),
      timetable: structuredClone(this.state.timetable),
      enrollments: structuredClone(this.state.enrollments),
      seat_charts: structuredClone(this.state.seatCharts),
    };
  }
}

class AppsScriptRepository {
  constructor(apiUrl) {
    this.apiUrl = apiUrl;
  }

  async request(action, payload = {}) {
    if (!this.apiUrl || this.apiUrl.includes("__APPS_SCRIPT")) {
      throw new Error(
        "학교 백엔드 주소가 아직 설정되지 않았습니다. 데모로 먼저 확인하거나 config.js를 연결하세요.",
      );
    }
    const body = {
      action,
      ...payload,
      school_session: sessionStorage.getItem(SCHOOL_SESSION_KEY) || "",
      admin_session: sessionStorage.getItem(ADMIN_SESSION_KEY) || "",
      request_id: payload.request_id || crypto.randomUUID(),
    };
    const delays = [0, 300, 900];
    let lastError;

    for (const delay of delays) {
      if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
      try {
        const response = await fetch(this.apiUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(body),
          redirect: "follow",
        });
        const result = await response.json();
        if (!result.ok) throw new Error(result.error?.message || "요청을 처리하지 못했습니다.");
        return result.data;
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("학교 서버에 연결하지 못했습니다.");
  }
}

const demoMode = new URLSearchParams(location.search).get("demo") === "1";
const repository = demoMode
  ? new DemoRepository()
  : new AppsScriptRepository(readConfig().API_URL || "");

export const appMode = {
  demo: demoMode,
  name: demoMode ? "데모" : "운영",
};

export function saveSchoolSession(value) {
  if (value) sessionStorage.setItem(SCHOOL_SESSION_KEY, value);
  else sessionStorage.removeItem(SCHOOL_SESSION_KEY);
}

export function saveAdminSession(value) {
  if (value) sessionStorage.setItem(ADMIN_SESSION_KEY, value);
  else sessionStorage.removeItem(ADMIN_SESSION_KEY);
}

export function hasSchoolSession() {
  return Boolean(sessionStorage.getItem(SCHOOL_SESSION_KEY));
}

export function hasAdminSession() {
  return Boolean(sessionStorage.getItem(ADMIN_SESSION_KEY));
}

export function clearSessions() {
  sessionStorage.removeItem(SCHOOL_SESSION_KEY);
  sessionStorage.removeItem(ADMIN_SESSION_KEY);
}

export async function apiRequest(action, payload) {
  return repository.request(action, payload);
}
