import {
  aggregateAbsenceReport,
  absenceKey,
  classLabel,
  completionKey,
  filterAbsences,
  makeId,
  seatChartKey,
  sortClasses,
  sortStudents,
  summarizeAbsences,
} from "./core.js";
import { createDemoState } from "./demo-data.js";

const SCHOOL_SESSION_KEY = "exam-management:school-session";
const ADMIN_SESSION_KEY = "exam-management:admin-session";
const CLIENT_ID_KEY = "exam-management:client-id";
// Apps Script mutations may wait up to 28 seconds for the spreadsheet lock.
// Keep the browser deadline comfortably above that ceiling so a request cannot
// be aborted while the server is still waiting to commit it.
const REQUEST_TIMEOUT_MS = 45_000;
const PRESENCE_TIMEOUT_MS = 8_000;
const TRANSIENT_CODES = new Set(["NETWORK_ERROR", "TIMEOUT", "LOCK_TIMEOUT", "BUSY", "TEMPORARY_ERROR"]);
const SESSION_CODES = new Set(["SESSION_EXPIRED", "INVALID_SESSION", "ADMIN_SESSION_EXPIRED"]);

function getClientId() {
  let value = sessionStorage.getItem(CLIENT_ID_KEY);
  if (!value) {
    value = crypto.randomUUID?.() ||
      `client-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    sessionStorage.setItem(CLIENT_ID_KEY, value);
  }
  return value;
}

export class ApiError extends Error {
  constructor(message, { code = "UNKNOWN_ERROR", retryable = false, details = null } = {}) {
    super(message);
    this.name = "ApiError";
    this.code = code;
    this.retryable = retryable;
    this.details = details;
  }
}

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

function validateDemoSeatChart(chart = {}) {
  if (chart.mode !== "own" && (!String(chart.subject_name || "").trim() || !String(chart.room_name || "").trim())) {
    throw new ApiError("별실 자리배치에는 과목과 호실이 필요합니다.", {
      code: "SEAT_SUBJECT_ROOM_REQUIRED",
    });
  }
  const rows = Number(chart.rows) || 6;
  const cols = Number(chart.cols) || 5;
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) {
    throw new ApiError("자리배치 행과 열을 확인하세요.", { code: "INVALID_SEAT_SIZE" });
  }
}

function parseClassIds(value) {
  if (Array.isArray(value)) return value.map(String);
  if (typeof value !== "string") return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed.map(String);
  } catch (ignored) {
    // Pipe/comma values are supported for older demo snapshots.
  }
  return value.split(/[|,]/).map((item) => item.trim()).filter(Boolean);
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

      case "presence_ping":
        return {
          connected: true,
          online_count: 0,
          observed_at: new Date().toISOString(),
          active_window_seconds: 90,
          approximate: true,
          stale: false,
        };

      case "get_absence_context": {
        const students = sortStudents(
          state.students.filter(
            (item) =>
              item.active !== false &&
              String(item.class_id) === String(payload.class_id),
          ),
        );
        const activeStudentIds = new Set(students.map((item) => String(item.id)));
        const allAbsences = state.absences.filter(
          (item) =>
            item.exam_date === payload.exam_date &&
            Number(item.period) === Number(payload.period) &&
            String(item.class_id) === String(payload.class_id),
        );
        const absences = allAbsences.filter((item) => activeStudentIds.has(String(item.student_id)));
        const completion = state.completions.find(
          (item) =>
            item.exam_date === payload.exam_date &&
            Number(item.period) === Number(payload.period) &&
            String(item.class_id) === String(payload.class_id),
        );
        return {
          students,
          absences,
          completed: Boolean(completion),
          submitted: Boolean(completion),
          revision: Number(completion?.revision || 0),
          archived_absence_count: allAbsences.length - absences.length,
        };
      }

      case "submit_absence_input": {
        const key = completionKey(payload);
        const current = state.completions.find((item) => item.key === key);
        const currentRevision = Number(current?.revision || 0);
        if (Number(payload.expected_revision || 0) !== currentRevision) {
          throw new ApiError("다른 사용자가 먼저 수정했습니다.", {
            code: "REVISION_CONFLICT",
            details: { current_revision: currentRevision },
          });
        }
        const studentIds = new Set(
          state.students
            .filter(
              (item) =>
                item.active !== false &&
                String(item.class_id) === String(payload.class_id),
            )
            .map((item) => String(item.id)),
        );
        const nextAbsences = [];
        const seen = new Set();
        const allowedReasons = new Set(["인정", "미인정", "질병", "기타"]);
        (payload.absences || []).forEach((item) => {
          const studentId = String(item.student_id || "");
          if (!studentIds.has(studentId) || seen.has(studentId)) {
            throw new ApiError("결시 명단에 잘못된 학생이 포함되어 있습니다.", {
              code: "VALIDATION_ERROR",
            });
          }
          seen.add(studentId);
          const reason = String(item.reason || "");
          if (!allowedReasons.has(reason)) {
            throw new ApiError("결시 사유가 올바르지 않습니다.", { code: "INVALID_ABSENCE_REASON" });
          }
          if (reason === "기타" && !String(item.reason_detail || "").trim()) {
            throw new ApiError("기타 사유의 상세 내용을 입력하세요.", {
              code: "ABSENCE_REASON_DETAIL_REQUIRED",
            });
          }
          const existing = state.absences.find((row) =>
            row.exam_date === payload.exam_date &&
            Number(row.period) === Number(payload.period) &&
            String(row.class_id) === String(payload.class_id) &&
            String(row.student_id) === studentId
          );
          nextAbsences.push({
            key: absenceKey({ ...payload, student_id: item.student_id }),
            id: existing?.id || makeId("absence"),
            exam_date: payload.exam_date,
            period: Number(payload.period),
            class_id: payload.class_id,
            student_id: item.student_id,
            reason,
            reason_detail: reason === "기타" ? String(item.reason_detail || "").trim() : "",
            updated_at: new Date().toISOString(),
          });
        });
        const preservedInactive = state.absences.filter((item) =>
          item.exam_date === payload.exam_date &&
          Number(item.period) === Number(payload.period) &&
          String(item.class_id) === String(payload.class_id) &&
          !studentIds.has(String(item.student_id))
        );
        state.absences = state.absences
          .filter(
            (item) =>
              !(
                item.exam_date === payload.exam_date &&
                Number(item.period) === Number(payload.period) &&
                String(item.class_id) === String(payload.class_id)
              ),
          )
          .concat(preservedInactive, nextAbsences);
        const completion = {
          key,
          id: current?.id || makeId("completion"),
          exam_date: payload.exam_date,
          period: Number(payload.period),
          class_id: payload.class_id,
          completed_at: new Date().toISOString(),
          revision: currentRevision + 1,
        };
        state.completions = state.completions.filter((item) => item.key !== key);
        state.completions.push(completion);
        return {
          submitted: true,
          completed: true,
          revision: completion.revision,
          absences: structuredClone(nextAbsences),
          archived_absence_count: preservedInactive.length,
        };
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
        const targetClasses = state.classes.filter(
          (item) =>
            item.active !== false &&
            (!payload.grade || Number(item.grade) === Number(payload.grade)) &&
            (!payload.class_id || String(item.id) === String(payload.class_id)),
        );
        const targetClassIds = new Set(targetClasses.map((item) => String(item.id)));
        const completions = state.completions.filter(
          (item) =>
            item.exam_date === payload.exam_date &&
            Number(item.period) === Number(payload.period) &&
            targetClassIds.has(String(item.class_id)),
        );
        const activeStudents = state.students.filter(
          (item) => item.active !== false && targetClassIds.has(String(item.class_id)),
        );
        const activeStudentIds = new Set(activeStudents.map((item) => String(item.id)));
        const activeAbsences = absences.filter((item) => activeStudentIds.has(String(item.student_id)));
        const enrolled = activeStudents.length;
        return {
          absences,
          completions: structuredClone(completions),
          summary: {
            enrolled,
            absent: activeAbsences.length,
            present: Math.max(0, enrolled - activeAbsences.length),
            archived_absences: absences.length - activeAbsences.length,
            reasons: summarizeAbsences(activeAbsences),
          },
        };
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
        return { absences, ...aggregateAbsenceReport(absences) };
      }

      case "save_exam_dates": {
        const incomingDates = new Set((payload.exam_dates || []).map((item) => item.exam_date));
        if (state.examDates.some((item) => item.active !== false && !incomingDates.has(item.exam_date))) {
          throw new ApiError("고사일 삭제는 삭제 버튼을 사용하세요.", {
            code: "EXAM_DATE_DELETE_ACTION_REQUIRED",
          });
        }
        const oldByDate = new Map(state.examDates.map((item) => [item.exam_date, item]));
        state.examDates = (payload.exam_dates || []).map((item) => ({
          ...item,
          id: oldByDate.get(item.exam_date)?.id || item.id || makeId("date"),
        }));
        return this.adminBootstrap();
      }

      case "delete_exam_date": {
        const impact = {
          timetable: state.timetable.filter((item) => item.exam_date === payload.exam_date).length,
          absences: state.absences.filter((item) => item.exam_date === payload.exam_date).length,
          completions: state.completions.filter((item) => item.exam_date === payload.exam_date).length,
          seat_charts: state.seatCharts.filter((item) => item.exam_date === payload.exam_date).length,
        };
        if (payload.preview_only) return { deleted: false, preview: true, impact };
        state.examDates = state.examDates.map((item) =>
          item.exam_date === payload.exam_date ? { ...item, active: false } : item,
        );
        state.timetable = state.timetable.filter((item) => item.exam_date !== payload.exam_date);
        state.absences = state.absences.filter((item) => item.exam_date !== payload.exam_date);
        state.completions = state.completions.filter((item) => item.exam_date !== payload.exam_date);
        state.seatCharts = state.seatCharts.filter((item) => item.exam_date !== payload.exam_date);
        return { ...this.adminBootstrap(), deleted: true, impact };
      }

      case "save_class": {
        const incoming = payload.class_item || {};
        const duplicate = state.classes.find((item) =>
          Number(item.grade) === Number(incoming.grade) &&
          Number(item.class_num) === Number(incoming.class_num) &&
          String(item.id) !== String(incoming.id || "")
        );
        if (duplicate?.active !== false) {
          throw new ApiError("같은 학년과 반이 이미 있습니다.", { code: "DUPLICATE_CLASS" });
        }
        const existing = state.classes.find((item) => item.id === incoming.id) || duplicate;
        const saved = existing || { id: makeId("class") };
        Object.assign(saved, incoming, { active: incoming.active !== false });
        if (!existing) state.classes.push(saved);
        state.classes = sortClasses(state.classes);
        return { ...this.adminBootstrap(), saved_class_id: saved.id };
      }

      case "delete_class": {
        const studentIds = new Set(
          state.students
            .filter((item) => String(item.class_id) === String(payload.class_id))
            .map((item) => String(item.id)),
        );
        const relatedCharts = state.seatCharts.filter(
          (chart) =>
            String(chart.class_id) === String(payload.class_id) ||
            (chart.examinee_ids || []).some((id) => studentIds.has(String(id))) ||
            (chart.assignment || []).some((item) => studentIds.has(String(item.student_id))),
        );
        const impact = {
          students: studentIds.size,
          absences: state.absences.filter((item) => String(item.class_id) === String(payload.class_id)).length,
          completions: state.completions.filter((item) => String(item.class_id) === String(payload.class_id)).length,
          enrollments: state.enrollments.filter((item) => studentIds.has(String(item.student_id))).length,
          timetable: state.timetable.filter((item) => {
            const ids = Array.isArray(item.class_ids)
              ? item.class_ids
              : String(item.class_ids || "").split(/[|,]/);
            return ids.map(String).includes(String(payload.class_id));
          }).length,
          seat_charts: relatedCharts.length,
        };
        if (payload.preview_only) return { deleted: false, preview: true, impact };
        state.classes = state.classes.map((item) =>
          String(item.id) === String(payload.class_id) ? { ...item, active: false } : item,
        );
        state.students = state.students.map((item) =>
          studentIds.has(String(item.id)) ? { ...item, active: false } : item,
        );
        state.enrollments = state.enrollments.filter(
          (item) => !studentIds.has(String(item.student_id)),
        );
        state.timetable = state.timetable.flatMap((item) => {
          const ids = Array.isArray(item.class_ids)
            ? item.class_ids.map(String)
            : String(item.class_ids || "").split(/[|,]/).filter(Boolean);
          if (!ids.includes(String(payload.class_id))) return [item];
          const nextIds = ids.filter((id) => id !== String(payload.class_id));
          return nextIds.length ? [{ ...item, class_ids: nextIds.join("|") }] : [];
        });
        state.seatCharts = state.seatCharts.filter((chart) => !relatedCharts.includes(chart));
        return { ...this.adminBootstrap(), deleted: true, impact };
      }

      case "replace_students":
        {
        const existingByNumber = new Map(
          state.students
            .filter((item) => String(item.class_id) === String(payload.class_id))
            .map((item) => [Number(item.number), item]),
        );
        const incomingNumbers = new Set(payload.students.map((item) => Number(item.number)));
        state.students = state.students.map((item) =>
          String(item.class_id) === String(payload.class_id) && !incomingNumbers.has(Number(item.number))
            ? { ...item, active: false }
            : item,
        );
        state.students = state.students.filter(
          (item) =>
            !(
              String(item.class_id) === String(payload.class_id) &&
              incomingNumbers.has(Number(item.number))
            ),
        );
        state.students.push(
          ...payload.students.map((item) => ({
            ...item,
            id: item.id || existingByNumber.get(Number(item.number))?.id || makeId("student"),
            class_id: payload.class_id,
            active: true,
          })),
        );
        return this.adminBootstrap();
        }

      case "save_student": {
        const student = payload.student || payload.student_item || payload;
        const item = {
          ...student,
          id: student.id || makeId("student"),
          number: Number(student.number),
          active: true,
        };
        state.students = state.students.filter((row) => String(row.id) !== String(item.id));
        state.students.push(item);
        return this.adminBootstrap();
      }

      case "delete_student": {
        const affectedCharts = state.seatCharts.filter(
          (chart) =>
            (chart.examinee_ids || []).some((id) => String(id) === String(payload.student_id)) ||
            (chart.assignment || []).some((item) => String(item.student_id) === String(payload.student_id)),
        );
        const impact = {
          absences: state.absences.filter(
            (item) => String(item.student_id) === String(payload.student_id),
          ).length,
          enrollments: state.enrollments.filter(
            (item) => String(item.student_id) === String(payload.student_id),
          ).length,
          seat_charts: affectedCharts.length,
        };
        if (payload.preview_only) return { deleted: false, preview: true, impact };
        state.students = state.students.map((item) =>
          String(item.id) === String(payload.student_id) ? { ...item, active: false } : item,
        );
        state.enrollments = state.enrollments.filter(
          (item) => String(item.student_id) !== String(payload.student_id),
        );
        state.seatCharts = state.seatCharts.filter((chart) => !affectedCharts.includes(chart));
        return { ...this.adminBootstrap(), deleted: true, impact };
      }

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
          const existing = state.students.find(
            (item) =>
              item.class_id === classInfo.id &&
              Number(item.number) === Number(row.number),
          );
          state.students = state.students.filter((item) => item !== existing);
          state.students.push({
            id: existing?.id || makeId("student"),
            class_id: classInfo.id,
            number: Number(row.number),
            name: row.name,
            active: true,
          });
        });
        state.classes = sortClasses(state.classes);
        return this.adminBootstrap();
      }

      case "save_timetable": {
        const incomingIds = new Set((payload.timetable || []).map((item) => String(item.id)));
        if (state.timetable.some((item) => !incomingIds.has(String(item.id)))) {
          throw new ApiError("시간표 삭제는 삭제 버튼을 사용하세요.", {
            code: "TIMETABLE_DELETE_ACTION_REQUIRED",
          });
        }
        state.timetable = payload.timetable;
        return this.adminBootstrap();
      }

      case "delete_timetable": {
        const target = state.timetable.find((item) => String(item.id) === String(payload.id));
        if (!target) {
          throw new ApiError("시간표 항목을 찾지 못했습니다.", { code: "TIMETABLE_NOT_FOUND" });
        }
        const remainingSubjectRows = state.timetable.filter((item) =>
          String(item.id) !== String(target.id) &&
          Number(item.grade) === Number(target.grade) &&
          String(item.subject_name) === String(target.subject_name) &&
          item.subject_type === "elective"
        );
        const removeEnrollments = target.subject_type === "elective" && remainingSubjectRows.length === 0;
        const impact = {
          enrollments: removeEnrollments
            ? state.enrollments.filter((item) =>
                Number(item.grade) === Number(target.grade) &&
                item.subject_name === target.subject_name
              ).length
            : 0,
          seat_charts: state.seatCharts.filter((item) =>
            item.exam_date === target.exam_date &&
            Number(item.period) === Number(target.period) &&
            Number(item.grade) === Number(target.grade) &&
            item.subject_name === target.subject_name
          ).length,
          remove_subject_enrollments: removeEnrollments,
        };
        if (payload.preview_only) return { deleted: false, preview: true, impact };
        state.timetable = state.timetable.filter((item) => String(item.id) !== String(target.id));
        if (removeEnrollments) {
          state.enrollments = state.enrollments.filter((item) => !(
            Number(item.grade) === Number(target.grade) &&
            item.subject_name === target.subject_name
          ));
        }
        state.seatCharts = state.seatCharts.filter((item) => !(
          item.exam_date === target.exam_date &&
          Number(item.period) === Number(target.period) &&
          Number(item.grade) === Number(target.grade) &&
          item.subject_name === target.subject_name
        ));
        return { ...this.adminBootstrap(), deleted: true, impact };
      }

      case "save_enrollments": {
        const targetGrade = Number(payload.grade);
        const classByNumber = new Map(
          state.classes
            .filter((item) => Number(item.grade) === targetGrade && item.active !== false)
            .map((item) => [Number(item.class_num), item]),
        );
        const studentByClassNumber = new Map(
          state.students
            .filter((item) => item.active !== false)
            .map((item) => [`${item.class_id}|${Number(item.number)}`, item]),
        );
        const electiveSubjects = new Map();
        state.timetable
          .filter(
            (item) =>
              Number(item.grade) === targetGrade &&
              item.subject_type === "elective",
          )
          .forEach((item) => {
            const classes = electiveSubjects.get(item.subject_name) || new Set();
            parseClassIds(item.class_ids).forEach((id) => classes.add(String(id)));
            electiveSubjects.set(item.subject_name, classes);
          });
        const errors = [];
        const valid = [];
        const seen = new Set();
        (payload.enrollments || []).forEach((item, index) => {
          const rowNumber = Number(item.row_number) || index + 2;
          const classInfo = classByNumber.get(Number(item.class_num));
          if (!classInfo) {
            errors.push({ row_number: rowNumber, code: "CLASS_NOT_FOUND" });
            return;
          }
          const student = studentByClassNumber.get(
            `${classInfo.id}|${Number(item.number)}`,
          );
          if (!student) {
            errors.push({ row_number: rowNumber, code: "STUDENT_NOT_FOUND" });
            return;
          }
          if (String(student.name) !== String(item.name || "").trim()) {
            errors.push({ row_number: rowNumber, code: "STUDENT_NAME_MISMATCH" });
            return;
          }
          const applicableClasses = electiveSubjects.get(item.subject_name);
          if (!applicableClasses) {
            errors.push({ row_number: rowNumber, code: "SUBJECT_NOT_FOUND" });
            return;
          }
          if (!applicableClasses.has(String(classInfo.id))) {
            errors.push({ row_number: rowNumber, code: "SUBJECT_CLASS_MISMATCH" });
            return;
          }
          const duplicateKey = `${student.id}|${item.subject_name}`;
          if (seen.has(duplicateKey)) {
            errors.push({ row_number: rowNumber, code: "DUPLICATE_ENROLLMENT" });
            return;
          }
          seen.add(duplicateKey);
          const existing = state.enrollments.find(
            (row) =>
              String(row.student_id) === String(student.id) &&
              row.subject_name === item.subject_name,
          );
          valid.push({
            id: existing?.id || makeId("enrollment"),
            grade: targetGrade,
            class_id: classInfo.id,
            student_id: student.id,
            subject_name: item.subject_name,
            room_name: item.room_name || "",
          });
        });
        if (errors.length) {
          throw new ApiError("선택과목 명단에 저장할 수 없는 행이 있습니다.", {
            code: "ENROLLMENT_VALIDATION_FAILED",
            details: { errors, unmatched_count: errors.length },
          });
        }
        state.enrollments = state.enrollments
          .filter((item) => Number(item.grade) !== targetGrade)
          .concat(valid);
        return {
          ...this.adminBootstrap(),
          enrollment_result: {
            matched_count: valid.length,
            unmatched_count: errors.length,
            errors,
          },
        };
      }

      case "save_seat_chart": {
        validateDemoSeatChart(payload.chart);
        const item = { ...payload.chart, id: payload.chart.id || makeId("seat") };
        const key = seatChartKey(item);
        const existing = state.seatCharts.find((chart) => seatChartKey(chart) === key);
        item.id = payload.chart.id || existing?.id || item.id;
        state.seatCharts = state.seatCharts.filter(
          (chart) => chart.id !== item.id && seatChartKey(chart) !== key,
        );
        state.seatCharts.unshift(item);
        return {
          chart: item,
          seat_charts: state.seatCharts,
          created_count: existing ? 0 : 1,
          updated_count: existing ? 1 : 0,
        };
      }

      case "save_seat_charts_batch": {
        let createdCount = 0;
        let updatedCount = 0;
        payload.charts.forEach((chart) => {
          validateDemoSeatChart(chart);
          const key = seatChartKey(chart);
          const existing = state.seatCharts.find((item) => seatChartKey(item) === key);
          if (existing) updatedCount += 1;
          else createdCount += 1;
          const next = { ...chart, id: chart.id || existing?.id || makeId("seat") };
          state.seatCharts = state.seatCharts.filter((item) => seatChartKey(item) !== key);
          state.seatCharts.unshift(next);
        });
        return {
          seat_charts: state.seatCharts,
          created_count: createdCount,
          updated_count: updatedCount,
        };
      }

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
      classes: sortClasses(this.state.classes.filter((item) => item.active !== false)),
      exam_dates: structuredClone(this.state.examDates.filter((item) => item.active !== false)),
    };
  }

  adminBootstrap() {
    return {
      ...this.teacherBootstrap(),
      students: sortStudents(this.state.students.filter((item) => item.active !== false)),
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
      client_id: getClientId(),
      request_id: payload.request_id || crypto.randomUUID?.() || makeId("request"),
    };
    const presenceRequest = action === "presence_ping";
    const delays = presenceRequest ? [0] : [0, 300, 900];
    const timeoutMs = presenceRequest ? PRESENCE_TIMEOUT_MS : REQUEST_TIMEOUT_MS;
    let lastError;

    for (const delay of delays) {
      if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), timeoutMs);
      try {
        const response = await fetch(this.apiUrl, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(body),
          redirect: "follow",
          signal: controller.signal,
        });
        if (!response.ok) {
          throw new ApiError("학교 서버가 요청을 처리하지 못했습니다.", {
            code: response.status >= 500 ? "TEMPORARY_ERROR" : "HTTP_ERROR",
            retryable: response.status >= 500,
          });
        }
        const result = await response.json();
        if (!result.ok) {
          const code = result.error?.code || "REQUEST_FAILED";
          throw new ApiError(result.error?.message || "요청을 처리하지 못했습니다.", {
            code,
            retryable: Boolean(result.error?.retryable) || TRANSIENT_CODES.has(code),
            details: result.error?.details || null,
          });
        }
        return result.data;
      } catch (error) {
        if (error?.name === "AbortError") {
          lastError = new ApiError("요청 시간이 초과되었습니다. 잠시 후 다시 시도하세요.", {
            code: "TIMEOUT",
            retryable: true,
          });
        } else if (error instanceof ApiError) {
          lastError = error;
        } else {
          lastError = new ApiError("학교 서버에 연결하지 못했습니다.", {
            code: "NETWORK_ERROR",
            retryable: true,
          });
        }
      } finally {
        clearTimeout(timeout);
      }
      if (!lastError?.retryable) break;
    }
    throw lastError || new ApiError("학교 서버에 연결하지 못했습니다.", {
      code: "NETWORK_ERROR",
      retryable: true,
    });
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
  try {
    return await repository.request(action, payload);
  } catch (error) {
    if (SESSION_CODES.has(error?.code)) {
      if (error.code === "ADMIN_SESSION_EXPIRED") {
        sessionStorage.removeItem(ADMIN_SESSION_KEY);
      } else {
        clearSessions();
      }
      window.dispatchEvent(new CustomEvent("exam-session-expired", {
        detail: { code: error.code },
      }));
    }
    throw error;
  }
}
