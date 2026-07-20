export const ABSENCE_REASONS = ["질병", "인정", "미인정", "기타"];

export function compactText(value, maxLength = 100) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

export function makeId(prefix = "id") {
  if (globalThis.crypto?.randomUUID) return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

export function classLabel(item) {
  return item ? `${Number(item.grade)}학년 ${Number(item.class_num)}반` : "";
}

export function sortClasses(items = []) {
  return [...items].sort(
    (left, right) =>
      Number(left.grade) - Number(right.grade) ||
      Number(left.class_num) - Number(right.class_num),
  );
}

export function sortStudents(items = []) {
  return [...items].sort(
    (left, right) =>
      Number(left.number) - Number(right.number) ||
      compactText(left.name).localeCompare(compactText(right.name), "ko"),
  );
}

function headerIndex(header, candidates) {
  return header.findIndex((value) =>
    candidates.includes(compactText(value, 40).replace(/\s/g, "")),
  );
}

export function normalizeStudentRows(rows = []) {
  if (!Array.isArray(rows) || !rows.length) return [];
  const matrix = rows.map((row) => (Array.isArray(row) ? row : Object.values(row)));
  const header = matrix[0].map((value) => compactText(value, 40).replace(/\s/g, ""));
  const gradeIndex = headerIndex(header, ["학년"]);
  const classIndex = headerIndex(header, ["반", "학급"]);
  const numberIndex = headerIndex(header, ["번호", "출석번호"]);
  const nameIndex = headerIndex(header, ["이름", "성명", "학생명"]);
  const hasHeader = [gradeIndex, classIndex, numberIndex, nameIndex].every((index) => index >= 0);
  const indexes = hasHeader
    ? { gradeIndex, classIndex, numberIndex, nameIndex }
    : { gradeIndex: 0, classIndex: 1, numberIndex: 2, nameIndex: 3 };
  const seen = new Set();
  const result = [];

  matrix.slice(hasHeader ? 1 : 0).forEach((row) => {
    const grade = Number(row[indexes.gradeIndex]);
    const classNum = Number(row[indexes.classIndex]);
    const number = Number(row[indexes.numberIndex]);
    const name = compactText(row[indexes.nameIndex], 50);
    if (![grade, classNum, number].every(Number.isInteger) || !name) return;
    if (grade < 1 || grade > 9 || classNum < 1 || number < 1) return;
    const key = `${grade}|${classNum}|${number}`;
    if (seen.has(key)) return;
    seen.add(key);
    result.push({ grade, class_num: classNum, number, name });
  });

  return result.sort(
    (left, right) =>
      left.grade - right.grade ||
      left.class_num - right.class_num ||
      left.number - right.number,
  );
}

export function normalizeEnrollmentRows(rows = []) {
  if (!Array.isArray(rows) || !rows.length) return [];
  const matrix = rows.map((row) => (Array.isArray(row) ? row : Object.values(row)));
  const header = matrix[0].map((value) => compactText(value, 40).replace(/\s/g, ""));
  const subjectIndex = headerIndex(header, ["과목명", "과목"]);
  const classIndex = headerIndex(header, ["반", "학급"]);
  const numberIndex = headerIndex(header, ["번호", "출석번호"]);
  const nameIndex = headerIndex(header, ["이름", "성명", "학생명"]);
  const roomIndex = headerIndex(header, ["호실", "고사실", "교실"]);
  const hasHeader = [subjectIndex, classIndex, numberIndex, nameIndex].every(
    (index) => index >= 0,
  );
  const indexes = hasHeader
    ? { subjectIndex, classIndex, numberIndex, nameIndex, roomIndex }
    : { subjectIndex: 0, classIndex: 1, numberIndex: 2, nameIndex: 3, roomIndex: 4 };
  const seen = new Set();
  const result = [];

  matrix.slice(hasHeader ? 1 : 0).forEach((row) => {
    const subject_name = compactText(row[indexes.subjectIndex], 80);
    const class_num = Number(row[indexes.classIndex]);
    const number = Number(row[indexes.numberIndex]);
    const name = compactText(row[indexes.nameIndex], 50);
    const room_name =
      indexes.roomIndex >= 0 ? compactText(row[indexes.roomIndex], 80) : "";
    if (!subject_name || !Number.isInteger(class_num) || !Number.isInteger(number) || !name) {
      return;
    }
    const key = `${subject_name}|${class_num}|${number}`;
    if (seen.has(key)) return;
    seen.add(key);
    result.push({ subject_name, class_num, number, name, room_name });
  });

  return result.sort(
    (left, right) =>
      left.subject_name.localeCompare(right.subject_name, "ko") ||
      left.class_num - right.class_num ||
      left.number - right.number,
  );
}

export function normalizeManualSeatRows(rows = []) {
  if (!Array.isArray(rows) || !rows.length) return [];
  const matrix = rows.map((row) => (Array.isArray(row) ? row : Object.values(row)));
  const header = matrix[0].map((value) => compactText(value, 40).replace(/\s/g, ""));
  const classIndex = headerIndex(header, ["반", "학급"]);
  const numberIndex = headerIndex(header, ["번호", "출석번호"]);
  const nameIndex = headerIndex(header, ["이름", "성명", "학생명"]);
  const hasHeader = [classIndex, numberIndex, nameIndex].every((index) => index >= 0);
  const indexes = hasHeader
    ? { classIndex, numberIndex, nameIndex }
    : { classIndex: 0, numberIndex: 1, nameIndex: 2 };
  const seen = new Set();
  const result = [];
  matrix.slice(hasHeader ? 1 : 0).forEach((row) => {
    const class_num = Number(row[indexes.classIndex]);
    const number = Number(row[indexes.numberIndex]);
    const name = compactText(row[indexes.nameIndex], 50);
    if (!Number.isInteger(class_num) || !Number.isInteger(number) || class_num < 1 || number < 1 || !name) return;
    const key = `${class_num}|${number}`;
    if (seen.has(key)) return;
    seen.add(key);
    result.push({ class_num, number, name });
  });
  return result.sort((left, right) => left.class_num - right.class_num || left.number - right.number);
}

export function absenceKey({ exam_date, period, student_id }) {
  return `${exam_date}|${Number(period)}|${student_id}`;
}

export function completionKey({ exam_date, period, class_id }) {
  return `${exam_date}|${Number(period)}|${class_id}`;
}

export function buildSeatSlots(rows, cols, startSide = "window", disabledSeats = []) {
  const safeRows = Math.min(12, Math.max(1, Number(rows) || 1));
  const safeCols = Math.min(12, Math.max(1, Number(cols) || 1));
  const disabled = new Set(disabledSeats);
  const columns = Array.from({ length: safeCols }, (_, index) => index);
  if (startSide === "aisle") columns.reverse();
  const slots = [];

  for (let row = 0; row < safeRows; row += 1) {
    columns.forEach((col) => {
      const key = `${row}-${col}`;
      if (!disabled.has(key)) slots.push({ key, row, col });
    });
  }
  return slots;
}

export function generateSeatAssignment({
  students = [],
  rows = 6,
  cols = 5,
  startSide = "window",
  disabledSeats = [],
  mode = "separate",
  selectedIds = [],
  absentIds = [],
}) {
  const selected = new Set(selectedIds);
  const absent = new Set(absentIds);
  const sorted = sortStudents(students);
  const queue =
    mode === "own"
      ? [
          ...sorted.filter((student) => selected.has(student.id)),
          ...sorted.filter((student) => !selected.has(student.id)),
        ]
      : sorted.filter((student) => selected.has(student.id));
  const slots = buildSeatSlots(rows, cols, startSide, disabledSeats);
  if (queue.length > slots.length) {
    return {
      ok: false,
      error: `응시자 ${queue.length}명보다 사용 가능한 좌석 ${slots.length}석이 적습니다.`,
      assignments: [],
      slots,
    };
  }

  const assignments = slots.map((slot, index) => {
    const student = queue[index] || null;
    return {
      ...slot,
      student,
      selected: Boolean(student && selected.has(student.id)),
      absent: Boolean(student && absent.has(student.id)),
    };
  });
  return { ok: true, error: "", assignments, slots };
}

function appliesToClass(item, classId) {
  const raw = item.class_ids;
  if (!raw) return true;
  const values = Array.isArray(raw)
    ? raw
    : String(raw)
        .split(/[|,]/)
        .map((value) => value.trim())
        .filter(Boolean);
  return values.includes(String(classId));
}

export function buildPersonalTimetable({
  student,
  classInfo,
  timetable = [],
  enrollments = [],
}) {
  if (!student || !classInfo) return [];
  const enrollmentMap = new Map(
    enrollments
      .filter((item) => String(item.student_id) === String(student.id))
      .map((item) => [compactText(item.subject_name, 80), item]),
  );
  const relevant = timetable
    .filter(
      (item) =>
        Number(item.grade) === Number(classInfo.grade) &&
        appliesToClass(item, classInfo.id),
    )
    .sort(
      (left, right) =>
        String(left.exam_date).localeCompare(String(right.exam_date)) ||
        Number(left.period) - Number(right.period),
    );
  const grouped = new Map();
  relevant.forEach((item) => {
    const key = `${item.exam_date}|${Number(item.period)}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(item);
  });

  return [...grouped.entries()].map(([key, items]) => {
    const [exam_date, period] = key.split("|");
    const common = items.find((item) => item.subject_type === "common");
    const elective = items.find((item) =>
      enrollmentMap.has(compactText(item.subject_name, 80)),
    );
    const chosen = elective || common || null;
    const enrollment = chosen
      ? enrollmentMap.get(compactText(chosen.subject_name, 80))
      : null;
    return {
      exam_date,
      period: Number(period),
      time:
        chosen?.start_time || chosen?.end_time
          ? `${chosen?.start_time || ""}${chosen?.end_time ? `~${chosen.end_time}` : ""}`
          : "",
      subject_name: chosen?.subject_name || "자습",
      room_name:
        enrollment?.room_name ||
        chosen?.room_name ||
        (chosen?.subject_type === "common" ? classLabel(classInfo) : ""),
    };
  });
}

export function filterAbsences(absences = [], filters = {}) {
  const { startDate, endDate, examDate, period, grade, classId } = filters;
  return absences.filter((item) => {
    if (examDate && item.exam_date !== examDate) return false;
    if (startDate && item.exam_date < startDate) return false;
    if (endDate && item.exam_date > endDate) return false;
    if (period && Number(item.period) !== Number(period)) return false;
    if (grade && Number(item.grade) !== Number(grade)) return false;
    if (classId && String(item.class_id) !== String(classId)) return false;
    return true;
  });
}

export function toCsvSafeFileName(value) {
  return compactText(value, 80).replace(/[\\/:*?"<>|]/g, "_") || "download";
}

export function formatKoreanDate(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(value || "");
  return `${Number(match[1])}년 ${Number(match[2])}월 ${Number(match[3])}일`;
}
