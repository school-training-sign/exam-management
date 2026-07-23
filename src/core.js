export const ABSENCE_REASONS = ["인정", "미인정", "질병", "기타"];

export function compactText(value, maxLength = 100) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

export function normalizeLoginName(value, maxLength = 50) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLength);
}

export function isSixDigitPin(value) {
  return /^\d{6}$/.test(String(value ?? ""));
}

export function makeId(prefix = "id") {
  if (globalThis.crypto?.randomUUID) return `${prefix}_${crypto.randomUUID()}`;
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

export function formatPresenceLabel(status = "connecting", onlineCount = 0) {
  if (status === "connected") {
    const count = Math.max(0, Math.floor(Number(onlineCount) || 0));
    return `연결됨, ${count}명 접속`;
  }
  if (status === "disconnected") return "연결 끊김";
  return "연결 확인 중";
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

  matrix.slice(hasHeader ? 1 : 0).forEach((row, index) => {
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
    result.push({
      subject_name,
      class_num,
      number,
      name,
      room_name,
      row_number: index + (hasHeader ? 2 : 1),
    });
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

function analyzeFixedRows(
  rows,
  columns,
  rowCode,
  duplicateCode,
  build,
  keyFor,
  sort,
  { skipBlankRows = true } = {},
) {
  if (!Array.isArray(rows) || !rows.length) return { rows: [], errors: [] };
  const matrix = rows.map((row) => (Array.isArray(row) ? row : Object.values(row)));
  const header = matrix[0].map((value) => compactText(value, 40).replace(/\s/g, ""));
  const indexes = Object.fromEntries(
    columns.map(({ key, aliases }) => [key, headerIndex(header, aliases)]),
  );
  if (Object.values(indexes).some((index) => index < 0)) {
    return { rows: [], errors: [{ row_number: 1, code: "INVALID_HEADERS" }] };
  }

  const seen = new Set();
  const valid = [];
  const errors = [];
  matrix.slice(1).forEach((row, index) => {
    if (skipBlankRows && row.every((value) => compactText(value, 10) === "")) return;
    const rowNumber = index + 2;
    const item = build(row, indexes, rowNumber);
    if (!item) {
      errors.push({ row_number: rowNumber, code: rowCode });
      return;
    }
    const key = keyFor(item);
    if (seen.has(key)) {
      errors.push({ row_number: rowNumber, code: duplicateCode });
      return;
    }
    seen.add(key);
    valid.push(item);
  });
  return { rows: valid.sort(sort), errors };
}

export function analyzeStudentRows(rows = []) {
  return analyzeFixedRows(
    rows,
    [
      { key: "grade", aliases: ["학년"] },
      { key: "classNum", aliases: ["반", "학급"] },
      { key: "number", aliases: ["번호", "출석번호"] },
      { key: "name", aliases: ["이름", "성명", "학생명"] },
    ],
    "INVALID_STUDENT_ROW",
    "DUPLICATE_STUDENT_NUMBER",
    (row, indexes) => {
      const grade = Number(row[indexes.grade]);
      const class_num = Number(row[indexes.classNum]);
      const number = Number(row[indexes.number]);
      const name = compactText(row[indexes.name], 50);
      if (
        !Number.isInteger(grade) ||
        grade < 1 ||
        grade > 9 ||
        !Number.isInteger(class_num) ||
        class_num < 1 ||
        !Number.isInteger(number) ||
        number < 1 ||
        !name
      ) return null;
      return { grade, class_num, number, name };
    },
    (item) => `${item.grade}|${item.class_num}|${item.number}`,
    (left, right) =>
      left.grade - right.grade ||
      left.class_num - right.class_num ||
      left.number - right.number,
  );
}

export function analyzeEnrollmentRows(rows = []) {
  return analyzeFixedRows(
    rows,
    [
      { key: "subject", aliases: ["과목명", "과목"] },
      { key: "classNum", aliases: ["반", "학급"] },
      { key: "number", aliases: ["번호", "출석번호"] },
      { key: "name", aliases: ["이름", "성명", "학생명"] },
      { key: "room", aliases: ["호실", "고사실", "교실"] },
    ],
    "INVALID_ENROLLMENT_ROW",
    "DUPLICATE_ENROLLMENT",
    (row, indexes, rowNumber) => {
      const subject_name = compactText(row[indexes.subject], 80);
      const class_num = Number(row[indexes.classNum]);
      const number = Number(row[indexes.number]);
      const name = compactText(row[indexes.name], 50);
      const room_name = compactText(row[indexes.room], 80);
      if (
        !subject_name ||
        !Number.isInteger(class_num) ||
        class_num < 1 ||
        !Number.isInteger(number) ||
        number < 1 ||
        !name
      ) return null;
      return { subject_name, class_num, number, name, room_name, row_number: rowNumber };
    },
    (item) => `${item.subject_name}|${item.class_num}|${item.number}`,
    (left, right) =>
      left.subject_name.localeCompare(right.subject_name, "ko") ||
      left.class_num - right.class_num ||
      left.number - right.number,
  );
}

export function analyzeManualSeatRows(rows = []) {
  return analyzeFixedRows(
    rows,
    [
      { key: "classNum", aliases: ["반", "학급"] },
      { key: "number", aliases: ["번호", "출석번호"] },
      { key: "name", aliases: ["이름", "성명", "학생명"] },
    ],
    "INVALID_SEAT_ROW",
    "DUPLICATE_SEAT_STUDENT",
    (row, indexes, rowNumber) => {
      const class_num = Number(row[indexes.classNum]);
      const number = Number(row[indexes.number]);
      const name = compactText(row[indexes.name], 50);
      if (
        !Number.isInteger(class_num) ||
        class_num < 1 ||
        !Number.isInteger(number) ||
        number < 1 ||
        !name
      ) return null;
      return { class_num, number, name, row_number: rowNumber };
    },
    (item) => `${item.class_num}|${item.number}`,
    (left, right) => left.class_num - right.class_num || left.number - right.number,
  );
}

export function analyzeSubjectCatalogRows(rows = []) {
  return analyzeFixedRows(
    rows,
    [{ key: "subject", aliases: ["과목명", "과목"] }],
    "INVALID_SUBJECT_ROW",
    "DUPLICATE_SUBJECT_NAME",
    (row, indexes, rowNumber) => {
      const subject_name = compactText(row[indexes.subject], 80);
      if (!subject_name) return null;
      return { subject_name, row_number: rowNumber };
    },
    (item) => item.subject_name,
    (left, right) => left.subject_name.localeCompare(right.subject_name, "ko"),
    { skipBlankRows: false },
  );
}

export function absenceKey({ exam_date, period, student_id }) {
  return `${exam_date}|${Number(period)}|${student_id}`;
}

export function completionKey({ exam_date, period, class_id }) {
  return `${exam_date}|${Number(period)}|${class_id}`;
}

export function buildSeatSlots(
  rows,
  cols,
  startSide = "window",
  disabledSeats = [],
  seatOrder = "row",
) {
  const safeRows = Math.min(12, Math.max(1, Number(rows) || 1));
  const safeCols = Math.min(12, Math.max(1, Number(cols) || 1));
  const disabled = new Set(disabledSeats.map(String));
  const columns = Array.from({ length: safeCols }, (_, index) => index);
  const rowIndexes = Array.from({ length: safeRows }, (_, index) => index);
  if (startSide === "aisle") columns.reverse();
  const slots = [];

  const append = (row, col) => {
    const key = `${row}-${col}`;
    if (!disabled.has(key)) slots.push({ key, row, col });
  };
  if (seatOrder === "column") {
    columns.forEach((col) => {
      rowIndexes.forEach((row) => append(row, col));
    });
  } else {
    rowIndexes.forEach((row) => {
      columns.forEach((col) => append(row, col));
    });
  }
  return slots;
}

export function generateSeatAssignment({
  students = [],
  rows = 6,
  cols = 5,
  startSide = "window",
  seatOrder = "row",
  disabledSeats = [],
  mode = "separate",
  selectedIds = [],
  absentIds = [],
}) {
  const selected = new Set(selectedIds.map(String));
  const absent = new Set(absentIds.map(String));
  const sorted = sortStudents(students);
  const queue =
    mode === "own"
      ? [
          ...sorted.filter((student) => selected.has(String(student.id)) || absent.has(String(student.id))),
          ...sorted.filter((student) => !selected.has(String(student.id)) && !absent.has(String(student.id))),
        ]
      : sorted.filter((student) => selected.has(String(student.id)));
  const slots = buildSeatSlots(rows, cols, startSide, disabledSeats, seatOrder);
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
      sequence: index + 1,
      student,
      selected: Boolean(student && selected.has(String(student.id))),
      absent: Boolean(student && absent.has(String(student.id))),
    };
  });
  return { ok: true, error: "", assignments, slots };
}

export function resolveStudentRoom({
  studentId,
  subjectName,
  timetableItem,
  enrollments = [],
  classInfo,
}) {
  const enrollment = enrollments.find(
    (item) =>
      String(item.student_id) === String(studentId) &&
      compactText(item.subject_name, 80) === compactText(subjectName, 80),
  );
  return (
    compactText(enrollment?.room_name, 80) ||
    compactText(timetableItem?.room_name, 80) ||
    (timetableItem?.subject_type === "common" ? classLabel(classInfo) : "")
  );
}

export function seatChartKey(chart = {}) {
  return [
    chart.mode || "",
    chart.exam_date || "",
    Number(chart.period) || "",
    Number(chart.grade) || "",
    chart.class_id || "",
    compactText(chart.subject_name, 80),
    compactText(chart.room_name, 80),
  ].join("|");
}

export function summarizeAbsences(rows = []) {
  const result = Object.fromEntries(ABSENCE_REASONS.map((reason) => [reason, 0]));
  rows.forEach((item) => {
    const reason = ABSENCE_REASONS.includes(item.reason) ? item.reason : "기타";
    result[reason] += 1;
  });
  return result;
}

export function aggregateAbsenceReport(rows = []) {
  const increment = (map, key, seed) => {
    if (!map.has(key)) map.set(key, { ...seed, count: 0 });
    map.get(key).count += 1;
  };
  const reasons = new Map();
  const classes = new Map();
  const dates = new Map();
  const students = new Map();

  rows.forEach((item) => {
    increment(reasons, item.reason || "기타", { reason: item.reason || "기타" });
    increment(classes, String(item.class_id || item.class_label || ""), {
      class_id: item.class_id || "",
      class_label: item.class_label || "",
    });
    increment(dates, String(item.exam_date || ""), { exam_date: item.exam_date || "" });
    increment(students, String(item.student_id || ""), {
      student_id: item.student_id || "",
      class_label: item.class_label || "",
      student_number: item.student_number ?? "",
      student_name: item.student_name || "",
    });
  });

  return {
    reason_summary: [...reasons.values()].sort((a, b) => b.count - a.count || a.reason.localeCompare(b.reason, "ko")),
    class_summary: [...classes.values()].sort((a, b) => a.class_label.localeCompare(b.class_label, "ko")),
    date_summary: [...dates.values()].sort((a, b) => a.exam_date.localeCompare(b.exam_date)),
    student_summary: [...students.values()].sort(
      (a, b) =>
        a.class_label.localeCompare(b.class_label, "ko") ||
        Number(a.student_number) - Number(b.student_number),
    ),
  };
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

export function buildExamNoticeTitle(notice = {}) {
  const year = compactText(notice?.year, 10);
  const semester = compactText(notice?.semester, 10);
  const examName = compactText(notice?.exam_name, 80);
  if (!year || !semester || !examName) return "";
  return `${year}학년도 ${semester}학기 ${examName} 시간표`;
}

function parseNoticeClassIds(value) {
  const values = Array.isArray(value)
    ? value
    : String(value || "").split(/[|,]/);
  return [...new Set(values.map((item) => compactText(item, 80)).filter(Boolean))];
}

export function buildExamNoticeSchedule({
  examDates = [],
  timetable = [],
  classes = [],
} = {}) {
  const activeClasses = sortClasses(
    classes.filter(
      (item) =>
        item?.active !== false &&
        Number.isInteger(Number(item?.grade)) &&
        Number(item.grade) > 0 &&
        Number.isInteger(Number(item?.class_num)) &&
        Number(item.class_num) > 0,
    ),
  );
  const classesByGrade = new Map();
  activeClasses.forEach((item) => {
    const grade = Number(item.grade);
    if (!classesByGrade.has(grade)) classesByGrade.set(grade, []);
    classesByGrade.get(grade).push(item);
  });
  const grades = [...classesByGrade.keys()].sort((left, right) => left - right);

  const activeExamDates = new Set(
    examDates
      .filter((item) => item?.active !== false)
      .map((item) => compactText(item?.exam_date, 20))
      .filter(Boolean),
  );
  const grouped = new Map();

  timetable.forEach((item) => {
    if (item?.active === false) return;
    const examDate = compactText(item?.exam_date, 20);
    const grade = Number(item?.grade);
    const period = Number(item?.period);
    const subjectName = compactText(item?.subject_name, 80);
    const gradeClasses = classesByGrade.get(grade) || [];
    if (
      !activeExamDates.has(examDate) ||
      !Number.isInteger(period) ||
      period < 1 ||
      !subjectName ||
      !gradeClasses.length
    ) return;

    const classIds = parseNoticeClassIds(item?.class_ids);
    const hasExplicitScope = classIds.length > 0;
    const selectedIds = new Set(classIds);
    const selectedClasses = hasExplicitScope
      ? gradeClasses.filter((classItem) => selectedIds.has(String(classItem.id)))
      : gradeClasses;
    if (hasExplicitScope && !selectedClasses.length) return;

    const coversWholeGrade =
      !hasExplicitScope || selectedClasses.length === gradeClasses.length;
    const scopeLabel = coversWholeGrade
      ? ""
      : `${[
          ...new Set(selectedClasses.map((classItem) => Number(classItem.class_num))),
        ]
          .sort((left, right) => left - right)
          .join("·")}반`;
    const startTime = compactText(item?.start_time, 20);
    const endTime = compactText(item?.end_time, 20);
    const time = [startTime, endTime].filter(Boolean).join("~");
    const key = `${examDate}|${period}|${startTime}|${endTime}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        exam_date: examDate,
        period,
        time,
        subjects_by_grade: Object.fromEntries(grades.map((value) => [value, []])),
      });
    }
    grouped.get(key).subjects_by_grade[grade].push({
      subject_name: subjectName,
      subject_type: compactText(item?.subject_type, 20),
      scope_label: scopeLabel,
    });
  });

  const typeRank = (value) => {
    if (value === "common") return 0;
    if (value === "elective") return 1;
    return 2;
  };
  const rows = [...grouped.values()]
    .map((row) => {
      grades.forEach((grade) => {
        row.subjects_by_grade[grade].sort(
          (left, right) =>
            typeRank(left.subject_type) - typeRank(right.subject_type) ||
            left.subject_name.localeCompare(right.subject_name, "ko") ||
            left.scope_label.localeCompare(right.scope_label, "ko"),
        );
      });
      return row;
    })
    .sort(
      (left, right) =>
        left.exam_date.localeCompare(right.exam_date) ||
        left.period - right.period ||
        left.time.localeCompare(right.time),
    );

  return { grades, rows };
}

export const EXAM_PACKAGING_ANSWER_SHEET_TYPES = ["", "card", "a4"];

function normalizePackagingRow(row = {}, index = 0) {
  return {
    id: compactText(row.id, 80),
    kind: compactText(row.kind, 20) || "slot",
    period_label: compactText(row.period_label, 80),
    start_time: compactText(row.start_time, 10),
    end_time: compactText(row.end_time, 10),
    note: compactText(row.note, 160),
    sort_order: Number.isFinite(Number(row.sort_order))
      ? Number(row.sort_order)
      : index + 1,
  };
}

export function normalizeExamPackagingConfig(config = {}) {
  const source = config && typeof config === "object" && !Array.isArray(config)
    ? config
    : {};
  return {
    revision: Math.max(0, Math.floor(Number(source.revision) || 0)),
    input_deadline: compactText(source.input_deadline, 40),
    capacity: Number(source.capacity ?? 3),
    packaging_dates: Array.isArray(source.packaging_dates)
      ? source.packaging_dates.map((value) => compactText(value, 20)).filter(Boolean)
      : [],
    rows: Array.isArray(source.rows)
      ? source.rows.map(normalizePackagingRow)
      : [],
    staff_assignments: Array.isArray(source.staff_assignments)
      ? source.staff_assignments.map((item = {}) => ({
          packaging_date: compactText(item.packaging_date, 20),
          slot_id: compactText(item.slot_id, 80),
          staff_name: compactText(item.staff_name, 100),
        }))
      : [],
  };
}

function validIsoDate(value) {
  return Boolean(parseIsoDate(value));
}

function validTime(value) {
  const match = /^(\d{2}):(\d{2})$/.exec(String(value || ""));
  return Boolean(match && Number(match[1]) < 24 && Number(match[2]) < 60);
}

function clockMinutes(value) {
  if (!validTime(value)) return NaN;
  const [hour, minute] = String(value).split(":").map(Number);
  return hour * 60 + minute;
}

function validIsoDateTime(value) {
  if (!value) return true;
  const match = /^(\d{4}-\d{2}-\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.\d{1,3})?)?(?:Z|([+-])(\d{2}):(\d{2}))?$/.exec(value);
  if (!match || !validIsoDate(match[1])) {
    return false;
  }
  if (
    Number(match[2]) > 23 ||
    Number(match[3]) > 59 ||
    Number(match[4] || 0) > 59 ||
    Number(match[6] || 0) > 23 ||
    Number(match[7] || 0) > 59
  ) return false;
  return Number.isFinite(Date.parse(value));
}

export function validateExamPackagingConfig(config = {}) {
  const normalized = normalizeExamPackagingConfig(config);
  const errors = [];
  const push = (code, details = {}) => errors.push({ code, ...details });

  if (!Number.isInteger(normalized.capacity) || normalized.capacity < 1 || normalized.capacity > 20) {
    push("INVALID_PACKAGING_CAPACITY", { field: "capacity" });
  }
  if (!validIsoDateTime(normalized.input_deadline)) {
    push("INVALID_PACKAGING_DEADLINE", { field: "input_deadline" });
  }

  const dateSet = new Set();
  normalized.packaging_dates.forEach((value, index) => {
    if (!validIsoDate(value)) {
      push("INVALID_PACKAGING_DATE", { field: "packaging_dates", index });
    } else if (dateSet.has(value)) {
      push("DUPLICATE_PACKAGING_DATE", { field: "packaging_dates", index, value });
    }
    dateSet.add(value);
  });

  const rowIds = new Set();
  const timeRanges = new Set();
  normalized.rows.forEach((row, index) => {
    if (!row.id) {
      push("PACKAGING_ROW_ID_REQUIRED", { field: "rows", index });
    } else if (rowIds.has(row.id)) {
      push("DUPLICATE_PACKAGING_ROW_ID", { field: "rows", index, id: row.id });
    }
    rowIds.add(row.id);
    if (!new Set(["slot", "break"]).has(row.kind)) {
      push("INVALID_PACKAGING_ROW_KIND", { field: "rows", index, id: row.id });
    }
    if (!row.period_label) {
      push("PACKAGING_ROW_LABEL_REQUIRED", { field: "rows", index, id: row.id });
    }
    if (!validTime(row.start_time) || !validTime(row.end_time) || row.start_time >= row.end_time) {
      push("INVALID_PACKAGING_ROW_TIME", { field: "rows", index, id: row.id });
    } else {
      if (row.kind === "slot" && clockMinutes(row.end_time) - clockMinutes(row.start_time) !== 25) {
        push("INVALID_PACKAGING_SLOT_DURATION", { field: "rows", index, id: row.id });
      }
      const timeKey = `${row.start_time}|${row.end_time}`;
      if (timeRanges.has(timeKey)) {
        push("DUPLICATE_PACKAGING_TIME", { field: "rows", index, id: row.id });
      }
      timeRanges.add(timeKey);
    }
  });

  const chronologicalRows = normalized.rows
    .filter((row) => validTime(row.start_time) && validTime(row.end_time) && row.start_time < row.end_time)
    .slice()
    .sort((left, right) => left.start_time.localeCompare(right.start_time) || left.end_time.localeCompare(right.end_time));
  for (let index = 1; index < chronologicalRows.length; index += 1) {
    if (chronologicalRows[index].start_time < chronologicalRows[index - 1].end_time) {
      push("OVERLAPPING_PACKAGING_TIME", {
        field: "rows",
        id: chronologicalRows[index].id,
      });
    }
  }

  const staffKeys = new Set();
  normalized.staff_assignments.forEach((item, index) => {
    const key = `${item.packaging_date}|${item.slot_id}`;
    if (!dateSet.has(item.packaging_date) || !rowIds.has(item.slot_id) || !item.staff_name) {
      push("INVALID_PACKAGING_STAFF_ASSIGNMENT", {
        field: "staff_assignments",
        index,
      });
    } else if (staffKeys.has(key)) {
      push("DUPLICATE_PACKAGING_STAFF_ASSIGNMENT", {
        field: "staff_assignments",
        index,
      });
    }
    staffKeys.add(key);
  });

  return { ok: errors.length === 0, errors, config: normalized };
}

export function sortExamPackagingItems(items = []) {
  return [...items].sort(
    (left, right) =>
      String(left.exam_date || "").localeCompare(String(right.exam_date || "")) ||
      Number(left.period || 0) - Number(right.period || 0) ||
      String(left.start_time || "").localeCompare(String(right.start_time || "")) ||
      Number(left.grade || 0) - Number(right.grade || 0) ||
      compactText(left.subject_name, 80).localeCompare(compactText(right.subject_name, 80), "ko") ||
      String(left.timetable_id || left.id || "").localeCompare(String(right.timetable_id || right.id || "")),
  );
}

export function summarizeExamPackaging(items = []) {
  const summary = {
    total: 0,
    assigned: 0,
    unassigned: 0,
    answer_selected: 0,
    answer_unselected: 0,
    card: 0,
    a4: 0,
    representative_missing: 0,
  };
  items.forEach((item) => {
    summary.total += 1;
    const assigned = Boolean(item.packaging_date && item.packaging_slot_id);
    if (assigned) summary.assigned += 1;
    else summary.unassigned += 1;
    if (item.answer_sheet_type === "card" || item.answer_sheet_type === "a4") {
      summary.answer_selected += 1;
      summary[item.answer_sheet_type] += 1;
    } else {
      summary.answer_unselected += 1;
    }
    if (!compactText(item.representative_teacher, 100)) summary.representative_missing += 1;
  });
  return summary;
}

export function groupExamPackagingSchedule(items = [], config = {}) {
  const normalized = normalizeExamPackagingConfig(config);
  const sortedRows = [...normalized.rows].sort(
    (left, right) =>
      Number(left.sort_order) - Number(right.sort_order) ||
      left.start_time.localeCompare(right.start_time) ||
      left.id.localeCompare(right.id),
  );
  const staffMap = new Map(
    normalized.staff_assignments.map((item) => [
      `${item.packaging_date}|${item.slot_id}`,
      item.staff_name,
    ]),
  );
  const itemMap = new Map();
  const unassigned = [];
  sortExamPackagingItems(items).forEach((item) => {
    if (!item.packaging_date || !item.packaging_slot_id) {
      unassigned.push(item);
      return;
    }
    const key = `${item.packaging_date}|${item.packaging_slot_id}`;
    if (!itemMap.has(key)) itemMap.set(key, []);
    itemMap.get(key).push(item);
  });
  const dates = [...normalized.packaging_dates].sort().map((packagingDate) => ({
    packaging_date: packagingDate,
    rows: sortedRows.map((row) => ({
      ...row,
      staff_name: staffMap.get(`${packagingDate}|${row.id}`) || "",
      items: row.kind === "slot"
        ? (itemMap.get(`${packagingDate}|${row.id}`) || [])
        : [],
    })),
  }));
  return { dates, unassigned };
}

export function analyzeExamPackagingConfigImpact(currentConfig = {}, nextConfig = {}, items = []) {
  const current = normalizeExamPackagingConfig(currentConfig);
  const next = normalizeExamPackagingConfig(nextConfig);
  const nextDates = new Set(next.packaging_dates);
  const nextSlots = new Set(next.rows.filter((row) => row.kind === "slot").map((row) => row.id));
  const removedDates = current.packaging_dates.filter((value) => !nextDates.has(value));
  const currentSlots = new Set(current.rows.filter((row) => row.kind === "slot").map((row) => row.id));
  const removedSlotIds = [...currentSlots].filter((value) => !nextSlots.has(value));
  const affected = items.filter((item) =>
    Boolean(item.packaging_date || item.packaging_slot_id) &&
    (!nextDates.has(String(item.packaging_date || "")) ||
      !nextSlots.has(String(item.packaging_slot_id || ""))),
  );
  return {
    affected_count: affected.length,
    affected_timetable_ids: affected.map((item) => String(item.timetable_id || item.id || "")),
    removed_dates: removedDates,
    removed_slot_ids: removedSlotIds,
  };
}

export function isExamPackagingDeadlinePassed(config = {}, now = Date.now()) {
  const deadline = normalizeExamPackagingConfig(config).input_deadline;
  if (!deadline || !validIsoDateTime(deadline)) return false;
  const nowValue = now instanceof Date ? now.getTime() : Number(now);
  return Number.isFinite(nowValue) && nowValue >= Date.parse(deadline);
}

export function sanitizeSpreadsheetCell(value) {
  if (typeof value !== "string") return value;
  return /^[\s]*[=+\-@]/.test(value) ? `'${value}` : value;
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

const KOREAN_WEEKDAYS = ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"];

function parseIsoDate(value) {
  const source = String(value || "");
  const match = source.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) return null;

  return { year, month, day, weekday: KOREAN_WEEKDAYS[date.getUTCDay()] };
}

export function formatKoreanWeekday(value) {
  return parseIsoDate(value)?.weekday || "";
}

export function formatKoreanDate(value) {
  const source = String(value || "");
  const date = parseIsoDate(source);
  if (!date) return source;
  return `${date.year}년 ${date.month}월 ${date.day}일 (${date.weekday})`;
}

export const IDLE_LOGOUT_MS = 5 * 60 * 1000;

export function createIdleLogoutTimer(
  onTimeout,
  { setTimer = setTimeout, clearTimer = clearTimeout } = {},
) {
  let timer = null;
  return {
    reset() {
      if (timer != null) clearTimer(timer);
      timer = setTimer(onTimeout, IDLE_LOGOUT_MS);
    },
    dispose() {
      if (timer != null) clearTimer(timer);
      timer = null;
    },
  };
}
