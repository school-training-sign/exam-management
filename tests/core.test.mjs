import test from "node:test";
import assert from "node:assert/strict";
import {
  aggregateAbsenceReport,
  analyzeExamPackagingConfigImpact,
  analyzeEnrollmentRows,
  analyzeManualSeatRows,
  analyzeStudentRows,
  analyzeSubjectCatalogRows,
  absenceKey,
  buildExamNoticeSchedule,
  buildExamNoticeTitle,
  groupExamPackagingSchedule,
  buildPersonalTimetable,
  buildSeatSlots,
  createIdleLogoutTimer,
  filterAbsences,
  formatKoreanDate,
  formatKoreanWeekday,
  formatPresenceLabel,
  generateSeatAssignment,
  isSixDigitPin,
  isExamPackagingDeadlinePassed,
  normalizeExamPackagingConfig,
  normalizeLoginName,
  normalizeEnrollmentRows,
  normalizeManualSeatRows,
  normalizeStudentRows,
  resolveStudentRoom,
  sanitizeSpreadsheetCell,
  seatChartKey,
  sortExamPackagingItems,
  summarizeAbsences,
  summarizeExamPackaging,
  validateExamPackagingConfig,
} from "../src/core.js";

test("미사용 타이머는 5분 뒤 로그아웃하고 활동 때마다 다시 시작한다", () => {
  const scheduled = [];
  const cleared = [];
  let logoutCount = 0;
  const timer = createIdleLogoutTimer(
    () => { logoutCount += 1; },
    {
      setTimer(callback, delay) {
        const id = scheduled.length + 1;
        scheduled.push({ id, callback, delay });
        return id;
      },
      clearTimer(id) {
        cleared.push(id);
      },
    },
  );

  timer.reset();
  assert.equal(scheduled[0].delay, 300_000);
  timer.reset();
  assert.deepEqual(cleared, [1]);
  scheduled[1].callback();
  assert.equal(logoutCount, 1);
  timer.dispose();
  assert.deepEqual(cleared, [1, 2]);
});

test("접속 이름은 유니코드와 연속 공백을 정규화하고 PIN은 숫자 6자리만 허용한다", () => {
  assert.equal(normalizeLoginName("  테스트　교사  "), "테스트 교사");
  assert.equal(normalizeLoginName("ＫＩＭ  선생님"), "KIM 선생님");
  assert.equal(isSixDigitPin("123456"), true);
  assert.equal(isSixDigitPin("12345"), false);
  assert.equal(isSixDigitPin("12가456"), false);
});

test("교과목 엑셀은 과목명 헤더를 인식하고 이름을 정리해 정렬한다", () => {
  const longName = `심화${"가".repeat(90)}`;
  const result = analyzeSubjectCatalogRows([
    ["과목명"],
    ["  사회   문화  "],
    ["경제"],
    [longName],
  ]);

  assert.deepEqual(result.errors, []);
  assert.deepEqual(result.rows.map((item) => item.subject_name), [
    "경제",
    "사회 문화",
    longName.replace(/\s+/g, " ").trim().slice(0, 80),
  ]);
  assert.deepEqual(result.rows.map((item) => item.row_number), [3, 2, 4]);
});

test("교과목 엑셀은 잘못된 헤더와 빈 행, 중복 과목의 행 번호를 알려준다", () => {
  assert.deepEqual(analyzeSubjectCatalogRows([["교과"]]), {
    rows: [],
    errors: [{ row_number: 1, code: "INVALID_HEADERS" }],
  });

  const result = analyzeSubjectCatalogRows([
    ["과목"],
    ["국어"],
    ["   "],
    ["국어"],
  ]);
  assert.deepEqual(result.rows, [{ subject_name: "국어", row_number: 2 }]);
  assert.deepEqual(result.errors, [
    { row_number: 3, code: "INVALID_SUBJECT_ROW" },
    { row_number: 4, code: "DUPLICATE_SUBJECT_NAME" },
  ]);
});

test("가정통신문 제목은 학년도, 학기, 고사명이 모두 있을 때만 만든다", () => {
  assert.equal(
    buildExamNoticeTitle({ year: 2026, semester: 1, exam_name: "중간고사" }),
    "2026학년도 1학기 중간고사 시간표",
  );
  assert.equal(buildExamNoticeTitle({ year: 2026, semester: 1, exam_name: " " }), "");
});

test("가정통신문 시간표는 활성 고사일과 학년만 묶어 날짜, 교시, 시간순으로 정렬한다", () => {
  const result = buildExamNoticeSchedule({
    examDates: [
      { exam_date: "2026-07-21", active: true },
      { exam_date: "2026-07-19", active: false },
      { exam_date: "2026-07-20", active: true },
    ],
    classes: [
      { id: "c23", grade: 2, class_num: 3, active: true },
      { id: "c12", grade: 1, class_num: 2, active: true },
      { id: "c31", grade: 3, class_num: 1, active: false },
      { id: "c11", grade: 1, class_num: 1, active: true },
      { id: "c21", grade: 2, class_num: 1, active: true },
      { id: "c22", grade: 2, class_num: 2, active: true },
    ],
    timetable: [
      { exam_date: "2026-07-21", grade: 1, period: 2, start_time: "10:10", end_time: "11:00", subject_name: "수학", subject_type: "common", class_ids: "c11|c12" },
      { exam_date: "2026-07-20", grade: 1, period: 2, start_time: "10:00", end_time: "10:50", subject_name: "영어", subject_type: "common", class_ids: ["c11", "c12"] },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:30", end_time: "10:20", subject_name: "물리", subject_type: "common", class_ids: "c11|c12" },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "한국사", subject_type: "common", class_ids: "c11|c12" },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "사회문화", subject_type: "elective", class_ids: "c12" },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "경제", subject_type: "elective", class_ids: "c11" },
      { exam_date: "2026-07-20", grade: 2, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "세계사", subject_type: "elective", class_ids: "c21|c22" },
      { exam_date: "2026-07-19", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "제외 과목", subject_type: "common", class_ids: "c11|c12" },
      { exam_date: "2026-07-20", grade: 1, period: 3, start_time: "11:10", end_time: "12:00", subject_name: "비활성 과목", subject_type: "common", class_ids: "c11|c12", active: false },
    ],
  });

  assert.deepEqual(result.grades, [1, 2]);
  assert.deepEqual(
    result.rows.map((row) => [row.exam_date, row.period, row.time]),
    [
      ["2026-07-20", 1, "09:00~09:50"],
      ["2026-07-20", 1, "09:30~10:20"],
      ["2026-07-20", 2, "10:00~10:50"],
      ["2026-07-21", 2, "10:10~11:00"],
    ],
  );
});

test("가정통신문 시간표는 공통·선택 과목 순서와 전체·일부 학급 범위를 표시한다", () => {
  const result = buildExamNoticeSchedule({
    examDates: [{ exam_date: "2026-07-20", active: true }],
    classes: [
      { id: "c11", grade: 1, class_num: 1, active: true },
      { id: "c12", grade: 1, class_num: 2, active: true },
      { id: "c21", grade: 2, class_num: 1, active: true },
      { id: "c22", grade: 2, class_num: 2, active: true },
      { id: "c23", grade: 2, class_num: 3, active: true },
    ],
    timetable: [
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "사회문화", subject_type: "elective", class_ids: "c12" },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "한국사", subject_type: "common", class_ids: "c11|c12" },
      { exam_date: "2026-07-20", grade: 1, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "경제", subject_type: "elective", class_ids: "c11" },
      { exam_date: "2026-07-20", grade: 2, period: 1, start_time: "09:00", end_time: "09:50", subject_name: "세계사", subject_type: "elective", class_ids: "c21,c22" },
    ],
  });

  const row = result.rows[0];
  assert.deepEqual(row.subjects_by_grade[1], [
    { subject_name: "한국사", subject_type: "common", scope_label: "" },
    { subject_name: "경제", subject_type: "elective", scope_label: "1반" },
    { subject_name: "사회문화", subject_type: "elective", scope_label: "2반" },
  ]);
  assert.deepEqual(row.subjects_by_grade[2], [
    { subject_name: "세계사", subject_type: "elective", scope_label: "1·2반" },
  ]);
});

test("고사일을 한국어 요일과 함께 지역 시간대에 영향 없이 표시한다", () => {
  assert.equal(formatKoreanDate("2026-07-20"), "2026년 7월 20일 (월요일)");
  assert.equal(formatKoreanDate("2026-07-21"), "2026년 7월 21일 (화요일)");
  assert.equal(formatKoreanWeekday("2026-07-21"), "화요일");
  assert.equal(formatKoreanDate("잘못된 날짜"), "잘못된 날짜");
  assert.equal(formatKoreanWeekday("2026-02-30"), "");
});

test("상단 연결 상태는 현재 화면을 제외한 활성 접속자 수를 표시한다", () => {
  assert.equal(formatPresenceLabel("connected", 0), "연결됨, 0명 접속");
  assert.equal(formatPresenceLabel("connected", 2.9), "연결됨, 2명 접속");
  assert.equal(formatPresenceLabel("connecting", 0), "연결 확인 중");
  assert.equal(formatPresenceLabel("disconnected", 3), "연결 끊김");
});

test("학생 명단 헤더를 인식하고 중복 번호를 제거한다", () => {
  const rows = normalizeStudentRows([
    ["학년", "반", "번호", "이름"],
    [1, 2, 1, "김하늘"],
    [1, 2, 1, "중복학생"],
    [1, 2, 2, "박바다"],
    ["오류", 2, 3, "제외"],
  ]);
  assert.deepEqual(rows, [
    { grade: 1, class_num: 2, number: 1, name: "김하늘" },
    { grade: 1, class_num: 2, number: 2, name: "박바다" },
  ]);
});

test("잘못된 엑셀 행과 공백 이름을 제외한다", () => {
  assert.deepEqual(normalizeStudentRows([
    ["학년", "반", "번호", "이름"],
    [0, 1, 1, "범위 오류"],
    [1, 1, 0, "번호 오류"],
    [1, 1, 1, " "],
  ]), []);
});

test("대량 학생 명단을 순서대로 정규화한다", () => {
  const matrix = [["학년", "반", "번호", "이름"]];
  for (let number = 1000; number >= 1; number -= 1) matrix.push([1, 1, number, `학생${number}`]);
  const rows = normalizeStudentRows(matrix);
  assert.equal(rows.length, 1000);
  assert.equal(rows[0].number, 1);
  assert.equal(rows.at(-1).number, 1000);
});

test("선택과목 엑셀의 고정 열을 정규화한다", () => {
  assert.deepEqual(normalizeEnrollmentRows([
    ["과목명", "반", "번호", "이름", "호실"],
    ["경제", 1, 2, "김하늘", "별실 1"],
  ]), [{
    subject_name: "경제",
    class_num: 1,
    number: 2,
    name: "김하늘",
    room_name: "별실 1",
    row_number: 2,
  }]);
});

test("선택과목 정렬 뒤에도 미매칭 안내용 원본 행 번호를 유지한다", () => {
  const rows = normalizeEnrollmentRows([
    ["과목명", "반", "번호", "이름", "호실"],
    ["세계사", 1, 2, "김하늘", "별실 2"],
    ["경제", 1, 1, "박다온", "별실 1"],
  ]);
  assert.deepEqual(rows.map((item) => [item.subject_name, item.row_number]), [
    ["경제", 3],
    ["세계사", 2],
  ]);
});

test("자리배치 수동 명단의 반·번호·이름 열을 정규화한다", () => {
  assert.deepEqual(normalizeManualSeatRows([
    ["반", "번호", "이름"],
    [2, 3, "김하늘"],
    [2, 3, "중복학생"],
  ]), [{ class_num: 2, number: 3, name: "김하늘" }]);
});

test("잘못된 엑셀은 행 번호와 오류 코드를 보존하고 일부 저장을 막을 수 있다", () => {
  assert.deepEqual(analyzeStudentRows([
    ["학년", "반", "번호", "이름"],
    [1, 1, 1, "가"],
    [1, 1, 1, "중복"],
    [1, 1, 2, ""],
  ]).errors, [
    { row_number: 3, code: "DUPLICATE_STUDENT_NUMBER" },
    { row_number: 4, code: "INVALID_STUDENT_ROW" },
  ]);
  assert.deepEqual(analyzeEnrollmentRows([
    ["과목명", "반", "번호", "이름", "호실"],
    ["경제", 1, 1, "가", "별실"],
    ["경제", 1, 1, "가", "별실"],
  ]).errors, [{ row_number: 3, code: "DUPLICATE_ENROLLMENT" }]);
  assert.deepEqual(analyzeManualSeatRows([
    ["반", "번호", "이름"],
    [1, "오류", "가"],
  ]).errors, [{ row_number: 2, code: "INVALID_SEAT_ROW" }]);
});

test("결시 키는 같은 학생·고사일·교시에 대해 동일하다", () => {
  assert.equal(
    absenceKey({ exam_date: "2026-07-20", period: 1, student_id: "s1" }),
    absenceKey({ exam_date: "2026-07-20", period: "1", student_id: "s1" }),
  );
});

test("자리배치는 창가와 복도 시작 순서를 구분하고 제외 좌석을 건너뛴다", () => {
  assert.deepEqual(buildSeatSlots(2, 3, "window", ["0-1"]).map((item) => item.key), ["0-0", "0-2", "1-0", "1-1", "1-2"]);
  assert.deepEqual(buildSeatSlots(1, 3, "aisle", []).map((item) => item.key), ["0-2", "0-1", "0-0"]);
});

test("자리배치는 가로행 우선과 세로열 우선 순서를 모두 지원한다", () => {
  assert.deepEqual(
    buildSeatSlots(2, 3, "window", [], "row").map((item) => item.key),
    ["0-0", "0-1", "0-2", "1-0", "1-1", "1-2"],
  );
  assert.deepEqual(
    buildSeatSlots(2, 3, "window", [], "column").map((item) => item.key),
    ["0-0", "1-0", "0-1", "1-1", "0-2", "1-2"],
  );
  assert.deepEqual(
    buildSeatSlots(2, 3, "aisle", [], "column").map((item) => item.key),
    ["0-2", "1-2", "0-1", "1-1", "0-0", "1-0"],
  );
});

test("결시자 좌석을 제거하지 않고 표시 상태로 유지한다", () => {
  const students = [
    { id: "s1", number: 1, name: "가" },
    { id: "s2", number: 2, name: "나" },
  ];
  const result = generateSeatAssignment({
    students,
    rows: 1,
    cols: 2,
    mode: "separate",
    selectedIds: ["s1", "s2"],
    absentIds: ["s1"],
  });
  assert.equal(result.ok, true);
  assert.equal(result.assignments[0].student.id, "s1");
  assert.equal(result.assignments[0].absent, true);
});

test("좌석보다 응시자가 많으면 생성하지 않는다", () => {
  const result = generateSeatAssignment({
    students: [{ id: "s1", number: 1 }, { id: "s2", number: 2 }],
    rows: 1,
    cols: 1,
    mode: "separate",
    selectedIds: ["s1", "s2"],
  });
  assert.equal(result.ok, false);
  assert.match(result.error, /좌석/);
});

test("각자 교실은 응시자와 결시자를 먼저 배치하고 미응시자를 뒤에 둔다", () => {
  const result = generateSeatAssignment({
    students: [
      { id: "s1", number: 1, name: "가" },
      { id: "s2", number: 2, name: "나" },
      { id: "s3", number: 3, name: "다" },
    ],
    rows: 1,
    cols: 3,
    mode: "own",
    selectedIds: ["s2"],
    absentIds: ["s3"],
  });
  assert.deepEqual(result.assignments.map((item) => item.student.id), ["s2", "s3", "s1"]);
  assert.equal(result.assignments[1].absent, true);
});

test("학생별 호실을 시간표 기본 호실보다 우선한다", () => {
  const common = {
    subject_name: "경제",
    subject_type: "elective",
    room_name: "기본 별실",
  };
  assert.equal(resolveStudentRoom({
    studentId: "s1",
    subjectName: "경제",
    timetableItem: common,
    enrollments: [{ student_id: "s1", subject_name: "경제", room_name: "학생 별실" }],
  }), "학생 별실");
  assert.equal(resolveStudentRoom({
    studentId: "s2",
    subjectName: "경제",
    timetableItem: common,
    enrollments: [],
  }), "기본 별실");
});

test("같은 자리배치 조합은 같은 갱신 키를 사용한다", () => {
  const left = { mode: "separate", exam_date: "2026-07-20", period: 1, grade: 3, subject_name: "경제", room_name: "별실 1" };
  const right = { ...left, id: "다른-id", updated_at: "나중" };
  assert.equal(seatChartKey(left), seatChartKey(right));
});

test("결시 사유와 기간 통계를 사유·학급·일자·학생별로 집계한다", () => {
  const rows = [
    { reason: "질병", class_id: "c1", class_label: "1학년 1반", exam_date: "2026-07-20", student_id: "s1", student_number: 1, student_name: "가" },
    { reason: "질병", class_id: "c1", class_label: "1학년 1반", exam_date: "2026-07-21", student_id: "s1", student_number: 1, student_name: "가" },
    { reason: "인정", class_id: "c2", class_label: "1학년 2반", exam_date: "2026-07-21", student_id: "s2", student_number: 2, student_name: "나" },
  ];
  assert.equal(summarizeAbsences(rows).질병, 2);
  const summary = aggregateAbsenceReport(rows);
  assert.deepEqual(summary.reason_summary.map((item) => [item.reason, item.count]), [["질병", 2], ["인정", 1]]);
  assert.equal(summary.class_summary.length, 2);
  assert.equal(summary.date_summary.length, 2);
  assert.equal(summary.student_summary[0].count, 2);
});

test("선택과목 등록과 시간표 과목명을 연결해 개인 시간표를 만든다", () => {
  const student = { id: "s1", class_id: "c1", number: 1, name: "가" };
  const rows = buildPersonalTimetable({
    student,
    classInfo: { id: "c1", grade: 1, class_num: 1 },
    timetable: [
      { exam_date: "2026-07-20", grade: 1, period: 1, subject_name: "국어", subject_type: "common", start_time: "09:00", end_time: "09:50", class_ids: "c1" },
      { exam_date: "2026-07-20", grade: 1, period: 2, subject_name: "경제", subject_type: "elective", start_time: "10:10", end_time: "11:00", class_ids: "c1" },
    ],
    enrollments: [{ student_id: "s1", subject_name: "경제", room_name: "별실 1" }],
  });
  assert.deepEqual(rows.map((item) => [item.subject_name, item.room_name]), [["국어", "1학년 1반"], ["경제", "별실 1"]]);
});

test("기간 합산 통계 필터가 시작일과 종료일을 포함한다", () => {
  const rows = filterAbsences([
    { exam_date: "2026-07-19", grade: 1 },
    { exam_date: "2026-07-20", grade: 1 },
    { exam_date: "2026-07-21", grade: 1 },
    { exam_date: "2026-07-22", grade: 1 },
  ], { startDate: "2026-07-20", endDate: "2026-07-21" });
  assert.equal(rows.length, 2);
});

test("25개 동시 재시도 요청은 같은 request_id에서 한 번만 커밋된다", async () => {
  const completed = new Map();
  let writes = 0;
  let tail = Promise.resolve();
  const mutate = (requestId) => {
    const task = tail.then(async () => {
      if (completed.has(requestId)) return completed.get(requestId);
      await Promise.resolve();
      writes += 1;
      const result = { saved: true, writes };
      completed.set(requestId, result);
      return result;
    });
    tail = task.catch(() => {});
    return task;
  };
  const results = await Promise.all(Array.from({ length: 25 }, () => mutate("same-request")));
  assert.equal(writes, 1);
  assert.ok(results.every((item) => item.saved));
});

test("원안 포장 설정은 포장일·시간행·담당자를 정규화하고 중복을 검증한다", () => {
  const valid = validateExamPackagingConfig({
    revision: 2,
    input_deadline: "2026-07-30T17:00:00+09:00",
    capacity: 3,
    packaging_dates: ["2026-07-28"],
    rows: [
      { id: "slot-2", kind: "slot", period_label: "2회", start_time: "13:35", end_time: "14:00", sort_order: 3 },
      { id: "break-1", kind: "break", period_label: "쉬는 시간", start_time: "13:25", end_time: "13:35", sort_order: 2 },
      { id: "slot-1", kind: "slot", period_label: "1회", start_time: "13:00", end_time: "13:25", sort_order: 1 },
    ],
    staff_assignments: [
      { packaging_date: "2026-07-28", slot_id: "slot-1", staff_name: "담당 가" },
      { packaging_date: "2026-07-28", slot_id: "break-1", staff_name: "휴식 담당" },
    ],
  });
  assert.equal(valid.ok, true);
  assert.equal(valid.config.capacity, 3);
  assert.equal(valid.config.staff_assignments[1].slot_id, "break-1");

  const invalid = validateExamPackagingConfig({
    input_deadline: "잘못된 날짜",
    capacity: 0,
    packaging_dates: ["2026-07-28", "2026-07-28"],
    rows: [
      { id: "same", kind: "slot", period_label: "1회", start_time: "13:00", end_time: "13:25" },
      { id: "same", kind: "break", period_label: "쉬는 시간", start_time: "13:00", end_time: "13:25" },
      { id: "long-slot", kind: "slot", period_label: "겹침", start_time: "13:20", end_time: "13:50" },
    ],
    staff_assignments: [
      { packaging_date: "2026-07-28", slot_id: "same", staff_name: "담당 가" },
      { packaging_date: "2026-07-28", slot_id: "same", staff_name: "담당 나" },
    ],
  });
  assert.equal(invalid.ok, false);
  const codes = new Set(invalid.errors.map((item) => item.code));
  assert.ok(codes.has("INVALID_PACKAGING_DEADLINE"));
  assert.ok(codes.has("INVALID_PACKAGING_CAPACITY"));
  assert.ok(codes.has("DUPLICATE_PACKAGING_DATE"));
  assert.ok(codes.has("DUPLICATE_PACKAGING_ROW_ID"));
  assert.ok(codes.has("DUPLICATE_PACKAGING_TIME"));
  assert.ok(codes.has("INVALID_PACKAGING_SLOT_DURATION"));
  assert.ok(codes.has("OVERLAPPING_PACKAGING_TIME"));
  assert.ok(codes.has("DUPLICATE_PACKAGING_STAFF_ASSIGNMENT"));
});

test("원안 포장 항목은 시험일·교시·학년순으로 정렬하고 슬롯별로 묶는다", () => {
  const items = [
    { timetable_id: "b", exam_date: "2026-07-21", period: 1, grade: 1, subject_name: "수학", answer_sheet_type: "", packaging_date: "", packaging_slot_id: "", representative_teacher: "" },
    { timetable_id: "c", exam_date: "2026-07-20", period: 1, grade: 2, subject_name: "영어", answer_sheet_type: "a4", packaging_date: "2026-07-18", packaging_slot_id: "slot-1", representative_teacher: "담당 나" },
    { timetable_id: "a", exam_date: "2026-07-20", period: 1, grade: 1, subject_name: "국어", answer_sheet_type: "card", packaging_date: "2026-07-18", packaging_slot_id: "slot-1", representative_teacher: "담당 가" },
  ];
  assert.deepEqual(sortExamPackagingItems(items).map((item) => item.timetable_id), ["a", "c", "b"]);
  assert.deepEqual(summarizeExamPackaging(items), {
    total: 3,
    assigned: 2,
    unassigned: 1,
    answer_selected: 2,
    answer_unselected: 1,
    card: 1,
    a4: 1,
    representative_missing: 1,
  });
  const grouped = groupExamPackagingSchedule(items, {
    capacity: 3,
    packaging_dates: ["2026-07-18"],
    rows: [
      { id: "break-1", kind: "break", period_label: "휴식", start_time: "13:25", end_time: "13:35", sort_order: 2 },
      { id: "slot-1", kind: "slot", period_label: "1회", start_time: "13:00", end_time: "13:25", sort_order: 1 },
    ],
    staff_assignments: [
      { packaging_date: "2026-07-18", slot_id: "slot-1", staff_name: "교무 가" },
    ],
  });
  assert.equal(grouped.dates[0].rows[0].items.length, 2);
  assert.equal(grouped.dates[0].rows[0].staff_name, "교무 가");
  assert.deepEqual(grouped.dates[0].rows[1].items, []);
  assert.deepEqual(grouped.unassigned.map((item) => item.timetable_id), ["b"]);
});

test("원안 포장 설정에서 사용 중인 날짜나 슬롯을 없애면 영향 항목을 계산한다", () => {
  const current = normalizeExamPackagingConfig({
    capacity: 3,
    packaging_dates: ["2026-07-18", "2026-07-19"],
    rows: [
      { id: "slot-1", kind: "slot", period_label: "1회", start_time: "13:00", end_time: "13:25" },
      { id: "slot-2", kind: "slot", period_label: "2회", start_time: "13:35", end_time: "14:00" },
    ],
  });
  const next = { ...current, packaging_dates: ["2026-07-18"], rows: [current.rows[0]] };
  const impact = analyzeExamPackagingConfigImpact(current, next, [
    { timetable_id: "a", packaging_date: "2026-07-18", packaging_slot_id: "slot-1" },
    { timetable_id: "b", packaging_date: "2026-07-19", packaging_slot_id: "slot-1" },
    { timetable_id: "c", packaging_date: "2026-07-18", packaging_slot_id: "slot-2" },
  ]);
  assert.equal(impact.affected_count, 2);
  assert.deepEqual(impact.affected_timetable_ids, ["b", "c"]);
  assert.deepEqual(impact.removed_dates, ["2026-07-19"]);
  assert.deepEqual(impact.removed_slot_ids, ["slot-2"]);
});

test("비밀 링크 마감 판정은 지정 시각을 포함하고 Excel 수식 시작값을 무력화한다", () => {
  const config = { input_deadline: "2026-07-23T17:00:00+09:00" };
  assert.equal(isExamPackagingDeadlinePassed(config, Date.parse("2026-07-23T16:59:59+09:00")), false);
  assert.equal(isExamPackagingDeadlinePassed(config, Date.parse("2026-07-23T17:00:00+09:00")), true);
  assert.equal(sanitizeSpreadsheetCell("=HYPERLINK(\"https://example.invalid\")"), "'=HYPERLINK(\"https://example.invalid\")");
  assert.equal(sanitizeSpreadsheetCell("  +1+1"), "'  +1+1");
  assert.equal(sanitizeSpreadsheetCell("국어"), "국어");
  assert.equal(sanitizeSpreadsheetCell(123), 123);
});
