import test from "node:test";
import assert from "node:assert/strict";
import {
  absenceKey,
  buildPersonalTimetable,
  buildSeatSlots,
  filterAbsences,
  generateSeatAssignment,
  normalizeEnrollmentRows,
  normalizeManualSeatRows,
  normalizeStudentRows,
} from "../src/core.js";

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
  ]), [{ subject_name: "경제", class_num: 1, number: 2, name: "김하늘", room_name: "별실 1" }]);
});

test("자리배치 수동 명단의 반·번호·이름 열을 정규화한다", () => {
  assert.deepEqual(normalizeManualSeatRows([
    ["반", "번호", "이름"],
    [2, 3, "김하늘"],
    [2, 3, "중복학생"],
  ]), [{ class_num: 2, number: 3, name: "김하늘" }]);
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
