import { absenceKey, completionKey } from "./core.js";

const classes = [
  { id: "c11", grade: 1, class_num: 1, active: true },
  { id: "c12", grade: 1, class_num: 2, active: true },
  { id: "c21", grade: 2, class_num: 1, active: true },
];

const students = [
  { id: "s111", class_id: "c11", number: 1, name: "강하늘" },
  { id: "s112", class_id: "c11", number: 2, name: "김도윤" },
  { id: "s113", class_id: "c11", number: 3, name: "최다온" },
  { id: "s114", class_id: "c11", number: 4, name: "문서아" },
  { id: "s121", class_id: "c12", number: 1, name: "김보람" },
  { id: "s122", class_id: "c12", number: 2, name: "박누리" },
  { id: "s123", class_id: "c12", number: 3, name: "정하람" },
  { id: "s211", class_id: "c21", number: 1, name: "서여름" },
  { id: "s212", class_id: "c21", number: 2, name: "오마루" },
  { id: "s213", class_id: "c21", number: 3, name: "윤새봄" },
];

const examDates = [
  { id: "d1", exam_date: "2026-07-20", label: "1일차", active: true },
  { id: "d2", exam_date: "2026-07-21", label: "2일차", active: true },
];

const timetable = [
  {
    id: "tt1",
    exam_date: "2026-07-20",
    grade: 1,
    period: 1,
    start_time: "09:00",
    end_time: "09:50",
    subject_name: "국어",
    subject_type: "common",
    room_name: "",
    class_ids: "c11|c12",
  },
  {
    id: "tt2",
    exam_date: "2026-07-20",
    grade: 1,
    period: 2,
    start_time: "10:10",
    end_time: "11:00",
    subject_name: "경제",
    subject_type: "elective",
    room_name: "별실 1",
    class_ids: "c11|c12",
  },
  {
    id: "tt3",
    exam_date: "2026-07-20",
    grade: 1,
    period: 2,
    start_time: "10:10",
    end_time: "11:00",
    subject_name: "세계사와 지리",
    subject_type: "elective",
    room_name: "별실 2",
    class_ids: "c11|c12",
  },
  {
    id: "tt4",
    exam_date: "2026-07-21",
    grade: 1,
    period: 1,
    start_time: "09:00",
    end_time: "09:50",
    subject_name: "수학",
    subject_type: "common",
    room_name: "",
    class_ids: "c11|c12",
  },
  {
    id: "tt5",
    exam_date: "2026-07-20",
    grade: 2,
    period: 1,
    start_time: "09:00",
    end_time: "09:50",
    subject_name: "영어",
    subject_type: "common",
    room_name: "",
    class_ids: "c21",
  },
].map((item) => ({
  ...item,
  representative_teacher: "",
  answer_sheet_type: "",
  packaging_date: "",
  packaging_slot_id: "",
  packaging_revision: 0,
  packaging_updated_at: "",
}));

Object.assign(timetable[0], {
  representative_teacher: "가상교사 가",
  answer_sheet_type: "card",
  packaging_date: "2026-07-17",
  packaging_slot_id: "slot-1",
  packaging_revision: 1,
  packaging_updated_at: "2026-07-15T15:10:00+09:00",
});
Object.assign(timetable[4], {
  representative_teacher: "가상교사 나",
  answer_sheet_type: "a4",
  packaging_date: "2026-07-17",
  packaging_slot_id: "slot-1",
  packaging_revision: 1,
  packaging_updated_at: "2026-07-15T15:12:00+09:00",
});

const subjectCatalog = {
  1: ["국어", "경제", "세계사와 지리", "수학"],
  2: ["영어"],
  3: [],
};

const examNotice = {
  year: 2026,
  semester: 1,
  exam_name: "기말고사",
  greeting: "",
  notes: [],
  issue_date: "2026-07-10",
  issuer: "한양대학교사범대학부속고등학교장",
};

const examPackagingConfig = {
  revision: 1,
  input_deadline: "2099-12-31T18:00:00+09:00",
  capacity: 3,
  packaging_dates: ["2026-07-17", "2026-07-18"],
  rows: [
    {
      id: "slot-1",
      kind: "slot",
      period_label: "1회",
      start_time: "13:00",
      end_time: "13:25",
      note: "",
      sort_order: 1,
    },
    {
      id: "break-1",
      kind: "break",
      period_label: "쉬는 시간",
      start_time: "13:25",
      end_time: "13:35",
      note: "",
      sort_order: 2,
    },
    {
      id: "slot-2",
      kind: "slot",
      period_label: "2회",
      start_time: "13:35",
      end_time: "14:00",
      note: "",
      sort_order: 3,
    },
    {
      id: "slot-3",
      kind: "slot",
      period_label: "3회",
      start_time: "14:00",
      end_time: "14:25",
      note: "",
      sort_order: 4,
    },
  ],
  staff_assignments: [
    { packaging_date: "2026-07-17", slot_id: "slot-1", staff_name: "교무담당 가" },
    { packaging_date: "2026-07-17", slot_id: "slot-2", staff_name: "교무담당 나" },
  ],
};

const enrollments = [
  { id: "e1", grade: 1, class_id: "c11", student_id: "s111", subject_name: "경제", room_name: "별실 1" },
  { id: "e2", grade: 1, class_id: "c11", student_id: "s112", subject_name: "세계사와 지리", room_name: "별실 2" },
  { id: "e3", grade: 1, class_id: "c12", student_id: "s121", subject_name: "경제", room_name: "별실 1" },
];

const absence = {
  id: "a1",
  exam_date: "2026-07-20",
  period: 1,
  class_id: "c11",
  student_id: "s113",
  reason: "질병",
  reason_detail: "",
  updated_at: "2026-07-20T09:02:00+09:00",
};
absence.key = absenceKey(absence);

const completion = {
  id: "cp1",
  exam_date: "2026-07-20",
  period: 1,
  class_id: "c12",
  completed_at: "2026-07-20T09:55:00+09:00",
};
completion.key = completionKey(completion);

export function createDemoState() {
  return structuredClone({
    settings: {
      school_name: "한양대학교사범대학부속고등학교",
      app_name: "한양대학교사범대학부속고등학교 통합 시스템",
      timezone: "Asia/Seoul",
    },
    classes,
    students,
    examDates,
    timetable,
    subjectCatalog,
    examNotice,
    examPackagingConfig,
    enrollments,
    absences: [absence],
    completions: [completion],
    seatCharts: [],
  });
}
