import test from "node:test";
import assert from "node:assert/strict";

class MemoryStorage {
  constructor() {
    this.values = new Map();
  }

  getItem(key) {
    return this.values.has(key) ? this.values.get(key) : null;
  }

  setItem(key, value) {
    this.values.set(key, String(value));
  }

  removeItem(key) {
    this.values.delete(key);
  }

  clear() {
    this.values.clear();
  }
}

globalThis.sessionStorage = new MemoryStorage();
globalThis.location = { search: "?demo=1" };
globalThis.window = { EXAM_MANAGEMENT_CONFIG: {} };

const {
  DemoRepository,
  clearPackagingSession,
  getPackagingSession,
  hasPackagingSession,
  savePackagingSession,
  sessionErrorBelongsToCurrentSession,
} = await import("../src/api.js");

function setUserSession(value) {
  sessionStorage.setItem("exam-management:user-session", value);
}

function setAdminSession(value) {
  sessionStorage.setItem("exam-management:admin-session", value);
}

function setPackagingSession(value) {
  sessionStorage.setItem("exam-management:packaging-session", value);
}

test("데모 로그인은 허용된 접속 이름과 숫자 6자리 PIN만 받는다", async () => {
  const repository = new DemoRepository();

  await assert.rejects(
    repository.request("user_login", { login_name: "없는교사", pin: "123456" }),
    (error) => error.code === "INVALID_LOGIN" && error.message === "접속 이름 또는 PIN을 확인하세요.",
  );
  await assert.rejects(
    repository.request("user_login", { login_name: "테스트교사", pin: "000000" }),
    (error) => error.code === "INVALID_LOGIN" && error.message === "접속 이름 또는 PIN을 확인하세요.",
  );

  const result = await repository.request("user_login", {
    login_name: "  테스트교사  ",
    pin: "123456",
  });
  assert.match(result.user_session, /^demo-user-session:demo-access-user:1$/);
  assert.equal(result.current_user.login_name, "테스트교사");
  assert.equal(result.bootstrap.current_user.id, "demo-access-user");
});

test("관리자 직접 진입은 사용자 세션과 관리자 세션을 함께 발급하고 PIN을 반환하지 않는다", async () => {
  const repository = new DemoRepository();
  await assert.rejects(
    repository.request("admin_entry_login", { password: "wrong" }),
    (error) => error.code === "INVALID_ADMIN_CREDENTIALS",
  );

  const result = await repository.request("admin_entry_login", { password: "demo-admin" });
  assert.equal(result.user_session, "demo-user-session:system-admin");
  assert.equal(result.admin_session, "demo-admin-session:system-admin");
  assert.equal(result.current_user.login_name, "관리자");
  assert.equal(result.bootstrap.access_users.length, 1);
  assert.equal(Object.hasOwn(result.bootstrap.access_users[0], "pin"), false);

  const emptyRepository = new DemoRepository();
  emptyRepository.accessUsers = [];
  const emptyResult = await emptyRepository.request("admin_entry_login", { password: "demo-admin" });
  assert.equal(emptyResult.bootstrap.access_users.length, 0);
  assert.equal(emptyResult.current_user.id, "system-admin");
});

test("로그아웃 요청은 브라우저 세션을 먼저 지운 뒤에도 만료 오류를 만들지 않는다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  const result = await repository.request("logout");
  assert.deepEqual(result, { logged_out: true });
});

test("이전 요청의 만료 응답은 새 로그인 세션을 끝내지 않는다", () => {
  sessionStorage.clear();
  setUserSession("new-user-session");
  setAdminSession("new-admin-session");
  assert.equal(
    sessionErrorBelongsToCurrentSession("SESSION_EXPIRED", "old-user-session", ""),
    false,
  );
  assert.equal(
    sessionErrorBelongsToCurrentSession("SESSION_EXPIRED", "new-user-session", ""),
    true,
  );
  assert.equal(
    sessionErrorBelongsToCurrentSession("ADMIN_SESSION_EXPIRED", "new-user-session", "old-admin-session"),
    false,
  );
  assert.equal(
    sessionErrorBelongsToCurrentSession("ADMIN_SESSION_EXPIRED", "new-user-session", "new-admin-session"),
    true,
  );
  setPackagingSession("new-packaging-session");
  assert.equal(
    sessionErrorBelongsToCurrentSession("PACKAGING_LINK_EXPIRED", "", "", "new-packaging-session"),
    true,
  );
});

test("관리자는 접속 사용자를 추가·수정·비활성화하고 PIN 재설정으로 이전 세션을 만료시킨다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");

  const created = await repository.request("save_access_user", {
    login_name: "새 교사",
    pin: "654321",
  });
  const user = created.access_users.find((item) => item.login_name === "새 교사");
  assert.ok(user);

  await assert.rejects(
    repository.request("save_access_user", { login_name: "  새　교사 ", pin: "111111" }),
    (error) => error.code === "DUPLICATE_LOGIN_NAME",
  );

  const firstLogin = await repository.request("user_login", {
    login_name: "새 교사",
    pin: "654321",
  });
  setUserSession(firstLogin.user_session);
  assert.equal((await repository.request("get_bootstrap")).current_user.id, user.id);

  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  await repository.request("reset_access_user_pin", { id: user.id, pin: "111111" });
  setUserSession(firstLogin.user_session);
  await assert.rejects(
    repository.request("get_bootstrap"),
    (error) => error.code === "SESSION_EXPIRED",
  );

  const secondLogin = await repository.request("user_login", {
    login_name: "새 교사",
    pin: "111111",
  });
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  await repository.request("set_access_user_active", { id: user.id, active: false });
  setUserSession(secondLogin.user_session);
  await assert.rejects(
    repository.request("get_bootstrap"),
    (error) => error.code === "SESSION_EXPIRED",
  );
  await assert.rejects(
    repository.request("user_login", { login_name: "새 교사", pin: "111111" }),
    (error) => error.code === "INVALID_LOGIN",
  );
});

test("데모 관리자 작업도 현재 사용자와 연결된 관리자 세션을 요구한다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  const login = await repository.request("user_login", {
    login_name: "테스트교사",
    pin: "123456",
  });
  setUserSession(login.user_session);

  await assert.rejects(
    repository.request("get_admin_bootstrap"),
    (error) => error.code === "ADMIN_SESSION_EXPIRED",
  );

  const unlocked = await repository.request("admin_login", { password: "demo-admin" });
  assert.equal(unlocked.admin_session, "demo-admin-session:demo-access-user");
  setAdminSession(unlocked.admin_session);
  const bootstrap = await repository.request("get_admin_bootstrap");
  assert.equal(bootstrap.current_user.id, "demo-access-user");
});

test("비밀 원안 포장 링크는 로그인 없이 전용 최소 컨텍스트로 교환된다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  await assert.rejects(
    repository.request("redeem_exam_packaging_invite", { token: "wrong" }),
    (error) => error.code === "INVALID_PACKAGING_LINK",
  );

  const redeemed = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  assert.match(redeemed.packaging_session, /^demo-packaging-session:/);
  assert.equal(redeemed.context.settings.school_name, "한양대학교사범대학부속고등학교");
  assert.equal(redeemed.context.summary.total, 5);
  assert.equal(Object.hasOwn(redeemed.context, "students"), false);
  assert.equal(Object.hasOwn(redeemed.context, "classes"), false);
  assert.equal(Object.hasOwn(redeemed.context, "access_users"), false);

  savePackagingSession(redeemed.packaging_session);
  assert.equal(hasPackagingSession(), true);
  assert.equal(getPackagingSession(), redeemed.packaging_session);
  const context = await repository.request("get_exam_packaging");
  assert.equal(context.items[0].exam_date, "2026-07-20");
  clearPackagingSession();
  assert.equal(hasPackagingSession(), false);
});

test("비밀 링크 사용자는 답안지와 포장 배정을 저장하고 항목 revision 충돌을 확인한다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  const redeemed = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(redeemed.packaging_session);
  const before = redeemed.context.items.find((item) => item.timetable_id === "tt2");
  const saved = await repository.request("save_exam_packaging_items", {
    expected_config_revision: redeemed.context.config.revision,
    items: [{
      timetable_id: "tt2",
      expected_revision: before.revision,
      representative_teacher: "가상교사 다",
      answer_sheet_type: "card",
      packaging_date: "2026-07-17",
      packaging_slot_id: "slot-2",
    }],
  });
  assert.equal(saved.changed_items[0].revision, 1);
  assert.equal(saved.changed_items[0].representative_teacher, "가상교사 다");
  assert.equal(saved.summary.assigned, 3);

  await assert.rejects(
    repository.request("save_exam_packaging_items", {
      expected_config_revision: saved.config.revision,
      items: [{
        timetable_id: "tt2",
        expected_revision: 0,
        representative_teacher: "다른 수정",
        answer_sheet_type: "a4",
        packaging_date: "2026-07-17",
        packaging_slot_id: "slot-2",
      }],
    }),
    (error) =>
      error.code === "PACKAGING_REVISION_CONFLICT" &&
      error.details.current_revision === 1,
  );
});

test("쉬는 시간에는 배정할 수 없고 같은 3자리 슬롯의 네 동시 요청은 정확히 세 건만 성공한다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  const redeemed = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(redeemed.packaging_session);
  await assert.rejects(
    repository.request("save_exam_packaging_items", {
      expected_config_revision: 1,
      items: [{
        timetable_id: "tt2",
        expected_revision: 0,
        representative_teacher: "가상교사",
        answer_sheet_type: "card",
        packaging_date: "2026-07-17",
        packaging_slot_id: "break-1",
      }],
    }),
    (error) => error.code === "PACKAGING_BREAK_NOT_ASSIGNABLE",
  );

  const ids = ["tt1", "tt2", "tt3", "tt4"];
  const requests = ids.map((id) => {
    const item = repository.state.timetable.find((row) => row.id === id);
    return repository.request("save_exam_packaging_items", {
      expected_config_revision: 1,
      items: [{
        timetable_id: id,
        expected_revision: item.packaging_revision,
        representative_teacher: `가상-${id}`,
        answer_sheet_type: "card",
        packaging_date: "2026-07-18",
        packaging_slot_id: "slot-3",
      }],
    });
  });
  const settled = await Promise.allSettled(requests);
  assert.equal(settled.filter((item) => item.status === "fulfilled").length, 3);
  const rejected = settled.find((item) => item.status === "rejected");
  assert.equal(rejected.reason.code, "PACKAGING_SLOT_FULL");
  assert.equal(
    repository.state.timetable.filter(
      (item) => item.packaging_date === "2026-07-18" && item.packaging_slot_id === "slot-3",
    ).length,
    3,
  );
});

test("비밀 링크는 30분 미사용 뒤 만료되고 입력 마감 뒤에는 읽기 전용이다", async () => {
  let now = Date.parse("2026-07-23T09:00:00+09:00");
  const repository = new DemoRepository({ now: () => now });
  sessionStorage.clear();
  const redeemed = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(redeemed.packaging_session);
  now += 29 * 60 * 1000;
  await repository.request("get_exam_packaging");
  now += 60 * 1000 + 1;
  await assert.rejects(
    repository.request("get_exam_packaging"),
    (error) => error.code === "PACKAGING_SESSION_EXPIRED",
  );

  const readOnlyRepository = new DemoRepository({
    now: () => Date.parse("2100-01-01T09:00:00+09:00"),
  });
  readOnlyRepository.packagingInvite.expires_at = "2100-12-31T23:59:59+09:00";
  sessionStorage.clear();
  const readOnlyRedeemed = await readOnlyRepository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(readOnlyRedeemed.packaging_session);
  assert.equal(readOnlyRedeemed.context.read_only, true);
  await assert.rejects(
    readOnlyRepository.request("save_exam_packaging_items", {
      expected_config_revision: 1,
      items: [],
    }),
    (error) => error.code === "PACKAGING_READ_ONLY",
  );

  sessionStorage.clear();
  const login = await readOnlyRepository.request("user_login", {
    login_name: "테스트교사",
    pin: "123456",
  });
  setUserSession(login.user_session);
  assert.equal((await readOnlyRepository.request("get_exam_packaging")).read_only, false);
  assert.equal(
    (await readOnlyRepository.request("save_exam_packaging_items", {
      expected_config_revision: 1,
      items: [],
    })).read_only,
    false,
  );
});

test("원안 포장 전용 접속 현황은 본인 화면을 제외하고 실제 활동 때만 링크 세션을 연장한다", async () => {
  let now = Date.parse("2026-07-23T09:00:00+09:00");
  const repository = new DemoRepository({ now: () => now });
  sessionStorage.clear();
  const first = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(first.packaging_session);
  assert.equal((await repository.request("presence_ping", { scope: "exam_packaging" })).online_count, 0);

  const second = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setPackagingSession(second.packaging_session);
  assert.equal((await repository.request("presence_ping", { scope: "exam_packaging" })).online_count, 1);

  now += 29 * 60 * 1000;
  await repository.request("presence_ping", { scope: "exam_packaging", touch: true });
  now += 29 * 60 * 1000;
  assert.equal((await repository.request("get_exam_packaging")).summary.total, 5);
});

test("관리자는 비밀 링크를 재발급·폐기하고 기존 링크 세션을 즉시 끝낸다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  const first = await repository.request("redeem_exam_packaging_invite", {
    token: "demo-packaging-invite",
  });
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  const created = await repository.request("create_exam_packaging_invite", {
    expires_at: "2098-12-31T23:59:59+09:00",
  });
  assert.match(created.token, /^[a-f0-9]{64}$/);
  assert.equal(Object.hasOwn(created.invite_status, "token"), false);

  sessionStorage.clear();
  setPackagingSession(first.packaging_session);
  await assert.rejects(
    repository.request("get_exam_packaging"),
    (error) => error.code === "PACKAGING_SESSION_EXPIRED",
  );

  sessionStorage.clear();
  const second = await repository.request("redeem_exam_packaging_invite", {
    token: created.token,
  });
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  await repository.request("disable_exam_packaging_invite");
  sessionStorage.clear();
  setPackagingSession(second.packaging_session);
  await assert.rejects(
    repository.request("get_exam_packaging"),
    (error) => error.code === "PACKAGING_SESSION_EXPIRED",
  );
  await assert.rejects(
    repository.request("redeem_exam_packaging_invite", { token: created.token }),
    (error) => error.code === "INVALID_PACKAGING_LINK",
  );
});

test("원안 포장 설정 변경은 영향 건수를 미리 보여주고 강제 확인 때만 배정을 비운다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  const current = structuredClone(repository.state.examPackagingConfig);
  const next = {
    ...current,
    packaging_dates: ["2026-07-18"],
    staff_assignments: [],
  };
  const preview = await repository.request("save_exam_packaging_config", {
    expected_revision: current.revision,
    config: next,
    preview_only: true,
  });
  assert.equal(preview.impact.affected_count, 2);
  await assert.rejects(
    repository.request("save_exam_packaging_config", {
      expected_revision: current.revision,
      config: next,
    }),
    (error) => error.code === "PACKAGING_CONFIG_IN_USE",
  );
  const saved = await repository.request("save_exam_packaging_config", {
    expected_revision: current.revision,
    config: next,
    force: true,
  });
  assert.equal(saved.exam_packaging_config.revision, current.revision + 1);
  assert.equal(saved.impact.affected_count, 2);
  assert.equal(repository.state.timetable.find((item) => item.id === "tt1").packaging_date, "");
  assert.equal(repository.state.timetable.find((item) => item.id === "tt1").packaging_revision, 2);
});

test("관리자 시간표 일괄 저장은 서버의 원안 포장 필드를 보존한다", async () => {
  const repository = new DemoRepository();
  sessionStorage.clear();
  setUserSession("demo-user-session:system-admin");
  setAdminSession("demo-admin-session:system-admin");
  const malicious = repository.state.timetable.map((item) => ({
    ...item,
    subject_name: item.id === "tt1" ? "변경된 국어" : item.subject_name,
    representative_teacher: "덮어쓰기 시도",
    answer_sheet_type: "a4",
    packaging_date: "2026-07-18",
    packaging_slot_id: "slot-3",
    packaging_revision: 999,
  }));
  await repository.request("save_timetable", { timetable: malicious });
  const saved = repository.state.timetable.find((item) => item.id === "tt1");
  assert.equal(saved.subject_name, "변경된 국어");
  assert.equal(saved.representative_teacher, "가상교사 가");
  assert.equal(saved.answer_sheet_type, "card");
  assert.equal(saved.packaging_slot_id, "slot-1");
  assert.equal(saved.packaging_revision, 1);
});
