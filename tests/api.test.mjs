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

const { DemoRepository, sessionErrorBelongsToCurrentSession } = await import("../src/api.js");

function setUserSession(value) {
  sessionStorage.setItem("exam-management:user-session", value);
}

function setAdminSession(value) {
  sessionStorage.setItem("exam-management:admin-session", value);
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
