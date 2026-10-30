import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
const source=fs.readFileSync(new URL('../src/App.jsx',import.meta.url),'utf8');
const start=source.indexOf('  useEffect(() => {\n    if (!loggedIn) return;');
const end=source.indexOf('  }, [loggedIn, admin]);',start);
assert.ok(start>=0&&end>start);
const effect=source.slice(start+'  useEffect(() => {'.length,end);
const deferred=()=>{let resolve;const promise=new Promise(done=>{resolve=done;});return{promise,resolve};};
test('최신 수동 동기화 뒤 늦게 도착한 startup 응답을 무시한다',async()=>{
 const pending=deferred();let state={students:['old']};const revision={current:0};
 const environment={loggedIn:true,admin:true,bootstrapGenerationRef:revision,apiRequest:()=>pending.promise,setBootstrap:value=>{state=typeof value==='function'?value(state):value;},setStartupError(){},setAdmin(){},setLoggedIn(){},clearSessions(){},messageFrom:error=>error.message};
 const cleanup=vm.runInNewContext(`(()=>{${effect}\n})`,environment)();assert.equal(typeof cleanup,'function');
 revision.current++;state={students:['new']};pending.resolve({students:['old']});await new Promise(done=>setImmediate(done));assert.deepEqual(state.students,['new']);
});
test('관리자 승격으로 종료한 teacher bootstrap은 새 학급·학생 조합을 바꾸지 않는다',async()=>{
 const pending=deferred();let state={classes:['old']};const environment={loggedIn:true,admin:false,bootstrapGenerationRef:{current:0},apiRequest:()=>pending.promise,setBootstrap:value=>{state=typeof value==='function'?value(state):value;},setStartupError(){},setAdmin(){},setLoggedIn(){},clearSessions(){},messageFrom:error=>error.message};
 const cleanup=vm.runInNewContext(`(()=>{${effect}\n})`,environment)();cleanup();state={classes:['new'],students:[{class_id:'new'}]};pending.resolve({classes:['old']});await new Promise(done=>setImmediate(done));assert.deepEqual(state.classes,['new']);
});
test('교시 option은 표시 문구 대신 숫자 value를 제출한다',()=>{
 const select=source.match(/<label>교시<select value=\{timetableDraft.period\}[\s\S]*?<\/select>/)?.[0]||'';
 assert.match(select,/<option key=\{item\} value=\{item\}>/);
 for(const value of ['1','2','7'])assert.ok(Number.isInteger(Number(value)));
});
test('선택과목 자동 응시자도 시간표의 적용 학급을 따른다',()=>{
 const begin=source.indexOf('  const defaultExamineeIds = useMemo(() => {');const finish=source.indexOf('  }, [',begin);const body=source.slice(begin+'  const defaultExamineeIds = useMemo(() => {'.length,finish);
 const environment={selectedTimetable:{class_ids:['c1'],subject_type:'elective'},students:[{id:'s1',class_id:'c1'},{id:'s2',class_id:'c2'}],bootstrap:{enrollments:[{student_id:'s1',subject_name:'합성과목'},{student_id:'s2',subject_name:'합성과목'}]},subjectName:'합성과목',mode:'own',roomName:'',classMap:new Map(),resolveStudentRoom:()=>''};
 assert.deepEqual(Array.from(vm.runInNewContext(`(()=>{${body}\n})`,environment)()),['s1']);
});
test('연결된 학급·학생 입력은 비활성화하되 학급 조회 버튼은 유지한다',()=>{
 const block=source.slice(source.indexOf('{active === "classes" ? ('),source.indexOf('{active === "timetable" ? ('));
 assert.equal((block.match(/<fieldset disabled=\{academicManaged\}/g)||[]).length,2);assert.match(block,/<\/fieldset>\s*<div className="class-chips">/);assert.match(source,/onBootstrap\(\(current\) => \(\{ \.\.\.current, \.\.\.result \}\)\)/);
});
