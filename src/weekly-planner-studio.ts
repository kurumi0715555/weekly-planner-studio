export {};

type Day = 'monday' | 'tuesday' | 'wednesday' | 'thursday' | 'friday';
type WeekKey = 'A' | 'B' | 'B1' | 'B2';
type RotationType = 'single' | 'AB' | 'ABC';
type CheckLevel = '' | 'error' | 'warning';

interface Homeroom { grade: number; section: number; }
interface TeachingClass extends Homeroom { id: string; subjects: string[]; }
interface Holiday { date: string; label: string; }
interface Vacation { label: string; start: string; end: string; }
interface TimetableWeek { [day: string]: string[]; }
interface PlannerState {
  version: string;
  teacherName: string;
  year: number;
  homeroom: Homeroom | null;
  teachingClasses: TeachingClass[];
  rotation: { type: RotationType; weeks: WeekKey[]; startDate: string };
  timetable: Record<string, TimetableWeek>;
  calendar: { termStart: string; termEnd: string; holidays: Holiday[]; vacations: Vacation[] };
  updatedAt: string | null;
}
interface WeekPlan { monday: string; weekKey: WeekKey; }
interface CheckItem { level: CheckLevel; title: string; body: string; }
interface LessonCombo { classId: string; subject: string; }
interface ModalApi { confirm(message: string, options?: { danger?: boolean }): Promise<boolean>; alert(message: string): Promise<void>; }
interface ExcelCell { value: unknown; font?: Record<string, unknown>; alignment?: Record<string, unknown>; }
interface ExcelWorksheet { columns: Array<{ width: number }>; getCell(address: string | number, column?: number): ExcelCell; }
interface ExcelWorkbook {
  creator: string;
  created: Date;
  addWorksheet(name: string): ExcelWorksheet;
  xlsx: { writeBuffer(): Promise<ArrayBuffer> };
}
interface ExcelJsApi { Workbook: new () => ExcelWorkbook; }

const STORAGE_KEY = 'weekly-planner-studio/v1/state';
const VERSION = '1.0';
const DAYS: Day[] = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'];
const DAY_LABELS = ['月', '火', '水', '木', '金'];
const PERIODS = 6;
const ACADEMIC_SUBJECTS = ['国語', '数学', '英語', '理科', '社会', '音楽', '美術', '保体', '技術', '家庭科'];
const HOMEROOM_SUBJECTS = ['道徳', '学活', '総合', '学校行事', '生徒会'];
const ROTATION_WEEKS: Record<RotationType, WeekKey[]> = {
  single: ['A'], AB: ['A', 'B'], ABC: ['A', 'B1', 'B2'],
};
const HOLIDAY_PRESETS: Record<number, Array<[string, string]>> = {
  2026: [['2026-04-29', '昭和の日'], ['2026-05-03', '憲法記念日'], ['2026-05-04', 'みどりの日'], ['2026-05-05', 'こどもの日'], ['2026-05-06', '振替休日'], ['2026-07-20', '海の日'], ['2026-08-11', '山の日'], ['2026-09-21', '敬老の日'], ['2026-09-23', '秋分の日'], ['2026-10-12', 'スポーツの日'], ['2026-11-03', '文化の日'], ['2026-11-23', '勤労感謝の日'], ['2027-01-01', '元日'], ['2027-01-11', '成人の日'], ['2027-02-11', '建国記念の日'], ['2027-02-23', '天皇誕生日'], ['2027-03-21', '春分の日'], ['2027-03-22', '振替休日']],
  2027: [['2027-04-29', '昭和の日'], ['2027-05-03', '憲法記念日'], ['2027-05-04', 'みどりの日'], ['2027-05-05', 'こどもの日'], ['2027-07-19', '海の日'], ['2027-08-11', '山の日'], ['2027-09-20', '敬老の日'], ['2027-09-23', '秋分の日'], ['2027-10-11', 'スポーツの日'], ['2027-11-03', '文化の日'], ['2027-11-23', '勤労感謝の日'], ['2028-01-01', '元日'], ['2028-01-10', '成人の日'], ['2028-02-11', '建国記念の日'], ['2028-02-23', '天皇誕生日'], ['2028-03-20', '春分の日']],
};

let activeView = 'setup';
let activeWeek: WeekKey = 'A';
let saveTimer: number | undefined;
let toastTimer: number | undefined;
let state = createInitialState();

function qs<T extends Element>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error(`必要な画面要素が見つかりません: ${selector}`);
  return element;
}
function qsa<T extends Element>(selector: string): T[] { return Array.from(document.querySelectorAll<T>(selector)); }
function optional<T extends Element>(selector: string): T | null { return document.querySelector<T>(selector); }
function isRecord(value: unknown): value is Record<string, unknown> { return typeof value === 'object' && value !== null; }
function textValue(value: unknown): string { return typeof value === 'string' ? value : ''; }
function finiteNumber(value: unknown, fallback: number): number { return typeof value === 'number' && Number.isFinite(value) ? value : fallback; }

function createEmptyTimetable(weeks: WeekKey[]): Record<string, TimetableWeek> {
  const table: Record<string, TimetableWeek> = {};
  weeks.forEach(week => {
    const weekTable: TimetableWeek = {};
    table[week] = weekTable;
    DAYS.forEach(day => { weekTable[day] = new Array<string>(PERIODS).fill(''); });
  });
  return table;
}
function buildHolidayPreset(year: number): Holiday[] { return (HOLIDAY_PRESETS[year] ?? []).map(([date, label]) => ({ date, label })); }
function buildVacationDefaults(year: number): Vacation[] {
  return [{ label: '夏休み', start: `${year}-07-21`, end: `${year}-08-31` }, { label: '冬休み', start: `${year}-12-25`, end: `${year + 1}-01-07` }, { label: '春休み', start: `${year + 1}-03-25`, end: `${year + 1}-04-07` }];
}
function createInitialState(): PlannerState {
  const year = new Date().getFullYear();
  return { version: VERSION, teacherName: '', year, homeroom: null, teachingClasses: [], rotation: { type: 'single', weeks: ['A'], startDate: '' }, timetable: createEmptyTimetable(['A']), calendar: { termStart: `${year}-04-01`, termEnd: `${year + 1}-03-31`, holidays: buildHolidayPreset(year), vacations: buildVacationDefaults(year) }, updatedAt: null };
}
function escapeText(value: unknown): string {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch] ?? ch);
}
function getWeekData(week: WeekKey): TimetableWeek {
  const data = state.timetable[week];
  if (data) return data;
  const created: TimetableWeek = {};
  DAYS.forEach(day => { created[day] = new Array<string>(PERIODS).fill(''); });
  state.timetable[week] = created;
  return created;
}
function getDayData(week: WeekKey, day: Day): string[] {
  const data = getWeekData(week);
  const cells = data[day];
  if (cells) return cells;
  const created = new Array<string>(PERIODS).fill('');
  data[day] = created;
  return created;
}
function getModal(): ModalApi | null { return (window as unknown as { Modal?: ModalApi }).Modal ?? null; }

function init(): void { loadSavedState(); populateSelects(); bindEvents(); renderAll(); }
function populateSelects(): void {
  const yearSelect = qs<HTMLSelectElement>('#schoolYear');
  const current = new Date().getFullYear();
  yearSelect.innerHTML = '';
  for (let year = current - 1; year <= current + 3; year += 1) yearSelect.appendChild(new Option(`${year}年度`, String(year)));
  ['homeroomSection', 'addSectionStart', 'addSectionEnd'].forEach(id => {
    const select = qs<HTMLSelectElement>(`#${id}`); select.innerHTML = '';
    for (let section = 1; section <= 10; section += 1) select.appendChild(new Option(`${section}組`, String(section)));
  });
  const subject = qs<HTMLSelectElement>('#addSubject'); subject.innerHTML = '';
  ACADEMIC_SUBJECTS.forEach(name => subject.appendChild(new Option(name, name)));
}
function bindEvents(): void {
  qsa<HTMLButtonElement>('.studio-tab').forEach(button => button.addEventListener('click', () => showView(button.dataset.view ?? 'setup')));
  bindInput('#teacherName', value => { state.teacherName = value; });
  bindInput('#schoolYear', value => {
    const nextYear = Number(value); if (nextYear === state.year || !Number.isFinite(nextYear)) return;
    state.year = nextYear; state.calendar.holidays = buildHolidayPreset(nextYear); state.calendar.vacations = buildVacationDefaults(nextYear);
    if (!state.calendar.termStart || !state.calendar.termEnd) { state.calendar.termStart = `${nextYear}-04-01`; state.calendar.termEnd = `${nextYear + 1}-03-31`; }
    renderCalendar();
  }, 'change');
  qs<HTMLInputElement>('#hasHomeroom').addEventListener('change', event => {
    const checked = (event.currentTarget as HTMLInputElement).checked;
    state.homeroom = checked ? { grade: Number(qs<HTMLSelectElement>('#homeroomGrade').value), section: Number(qs<HTMLSelectElement>('#homeroomSection').value) } : null;
    renderSetup(); scheduleSave();
  });
  ['#homeroomGrade', '#homeroomSection'].forEach(selector => bindInput(selector, () => {
    if (!state.homeroom) return; state.homeroom.grade = Number(qs<HTMLSelectElement>('#homeroomGrade').value); state.homeroom.section = Number(qs<HTMLSelectElement>('#homeroomSection').value); renderSetup();
  }, 'change'));
  qs<HTMLButtonElement>('#btnAddClass').addEventListener('click', addTeachingClassesFromForm);
  bindInput('#rotationType', value => applyRotation(value), 'change');
  bindInput('#rotationStartDate', value => { state.rotation.startDate = value; }, 'change');
  qs<HTMLButtonElement>('#btnFillNone').addEventListener('click', fillActiveWeekNone);
  qs<HTMLButtonElement>('#btnClearWeek').addEventListener('click', () => { void clearActiveWeek(); });
  bindInput('#termStart', value => { state.calendar.termStart = value; }, 'change');
  bindInput('#termEnd', value => { state.calendar.termEnd = value; }, 'change');
  qs<HTMLButtonElement>('#btnAddHoliday').addEventListener('click', addHoliday);
  qs<HTMLButtonElement>('#btnAddVacation').addEventListener('click', addVacation);
  qs<HTMLButtonElement>('#btnPreviewTop').addEventListener('click', () => showView('preview'));
  qs<HTMLButtonElement>('#btnGenerateBottom').addEventListener('click', () => { void generateExcel(); });
  qs<HTMLButtonElement>('#btnSaveConfig').addEventListener('click', saveConfigFile);
  qs<HTMLButtonElement>('#btnLoadConfig').addEventListener('click', () => qs<HTMLInputElement>('#configInput').click());
  qs<HTMLInputElement>('#configInput').addEventListener('change', event => { void handleConfigFile(event); });
  qs<HTMLButtonElement>('#btnResetAll').addEventListener('click', () => { void resetAll(); });
}
function bindInput(selector: string, setter: (value: string) => void, eventName = 'input'): void {
  const element = optional<HTMLInputElement | HTMLSelectElement>(selector); if (!element) return;
  element.addEventListener(eventName, () => { setter(element.value); scheduleSave(); renderDashboard(); if (activeView === 'preview') renderPreview(); });
}
function showView(view: string): void { activeView = view; qsa<HTMLElement>('.studio-tab').forEach(tab => tab.classList.toggle('active', tab.dataset.view === view)); qsa<HTMLElement>('.studio-view').forEach(panel => panel.classList.toggle('active', panel.dataset.view === view)); if (view === 'preview') renderPreview(); window.scrollTo({ top: 0, behavior: 'smooth' }); }
function renderAll(): void { renderSetup(); renderTimetable(); renderCalendar(); renderDashboard(); if (activeView === 'preview') renderPreview(); }
function renderDashboard(): void {
  const title = state.teacherName ? `${state.year}年度 ${state.teacherName}先生の週案` : `${state.year}年度 週案`;
  const classes = state.teachingClasses.length; const weeks = computeYearWeeks(); const missing = findMissingCells().length;
  qs<HTMLElement>('#dashboardTitle').textContent = title; qs<HTMLElement>('#dashboardSummary').textContent = classes ? `担当クラス ${classes}件。生成前プレビューで ${weeks.length}週分を確認できます。` : '基本設定から担当クラスを追加してください。';
  qs<HTMLElement>('#statusGrid').innerHTML = [statusCard('担当クラス', classes, state.homeroom ? `担任 ${formatClass(state.homeroom)}` : '担任なし'), statusCard('週パターン', state.rotation.weeks.length, state.rotation.weeks.map(w => `${w}週`).join(' / ')), statusCard('生成予定', weeks.length, '授業日を含む週だけ作成'), statusCard('未入力', missing, missing ? '空欄があります' : '入力漏れなし')].join('');
}
function statusCard(label: string, value: number, note: string): string { return `<div class="status-card"><span class="status-label">${escapeText(label)}</span><span class="status-value">${escapeText(value)}</span><span class="status-note">${escapeText(note)}</span></div>`; }
function renderSetup(): void {
  qs<HTMLInputElement>('#teacherName').value = state.teacherName; qs<HTMLSelectElement>('#schoolYear').value = String(state.year); qs<HTMLInputElement>('#hasHomeroom').checked = !!state.homeroom; qs<HTMLElement>('#homeroomFields').hidden = !state.homeroom;
  if (state.homeroom) { qs<HTMLSelectElement>('#homeroomGrade').value = String(state.homeroom.grade); qs<HTMLSelectElement>('#homeroomSection').value = String(state.homeroom.section); }
  const list = qs<HTMLElement>('#classList');
  if (state.teachingClasses.length === 0) { list.innerHTML = '<p class="hint">まだ担当クラスがありません。</p>'; return; }
  list.innerHTML = state.teachingClasses.map(item => `<div class="class-row" data-id="${escapeText(item.id)}"><div><strong>${escapeText(formatClass(item))}</strong><small>${escapeText(item.subjects.join('・'))}</small></div><button class="icon-button" type="button" data-remove-class="${escapeText(item.id)}" title="削除">×</button></div>`).join('');
  qsa<HTMLButtonElement>('[data-remove-class]').forEach(button => button.addEventListener('click', () => { void removeTeachingClass(button.dataset.removeClass ?? ''); }));
}
function addTeachingClassesFromForm(): void {
  const grade = Number(qs<HTMLSelectElement>('#addGrade').value); const start = Number(qs<HTMLSelectElement>('#addSectionStart').value); const end = Number(qs<HTMLSelectElement>('#addSectionEnd').value); const subject = qs<HTMLSelectElement>('#addSubject').value;
  if (!subject) return showToast('教科を選んでください。'); if (start > end) return showToast('終了組は開始組以降にしてください。');
  let changed = 0;
  for (let section = start; section <= end; section += 1) { const id = `${grade}-${section}`; const existing = state.teachingClasses.find(item => item.id === id); if (existing) { if (!existing.subjects.includes(subject)) { existing.subjects.push(subject); changed += 1; } } else { state.teachingClasses.push({ id, grade, section, subjects: [subject] }); changed += 1; } }
  state.teachingClasses.sort((a, b) => (a.grade - b.grade) || (a.section - b.section)); renderAll(); scheduleSave(); showToast(changed ? `${changed}件を更新しました。` : 'すでに追加済みです。');
}
async function removeTeachingClass(id: string): Promise<void> {
  const modal = getModal(); if (modal && !(await modal.confirm('このクラスを削除します。時間割の該当コマも空欄になります。よろしいですか？', { danger: true }))) return;
  state.teachingClasses = state.teachingClasses.filter(item => item.id !== id); Object.values(state.timetable).forEach(week => DAYS.forEach(day => { week[day] = (week[day] ?? []).map(value => value.startsWith(`${id}|`) ? '' : value); })); renderAll(); scheduleSave();
}
function renderTimetable(): void { qs<HTMLSelectElement>('#rotationType').value = state.rotation.type; qs<HTMLElement>('#rotationStartGroup').hidden = state.rotation.type === 'single'; qs<HTMLInputElement>('#rotationStartDate').value = state.rotation.startDate; renderWeekTabs(); renderPlannerGrid(); }
function asRotationType(value: string): RotationType { return value === 'AB' || value === 'ABC' ? value : 'single'; }
function applyRotation(value: string): void { const type = asRotationType(value); const nextWeeks = ROTATION_WEEKS[type]; const nextTable = createEmptyTimetable(nextWeeks); nextWeeks.forEach(week => { if (state.timetable[week]) nextTable[week] = state.timetable[week]; }); state.rotation = { type, weeks: [...nextWeeks], startDate: state.rotation.startDate }; state.timetable = nextTable; activeWeek = nextWeeks.includes(activeWeek) ? activeWeek : (nextWeeks[0] ?? 'A'); renderTimetable(); }
function renderWeekTabs(): void { qs<HTMLElement>('#weekTabs').innerHTML = state.rotation.weeks.map(week => `<button class="week-tab ${week === activeWeek ? 'active' : ''}" data-week="${week}">${week}週</button>`).join(''); qsa<HTMLButtonElement>('.week-tab').forEach(button => button.addEventListener('click', () => { const value = button.dataset.week; if (value === 'A' || value === 'B' || value === 'B1' || value === 'B2') activeWeek = value; renderTimetable(); })); }
function buildCellOptions(): Array<{ value: string; label: string }> { const options = [{ value: '', label: '選択' }]; state.teachingClasses.forEach(item => item.subjects.forEach(subject => options.push({ value: `${item.id}|${subject}`, label: `${formatClass(item)} ${subject}` }))); if (state.homeroom) HOMEROOM_SUBJECTS.forEach(subject => options.push({ value: `${state.homeroom?.grade}-${state.homeroom?.section}|${subject}`, label: `${formatClass(state.homeroom as Homeroom)} ${subject}` })); options.push({ value: '__undecided__', label: '未定' }, { value: '__none__', label: '持ちコマなし' }); return options; }
function renderPlannerGrid(): void {
  const grid = qs<HTMLTableElement>('#plannerGrid'); const data = getWeekData(activeWeek); const options = buildCellOptions(); let html = '<thead><tr><th>時限</th>'; DAY_LABELS.forEach(label => { html += `<th>${label}</th>`; }); html += '</tr></thead><tbody>';
  for (let period = 0; period < PERIODS; period += 1) { html += `<tr><th>${period + 1}限</th>`; DAYS.forEach(day => { const value = getDayData(activeWeek, day)[period] ?? ''; const tdClass = value === '__none__' ? 'none' : value ? 'filled' : ''; html += `<td class="${tdClass}"><select data-day="${day}" data-period="${period}">${options.map(option => `<option value="${escapeText(option.value)}"${option.value === value ? ' selected' : ''}>${escapeText(option.label)}</option>`).join('')}</select></td>`; }); html += '</tr>'; }
  grid.innerHTML = `${html}</tbody>`; qsa<HTMLSelectElement>('#plannerGrid select').forEach(select => select.addEventListener('change', () => { const day = select.dataset.day; const period = Number(select.dataset.period); if (day && DAYS.includes(day as Day) && Number.isInteger(period) && period >= 0 && period < PERIODS) getDayData(activeWeek, day as Day)[period] = select.value; scheduleSave(); renderPlannerGrid(); renderDashboard(); }));
}
function fillActiveWeekNone(): void { DAYS.forEach(day => { getDayData(activeWeek, day).forEach((value, index, cells) => { cells[index] = value || '__none__'; }); }); renderAll(); scheduleSave(); }
async function clearActiveWeek(): Promise<void> { const modal = getModal(); if (modal && !(await modal.confirm(`${activeWeek}週の時間割を空欄にします。よろしいですか？`, { danger: true }))) return; DAYS.forEach(day => { getWeekData(activeWeek)[day] = new Array<string>(PERIODS).fill(''); }); renderAll(); scheduleSave(); }
function renderCalendar(): void { qs<HTMLInputElement>('#termStart').value = state.calendar.termStart; qs<HTMLInputElement>('#termEnd').value = state.calendar.termEnd; renderEvents('#holidayList', state.calendar.holidays, 'holiday'); renderEvents('#vacationList', state.calendar.vacations, 'vacation'); }
function renderEvents(selector: string, list: Holiday[] | Vacation[], type: 'holiday' | 'vacation'): void { const target = qs<HTMLElement>(selector); if (list.length === 0) { target.innerHTML = '<p class="hint">登録なし</p>'; return; } target.innerHTML = list.map((item, index) => { const dateText = type === 'holiday' ? (item as Holiday).date : `${(item as Vacation).start} 〜 ${(item as Vacation).end}`; return `<div class="event-row"><div><strong>${escapeText(item.label)}</strong><span>${escapeText(dateText)}</span></div><button class="icon-button" type="button" data-remove-event="${type}:${index}" title="削除">×</button></div>`; }).join(''); qsa<HTMLButtonElement>('[data-remove-event]').forEach(button => button.addEventListener('click', () => { const [eventType, indexText] = (button.dataset.removeEvent ?? '').split(':'); const index = Number(indexText); if (eventType === 'holiday' && Number.isInteger(index)) state.calendar.holidays.splice(index, 1); if (eventType === 'vacation' && Number.isInteger(index)) state.calendar.vacations.splice(index, 1); renderAll(); scheduleSave(); })); }
function addHoliday(): void { const date = qs<HTMLInputElement>('#holidayDate').value; const label = qs<HTMLInputElement>('#holidayLabel').value.trim() || '休業日'; if (!date) return showToast('日付を選んでください。'); state.calendar.holidays.push({ date, label }); state.calendar.holidays.sort((a, b) => a.date.localeCompare(b.date)); qs<HTMLInputElement>('#holidayDate').value = ''; qs<HTMLInputElement>('#holidayLabel').value = ''; renderAll(); scheduleSave(); }
function addVacation(): void { const label = qs<HTMLInputElement>('#vacationLabel').value.trim() || '休業'; const start = qs<HTMLInputElement>('#vacationStart').value; const end = qs<HTMLInputElement>('#vacationEnd').value; if (!start || !end) return showToast('開始日と終了日を選んでください。'); if (start > end) return showToast('終了日は開始日以降にしてください。'); state.calendar.vacations.push({ label, start, end }); state.calendar.vacations.sort((a, b) => a.start.localeCompare(b.start)); qs<HTMLInputElement>('#vacationLabel').value = ''; qs<HTMLInputElement>('#vacationStart').value = ''; qs<HTMLInputElement>('#vacationEnd').value = ''; renderAll(); scheduleSave(); }
function renderPreview(): void { const checks = buildChecks(); qs<HTMLElement>('#checkList').innerHTML = checks.map(item => `<div class="check-item ${item.level}"><strong>${escapeText(item.title)}</strong><span>${escapeText(item.body)}</span></div>`).join(''); const weeks = computeYearWeeks(); qs<HTMLElement>('#weekPreview').innerHTML = weeks.length ? weeks.map((week, index) => `<div class="week-chip"><strong>${index + 1}週目</strong><span>${escapeText(week.monday)} / ${escapeText(week.weekKey)}週</span></div>`).join('') : '<p class="hint">生成される週がありません。</p>'; }
function buildChecks(): CheckItem[] { const checks: CheckItem[] = []; if (!state.teacherName.trim()) checks.push({ level: 'error', title: '先生名が未入力', body: 'Excelファイル名と表紙に使うため、先生名を入力してください。' }); if (state.teachingClasses.length === 0) checks.push({ level: 'error', title: '担当クラスが未登録', body: '基本設定で授業担当クラスを追加してください。' }); if (!state.calendar.termStart || !state.calendar.termEnd || state.calendar.termStart > state.calendar.termEnd) checks.push({ level: 'error', title: '生成対象期間を確認', body: '開始日と終了日を正しい順序で入力してください。' }); const weeks = computeYearWeeks(); if (weeks.length === 0) checks.push({ level: 'error', title: '生成週がありません', body: '対象期間と休業期間を確認してください。' }); const missing = findMissingCells(); if (missing.length) checks.push({ level: 'warning', title: '時間割に空欄あり', body: `${missing.length}コマが空欄です。未定または持ちコマなしにすると意図が明確です。` }); const comboCount = getLessonCombos().length; if (comboCount > 17) checks.push({ level: 'warning', title: '集計欄の上限超過', body: `集計できる組み合わせは17件までです。現在 ${comboCount}件あります。` }); return checks.length ? checks : [{ level: '', title: '生成できます', body: '大きな問題は見つかりません。Excelを作成できます。' }]; }
function findMissingCells(): Array<{ week: WeekKey; day: Day; period: number }> { const missing: Array<{ week: WeekKey; day: Day; period: number }> = []; state.rotation.weeks.forEach(week => DAYS.forEach(day => { for (let period = 0; period < PERIODS; period += 1) if (!getDayData(week, day)[period]) missing.push({ week, day, period }); })); return missing; }
function parseDate(value: string): Date { return new Date(`${value}T00:00:00`); }
function formatDate(date: Date): string { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }
function addDays(date: Date, days: number): Date { const result = new Date(date); result.setDate(result.getDate() + days); return result; }
function getMondayOfWeek(date: Date): Date { const result = new Date(date); const dow = result.getDay(); result.setDate(result.getDate() + (dow === 0 ? -6 : 1 - dow)); return result; }
function computeRotationIndex(monday: Date): number { const startValue = state.rotation.startDate || formatDate(getMondayOfWeek(parseDate(state.calendar.termStart))); const diff = Math.floor((monday.getTime() - parseDate(startValue).getTime()) / 86400000); return diff < 0 ? 0 : Math.floor(diff / 7); }
function computeYearWeeks(): WeekPlan[] { if (!state.calendar.termStart || !state.calendar.termEnd || state.calendar.termStart > state.calendar.termEnd || state.rotation.weeks.length === 0) return []; const start = parseDate(state.calendar.termStart); const end = parseDate(state.calendar.termEnd); const weeks: WeekPlan[] = []; let monday = getMondayOfWeek(start); let safety = 0; while (monday <= end && safety < 80) { const weekKey = state.rotation.weeks[computeRotationIndex(monday) % state.rotation.weeks.length]; if (weekKey && hasTeachingDayInWeek(monday, start, end)) weeks.push({ monday: formatDate(monday), weekKey }); monday = addDays(monday, 7); safety += 1; } return weeks; }
function hasTeachingDayInWeek(monday: Date, termStart: Date, termEnd: Date): boolean { for (let i = 0; i < 5; i += 1) if (isTeachingDate(addDays(monday, i), termStart, termEnd)) return true; return false; }
function isTeachingDate(date: Date, termStart: Date, termEnd: Date): boolean { return date >= termStart && date <= termEnd && !state.calendar.holidays.some(item => item.date === formatDate(date)) && !state.calendar.vacations.some(item => date >= parseDate(item.start) && date <= parseDate(item.end)); }
function getLessonCombos(): LessonCombo[] { const combos: LessonCombo[] = []; state.teachingClasses.forEach(item => item.subjects.forEach(subject => combos.push({ classId: item.id, subject }))); if (state.homeroom) HOMEROOM_SUBJECTS.forEach(subject => combos.push({ classId: `${state.homeroom?.grade}-${state.homeroom?.section}`, subject })); return combos; }
function labelForCell(value: string): string { if (!value || value === '__none__') return ''; if (value === '__undecided__') return '未定'; const [classId, subject] = value.split('|'); return `${classId ?? ''} ${subject ?? ''}`.trim(); }
function safeCell(value: unknown): string { const text = String(value ?? ''); return /^[=+\-@]/.test(text) ? `'${text}` : text; }
function sanitizeFilename(value: string): string { return value.replace(/[\\/:*?"<>|]/g, '_').slice(0, 30); }
function buildExcelFilename(): string { return `週案スタジオ_${state.year}年度${state.teacherName ? `_${sanitizeFilename(state.teacherName)}` : ''}.xlsx`; }
function buildSummarySheet(wb: ExcelWorkbook): void { const ws = wb.addWorksheet('設定とプレビュー'); ws.columns = [{ width: 18 }, { width: 28 }, { width: 18 }, { width: 28 }]; ws.getCell('A1').value = '週案スタジオ 設定'; ws.getCell('A1').font = { bold: true, size: 16 }; ws.getCell('A3').value = '先生名'; ws.getCell('B3').value = safeCell(state.teacherName); ws.getCell('A4').value = '年度'; ws.getCell('B4').value = `${state.year}年度`; ws.getCell('A5').value = '生成期間'; ws.getCell('B5').value = `${state.calendar.termStart} 〜 ${state.calendar.termEnd}`; ws.getCell('A7').value = '担当クラス'; state.teachingClasses.forEach((item, index) => { ws.getCell(`A${8 + index}`).value = formatClass(item); ws.getCell(`B${8 + index}`).value = safeCell(item.subjects.join('・')); }); ws.getCell('C7').value = '生成週'; computeYearWeeks().forEach((week, index) => { ws.getCell(`C${8 + index}`).value = `${index + 1}週目`; ws.getCell(`D${8 + index}`).value = `${week.monday} / ${week.weekKey}週`; }); }
function buildWeekSheet(wb: ExcelWorkbook, week: WeekPlan, index: number): void { const ws = wb.addWorksheet(`${index + 1}週目`); ws.columns = [{ width: 8 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }]; ws.getCell('A1').value = `${index + 1}週目`; ws.getCell('B1').value = `${week.monday} / ${week.weekKey}週`; ws.getCell('A1').font = { bold: true, size: 14 }; ['時限', ...DAY_LABELS].forEach((label, col) => { const cell = ws.getCell(2, col + 1); cell.value = label; cell.font = { bold: true }; cell.alignment = { horizontal: 'center' }; }); const monday = parseDate(week.monday); for (let period = 0; period < PERIODS; period += 1) { ws.getCell(period + 3, 1).value = `${period + 1}限`; DAYS.forEach((day, dayIndex) => { const value = isTeachingDate(addDays(monday, dayIndex), parseDate(state.calendar.termStart), parseDate(state.calendar.termEnd)) ? getDayData(week.weekKey, day)[period] : '__none__'; const cell = ws.getCell(period + 3, dayIndex + 2); cell.value = safeCell(labelForCell(value ?? '')); cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; }); } }
async function generateExcel(): Promise<void> { renderPreview(); if (buildChecks().some(item => item.level === 'error')) return showToast('未入力があります。プレビューの赤い項目を確認してください。'); const excel = (window as unknown as { ExcelJS?: ExcelJsApi }).ExcelJS; const modal = getModal(); if (!excel) { if (modal) await modal.alert('Excel作成ライブラリを読み込めませんでした。ページを再読み込みしてからお試しください。'); return; } const button = qs<HTMLButtonElement>('#btnGenerateBottom'); const originalText = button.innerHTML; button.disabled = true; button.innerHTML = '<span aria-hidden="true">…</span> 作成中'; try { const wb = new excel.Workbook(); wb.creator = 'nobatasu 週案スタジオ'; wb.created = new Date(); buildSummarySheet(wb); computeYearWeeks().forEach((week, index) => buildWeekSheet(wb, week, index)); const buffer = await wb.xlsx.writeBuffer(); downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), buildExcelFilename()); state.updatedAt = new Date().toISOString(); scheduleSave(); showToast('Excelファイルを作成しました。'); } catch (error: unknown) { if (modal) await modal.alert(`Excelファイルを作成できませんでした。\n${error instanceof Error ? error.message : '不明なエラー'}`); } finally { button.disabled = false; button.innerHTML = originalText; } }
function downloadBlob(blob: Blob, filename: string): void { const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }
function saveConfigFile(): void { downloadBlob(new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' }), `週案スタジオ設定_${state.year}年度${state.teacherName ? `_${sanitizeFilename(state.teacherName)}` : ''}.json`); }
async function handleConfigFile(event: Event): Promise<void> { const input = event.currentTarget as HTMLInputElement; const file = input.files?.[0]; if (!file) return; const modal = getModal(); try { const loaded: unknown = JSON.parse(await file.text()); if (!isRecord(loaded) || loaded.version !== VERSION) throw new Error('対応していない設定ファイルです。'); state = normalizeState(loaded); activeWeek = state.rotation.weeks[0] ?? 'A'; renderAll(); scheduleSave(); showToast('設定ファイルを読み込みました。'); } catch (error: unknown) { if (modal) await modal.alert(`設定ファイルを読み込めませんでした。\n${error instanceof Error ? error.message : '不明なエラー'}`); } finally { input.value = ''; } }
function normalizeState(value: Record<string, unknown>): PlannerState {
  const base = createInitialState();
  const rotation = isRecord(value.rotation) ? value.rotation : {};
  const rotationType: RotationType = rotation.type === 'AB' || rotation.type === 'ABC' ? rotation.type : 'single';
  base.teacherName = textValue(value.teacherName).slice(0, 20);
  base.year = finiteNumber(value.year, base.year);
  const home = isRecord(value.homeroom) ? value.homeroom : null;
  base.homeroom = home ? { grade: finiteNumber(home.grade, 1), section: finiteNumber(home.section, 1) } : null;
  const classes = Array.isArray(value.teachingClasses) ? value.teachingClasses : [];
  base.teachingClasses = classes.filter(isRecord).map(item => ({ id: textValue(item.id), grade: finiteNumber(item.grade, 1), section: finiteNumber(item.section, 1), subjects: Array.isArray(item.subjects) ? item.subjects.filter((subject): subject is string => typeof subject === 'string' && subject.length > 0) : [] })).filter(item => item.id);
  base.rotation = { type: rotationType, weeks: [...ROTATION_WEEKS[rotationType]], startDate: textValue(rotation.startDate) };
  base.timetable = createEmptyTimetable(base.rotation.weeks);
  const timetable = isRecord(value.timetable) ? value.timetable : null;
  if (timetable) base.rotation.weeks.forEach(week => {
    const rawWeek = timetable[week];
    if (!isRecord(rawWeek)) return;
    DAYS.forEach(day => { const arr = rawWeek[day]; if (Array.isArray(arr)) getDayDataFromTable(base.timetable, week, day).splice(0, PERIODS, ...Array.from({ length: PERIODS }, (_, index) => textValue(arr[index]))); });
  });
  const calendar = isRecord(value.calendar) ? value.calendar : {};
  base.calendar.termStart = textValue(calendar.termStart) || base.calendar.termStart;
  base.calendar.termEnd = textValue(calendar.termEnd) || base.calendar.termEnd;
  base.calendar.holidays = Array.isArray(calendar.holidays) ? calendar.holidays.filter(isRecord).map(item => ({ date: textValue(item.date), label: textValue(item.label).slice(0, 30) })).filter(item => item.date) : [];
  base.calendar.vacations = Array.isArray(calendar.vacations) ? calendar.vacations.filter(isRecord).map(item => ({ label: textValue(item.label).slice(0, 30), start: textValue(item.start), end: textValue(item.end) })).filter(item => item.start && item.end) : [];
  return base;
}
function getDayDataFromTable(table: Record<string, TimetableWeek>, week: WeekKey, day: Day): string[] { const weekData = table[week] ?? (table[week] = {}); return weekData[day] ?? (weekData[day] = new Array<string>(PERIODS).fill('')); }
async function resetAll(): Promise<void> { const modal = getModal(); if (modal && !(await modal.confirm('入力内容をすべて消して、新規状態に戻します。よろしいですか？', { danger: true }))) return; state = createInitialState(); activeWeek = 'A'; clearSavedState(); renderAll(); showView('setup'); }
function scheduleSave(): void { state.updatedAt = new Date().toISOString(); if (saveTimer !== undefined) window.clearTimeout(saveTimer); saveTimer = window.setTimeout(saveState, 300); }
function saveState(): void { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch { /* storage unavailable */ } }
function loadSavedState(): void { try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) { const parsed: unknown = JSON.parse(raw); if (isRecord(parsed)) state = normalizeState(parsed); } } catch { state = createInitialState(); } }
function clearSavedState(): void { try { localStorage.removeItem(STORAGE_KEY); } catch { /* storage unavailable */ } }
function showToast(message: string): void { const toast = qs<HTMLElement>('#toast'); toast.textContent = message; toast.hidden = false; if (toastTimer !== undefined) window.clearTimeout(toastTimer); toastTimer = window.setTimeout(() => { toast.hidden = true; }, 2200); }
function formatClass(item: Homeroom): string { return `${item.grade}年${item.section}組`; }

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
