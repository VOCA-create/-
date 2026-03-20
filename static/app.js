// 백엔드 API 서버 주소 (Cloudflare Pages 등 정적 호스팅 시 설정 필요)
// 같은 서버에서 서비스할 경우 빈 문자열("") 유지
const API_BASE = "";

const state = {
  token: localStorage.getItem("approval_token") || "",
  me: null,
  users: [],
  selectedDocId: null,
  approverIds: [],
  referenceIds: [],
  picker: {
    target: "approver",
    selectedIds: [],
    query: "",
  },
  google: {
    config: null,
    configError: "",
    scriptsReady: false,
    pickerReady: false,
    tokenClient: null,
    accessToken: "",
    tokenExpiryMs: 0,
    initialized: false,
    initPromise: null,
  },
  ui: {
    lastFocusedBeforeDetail: null,
    dashboardPages: {
      draft: 1,
      in_review: 1,
      to_approve: 1,
      completed: 1,
    },
    dashboardSummaryMonth: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(), // 0-11
    },
    leaveCalendar: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(), // 0-11
      filter: "all", // all | mine | others
      department: "all",
    },
    overtimeDashboard: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(), // 0-11
    },
    businessTripDashboard: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(), // 0-11
    },
    educationDashboard: {
      year: new Date().getFullYear(),
      month: new Date().getMonth(), // 0-11
    },
    tripResultModal: {
      sourceDocId: null,
      sourceRow: null,
      approverIds: [],
      referenceIds: [],
    },
    educationResultModal: {
      sourceDocId: null,
      sourceRow: null,
      approverIds: [],
      referenceIds: [],
    },
    tabEdit: {
      editing: false,
      originalOrder: [],
      draggingView: "",
    },
    loginGoogleAuth: {
      initialized: false,
      clientId: "",
    },
  },
};

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

const statusLabel = {
  draft: "임시저장",
  in_review: "진행중",
  approved: "승인",
  rejected: "반려",
  waiting: "대기",
  pending: "처리필요",
  skipped: "건너뜀",
};

const editorLabel = {
  internal: "내장 편집기",
  google_docs: "Google Docs",
};

const templateTypeLabel = {
  internal: "내부문서",
  outbound: "외부발신",
  general: "일반",
  expense: "지출결의",
  leave: "휴가계",
  overtime: "연장근로",
  business_trip: "출장신청서",
  business_trip_result: "출장보고서",
  education: "교육신청서",
  education_result: "교육보고서",
  purchase: "구매품의",
};

const visibilityLabel = {
  public: "공개",
  private: "비공개",
  department: "부분공개",
};

const DEFAULT_GOOGLE_SCOPE = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/documents.readonly";
const GOOGLE_DOCS_MIME = "application/vnd.google-apps.document";
const APPROVAL_TEMPLATE_TOTAL_SLOTS = 4;
const MAX_APPROVER_STEPS = APPROVAL_TEMPLATE_TOTAL_SLOTS - 1;
const LEAVE_APPROVAL_TEMPLATE_TOTAL_SLOTS = 2;
const WORK_START_HOUR = 9;
const LUNCH_START_HOUR = 12;
const LUNCH_END_HOUR = 13;
const WORK_END_HOUR = 18;
const WORK_HOURS_PER_DAY = 8;
const GOOGLE_TEMPLATE_FOLDER_ID = "1H_CL_PkfWzVy_DXBBNkAxAariut5u0iX";
const GOOGLE_DRAFT_OUTPUT_FOLDER_ID = "15SEsuFf4kwq5G4oerk5cIUDDrgMp6gpl";
const GOOGLE_LEAVE_DRAFT_OUTPUT_FOLDER_ID = "1zBW2qkg3pwekUCg2eonxZl4CdY5LMKe0";
const GOOGLE_OVERTIME_DRAFT_OUTPUT_FOLDER_ID = "1Uz-iTjuqFMgRgUiIRBOji2esU3Obi3uu";
const GOOGLE_BUSINESS_TRIP_DRAFT_OUTPUT_FOLDER_ID = "1K0UaEg-t39aSLUaXwO4g91snBHBtdFhz";
const GOOGLE_EDUCATION_DRAFT_OUTPUT_FOLDER_ID = "1MDucU8ZCrWl9O8PEomu4ig_OoVuK6-Yv";
const GOOGLE_OUTBOUND_DRAFT_OUTPUT_FOLDER_ID = "1k2GKZ6xIS2oltLUYenkvqHdkUOPWDVpe";
const GOOGLE_ATTACHMENT_FOLDER_ID = "1BrjhQsV8kr8CGPM28KDkgzJ-IqMmzMCV";
const ISSUE_DEPARTMENT_OPTIONS = [
  "복지관",
  "함께배움팀",
  "성장이음팀",
  "건강채움팀",
  "같이돌봄팀",
  "내일이룸팀",
  "미래그림팀",
  "행복동네팀",
  "일상동행팀",
  "누구나운동센터",
];
const DASHBOARD_LIST_PAGE_SIZE = 5;
const TAB_ORDER_DEFAULT = ["dashboard", "leaveMgmt", "overtimeMgmt", "tripMgmt", "educationMgmt", "newDoc", "myDocs", "admin"];

async function api(path, options = {}) {
  const headers = { "Content-Type": "application/json", ...(options.headers || {}) };
  if (state.token) headers.Authorization = `Bearer ${state.token}`;
  const resp = await fetch(API_BASE + path, { ...options, headers });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.error || `요청 실패 (${resp.status})`);
  }
  return data;
}

function fmt(value) {
  if (!value) return "-";
  return String(value).replace("T", " ");
}

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function applyDraftIssueDefaults(force = false) {
  const form = $("#draftForm");
  if (!form) return;
  const yearInput = form.issue_year;
  const deptSelect = form.issue_department;

  if (yearInput && (force || !String(yearInput.value || "").trim())) {
    yearInput.value = String(new Date().getFullYear());
  }
  if (deptSelect && state.me) {
    const myDept = String(state.me.department || "").trim();
    if ((force || !String(deptSelect.value || "").trim()) && myDept && ISSUE_DEPARTMENT_OPTIONS.includes(myDept)) {
      deptSelect.value = myDept;
    }
  }
}

function getDraftApprovalTemplateTotalSlots() {
  const templateType = ($("#draftForm [name='template_type']")?.value || "").trim();
  return templateType === "leave" ? LEAVE_APPROVAL_TEMPLATE_TOTAL_SLOTS : APPROVAL_TEMPLATE_TOTAL_SLOTS;
}

function getDraftMaxApproverSteps() {
  return Math.max(1, getDraftApprovalTemplateTotalSlots() - 1);
}

function isElementVisible(el) {
  return !!(el && !el.classList.contains("hidden"));
}

function syncBodyModalOpenClass() {
  const hasOpenModal =
    isElementVisible($("#detailPanel")) ||
    isElementVisible($("#userPickerModal")) ||
    isElementVisible($("#leaveInfoModal")) ||
    isElementVisible($("#overtimeInfoModal")) ||
    isElementVisible($("#businessTripInfoModal")) ||
    isElementVisible($("#educationInfoModal")) ||
    isElementVisible($("#tripResultModal")) ||
    isElementVisible($("#educationResultModal"));
  document.body.classList.toggle("modal-open", hasOpenModal);
}

function ensureTripResultModalMountedToBody() {
  const modal = $("#tripResultModal");
  if (!modal) return;
  if (modal.parentElement !== document.body) {
    document.body.appendChild(modal);
  }
}

function ensureEducationResultModalMountedToBody() {
  const modal = $("#educationResultModal");
  if (!modal) return;
  if (modal.parentElement !== document.body) {
    document.body.appendChild(modal);
  }
}

function ensureUserPickerModalMountedToBody() {
  const modal = $("#userPickerModal");
  if (!modal) return;
  if (modal.parentElement !== document.body) {
    document.body.appendChild(modal);
  }
}

function parseLocalDateTimeInput(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const dt = new Date(raw);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function toDateTimeLocalValue(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function formatLocalDateTimeDisplay(value) {
  const dt = parseLocalDateTimeInput(value);
  if (!dt) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())} ${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
}

function fmtDateOnly(value) {
  if (!value) return "";
  return String(value).slice(0, 10);
}

function parseLeaveRecordStart(record) {
  return parseLocalDateTimeInput(record?.start_date) || (record?.start_date ? new Date(record.start_date) : null);
}

function parseLeaveRecordEnd(record) {
  return parseLocalDateTimeInput(record?.end_date) || (record?.end_date ? new Date(record.end_date) : null);
}

function toYmdKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function eachDateInclusive(startDate, endDate, fn) {
  let cur = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const last = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (cur <= last) {
    fn(new Date(cur));
    cur.setDate(cur.getDate() + 1);
  }
}

function getLeaveCalendarMonthLabel(year, month) {
  return `${year}년 ${month + 1}월`;
}

function normalizeLeaveTypeToken(value) {
  const raw = String(value || "").trim();
  if (!raw) return "unknown";
  if (raw === "연차") return "annual";
  if (raw === "오전반차") return "half-am";
  if (raw === "오후반차") return "half-pm";
  if (raw === "병가") return "sick";
  if (raw === "공가") return "official";
  if (raw === "특별휴가") return "special";
  return "other";
}

function buildLeaveCalendarEventMap(myRecords, otherRecords) {
  const map = new Map();
  const addRecord = (record, ownerType) => {
    const start = parseLeaveRecordStart(record);
    const end = parseLeaveRecordEnd(record);
    if (!start || !end) return;
    const safeEnd = end < start ? start : end;
    eachDateInclusive(start, safeEnd, (day) => {
      const key = toYmdKey(day);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push({
        id: record.document_id,
        ownerType,
        name: record.user?.name || "-",
        department: record.user?.department || "-",
        leaveType: String(record.leave_type || "").trim(),
        leaveTypeKey: normalizeLeaveTypeToken(record.leave_type),
        title: record.document_title || "-",
        leaveDays: record.leave_days,
        status: record.document_status || "",
        startDate: record.start_date || "",
        endDate: record.end_date || "",
      });
    });
  };
  (myRecords || []).forEach((r) => addRecord(r, "mine"));
  (otherRecords || []).forEach((r) => addRecord(r, "others"));
  return map;
}

function renderLeaveCalendarPanel(data) {
  const myRecords = data.my_records || [];
  const otherRecords = data.other_records || [];
  const stateCal = state.ui.leaveCalendar || (state.ui.leaveCalendar = {
    year: new Date().getFullYear(),
    month: new Date().getMonth(),
    filter: "all",
  });
  const { year, month, filter } = stateCal;
  const selectedDepartment = String(stateCal.department || "all");
  const eventMap = buildLeaveCalendarEventMap(myRecords, otherRecords);
  const departments = Array.from(new Set(
    [...myRecords, ...otherRecords]
      .map((r) => String(r?.user?.department || "").trim())
      .filter(Boolean)
  )).sort((a, b) => a.localeCompare(b, "ko"));

  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const startWeekday = firstDay.getDay(); // 0 Sun
  const cellStart = addDays(firstDay, -startWeekday);
  const todayKey = toYmdKey(new Date());
  const weekdayLabels = ["일", "월", "화", "수", "목", "금", "토"];

  let cellsHtml = "";
  for (let i = 0; i < 42; i++) {
    const cellDate = addDays(cellStart, i);
    const key = toYmdKey(cellDate);
    const inMonth = cellDate.getMonth() === month;
    const rawEvents = eventMap.get(key) || [];
    const events = rawEvents.filter((e) => {
      if (filter === "mine") return e.ownerType === "mine";
      if (filter === "others") return e.ownerType === "others";
      return true;
    }).filter((e) => selectedDepartment === "all" || e.department === selectedDepartment);

    const classes = [
      "leave-cal-cell",
      inMonth ? "" : "is-other-month",
      key === todayKey ? "is-today" : "",
      events.length ? "has-event" : "",
    ].filter(Boolean).join(" ");

    const chips = events.slice(0, 2).map((e) => {
      const ownerCls = e.ownerType === "mine" ? "mine" : "others";
      const typeCls = `type-${e.leaveTypeKey || "unknown"}`;
      const typeText = e.leaveType || "휴가";
      const text = e.ownerType === "mine" ? `${typeText} · ${e.title}` : `${e.name} · ${typeText}`;
      const label = e.ownerType === "mine" ? `${typeText}` : `${e.name}`;
      return `<button type="button" class="leave-cal-chip ${ownerCls} ${typeCls}" data-leave-doc-id="${e.id}" title="${esc(text)}">${esc(label)}</button>`;
    }).join("");

    const extra = events.length > 2 ? `<div class="leave-cal-more">+${events.length - 2}</div>` : "";

    cellsHtml += `
      <div class="${classes}" data-date="${key}">
        <div class="leave-cal-date">${cellDate.getDate()}</div>
        <div class="leave-cal-events">${chips}${extra}</div>
      </div>
    `;
  }

  return `
    <div class="leave-calendar-card">
        <div class="leave-calendar-head">
          <div class="leave-calendar-nav">
          <button type="button" class="btn" data-leave-cal-nav="-1">이전</button>
          <strong>${getLeaveCalendarMonthLabel(year, month)}</strong>
          <button type="button" class="btn" data-leave-cal-nav="1">다음</button>
          <button type="button" class="btn" data-leave-cal-nav="0">이번달</button>
        </div>
          <div class="leave-cal-filter" role="tablist" aria-label="휴가 캘린더 필터">
            <button type="button" class="doc-filter-btn ${filter === "all" ? "is-active" : ""}" data-leave-cal-filter="all">전체</button>
            <button type="button" class="doc-filter-btn ${filter === "mine" ? "is-active" : ""}" data-leave-cal-filter="mine">내 휴가</button>
            <button type="button" class="doc-filter-btn ${filter === "others" ? "is-active" : ""}" data-leave-cal-filter="others">타인 휴가</button>
          </div>
        </div>
        <div class="leave-cal-toolbar">
          <label>부서 필터
            <select id="leaveCalDeptFilter">
              <option value="all">전체 부서</option>
              ${departments.map((dept) => `<option value="${esc(dept)}" ${selectedDepartment === dept ? "selected" : ""}>${esc(dept)}</option>`).join("")}
            </select>
          </label>
        </div>
        <div class="leave-cal-legend">
          <span class="leave-cal-legend-item"><span class="dot type-annual"></span>연차</span>
          <span class="leave-cal-legend-item"><span class="dot type-half-am"></span>오전반차</span>
          <span class="leave-cal-legend-item"><span class="dot type-half-pm"></span>오후반차</span>
          <span class="leave-cal-legend-item"><span class="dot type-sick"></span>병가</span>
          <span class="leave-cal-legend-item"><span class="dot type-official"></span>공가</span>
          <span class="leave-cal-legend-item"><span class="dot type-special"></span>특별휴가</span>
          <span class="leave-cal-legend-item"><span class="dot type-other"></span>기타</span>
        </div>
        <div class="leave-cal-grid">
        ${weekdayLabels.map((w) => `<div class="leave-cal-weekday">${w}</div>`).join("")}
        ${cellsHtml}
      </div>
    </div>
  `;
}

function overlapMs(startA, endA, startB, endB) {
  const s = Math.max(startA.getTime(), startB.getTime());
  const e = Math.min(endA.getTime(), endB.getTime());
  return Math.max(0, e - s);
}

function calculateLeaveDaysByWorkHours(startValue, endValue) {
  const start = parseLocalDateTimeInput(startValue);
  const end = parseLocalDateTimeInput(endValue);
  if (!start || !end || end <= start) return null;

  let totalMs = 0;
  const dayCursor = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const lastDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());

  while (dayCursor <= lastDay) {
    const y = dayCursor.getFullYear();
    const m = dayCursor.getMonth();
    const d = dayCursor.getDate();
    const morningStart = new Date(y, m, d, WORK_START_HOUR, 0, 0, 0);
    const morningEnd = new Date(y, m, d, LUNCH_START_HOUR, 0, 0, 0);
    const afternoonStart = new Date(y, m, d, LUNCH_END_HOUR, 0, 0, 0);
    const afternoonEnd = new Date(y, m, d, WORK_END_HOUR, 0, 0, 0);
    totalMs += overlapMs(start, end, morningStart, morningEnd);
    totalMs += overlapMs(start, end, afternoonStart, afternoonEnd);
    dayCursor.setDate(dayCursor.getDate() + 1);
  }

  const hours = totalMs / (1000 * 60 * 60);
  if (hours <= 0) return 0;
  return Math.round((hours / WORK_HOURS_PER_DAY) * 100) / 100;
}

function syncLeaveDaysAutoCalc() {
  const form = $("#draftForm");
  if (!form) return;
  if ((form.template_type?.value || "") !== "leave") {
    renderLeaveInfoSummary();
    return;
  }
  const startValue = form.leave_start_date?.value || "";
  const endValue = form.leave_end_date?.value || "";
  const days = calculateLeaveDaysByWorkHours(startValue, endValue);
  if (days === null) {
    if (form.leave_days) form.leave_days.value = "";
    renderLeaveInfoSummary();
    return;
  }
  if (form.leave_days) form.leave_days.value = String(days);
  renderLeaveInfoSummary();
}

function leaveInfoHasValue(form = $("#draftForm")) {
  if (!form) return false;
  return [
    form.leave_type?.value,
    form.leave_start_date?.value,
    form.leave_end_date?.value,
    form.leave_days?.value,
    form.leave_reason?.value,
    form.leave_substitute_name?.value,
    form.leave_substitute_work?.value,
  ].some((v) => String(v || "").trim());
}

function renderLeaveInfoSummary() {
  const form = $("#draftForm");
  const summaryEl = $("#leaveInfoSummary");
  const metaEl = $("#leaveInfoSummaryMeta");
  const openBtn = $("#openLeaveInfoModalBtn");
  const resetBtn = $("#resetLeaveInfoBtn");
  if (!form || !summaryEl || !metaEl) return;

  const hasValue = leaveInfoHasValue(form);
  if (openBtn) openBtn.textContent = hasValue ? "휴가정보 수정" : "휴가정보 입력";
  if (resetBtn) resetBtn.classList.toggle("hidden", !hasValue);

  if (!hasValue) {
    summaryEl.textContent = "휴가정보가 아직 입력되지 않았습니다. '휴가정보 입력' 버튼으로 등록하세요.";
    metaEl.classList.add("hidden");
    metaEl.innerHTML = "";
    return;
  }

  const leaveType = String(form.leave_type?.value || "").trim();
  const startText = formatLocalDateTimeDisplay(form.leave_start_date?.value || "");
  const endText = formatLocalDateTimeDisplay(form.leave_end_date?.value || "");
  const leaveDays = String(form.leave_days?.value || "").trim();
  const substitute = String(form.leave_substitute_name?.value || "").trim();
  const reason = String(form.leave_reason?.value || "").trim();

  const headline = [];
  if (leaveType) headline.push(leaveType);
  if (startText || endText) headline.push(`${startText || "-"} ~ ${endText || "-"}`);
  if (leaveDays) headline.push(`${leaveDays}일`);
  summaryEl.textContent = headline.length ? headline.join(" | ") : "휴가정보 입력값이 있습니다.";

  const badges = [];
  if (substitute) badges.push(`대직자: ${substitute}`);
  if (reason) badges.push(`사유: ${reason.length > 24 ? `${reason.slice(0, 24)}...` : reason}`);
  if (!badges.length) badges.push("상세 사유/대직자 정보는 모달에서 확인");
  metaEl.innerHTML = badges.map((t) => `<span class="pill">${esc(t)}</span>`).join("");
  metaEl.classList.remove("hidden");
}

function openLeaveInfoModal() {
  const form = $("#draftForm");
  const modal = $("#leaveInfoModal");
  if (!form || !modal) return;
  if ((form.template_type?.value || "") !== "leave") return;
  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  const focusTarget = modal.querySelector("[name='leave_type']") || $("#saveLeaveInfoModalBtn");
  if (focusTarget && typeof focusTarget.focus === "function") {
    setTimeout(() => {
      try { focusTarget.focus(); } catch (_) { }
    }, 0);
  }
}

function closeLeaveInfoModal() {
  const modal = $("#leaveInfoModal");
  if (!modal) return;
  modal.classList.add("hidden");
  renderLeaveInfoSummary();
  syncBodyModalOpenClass();
}

function resetLeaveInfoInputs() {
  const form = $("#draftForm");
  if (!form) return;
  const fields = [
    "leave_type",
    "leave_substitute_name",
    "leave_start_date",
    "leave_end_date",
    "leave_days",
    "leave_reason",
    "leave_substitute_work",
  ];
  for (const name of fields) {
    const el = form[name];
    if (!el) continue;
    el.value = "";
  }
  syncLeaveDaysAutoCalc();
  renderLeaveInfoSummary();
}

function applyLeaveTimePreset(startHour, endHour, presetType = "") {
  const form = $("#draftForm");
  if (!form) return;
  const base =
    parseLocalDateTimeInput(form.leave_start_date?.value || "") ||
    parseLocalDateTimeInput(form.leave_end_date?.value || "") ||
    (form.due_date?.value ? new Date(`${form.due_date.value}T09:00`) : new Date());
  if (Number.isNaN(base.getTime())) return;
  const y = base.getFullYear();
  const m = base.getMonth();
  const d = base.getDate();
  const start = new Date(y, m, d, startHour, 0, 0, 0);
  const end = new Date(y, m, d, endHour, 0, 0, 0);
  if (form.leave_start_date) form.leave_start_date.value = toDateTimeLocalValue(start);
  if (form.leave_end_date) form.leave_end_date.value = toDateTimeLocalValue(end);
  if (presetType && form.leave_type && !String(form.leave_type.value || "").trim()) {
    form.leave_type.value = presetType;
  }
  syncLeaveDaysAutoCalc();
}

function calculateOvertimeHoursByDatetimeRange(startValue, endValue) {
  const start = parseLocalDateTimeInput(startValue);
  const end = parseLocalDateTimeInput(endValue);
  if (!start || !end || end <= start) return null;
  const hours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
  if (hours <= 0) return 0;
  return Math.round(hours * 100) / 100;
}

function syncOvertimeHoursAutoCalc() {
  const form = $("#draftForm");
  if (!form) return;
  if ((form.template_type?.value || "") !== "overtime") {
    renderOvertimeInfoSummary();
    return;
  }
  const startValue = form.overtime_start_date?.value || "";
  const endValue = form.overtime_end_date?.value || "";
  const hours = calculateOvertimeHoursByDatetimeRange(startValue, endValue);
  if (hours === null) {
    if (form.overtime_hours) form.overtime_hours.value = "";
    renderOvertimeInfoSummary();
    return;
  }
  if (form.overtime_hours) form.overtime_hours.value = String(hours);
  renderOvertimeInfoSummary();
}

function overtimeInfoHasValue(form = $("#draftForm")) {
  if (!form) return false;
  return [
    form.overtime_type?.value,
    form.overtime_start_date?.value,
    form.overtime_end_date?.value,
    form.overtime_hours?.value,
    form.overtime_content?.value,
    form.overtime_etc?.value,
  ].some((v) => String(v || "").trim());
}

function renderOvertimeInfoSummary() {
  const form = $("#draftForm");
  const summaryEl = $("#overtimeInfoSummary");
  const metaEl = $("#overtimeInfoSummaryMeta");
  const openBtn = $("#openOvertimeInfoModalBtn");
  const resetBtn = $("#resetOvertimeInfoBtn");
  if (!form || !summaryEl || !metaEl) return;

  const hasValue = overtimeInfoHasValue(form);
  if (openBtn) openBtn.textContent = hasValue ? "연장근로정보 수정" : "연장근로정보 입력";
  if (resetBtn) resetBtn.classList.toggle("hidden", !hasValue);

  if (!hasValue) {
    summaryEl.textContent = "연장근로 정보가 아직 입력되지 않았습니다. '연장근로정보 입력' 버튼으로 등록하세요.";
    metaEl.classList.add("hidden");
    metaEl.innerHTML = "";
    return;
  }

  const type = String(form.overtime_type?.value || "").trim();
  const startText = formatLocalDateTimeDisplay(form.overtime_start_date?.value || "");
  const endText = formatLocalDateTimeDisplay(form.overtime_end_date?.value || "");
  const hours = String(form.overtime_hours?.value || "").trim();
  const content = String(form.overtime_content?.value || "").trim();
  const etc = String(form.overtime_etc?.value || "").trim();

  const headline = [];
  if (type) headline.push(type);
  if (startText || endText) headline.push(`${startText || "-"} ~ ${endText || "-"}`);
  if (hours) headline.push(`${hours}시간`);
  summaryEl.textContent = headline.length ? headline.join(" | ") : "연장근로 정보 입력값이 있습니다.";

  const badges = [];
  if (content) badges.push(`내용: ${content.length > 24 ? `${content.slice(0, 24)}...` : content}`);
  if (etc) badges.push(`기타: ${etc.length > 24 ? `${etc.slice(0, 24)}...` : etc}`);
  if (!badges.length) badges.push("상세 내용은 모달에서 확인");
  metaEl.innerHTML = badges.map((t) => `<span class=\"pill\">${esc(t)}</span>`).join("");
  metaEl.classList.remove("hidden");
}

function openOvertimeInfoModal() {
  const form = $("#draftForm");
  const modal = $("#overtimeInfoModal");
  if (!form || !modal) return;
  if ((form.template_type?.value || "") !== "overtime") return;
  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  const focusTarget = modal.querySelector("[name='overtime_type']") || $("#saveOvertimeInfoModalBtn");
  if (focusTarget && typeof focusTarget.focus === "function") {
    setTimeout(() => {
      try { focusTarget.focus(); } catch (_) { }
    }, 0);
  }
}

function closeOvertimeInfoModal() {
  const modal = $("#overtimeInfoModal");
  if (!modal) return;
  modal.classList.add("hidden");
  renderOvertimeInfoSummary();
  syncBodyModalOpenClass();
}

function resetOvertimeInfoInputs() {
  const form = $("#draftForm");
  if (!form) return;
  const fields = [
    "overtime_type",
    "overtime_start_date",
    "overtime_end_date",
    "overtime_hours",
    "overtime_content",
    "overtime_etc",
  ];
  for (const name of fields) {
    const el = form[name];
    if (!el) continue;
    el.value = "";
  }
  syncOvertimeHoursAutoCalc();
  renderOvertimeInfoSummary();
}

function businessTripInfoHasValue(form = $("#draftForm")) {
  if (!form) return false;
  return [
    form.trip_department?.value,
    form.trip_job_title?.value,
    form.trip_name?.value,
    form.trip_type?.value,
    form.trip_destination?.value,
    form.trip_start_date?.value,
    form.trip_end_date?.value,
    form.trip_transportation?.value,
    form.trip_expense?.value,
    form.trip_purpose?.value,
  ].some((v) => String(v || "").trim());
}

function renderBusinessTripInfoSummary() {
  const form = $("#draftForm");
  const summaryEl = $("#businessTripInfoSummary");
  const metaEl = $("#businessTripInfoSummaryMeta");
  const openBtn = $("#openBusinessTripInfoModalBtn");
  const resetBtn = $("#resetBusinessTripInfoBtn");
  if (!form || !summaryEl || !metaEl) return;

  const hasValue = businessTripInfoHasValue(form);
  if (openBtn) openBtn.textContent = hasValue ? "출장정보 수정" : "출장정보 입력";
  if (resetBtn) resetBtn.classList.toggle("hidden", !hasValue);

  if (!hasValue) {
    summaryEl.textContent = "출장 정보가 아직 입력되지 않았습니다. '출장정보 입력' 버튼으로 등록하세요.";
    metaEl.classList.add("hidden");
    metaEl.innerHTML = "";
    return;
  }

  const tripType = String(form.trip_type?.value || "").trim();
  const startText = formatLocalDateTimeDisplay(form.trip_start_date?.value || "");
  const endText = formatLocalDateTimeDisplay(form.trip_end_date?.value || "");
  const destination = String(form.trip_destination?.value || "").trim();
  const purpose = String(form.trip_purpose?.value || "").trim();

  const headline = [];
  if (tripType) headline.push(tripType);
  if (startText || endText) headline.push(`${startText || "-"} ~ ${endText || "-"}`);
  if (destination) headline.push(destination);
  summaryEl.textContent = headline.length ? headline.join(" | ") : "출장정보 입력값이 있습니다.";

  const badges = [];
  const dept = String(form.trip_department?.value || "").trim();
  const name = String(form.trip_name?.value || "").trim();
  const transportation = String(form.trip_transportation?.value || "").trim();
  const expense = String(form.trip_expense?.value || "").trim();
  if (dept || name) badges.push(`신청자: ${(dept || "-")} / ${(name || "-")}`);
  if (transportation) badges.push(`교통수단: ${transportation}`);
  if (expense) badges.push(`출장비: ${expense}`);
  if (purpose) badges.push(`목적: ${purpose.length > 24 ? `${purpose.slice(0, 24)}...` : purpose}`);
  if (!badges.length) badges.push("상세 내용은 모달에서 확인");
  metaEl.innerHTML = badges.map((t) => `<span class="pill">${esc(t)}</span>`).join("");
  metaEl.classList.remove("hidden");
}

function openBusinessTripInfoModal() {
  const form = $("#draftForm");
  const modal = $("#businessTripInfoModal");
  if (!form || !modal) return;
  if ((form.template_type?.value || "") !== "business_trip") return;

  if (state.me) {
    if (form.trip_department && !String(form.trip_department.value || "").trim()) {
      form.trip_department.value = String(state.me.department || "").trim();
    }
    if (form.trip_name && !String(form.trip_name.value || "").trim()) {
      form.trip_name.value = String(state.me.full_name || "").trim();
    }
    if (form.trip_job_title && !String(form.trip_job_title.value || "").trim()) {
      form.trip_job_title.value = String(state.me.job_title || "").trim();
    }
  }

  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  const focusTarget = modal.querySelector("[name='trip_type']") || $("#saveBusinessTripInfoModalBtn");
  if (focusTarget && typeof focusTarget.focus === "function") {
    setTimeout(() => {
      try { focusTarget.focus(); } catch (_) { }
    }, 0);
  }
}

function closeBusinessTripInfoModal() {
  const modal = $("#businessTripInfoModal");
  if (!modal) return;
  modal.classList.add("hidden");
  renderBusinessTripInfoSummary();
  syncBodyModalOpenClass();
}

function resetBusinessTripInfoInputs() {
  const form = $("#draftForm");
  if (!form) return;
  const fields = [
    "trip_department",
    "trip_job_title",
    "trip_name",
    "trip_type",
    "trip_destination",
    "trip_start_date",
    "trip_end_date",
    "trip_transportation",
    "trip_expense",
    "trip_purpose",
  ];
  for (const name of fields) {
    const el = form[name];
    if (!el) continue;
    el.value = "";
  }
  renderBusinessTripInfoSummary();
}

function parseEducationMoneyInput(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  const normalized = text.replaceAll(",", "").replaceAll("원", "").trim();
  const num = Number(normalized);
  if (!Number.isFinite(num)) return null;
  return Math.round(num * 100) / 100;
}

function formatEducationMoneyDisplay(raw) {
  const num = parseEducationMoneyInput(raw);
  if (num === null) return "";
  if (Number.isInteger(num)) return `${num.toLocaleString("ko-KR")}원`;
  return `${num.toLocaleString("ko-KR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}원`;
}

function educationInfoHasValue(form = $("#draftForm")) {
  if (!form) return false;
  return [
    form.education_department?.value,
    form.education_job_title?.value,
    form.education_name?.value,
    form.education_title?.value,
    form.education_category?.value,
    form.education_provider?.value,
    form.education_location?.value,
    form.education_start_date?.value,
    form.education_end_date?.value,
    form.education_purpose?.value,
    form.education_tuition_detail?.value,
    form.education_tuition_amount?.value,
    form.education_material_detail?.value,
    form.education_material_amount?.value,
    form.education_transport_detail?.value,
    form.education_transport_amount?.value,
    form.education_other_detail?.value,
    form.education_other_amount?.value,
    form.education_budget_subject?.value,
    form.education_funding_source?.value,
    form.education_payment_method?.value,
    form.education_support_budget?.value,
    form.education_used_budget?.value,
    form.education_remain_budget?.value,
    form.education_companion?.value,
    form.education_ordered?.value,
    form.education_suggestion?.value,
  ].some((v) => String(v || "").trim());
}

function renderEducationInfoSummary() {
  const form = $("#draftForm");
  const summaryEl = $("#educationInfoSummary");
  const metaEl = $("#educationInfoSummaryMeta");
  const openBtn = $("#openEducationInfoModalBtn");
  const resetBtn = $("#resetEducationInfoBtn");
  if (!form || !summaryEl || !metaEl) return;

  const hasValue = educationInfoHasValue(form);
  if (openBtn) openBtn.textContent = hasValue ? "교육정보 수정" : "교육정보 입력";
  if (resetBtn) resetBtn.classList.toggle("hidden", !hasValue);

  if (!hasValue) {
    summaryEl.textContent = "교육 정보가 아직 입력되지 않았습니다. '교육정보 입력' 버튼으로 등록하세요.";
    metaEl.classList.add("hidden");
    metaEl.innerHTML = "";
    return;
  }

  const title = String(form.education_title?.value || "").trim();
  const category = String(form.education_category?.value || "").trim();
  const provider = String(form.education_provider?.value || "").trim();
  const startText = formatLocalDateTimeDisplay(form.education_start_date?.value || "");
  const endText = formatLocalDateTimeDisplay(form.education_end_date?.value || "");

  const headline = [];
  if (title) headline.push(title);
  if (category) headline.push(category);
  if (startText || endText) headline.push(`${startText || "-"} ~ ${endText || "-"}`);
  if (provider) headline.push(provider);
  summaryEl.textContent = headline.length ? headline.join(" | ") : "교육정보 입력값이 있습니다.";

  const badges = [];
  const dept = String(form.education_department?.value || "").trim();
  const name = String(form.education_name?.value || "").trim();
  const location = String(form.education_location?.value || "").trim();
  const purpose = String(form.education_purpose?.value || "").trim();
  const paymentMethod = String(form.education_payment_method?.value || "").trim();
  const totalAmount =
    (parseEducationMoneyInput(form.education_tuition_amount?.value || "") || 0) +
    (parseEducationMoneyInput(form.education_material_amount?.value || "") || 0) +
    (parseEducationMoneyInput(form.education_transport_amount?.value || "") || 0) +
    (parseEducationMoneyInput(form.education_other_amount?.value || "") || 0);
  if (dept || name) badges.push(`신청자: ${(dept || "-")} / ${(name || "-")}`);
  if (location) badges.push(`장소: ${location}`);
  if (totalAmount > 0) badges.push(`총 교육비: ${formatEducationMoneyDisplay(totalAmount)}`);
  if (paymentMethod) badges.push(`결재방법: ${paymentMethod}`);
  if (purpose) badges.push(`목적: ${purpose.length > 24 ? `${purpose.slice(0, 24)}...` : purpose}`);
  if (!badges.length) badges.push("상세 내용은 모달에서 확인");
  metaEl.innerHTML = badges.map((t) => `<span class="pill">${esc(t)}</span>`).join("");
  metaEl.classList.remove("hidden");
}

function openEducationInfoModal() {
  const form = $("#draftForm");
  const modal = $("#educationInfoModal");
  if (!form || !modal) return;
  if ((form.template_type?.value || "") !== "education") return;

  if (state.me) {
    if (form.education_department && !String(form.education_department.value || "").trim()) {
      form.education_department.value = String(state.me.department || "").trim();
    }
    if (form.education_name && !String(form.education_name.value || "").trim()) {
      form.education_name.value = String(state.me.full_name || "").trim();
    }
    if (form.education_job_title && !String(form.education_job_title.value || "").trim()) {
      form.education_job_title.value = String(state.me.job_title || "").trim();
    }
  }

  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  const focusTarget = modal.querySelector("[name='education_title']") || $("#saveEducationInfoModalBtn");
  if (focusTarget && typeof focusTarget.focus === "function") {
    setTimeout(() => {
      try { focusTarget.focus(); } catch (_) { }
    }, 0);
  }
}

function closeEducationInfoModal() {
  const modal = $("#educationInfoModal");
  if (!modal) return;
  modal.classList.add("hidden");
  renderEducationInfoSummary();
  syncBodyModalOpenClass();
}

function resetEducationInfoInputs() {
  const form = $("#draftForm");
  if (!form) return;
  const fields = [
    "education_department",
    "education_job_title",
    "education_name",
    "education_title",
    "education_category",
    "education_provider",
    "education_location",
    "education_start_date",
    "education_end_date",
    "education_purpose",
    "education_tuition_detail",
    "education_tuition_amount",
    "education_material_detail",
    "education_material_amount",
    "education_transport_detail",
    "education_transport_amount",
    "education_other_detail",
    "education_other_amount",
    "education_budget_subject",
    "education_funding_source",
    "education_payment_method",
    "education_support_budget",
    "education_used_budget",
    "education_remain_budget",
    "education_companion",
    "education_ordered",
    "education_suggestion",
  ];
  for (const name of fields) {
    const el = form[name];
    if (!el) continue;
    el.value = "";
  }
  renderEducationInfoSummary();
}

function setTripResultModalMessage(message, isError = false) {
  const msgEl = $("#tripResultModalMsg");
  if (!msgEl) return;
  msgEl.textContent = String(message || "");
  msgEl.style.color = isError ? "var(--danger)" : "var(--muted)";
}

function closeTripResultModal() {
  const modal = $("#tripResultModal");
  if (!modal) return;
  modal.classList.add("hidden");
  const summaryEl = $("#tripResultModalSummary");
  const textEl = $("#tripResultTextInput");
  if (summaryEl) summaryEl.textContent = "대상 출장문서 정보를 불러오는 중입니다.";
  if (textEl) textEl.value = "";
  setTripResultModalMessage("");
  state.ui.tripResultModal = {
    sourceDocId: null,
    sourceRow: null,
    approverIds: [],
    referenceIds: [],
  };
  renderTripResultAssigneeLists();
  syncBodyModalOpenClass();
}

async function openTripResultModal(rowData) {
  const modal = $("#tripResultModal");
  const summaryEl = $("#tripResultModalSummary");
  const textEl = $("#tripResultTextInput");
  if (!modal || !summaryEl || !textEl) return;

  const sourceDocId = Number(rowData?.document_id || 0);
  if (!sourceDocId) {
    setTripResultModalMessage("대상 출장문서를 찾지 못했습니다.", true);
    return;
  }

  if ((!state.users || !state.users.length) && state.token) {
    try {
      await loadUsers();
    } catch (_) {
      // 사용자 목록 조회 실패 시에도 모달은 열어두고 수동 재시도 가능하게 둔다.
    }
  }

  const candidates = (state.users || []).filter((u) => !state.me || Number(u.id) !== Number(state.me.id));

  const statusText = statusLabel[rowData.document_status] || rowData.document_status || "-";
  const periodText = `${fmt(rowData.trip_start_date)} ~ ${fmt(rowData.trip_end_date)}`;
  const resultCount = Number(rowData.result_doc_count || 0);
  summaryEl.textContent = `#${sourceDocId} ${rowData.document_title || "-"} | 상태: ${statusText} | 출장기간: ${periodText} | 기존 출장결과 ${resultCount}건`;
  textEl.value = "";
  setTripResultModalMessage(candidates.length ? "" : "선택 가능한 결재자가 없습니다.", !candidates.length);

  state.ui.tripResultModal = {
    sourceDocId,
    sourceRow: rowData || null,
    approverIds: [],
    referenceIds: [],
  };
  renderTripResultAssigneeLists();
  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  setTimeout(() => {
    try { textEl.focus(); } catch (_) { }
  }, 0);
}

async function submitTripResultModal() {
  const textEl = $("#tripResultTextInput");
  const submitBtn = $("#submitTripResultBtn");
  const cancelBtn = $("#cancelTripResultModalBtn");
  const closeBtn = $("#closeTripResultModalBtn");
  const sourceDocId = Number(state.ui?.tripResultModal?.sourceDocId || 0);
  if (!textEl || !submitBtn || !cancelBtn || !closeBtn) return;
  if (!sourceDocId) {
    setTripResultModalMessage("대상 출장문서 정보가 없습니다. 다시 시도해 주세요.", true);
    return;
  }

  const tripResult = String(textEl.value || "").trim();
  const approverIds = [...(state.ui?.tripResultModal?.approverIds || [])];
  const referenceIds = [...(state.ui?.tripResultModal?.referenceIds || [])];
  if (!tripResult) {
    setTripResultModalMessage("출장결과 내용을 입력해 주세요.", true);
    textEl.focus();
    return;
  }
  if (!approverIds.length) {
    setTripResultModalMessage("결재선을 최소 1명 이상 지정해 주세요.", true);
    return;
  }

  setTripResultModalMessage("출장결과 문서를 생성/상신하는 중입니다...");
  submitBtn.disabled = true;
  cancelBtn.disabled = true;
  closeBtn.disabled = true;
  try {
    const data = await api(`/api/business-trips/${sourceDocId}/result`, {
      method: "POST",
      body: JSON.stringify({
        approver_ids: approverIds,
        reference_ids: referenceIds,
        trip_result: tripResult,
      }),
    });
    const warnings = Array.isArray(data?.warnings) ? data.warnings.filter((x) => String(x || "").trim()) : [];
    const createdDocId = Number(data?.document?.id || 0);
    closeTripResultModal();
    await refreshAll();
    if (createdDocId) {
      await openDocument(createdDocId);
    }
    if (warnings.length) {
      alert(`출장결과 상신은 완료되었지만 경고가 있습니다.\n- ${warnings.join("\n- ")}`);
    }
  } catch (err) {
    setTripResultModalMessage(err.message || "출장결과 상신에 실패했습니다.", true);
  } finally {
    submitBtn.disabled = false;
    cancelBtn.disabled = false;
    closeBtn.disabled = false;
  }
}

function setEducationResultModalMessage(message, isError = false) {
  const msgEl = $("#educationResultModalMsg");
  if (!msgEl) return;
  msgEl.textContent = String(message || "");
  msgEl.style.color = isError ? "var(--danger)" : "var(--muted)";
}

function closeEducationResultModal() {
  const modal = $("#educationResultModal");
  if (!modal) return;
  modal.classList.add("hidden");
  const summaryEl = $("#educationResultModalSummary");
  const contentEl = $("#educationResultContentInput");
  const applyPointEl = $("#educationResultApplyPointInput");
  if (summaryEl) summaryEl.textContent = "대상 교육문서 정보를 불러오는 중입니다.";
  if (contentEl) contentEl.value = "";
  if (applyPointEl) applyPointEl.value = "";
  setEducationResultModalMessage("");
  state.ui.educationResultModal = {
    sourceDocId: null,
    sourceRow: null,
    approverIds: [],
    referenceIds: [],
  };
  renderEducationResultAssigneeLists();
  syncBodyModalOpenClass();
}

async function openEducationResultModal(rowData) {
  const modal = $("#educationResultModal");
  const summaryEl = $("#educationResultModalSummary");
  const contentEl = $("#educationResultContentInput");
  const applyPointEl = $("#educationResultApplyPointInput");
  if (!modal || !summaryEl || !contentEl || !applyPointEl) return;

  const sourceDocId = Number(rowData?.document_id || 0);
  if (!sourceDocId) {
    setEducationResultModalMessage("대상 교육문서를 찾지 못했습니다.", true);
    return;
  }

  if ((!state.users || !state.users.length) && state.token) {
    try {
      await loadUsers();
    } catch (_) {
      // 사용자 목록 조회 실패 시에도 모달은 열어두고 수동 재시도 가능하게 둔다.
    }
  }

  const candidates = (state.users || []).filter((u) => !state.me || Number(u.id) !== Number(state.me.id));

  const statusText = statusLabel[rowData.document_status] || rowData.document_status || "-";
  const periodText = `${fmt(rowData.education_start_date)} ~ ${fmt(rowData.education_end_date)}`;
  const resultCount = Number(rowData.result_doc_count || 0);
  summaryEl.textContent = `#${sourceDocId} ${rowData.document_title || "-"} | 상태: ${statusText} | 교육기간: ${periodText} | 기존 교육결과 ${resultCount}건`;
  contentEl.value = "";
  applyPointEl.value = "";
  setEducationResultModalMessage(candidates.length ? "" : "선택 가능한 결재자가 없습니다.", !candidates.length);

  state.ui.educationResultModal = {
    sourceDocId,
    sourceRow: rowData || null,
    approverIds: [],
    referenceIds: [],
  };
  renderEducationResultAssigneeLists();
  modal.classList.remove("hidden");
  syncBodyModalOpenClass();
  setTimeout(() => {
    try { contentEl.focus(); } catch (_) { }
  }, 0);
}

async function submitEducationResultModal() {
  const contentEl = $("#educationResultContentInput");
  const applyPointEl = $("#educationResultApplyPointInput");
  const submitBtn = $("#submitEducationResultBtn");
  const cancelBtn = $("#cancelEducationResultModalBtn");
  const closeBtn = $("#closeEducationResultModalBtn");
  const sourceDocId = Number(state.ui?.educationResultModal?.sourceDocId || 0);
  if (!contentEl || !applyPointEl || !submitBtn || !cancelBtn || !closeBtn) return;
  if (!sourceDocId) {
    setEducationResultModalMessage("대상 교육문서 정보가 없습니다. 다시 시도해 주세요.", true);
    return;
  }

  const educationContent = String(contentEl.value || "").trim();
  const educationApplyPoint = String(applyPointEl.value || "").trim();
  const approverIds = [...(state.ui?.educationResultModal?.approverIds || [])];
  const referenceIds = [...(state.ui?.educationResultModal?.referenceIds || [])];
  if (!educationContent) {
    setEducationResultModalMessage("교육내용을 입력해 주세요.", true);
    contentEl.focus();
    return;
  }
  if (!educationApplyPoint) {
    setEducationResultModalMessage("적용점을 입력해 주세요.", true);
    applyPointEl.focus();
    return;
  }
  if (!approverIds.length) {
    setEducationResultModalMessage("결재선을 최소 1명 이상 지정해 주세요.", true);
    return;
  }

  setEducationResultModalMessage("교육결과 문서를 생성/상신하는 중입니다...");
  submitBtn.disabled = true;
  cancelBtn.disabled = true;
  closeBtn.disabled = true;
  try {
    const data = await api(`/api/educations/${sourceDocId}/result`, {
      method: "POST",
      body: JSON.stringify({
        approver_ids: approverIds,
        reference_ids: referenceIds,
        education_content: educationContent,
        education_apply_point: educationApplyPoint,
      }),
    });
    const warnings = Array.isArray(data?.warnings) ? data.warnings.filter((x) => String(x || "").trim()) : [];
    const createdDocId = Number(data?.document?.id || 0);
    closeEducationResultModal();
    await refreshAll();
    if (createdDocId) {
      await openDocument(createdDocId);
    }
    if (warnings.length) {
      alert(`교육결과 상신은 완료되었지만 경고가 있습니다.\n- ${warnings.join("\n- ")}`);
    }
  } catch (err) {
    setEducationResultModalMessage(err.message || "교육결과 상신에 실패했습니다.", true);
  } finally {
    submitBtn.disabled = false;
    cancelBtn.disabled = false;
    closeBtn.disabled = false;
  }
}

function handleDraftTemplateTypeChange() {
  const form = $("#draftForm");
  const wasVisible = isElementVisible($("#leaveInfoModal"));
  const wasOvertimeVisible = isElementVisible($("#overtimeInfoModal"));
  const wasBusinessTripVisible = isElementVisible($("#businessTripInfoModal"));
  const wasEducationVisible = isElementVisible($("#educationInfoModal"));
  syncDraftTemplateTypeUI();
  if (!form) return;
  const templateType = (form.template_type?.value || "");
  if (templateType === "leave") {
    if (!leaveInfoHasValue(form)) openLeaveInfoModal();
    if (wasOvertimeVisible) closeOvertimeInfoModal();
    if (wasBusinessTripVisible) closeBusinessTripInfoModal();
    if (wasEducationVisible) closeEducationInfoModal();
  } else if (templateType === "overtime") {
    if (!overtimeInfoHasValue(form)) openOvertimeInfoModal();
    if (wasVisible) closeLeaveInfoModal();
    if (wasBusinessTripVisible) closeBusinessTripInfoModal();
    if (wasEducationVisible) closeEducationInfoModal();
  } else if (templateType === "business_trip") {
    if (!businessTripInfoHasValue(form)) openBusinessTripInfoModal();
    if (wasVisible) closeLeaveInfoModal();
    if (wasOvertimeVisible) closeOvertimeInfoModal();
    if (wasEducationVisible) closeEducationInfoModal();
  } else if (templateType === "education") {
    if (!educationInfoHasValue(form)) openEducationInfoModal();
    if (wasVisible) closeLeaveInfoModal();
    if (wasOvertimeVisible) closeOvertimeInfoModal();
    if (wasBusinessTripVisible) closeBusinessTripInfoModal();
  } else {
    if (wasVisible) closeLeaveInfoModal();
    if (wasOvertimeVisible) closeOvertimeInfoModal();
    if (wasBusinessTripVisible) closeBusinessTripInfoModal();
    if (wasEducationVisible) closeEducationInfoModal();
  }
}

function syncDraftTemplateTypeUI() {
  const form = $("#draftForm");
  if (!form) return;
  const templateType = (form.template_type?.value || "");
  const isLeave = templateType === "leave";
  const isOvertime = templateType === "overtime";
  const isBusinessTrip = templateType === "business_trip";
  const isEducation = templateType === "education";
  const leaveWrap = $("#leaveFieldsWrap");
  const overtimeWrap = $("#overtimeFieldsWrap");
  const businessTripWrap = $("#businessTripFieldsWrap");
  const educationWrap = $("#educationFieldsWrap");
  if (leaveWrap) leaveWrap.classList.toggle("hidden", !isLeave);
  if (overtimeWrap) overtimeWrap.classList.toggle("hidden", !isOvertime);
  if (businessTripWrap) businessTripWrap.classList.toggle("hidden", !isBusinessTrip);
  if (educationWrap) educationWrap.classList.toggle("hidden", !isEducation);
  if (form.leave_start_date) form.leave_start_date.required = isLeave;
  if (form.leave_end_date) form.leave_end_date.required = isLeave;
  if (form.leave_type) form.leave_type.required = isLeave;
  if (form.overtime_start_date) form.overtime_start_date.required = isOvertime;
  if (form.overtime_end_date) form.overtime_end_date.required = isOvertime;
  if (form.overtime_type) form.overtime_type.required = isOvertime;
  if (form.trip_start_date) form.trip_start_date.required = isBusinessTrip;
  if (form.trip_end_date) form.trip_end_date.required = isBusinessTrip;
  if (form.trip_type) form.trip_type.required = isBusinessTrip;
  if (form.trip_destination) form.trip_destination.required = isBusinessTrip;
  if (form.trip_purpose) form.trip_purpose.required = isBusinessTrip;
  if (form.education_title) form.education_title.required = isEducation;
  if (form.education_category) form.education_category.required = isEducation;
  if (form.education_provider) form.education_provider.required = isEducation;
  if (form.education_location) form.education_location.required = isEducation;
  if (form.education_start_date) form.education_start_date.required = isEducation;
  if (form.education_end_date) form.education_end_date.required = isEducation;
  if (form.education_purpose) form.education_purpose.required = isEducation;
  if (!isLeave) $("#leaveInfoModal")?.classList.add("hidden");
  if (!isOvertime) $("#overtimeInfoModal")?.classList.add("hidden");
  if (!isBusinessTrip) $("#businessTripInfoModal")?.classList.add("hidden");
  if (!isEducation) $("#educationInfoModal")?.classList.add("hidden");
  if (isLeave) syncLeaveDaysAutoCalc();
  else renderLeaveInfoSummary();
  if (isOvertime) syncOvertimeHoursAutoCalc();
  else renderOvertimeInfoSummary();
  renderBusinessTripInfoSummary();
  renderEducationInfoSummary();
  syncBodyModalOpenClass();
}

function getCurrentDraftOutputFolderId() {
  const templateType = ($("#draftForm [name='template_type']")?.value || "").trim();
  if (templateType === "leave" && GOOGLE_LEAVE_DRAFT_OUTPUT_FOLDER_ID) {
    return GOOGLE_LEAVE_DRAFT_OUTPUT_FOLDER_ID;
  }
  if (templateType === "overtime" && GOOGLE_OVERTIME_DRAFT_OUTPUT_FOLDER_ID) {
    return GOOGLE_OVERTIME_DRAFT_OUTPUT_FOLDER_ID;
  }
  if (templateType === "business_trip" && GOOGLE_BUSINESS_TRIP_DRAFT_OUTPUT_FOLDER_ID) {
    return GOOGLE_BUSINESS_TRIP_DRAFT_OUTPUT_FOLDER_ID;
  }
  if (templateType === "education" && GOOGLE_EDUCATION_DRAFT_OUTPUT_FOLDER_ID) {
    return GOOGLE_EDUCATION_DRAFT_OUTPUT_FOLDER_ID;
  }
  if (templateType === "outbound" && GOOGLE_OUTBOUND_DRAFT_OUTPUT_FOLDER_ID) {
    return GOOGLE_OUTBOUND_DRAFT_OUTPUT_FOLDER_ID;
  }
  return GOOGLE_DRAFT_OUTPUT_FOLDER_ID;
}

function extractGoogleDocIdClient(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  if (/^[A-Za-z0-9_-]{20,}$/.test(raw)) return raw;

  const pathMatch = raw.match(/\/document\/d\/([A-Za-z0-9_-]{20,})/);
  if (pathMatch) return pathMatch[1];

  try {
    const parsed = new URL(raw);
    if (parsed.hostname.endsWith("google.com")) {
      const queryId = parsed.searchParams.get("id");
      if (queryId && /^[A-Za-z0-9_-]{20,}$/.test(queryId)) return queryId;
    }
  } catch (_) { }
  return null;
}

function googleDocLinksClient(value) {
  const docId = extractGoogleDocIdClient(value);
  if (!docId) return null;
  const base = `https://docs.google.com/document/d/${docId}`;
  return {
    docId,
    editUrl: `${base}/edit`,
    previewUrl: `${base}/preview`,
  };
}

function updateGoogleDocComposePanel(forceReload = false) {
  const links = googleDocLinksClient($("#googleDocUrl").value);
  const frame = $("#googleDocComposeFrame");
  const openBtn = $("#openGoogleDocBtn");
  const stateText = $("#googleDocState");

  if (!links) {
    openBtn.disabled = true;
    delete openBtn.dataset.href;
    stateText.textContent = "기안문서를 선택해 주세요.";
    if (forceReload || frame.dataset.docId) {
      frame.src = "about:blank";
      delete frame.dataset.docId;
    }
    setGoogleControlState("");
    return;
  }

  openBtn.disabled = false;
  openBtn.dataset.href = links.editUrl;
  stateText.textContent = `문서 ID: ${links.docId}`;
  if (forceReload || frame.dataset.docId !== links.docId) {
    frame.src = links.previewUrl;
    frame.dataset.docId = links.docId;
  }
  setGoogleControlState("");
}

function setGoogleControlState(message = "") {
  const cfg = state.google.config;
  const configError = String(state.google.configError || "").trim();
  const hasConfig = !!(cfg && cfg.enabled);
  const hasToken = !!state.google.accessToken && Date.now() < state.google.tokenExpiryMs;
  const isGoogleLogin = (state.me?.auth_provider || "") === "google";

  const connectBtn = $("#googleConnectBtn");
  const pickBtn = $("#googlePickBtn");
  if (!connectBtn || !pickBtn) return;

  connectBtn.classList.add("hidden");

  if (!cfg) {
    pickBtn.disabled = true;
    const status = $("#googleDocState");
    if (status) status.textContent = message || configError || "Google 설정 로딩 중...";
    return;
  }

  pickBtn.disabled = !hasConfig;

  if (!hasConfig) {
    $("#googleDocState").textContent = "Google 문서 설정이 아직 완료되지 않았습니다. 관리자에게 확인해 주세요.";
    return;
  }
  if (message) {
    $("#googleDocState").textContent = message;
  } else if (hasToken) {
    $("#googleDocState").textContent = "Google 문서 사용 준비가 완료되었습니다.";
  } else if (isGoogleLogin) {
    $("#googleDocState").textContent = "기안문서를 선택하면 필요한 문서 권한을 자동으로 확인합니다.";
  } else {
    $("#googleDocState").textContent = "기안문서 선택 시 Google 권한 확인이 필요할 수 있습니다.";
  }
}

function resetGoogleSession() {
  state.google.accessToken = "";
  state.google.tokenExpiryMs = 0;
  setGoogleControlState("");
}

function waitForGoogleScripts(timeoutMs = 12000) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const timer = setInterval(() => {
      const hasGis = !!window.google?.accounts?.oauth2;
      const hasGapi = !!window.gapi?.load;
      if (hasGis && hasGapi) {
        clearInterval(timer);
        resolve();
        return;
      }
      if (Date.now() - start > timeoutMs) {
        clearInterval(timer);
        reject(new Error("Google 스크립트를 로드하지 못했습니다. 네트워크와 브라우저 차단 설정을 확인해 주세요."));
      }
    }, 120);
  });
}

async function ensureGoogleIntegrationInitialized() {
  if (!state.google.config) {
    try {
      const data = await api("/api/integrations/google/config");
      state.google.config = data.config;
      state.google.configError = "";
    } catch (err) {
      const msg = `Google 설정 조회 실패: ${err.message}`;
      state.google.configError = msg;
      setGoogleControlState(msg);
      throw new Error(msg);
    }
  }

  if (!state.google.config?.enabled) {
    setGoogleControlState("");
    return false;
  }

  if (state.google.initialized) return true;
  if (state.google.initPromise) return state.google.initPromise;

  state.google.initPromise = (async () => {
    await waitForGoogleScripts();
    state.google.scriptsReady = true;

    await new Promise((resolve) => {
      window.gapi.load("picker", {
        callback: () => {
          state.google.pickerReady = true;
          resolve();
        },
      });
    });

    state.google.tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: state.google.config.client_id,
      scope: state.google.config.scope || DEFAULT_GOOGLE_SCOPE,
      callback: () => { },
    });
    state.google.initialized = true;
    state.google.configError = "";
    setGoogleControlState("");
    return true;
  })()
    .catch((err) => {
      state.google.initialized = false;
      state.google.configError = String(err?.message || "Google 초기화 실패");
      setGoogleControlState(err.message);
      throw err;
    })
    .finally(() => {
      state.google.initPromise = null;
    });

  return state.google.initPromise;
}

async function requestGoogleAccessToken(prompt = "") {
  await ensureGoogleIntegrationInitialized();
  if (!state.google.tokenClient) {
    throw new Error("Google OAuth 초기화에 실패했습니다.");
  }

  const response = await new Promise((resolve, reject) => {
    state.google.tokenClient.callback = (tokenResponse) => {
      if (tokenResponse?.error) {
        reject(new Error(tokenResponse.error_description || tokenResponse.error));
        return;
      }
      resolve(tokenResponse);
    };
    state.google.tokenClient.requestAccessToken(prompt ? { prompt } : {});
  });

  state.google.accessToken = response.access_token || "";
  const expiresSec = Number(response.expires_in || 3600);
  state.google.tokenExpiryMs = Date.now() + Math.max(expiresSec - 60, 60) * 1000;
  setGoogleControlState("Google 문서 사용 준비가 완료되었습니다.");
  return state.google.accessToken;
}

async function ensureGoogleToken() {
  if (state.google.accessToken && Date.now() < state.google.tokenExpiryMs) {
    return state.google.accessToken;
  }
  try {
    return await requestGoogleAccessToken("");
  } catch (_) {
    return requestGoogleAccessToken("consent");
  }
}

function applyGoogleDocSelection(docId, url, name = "") {
  const links = googleDocLinksClient(docId || url);
  if (!links) {
    throw new Error("선택한 문서에서 Google Docs ID를 추출하지 못했습니다.");
  }
  $("#googleDocUrl").value = url || links.editUrl;
  if (!$("#draftForm [name='title']").value.trim() && name) {
    $("#draftForm [name='title']").value = name;
  }
  updateGoogleDocComposePanel(true);
  setGoogleControlState(`선택됨: ${name || links.docId}`);
}

async function copyGoogleDocFromTemplate(sourceDocId, sourceName = "") {
  const token = await ensureGoogleToken();
  const draftTitle = $("#draftForm [name='title']").value.trim();
  const baseName = draftTitle || sourceName || "전자결재 문서";
  const copyName = `${baseName} (${new Date().toISOString().slice(0, 16).replace("T", " ")})`;

  const outputFolderId = getCurrentDraftOutputFolderId();
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(sourceDocId)}/copy?supportsAllDrives=true&fields=id,name,webViewLink,parents`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: copyName,
      ...(outputFolderId ? { parents: [outputFolderId] } : {}),
    }),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload?.error?.message || "템플릿 문서 사본 생성 실패");
  }
  return payload;
}

async function validateTemplateSourceDoc(docId) {
  const data = await api("/api/integrations/google/validate-template", {
    method: "POST",
    body: JSON.stringify({
      doc_id: docId,
      required_slots: getDraftApprovalTemplateTotalSlots(),
    }),
  });
  return data.result;
}

async function resetCopiedApprovalDoc(docId) {
  const data = await api("/api/integrations/google/reset-approval-doc", {
    method: "POST",
    body: JSON.stringify({
      doc_id: docId,
      total_slots: getDraftApprovalTemplateTotalSlots(),
    }),
  });
  return data.result;
}

function isTimeoutLikeErrorMessage(message) {
  const text = String(message || "").toLowerCase();
  return text.includes("timeout") || text.includes("timed out");
}

async function openDrivePicker() {
  await ensureGoogleIntegrationInitialized();
  const token = await ensureGoogleToken();
  if (!state.google.pickerReady || !window.google?.picker) {
    throw new Error("템플릿 선택창을 아직 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.");
  }

  const view = new window.google.picker.DocsView(window.google.picker.ViewId.DOCS)
    .setIncludeFolders(false)
    .setSelectFolderEnabled(false)
    .setMimeTypes(GOOGLE_DOCS_MIME);
  if (GOOGLE_TEMPLATE_FOLDER_ID && typeof view.setParent === "function") {
    try {
      view.setParent(GOOGLE_TEMPLATE_FOLDER_ID);
    } catch (_) { }
  }

  const pickerBuilder = new window.google.picker.PickerBuilder()
    .addView(view)
    .setOAuthToken(token)
    .setDeveloperKey(state.google.config.api_key)
    .setOrigin(window.location.origin)
    .setCallback((data) => {
      if (data.action !== window.google.picker.Action.PICKED) return;
      const picked = data.docs?.[0];
      if (!picked) return;
      const pickedId = picked[window.google.picker.Document.ID] || picked.id;
      const pickedName = picked[window.google.picker.Document.NAME] || picked.name || "";
      (async () => {
        try {
          setGoogleControlState("템플릿 확인 중...");
          await validateTemplateSourceDoc(pickedId);
          setGoogleControlState("선택한 템플릿을 복사하는 중...");
          const copied = await copyGoogleDocFromTemplate(pickedId, pickedName);
          setGoogleControlState("결재 칸을 정리하는 중...");
          let resetWarn = "";
          try {
            await resetCopiedApprovalDoc(copied.id);
          } catch (resetErr) {
            if (isTimeoutLikeErrorMessage(resetErr?.message)) {
              resetWarn = " (결재칸 초기화 지연 - 계속 진행)";
              console.warn("resetCopiedApprovalDoc timeout; continuing with copied template", resetErr);
            } else {
              throw resetErr;
            }
          }
          applyGoogleDocSelection(copied.id, copied.webViewLink, copied.name || pickedName);
          setGoogleControlState(`템플릿 연결 완료${resetWarn}: ${copied.name || pickedName}`);
        } catch (err) {
          setGoogleControlState(err.message);
        }
      })();
    });
  if (window.google.picker.Feature?.SUPPORT_DRIVES) {
    try { pickerBuilder.enableFeature(window.google.picker.Feature.SUPPORT_DRIVES); } catch (_) { }
  }
  if (window.google.picker.Feature?.NAV_HIDDEN) {
    try { pickerBuilder.enableFeature(window.google.picker.Feature.NAV_HIDDEN); } catch (_) { }
  }

  if (state.google.config.app_id) {
    pickerBuilder.setAppId(state.google.config.app_id);
  }
  pickerBuilder.build().setVisible(true);
}


function showView(viewId) {
  $$(".view").forEach((v) => v.classList.remove("active"));
  $(`#${viewId}`)?.classList.add("active");
  $$("#tabs button").forEach((b) => b.classList.toggle("active", b.dataset.view === viewId));
}

function hideLandingShowApp() {
  const lp = document.getElementById("landingPage");
  const main = document.querySelector("main.layout");
  if (lp) lp.style.display = "none";
  if (main) main.style.display = "";
}

function setLoggedIn(loggedIn) {
  hideLandingShowApp();
  $("#loginView").classList.toggle("hidden", loggedIn);
  $("#appView").classList.toggle("hidden", !loggedIn);
}

function setLoginGoogleHint(message) {
  const hint = $("#googleAuthHint");
  if (hint) hint.textContent = message || "";
}

function waitForGoogleIdentityLibrary(timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    if (window.google?.accounts?.id) {
      resolve(true);
      return;
    }
    const started = Date.now();
    const timer = setInterval(() => {
      if (window.google?.accounts?.id) {
        clearInterval(timer);
        resolve(true);
        return;
      }
      if (Date.now() - started >= timeoutMs) {
        clearInterval(timer);
        reject(new Error("Google 인증 스크립트 로딩 시간이 초과되었습니다."));
      }
    }, 120);
  });
}

async function fetchGoogleAuthPublicConfig() {
  const resp = await fetch(API_BASE + "/api/auth/google/config", { method: "GET" });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || `Google 설정 조회 실패 (${resp.status})`);
  return data;
}

async function completeLogin(data) {
  state.token = data.token;
  state.me = data.user;
  localStorage.setItem("approval_token", state.token);
  setLoggedIn(true);
  $("#whoami").textContent = `${state.me.full_name} (${state.me.department}, ${state.me.role})`;
  $("#adminTabBtn").classList.toggle("hidden", state.me.role !== "admin");
  $("#docFilterArchivedBtn")?.classList.toggle("hidden", state.me.role !== "admin");
  if (state.me.role !== "admin" && getCurrentDocFilter() === "archived") setCurrentDocFilter("all");
  await initTabsForCurrentUser();
  await refreshAll();
  const googleReady = await ensureGoogleIntegrationInitialized().catch(() => false);
  if (googleReady) setGoogleControlState("");
  showView("dashboard");
}

async function handleGoogleAuthCredential(credential) {
  $("#loginError").textContent = "";
  if (!credential) {
    $("#loginError").textContent = "Google 인증 토큰을 받지 못했습니다.";
    return;
  }
  setLoginGoogleHint("Google 계정 확인 중...");
  try {
    const data = await api("/api/auth/google", {
      method: "POST",
      body: JSON.stringify({ id_token: credential }),
    });
    await completeLogin(data);
    setLoginGoogleHint(
      data?.created
        ? "Google 계정으로 회원가입이 완료되었습니다."
        : "Google 계정 로그인 완료"
    );
  } catch (err) {
    $("#loginError").textContent = err.message;
    setLoginGoogleHint("Google 인증에 실패했습니다. 다시 시도해 주세요.");
  }
}

async function initLoginGoogleAuth(force = false) {
  const loginView = $("#loginView");
  const buttonWrap = $("#googleAuthButton");
  if (!loginView || !buttonWrap) return;
  if (state.token) return;

  if (force) {
    state.ui.loginGoogleAuth.initialized = false;
  }

  let cfg;
  try {
    cfg = await fetchGoogleAuthPublicConfig();
  } catch (err) {
    buttonWrap.innerHTML = "";
    setLoginGoogleHint(`Google 회원가입 설정을 불러오지 못했습니다. (${err.message})`);
    return;
  }

  if (!cfg?.enabled || !cfg?.client_id) {
    buttonWrap.innerHTML = "";
    setLoginGoogleHint("Google 회원가입 기능이 아직 설정되지 않았습니다. 관리자에게 문의해 주세요.");
    return;
  }

  if (
    state.ui.loginGoogleAuth.initialized &&
    state.ui.loginGoogleAuth.clientId === cfg.client_id &&
    !force
  ) {
    return;
  }

  try {
    await waitForGoogleIdentityLibrary();
    window.google.accounts.id.initialize({
      client_id: cfg.client_id,
      callback: (response) => {
        handleGoogleAuthCredential(response?.credential || "").catch((err) => {
          $("#loginError").textContent = err.message || "Google 인증 처리 중 오류가 발생했습니다.";
        });
      },
      auto_select: false,
      cancel_on_tap_outside: true,
    });
    buttonWrap.innerHTML = "";
    window.google.accounts.id.renderButton(buttonWrap, {
      type: "standard",
      theme: "outline",
      size: "large",
      text: "signin_with",
      shape: "rectangular",
      logo_alignment: "left",
      width: 360,
    });
    state.ui.loginGoogleAuth.initialized = true;
    state.ui.loginGoogleAuth.clientId = cfg.client_id;
    setLoginGoogleHint("Google 계정으로 회원가입/로그인이 가능합니다. 최초 1회 가입 후 같은 버튼으로 로그인하세요.");
  } catch (err) {
    buttonWrap.innerHTML = "";
    setLoginGoogleHint(`Google 인증 버튼 초기화 실패: ${err.message}`);
  }
}

function getCurrentTabOrder() {
  return $$("#tabs button[data-view]").map((btn) => btn.dataset.view);
}

function applyTabOrder(order) {
  const tabs = $("#tabs");
  if (!tabs || !Array.isArray(order) || !order.length) return;
  const buttons = Array.from(tabs.querySelectorAll("button[data-view]"));
  const byView = new Map(buttons.map((btn) => [String(btn.dataset.view || ""), btn]));
  const arranged = [];
  for (const view of order) {
    const key = String(view || "");
    if (!byView.has(key)) continue;
    arranged.push(byView.get(key));
    byView.delete(key);
  }
  for (const fallbackView of TAB_ORDER_DEFAULT) {
    if (!byView.has(fallbackView)) continue;
    arranged.push(byView.get(fallbackView));
    byView.delete(fallbackView);
  }
  arranged.push(...Array.from(byView.values()));
  arranged.forEach((btn) => tabs.appendChild(btn));
}

function normalizeTabOrder(order) {
  if (!Array.isArray(order) || !order.length) return [];
  const out = [];
  for (const item of order) {
    const key = String(item || "").trim();
    if (!key || out.includes(key)) continue;
    out.push(key);
  }
  return out;
}

async function loadSavedTabOrder() {
  let order = [];
  try {
    const data = await api("/api/ui/tab-order");
    order = normalizeTabOrder(data?.tab_order);
  } catch (_) {
    order = [];
  }
  applyTabOrder(order.length ? order : TAB_ORDER_DEFAULT);
}

async function saveCurrentTabOrder() {
  const order = getCurrentTabOrder();
  const normalized = normalizeTabOrder(order);
  if (!normalized.length) {
    throw new Error("저장할 탭 순서가 없습니다.");
  }
  await api("/api/ui/tab-order", {
    method: "POST",
    body: JSON.stringify({ tab_order: normalized }),
  });
}

function setTabEditMode(editing, options = {}) {
  const { restore = false } = options;
  const tabs = $("#tabs");
  const editBtn = $("#tabEditBtn");
  const cancelBtn = $("#tabEditCancelBtn");
  if (!tabs || !editBtn || !cancelBtn) return;
  const isAdmin = state.me?.role === "admin";
  if (!isAdmin) {
    editing = false;
  }
  if (restore && Array.isArray(state.ui?.tabEdit?.originalOrder) && state.ui.tabEdit.originalOrder.length) {
    applyTabOrder(state.ui.tabEdit.originalOrder);
  }
  state.ui.tabEdit.editing = !!editing;
  if (!editing) {
    state.ui.tabEdit.draggingView = "";
  }
  tabs.classList.toggle("is-editing", !!editing);
  tabs.querySelectorAll("button[data-view]").forEach((btn) => {
    const hidden = btn.classList.contains("hidden");
    btn.draggable = !!editing && !hidden;
  });
  editBtn.textContent = editing ? "탭 순서 저장" : "탭 위치 편집";
  editBtn.classList.toggle("primary", !!editing);
  cancelBtn.classList.toggle("hidden", !editing);
}

function syncTabEditControlsByRole() {
  const editBtn = $("#tabEditBtn");
  const cancelBtn = $("#tabEditCancelBtn");
  if (!editBtn || !cancelBtn) return;
  const isAdmin = state.me?.role === "admin";
  editBtn.classList.toggle("hidden", !isAdmin);
  if (!isAdmin) {
    cancelBtn.classList.add("hidden");
    setTabEditMode(false);
    return;
  }
  if (!state.ui?.tabEdit?.editing) {
    cancelBtn.classList.add("hidden");
  }
}

async function initTabsForCurrentUser() {
  if (!state.me) return;
  await loadSavedTabOrder();
  setTabEditMode(false);
  syncTabEditControlsByRole();
}

async function bootstrap() {
  if (!state.token) {
    resetGoogleSession();
    setLoggedIn(false);
    await initLoginGoogleAuth(true);
    return;
  }
  try {
    const me = await api("/api/auth/me");
    state.me = me.user;
    setLoggedIn(true);
    $("#whoami").textContent = `${state.me.full_name} (${state.me.department}, ${state.me.role})`;

    // Toggle Admin Tab
    $("#adminTabBtn").classList.toggle("hidden", state.me.role !== "admin");
    $("#docFilterArchivedBtn")?.classList.toggle("hidden", state.me.role !== "admin");
    if (state.me.role !== "admin" && getCurrentDocFilter() === "archived") setCurrentDocFilter("all");
    await initTabsForCurrentUser();
    applyDraftIssueDefaults(true);
    syncDraftTemplateTypeUI();

    await refreshAll();
    const googleReady = await ensureGoogleIntegrationInitialized().catch(() => false);
    if (googleReady) setGoogleControlState("");
  } catch (err) {
    state.token = "";
    state.me = null;
    localStorage.removeItem("approval_token");
    resetGoogleSession();
    syncTabEditControlsByRole();
    setLoggedIn(false);
    await initLoginGoogleAuth(true);
  }
}

async function refreshAll() {
  if (state.token && !state.ui?.tabEdit?.editing) {
    await loadSavedTabOrder().catch(() => { });
  }
  if (state.token) {
    try {
      const meResp = await api("/api/auth/me");
      state.me = meResp.user;
      $("#whoami").textContent = `${state.me.full_name} (${state.me.department}, ${state.me.role})`;
      applyDraftIssueDefaults(false);
      syncDraftTemplateTypeUI();
      syncTabEditControlsByRole();
    } catch (_) { }
  }
  const tasks = [loadUsers(), loadDashboard(), loadMyDocs(), loadLeaveMgmt(), loadOvertimeMgmt(), loadTripMgmt(), loadEducationMgmt()];
  const results = await Promise.allSettled(tasks);
  const firstRejected = results.find((r) => r.status === "rejected");
  if (firstRejected && firstRejected.reason) {
    console.warn("[refreshAll] partial failure:", firstRejected.reason);
  }
}

async function loadUsers() {
  const data = await api("/api/users");
  state.users = data.users;
  const availableIds = new Set(state.users.map((u) => u.id));
  state.approverIds = state.approverIds.filter((id) => availableIds.has(id));
  state.referenceIds = state.referenceIds.filter((id) => availableIds.has(id));
  if (state.picker.selectedIds.length) {
    state.picker.selectedIds = state.picker.selectedIds.filter((id) => availableIds.has(id));
  }
  renderAssigneeLists();
  renderPickerRows();
  renderAssigneeLists();
  renderPickerRows();
  renderAdminUserTable();
}

function renderAdminUserTable() {
  const tbody = $("#adminUserBody");
  if (!tbody) return;
  tbody.innerHTML = "";
  const sorted = [...state.users].sort((a, b) => a.username.localeCompare(b.username));
  for (const user of sorted) {
    const lockSelf = state.me && user.id === state.me.id;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${user.id}</td>
      <td>${esc(user.username)}</td>
      <td>${esc(user.full_name)}</td>
      <td>${esc(user.department)}</td>
      <td>${esc(user.role)}</td>
      <td>
        <button type="button" class="btn" data-edit-user-id="${user.id}">수정</button>
        <button type="button" class="btn danger" data-delete-user-id="${user.id}" ${lockSelf ? "disabled" : ""}>
          ${lockSelf ? "본인계정" : "삭제"}
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  }
}

let editingUserId = null;

function editAdminUser(userId) {
  const user = getUserById(userId);
  if (!user) return;

  editingUserId = userId;
  const form = $("#adminUserForm");
  form.username.value = user.username;
  form.username.disabled = true; // Cannot change username
  form.password.value = ""; // Don't show password, placeholders for change
  form.password.placeholder = "변경시에만 입력";
  form.full_name.value = user.full_name;
  form.department.value = user.department;
  form.role.value = user.role;

  $("#adminUserSubmitBtn").textContent = "정보 수정";
  $("#adminUserCancelBtn").classList.remove("hidden");
  $("#adminUserTitle").textContent = "사용자 정보 수정";
}

function cancelEditMode() {
  editingUserId = null;
  const form = $("#adminUserForm");
  form.reset();
  form.username.disabled = false;
  form.password.placeholder = "6자 이상";

  $("#adminUserSubmitBtn").textContent = "사용자 등록";
  $("#adminUserCancelBtn").classList.add("hidden");
  $("#adminUserTitle").textContent = "새 사용자 등록";
}

async function handleAdminUserSubmit() {
  const form = $("#adminUserForm");
  const msg = $("#adminUserMsg");
  msg.textContent = "";

  const payload = {
    username: form.username.value.trim(),
    password: form.password.value,
    full_name: form.full_name.value.trim(),
    department: form.department.value.trim(),
    role: form.role.value,
  };

  try {
    if (editingUserId) {
      // Update Mode
      await api(`/api/users/${editingUserId}/update`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      msg.textContent = "사용자 정보 수정 완료";
      cancelEditMode(); // Exit edit mode on success
    } else {
      // Create Mode
      await api("/api/users", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      msg.textContent = "사용자 등록 완료";
      form.reset();
    }

    msg.style.color = "var(--accent-2)";
    await loadUsers();
  } catch (err) {
    msg.textContent = err.message;
    msg.style.color = "var(--danger)";
  }
}

async function deleteAdminUser(userId) {
  const msg = $("#adminUserMsg");
  // Confirm dialog is handled, but let's be safe
  if (!window.confirm("정말 이 사용자를 삭제하시겠습니까? 삭제 후 복구할 수 없습니다.")) return;

  try {
    await api(`/api/users/${userId}/delete`, { method: "POST", body: "{}" });
    msg.textContent = "사용자 삭제 완료";
    msg.style.color = "var(--accent-2)";
    await loadUsers();
  } catch (err) {
    alert(`삭제 실패: ${err.message}`); // Use alert for better visibility of errors
    msg.textContent = `삭제 실패: ${err.message}`;
    msg.style.color = "var(--danger)";
  }
}

function getUserById(userId) {
  return state.users.find((u) => u.id === userId) || null;
}

function userDisplay(u) {
  return `${u.full_name} (${u.username} / ${u.department})`;
}

function renderAssigneeLists() {
  const approverHost = $("#approverSelectedList");
  const refHost = $("#refSelectedList");
  if (!approverHost || !refHost) return;

  approverHost.innerHTML = "";
  refHost.innerHTML = "";

  if (!state.approverIds.length) {
    approverHost.innerHTML = "<li>등록된 결재선이 없습니다.</li>";
  } else {
    state.approverIds.forEach((id, idx) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${idx + 1}. ${esc(userDisplay(user))}</span>
        <span class="row-gap">
          <button type="button" class="btn" data-approver-action="up" data-user-id="${id}">↑</button>
          <button type="button" class="btn" data-approver-action="down" data-user-id="${id}">↓</button>
          <button type="button" class="btn danger" data-approver-action="remove" data-user-id="${id}">삭제</button>
        </span>
      `;
      approverHost.appendChild(li);
    });
  }

  if (!state.referenceIds.length) {
    refHost.innerHTML = "<li>등록된 참조자가 없습니다.</li>";
  } else {
    state.referenceIds.forEach((id) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${esc(userDisplay(user))}</span>
        <button type="button" class="btn danger" data-ref-action="remove" data-user-id="${id}">삭제</button>
      `;
      refHost.appendChild(li);
    });
  }
}

function moveApprover(userId, direction) {
  const idx = state.approverIds.indexOf(userId);
  if (idx < 0) return;
  const swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= state.approverIds.length) return;
  [state.approverIds[idx], state.approverIds[swapIdx]] = [state.approverIds[swapIdx], state.approverIds[idx]];
  renderAssigneeLists();
}

function moveTripResultApprover(userId, direction) {
  const list = state.ui?.tripResultModal?.approverIds || [];
  const idx = list.indexOf(userId);
  if (idx < 0) return;
  const swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= list.length) return;
  [list[idx], list[swapIdx]] = [list[swapIdx], list[idx]];
  state.ui.tripResultModal.approverIds = list;
  renderTripResultAssigneeLists();
}

function moveEducationResultApprover(userId, direction) {
  const list = state.ui?.educationResultModal?.approverIds || [];
  const idx = list.indexOf(userId);
  if (idx < 0) return;
  const swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= list.length) return;
  [list[idx], list[swapIdx]] = [list[swapIdx], list[idx]];
  state.ui.educationResultModal.approverIds = list;
  renderEducationResultAssigneeLists();
}

function renderTripResultAssigneeLists() {
  const approverHost = $("#tripResultApproverSelectedList");
  const refHost = $("#tripResultRefSelectedList");
  if (!approverHost || !refHost) return;

  const approverIds = state.ui?.tripResultModal?.approverIds || [];
  const refIds = state.ui?.tripResultModal?.referenceIds || [];

  approverHost.innerHTML = "";
  refHost.innerHTML = "";

  if (!approverIds.length) {
    approverHost.innerHTML = "<li>등록된 결재선이 없습니다.</li>";
  } else {
    approverIds.forEach((id, idx) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${idx + 1}. ${esc(userDisplay(user))}</span>
        <span class="row-gap">
          <button type="button" class="btn" data-trip-result-approver-action="up" data-user-id="${id}">↑</button>
          <button type="button" class="btn" data-trip-result-approver-action="down" data-user-id="${id}">↓</button>
          <button type="button" class="btn danger" data-trip-result-approver-action="remove" data-user-id="${id}">삭제</button>
        </span>
      `;
      approverHost.appendChild(li);
    });
  }

  if (!refIds.length) {
    refHost.innerHTML = "<li>등록된 참조자가 없습니다.</li>";
  } else {
    refIds.forEach((id) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${esc(userDisplay(user))}</span>
        <button type="button" class="btn danger" data-trip-result-ref-action="remove" data-user-id="${id}">삭제</button>
      `;
      refHost.appendChild(li);
    });
  }
}

function renderEducationResultAssigneeLists() {
  const approverHost = $("#educationResultApproverSelectedList");
  const refHost = $("#educationResultRefSelectedList");
  if (!approverHost || !refHost) return;

  const approverIds = state.ui?.educationResultModal?.approverIds || [];
  const refIds = state.ui?.educationResultModal?.referenceIds || [];

  approverHost.innerHTML = "";
  refHost.innerHTML = "";

  if (!approverIds.length) {
    approverHost.innerHTML = "<li>등록된 결재선이 없습니다.</li>";
  } else {
    approverIds.forEach((id, idx) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${idx + 1}. ${esc(userDisplay(user))}</span>
        <span class="row-gap">
          <button type="button" class="btn" data-education-result-approver-action="up" data-user-id="${id}">↑</button>
          <button type="button" class="btn" data-education-result-approver-action="down" data-user-id="${id}">↓</button>
          <button type="button" class="btn danger" data-education-result-approver-action="remove" data-user-id="${id}">삭제</button>
        </span>
      `;
      approverHost.appendChild(li);
    });
  }

  if (!refIds.length) {
    refHost.innerHTML = "<li>등록된 참조자가 없습니다.</li>";
  } else {
    refIds.forEach((id) => {
      const user = getUserById(id);
      if (!user) return;
      const li = document.createElement("li");
      li.innerHTML = `
        <span>${esc(userDisplay(user))}</span>
        <button type="button" class="btn danger" data-education-result-ref-action="remove" data-user-id="${id}">삭제</button>
      `;
      refHost.appendChild(li);
    });
  }
}

function renderPickerRows() {
  const body = $("#pickerBody");
  if (!body) return;
  const query = (state.picker.query || "").trim().toLowerCase();
  const users = state.users
    .filter((u) => !state.me || u.id !== state.me.id)
    .filter((u) => {
      if (!query) return true;
      return (
        u.username.toLowerCase().includes(query)
        || u.full_name.toLowerCase().includes(query)
        || u.department.toLowerCase().includes(query)
      );
    })
    .slice(0, 300);

  body.innerHTML = "";
  if (!users.length) {
    body.innerHTML = `<tr><td colspan="5">일치하는 사용자가 없습니다.</td></tr>`;
    return;
  }

  for (const user of users) {
    const checked = state.picker.selectedIds.includes(user.id) ? "checked" : "";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input type="checkbox" data-picker-user-id="${user.id}" ${checked} /></td>
      <td>${esc(user.username)}</td>
      <td>${esc(user.full_name)}</td>
      <td>${esc(user.department)}</td>
      <td>${esc(user.role)}</td>
    `;
    body.appendChild(tr);
  }
}

async function openUserPicker(target) {
  ensureUserPickerModalMountedToBody();
  state.picker.target = target;
  state.picker.query = "";
  if (target === "approver") {
    state.picker.selectedIds = [...state.approverIds];
  } else if (target === "ref") {
    state.picker.selectedIds = [...state.referenceIds];
  } else if (target === "trip_result_approver") {
    state.picker.selectedIds = [...(state.ui?.tripResultModal?.approverIds || [])];
  } else if (target === "trip_result_ref") {
    state.picker.selectedIds = [...(state.ui?.tripResultModal?.referenceIds || [])];
  } else if (target === "education_result_approver") {
    state.picker.selectedIds = [...(state.ui?.educationResultModal?.approverIds || [])];
  } else if (target === "education_result_ref") {
    state.picker.selectedIds = [...(state.ui?.educationResultModal?.referenceIds || [])];
  } else {
    state.picker.selectedIds = [];
  }
  const pickerTitle =
    target === "approver" ? "결재선 사용자 선택"
      : target === "ref" ? "참조자 사용자 선택"
        : target === "trip_result_approver" ? "출장결과 결재선 선택"
          : target === "trip_result_ref" ? "출장결과 참조자 선택"
            : target === "education_result_approver" ? "교육결과 결재선 선택"
              : target === "education_result_ref" ? "교육결과 참조자 선택"
                : "사용자 선택";
  $("#pickerTitle").textContent = pickerTitle;
  $("#pickerSearchInput").value = "";
  $("#pickerMsg").textContent = "";
  if ((!state.users || state.users.length === 0) && state.token) {
    const body = $("#pickerBody");
    if (body) body.innerHTML = `<tr><td colspan="5">사용자 목록 불러오는 중...</td></tr>`;
    try {
      await loadUsers();
    } catch (err) {
      if (body) body.innerHTML = `<tr><td colspan="5">사용자 목록 조회 실패: ${esc(err?.message || "알 수 없는 오류")}</td></tr>`;
    }
  }
  renderPickerRows();
  $("#userPickerModal").classList.remove("hidden");
  syncBodyModalOpenClass();
}

function closeUserPicker() {
  $("#userPickerModal").classList.add("hidden");
  syncBodyModalOpenClass();
}

function applyPickerSelection() {
  const pickerMsg = $("#pickerMsg");
  if (pickerMsg) pickerMsg.textContent = "";
  if (state.picker.target === "approver") {
    const maxApprovers = getDraftMaxApproverSteps();
    const totalSlots = getDraftApprovalTemplateTotalSlots();
    if (state.picker.selectedIds.length > maxApprovers) {
      if (pickerMsg) {
        pickerMsg.textContent = `현재 템플릿은 결재자 최대 ${maxApprovers}명까지 지원합니다. (기안자 포함 ${totalSlots}칸)`;
      }
      return;
    }
    state.approverIds = [...state.picker.selectedIds];
  } else if (state.picker.target === "ref") {
    state.referenceIds = [...state.picker.selectedIds];
  } else if (state.picker.target === "trip_result_approver") {
    if (state.picker.selectedIds.length > MAX_APPROVER_STEPS) {
      if (pickerMsg) {
        pickerMsg.textContent = `출장결과 결재선은 최대 ${MAX_APPROVER_STEPS}명까지 지정할 수 있습니다.`;
      }
      return;
    }
    state.ui.tripResultModal.approverIds = [...state.picker.selectedIds];
    renderTripResultAssigneeLists();
  } else if (state.picker.target === "trip_result_ref") {
    state.ui.tripResultModal.referenceIds = [...state.picker.selectedIds];
    renderTripResultAssigneeLists();
  } else if (state.picker.target === "education_result_approver") {
    if (state.picker.selectedIds.length > MAX_APPROVER_STEPS) {
      if (pickerMsg) {
        pickerMsg.textContent = `교육결과 결재선은 최대 ${MAX_APPROVER_STEPS}명까지 지정할 수 있습니다.`;
      }
      return;
    }
    state.ui.educationResultModal.approverIds = [...state.picker.selectedIds];
    renderEducationResultAssigneeLists();
  } else if (state.picker.target === "education_result_ref") {
    state.ui.educationResultModal.referenceIds = [...state.picker.selectedIds];
    renderEducationResultAssigneeLists();
  }
  renderAssigneeLists();
  closeUserPicker();
}

async function loadDashboard() {
  const now = new Date();
  const dashSummaryMonth = state.ui.dashboardSummaryMonth || (state.ui.dashboardSummaryMonth = {
    year: now.getFullYear(),
    month: now.getMonth(),
  });
  const summaryYear = Number.isFinite(Number(dashSummaryMonth.year))
    ? Number(dashSummaryMonth.year)
    : now.getFullYear();
  const summaryMonthIndex = Number.isFinite(Number(dashSummaryMonth.month))
    ? Number(dashSummaryMonth.month)
    : now.getMonth();
  const summaryMonth = summaryMonthIndex + 1;
  const overtimeYear = summaryYear;
  const overtimeMonth = summaryMonth;
  const [d, overtimePayload, leavePayload, tripPayload, educationPayload, notificationPayload] = await Promise.all([
    api("/api/dashboard"),
    api(`/api/overtimes?year=${overtimeYear}&month=${overtimeMonth}`).catch(() => null),
    api("/api/leaves").catch(() => null),
    api(`/api/business-trips?year=${summaryYear}&month=${summaryMonth}`).catch(() => null),
    api(`/api/educations?year=${summaryYear}&month=${summaryMonth}`).catch(() => null),
    api("/api/notifications").catch(() => ({ notifications: [] })),
  ]);
  const me = state.me;
  const rawDrafts = d.my_drafts || [];
  const rawInReview = d.my_in_review || [];
  const rawPendingApprovals = d.my_pending_approvals || [];
  const rawCompleted = d.my_completed || [];
  const rejectedFromCompleted = rawCompleted.filter((x) => x && x.status === "rejected");
  const completedApprovedOnly = rawCompleted.filter((x) => !x || x.status !== "rejected");
  const inReviewMerged = [...rawInReview];
  for (const doc of rejectedFromCompleted) {
    if (!inReviewMerged.some((x) => Number(x?.id || 0) === Number(doc?.id || 0))) {
      inReviewMerged.push(doc);
    }
  }
  inReviewMerged.sort((a, b) => String(b?.updated_at || "").localeCompare(String(a?.updated_at || "")));

  const remainingLeave = (me.total_leave || 0) - (me.used_leave || 0);
  const initial = (me.full_name || "?").charAt(0);
  const fmtNum = (value) => {
    const n = Math.round((Number(value || 0) + Number.EPSILON) * 10) / 10;
    return Number.isInteger(n) ? String(n) : String(n);
  };
  const pad2 = (n) => String(n).padStart(2, "0");
  const summaryMonthLabel = `${summaryYear}.${pad2(summaryMonth)}`;
  const leaveRecords = leavePayload?.my_records || [];
  const leaveMonthRecords = leaveRecords.filter((r) => {
    const raw = String(r?.start_date || "").trim();
    if (!raw) return false;
    const dt = new Date(raw.replace(" ", "T"));
    if (Number.isNaN(dt.getTime())) return false;
    return dt.getFullYear() === summaryYear && (dt.getMonth() + 1) === summaryMonth;
  });
  const leaveMonthUsed = leaveMonthRecords.reduce((sum, r) => sum + Number(r?.leave_days || 0), 0);
  const leaveMonthCount = leaveMonthRecords.length;
  const leaveSummaryLoaded = !!leavePayload;
  const overtimeSummary = overtimePayload?.summary || null;
  const overtimeSel = overtimePayload?.selection || { year: overtimeYear, month: overtimeMonth };
  const overtimeMonthLabel = `${overtimeSel.year}.${pad2(overtimeSel.month)}`;
  const overtimeExceeded = Number(overtimeSummary?.month_used_hours || 0) > Number(overtimeSummary?.monthly_cap_hours || 15);
  const tripSummary = tripPayload?.summary || null;
  const tripSel = tripPayload?.selection || { year: summaryYear, month: summaryMonth };
  const tripMonthLabel = `${tripSel.year}.${pad2(tripSel.month)}`;
  const tripSummaryLoaded = !!tripPayload;
  const tripResultDocCount = Array.isArray(tripPayload?.result_history)
    ? tripPayload.result_history.reduce((sum, item) => sum + Number(item?.result_doc_count || 0), 0)
    : 0;
  const educationSummary = educationPayload?.summary || null;
  const educationSel = educationPayload?.selection || { year: summaryYear, month: summaryMonth };
  const educationMonthLabel = `${educationSel.year}.${pad2(educationSel.month)}`;
  const educationSummaryLoaded = !!educationPayload;
  const educationResultDocCount = Array.isArray(educationPayload?.result_history)
    ? educationPayload.result_history.reduce((sum, item) => sum + Number(item?.result_doc_count || 0), 0)
    : 0;
  const notifications = Array.isArray(notificationPayload?.notifications) ? notificationPayload.notifications : [];
  const unreadNotifications = notifications.filter((n) => !n?.is_read);
  const unreadNotificationItems = unreadNotifications.slice(0, 8);

  const renderList = (items, type, pageKey) => {
    if (!items || items.length === 0) return `<div class="dash-empty">등록된 문서가 없습니다.</div>`;
    const total = items.length;
    const totalPages = Math.max(1, Math.ceil(total / DASHBOARD_LIST_PAGE_SIZE));
    const currentPage = Math.min(
      Math.max(1, Number(state.ui?.dashboardPages?.[pageKey] || 1)),
      totalPages
    );
    if (state.ui?.dashboardPages) {
      state.ui.dashboardPages[pageKey] = currentPage;
    }
    const startIdx = (currentPage - 1) * DASHBOARD_LIST_PAGE_SIZE;
    const pageItems = items.slice(startIdx, startIdx + DASHBOARD_LIST_PAGE_SIZE);

    const listHtml = `<div class="dash-list-body">` + pageItems.map(doc => {
      let meta = "";
      if (type === "draft") meta = fmt(doc.updated_at);
      else if (type === "in_review") {
        if ((doc.status === "rejected" || doc.returned_for_resubmit) && doc.latest_rejection) {
          meta = `<span class="stat-rejected">반려됨: ${esc(doc.latest_rejection.approver?.name || "결재자")} (재상신 필요)</span>`;
        } else {
          const approverName = doc.current_step?.approver_name || doc.current_approver_name || "대기";
          meta = `<span class="stat-pending">결재중: ${esc(approverName)}</span>`;
        }
      }
      else if (type === "to_approve") {
        meta = `<span class="stat-pending">기안자: ${esc(doc.drafter?.name || doc.drafter_name || "-")}</span>`;
      }
      else if (type === "completed") meta = `<span class="${doc.status === 'approved' ? 'stat-approved' : 'stat-rejected'}">${doc.status === 'approved' ? '승인' : '반려'}</span>`;
      return `<div class="dash-list-item" onclick="openDocument(${doc.id})">
        <span class="title">${esc(doc.title)}</span>
        <span class="meta">${meta}</span>
      </div>`;
    }).join("") + `</div>`;

    if (totalPages <= 1) return listHtml;

    const pagerButtons = Array.from({ length: totalPages }, (_, idx) => {
      const page = idx + 1;
      const active = page === currentPage ? "is-active" : "";
      return `<button type="button" class="dash-page-btn ${active}" onclick="dashboardSetPage('${pageKey}', ${page})">${page}</button>`;
    }).join("");

    return `${listHtml}<div class="dash-pager"><span class="dash-pager-meta">${currentPage}/${totalPages}</span>${pagerButtons}</div>`;
  };

  const getNotificationIcon = (message) => {
    const msg = String(message || "");
    if (msg.includes("최종 승인") || msg.includes("승인")) return "✅";
    if (msg.includes("결재 요청") || msg.includes("결재 차례")) return "📥";
    if (msg.includes("반려")) return "⛔";
    return "🔔";
  };
  const unreadNotificationHtml = unreadNotificationItems.length
    ? `<ul id="dashAlertList" class="dash-alert-list">${unreadNotificationItems.map((n) => `
        <li class="dash-alert-item is-unread" data-notification-id="${Number(n.id || 0)}" data-notification-link="${esc(n.link || "")}">
          <button type="button" class="dash-alert-open" data-alert-open-id="${Number(n.id || 0)}">
            <span class="dash-alert-dot" aria-hidden="true"></span>
            <span class="dash-alert-msg">${esc(getNotificationIcon(n.message))} ${esc(n.message || "-")}</span>
            <span class="dash-alert-time">${esc(fmt(n.created_at))}</span>
          </button>
          <button type="button" class="btn btn-small dash-alert-read-btn" data-alert-read-id="${Number(n.id || 0)}">읽음</button>
        </li>
      `).join("")
    }</ul>${unreadNotifications.length > unreadNotificationItems.length
      ? `<p class="dash-alert-more">미확인 ${unreadNotifications.length - unreadNotificationItems.length}건이 더 있습니다.</p>`
      : ""
    }`
    : `<div class="dash-alert-empty">새 알림이 없습니다.</div>`;

  $("#dashboard").innerHTML = `
    <div class="dash-grid">
      <!-- Left Sidebar -->
      <div class="dash-sidebar">
        <!-- Profile Card -->
        <div class="profile-card">
          <div class="profile-header">
            ${me.profile_image_url
      ? `<div class="profile-avatar" style="background-image:url('${esc(me.profile_image_url)}');background-size:cover;background-position:center;color:transparent;">${esc(initial)}</div>`
      : `<div class="profile-avatar">${esc(initial)}</div>`
    }
            <div class="profile-name">${esc(me.full_name)}</div>
            <div class="profile-dept">${esc(me.department)} · ${esc(me.job_title || me.role)}</div>
          </div>
          <div class="profile-body">
            <div class="profile-info-row">
              <span class="info-label">아이디</span>
              <span class="info-value">${esc(me.username)}</span>
            </div>
            <div class="profile-info-row">
              <span class="info-label">부서</span>
              <span class="info-value">${esc(me.department)}</span>
            </div>
            <div class="profile-info-row">
              <span class="info-label">직급</span>
              <span class="info-value">${esc(me.job_title || "-")}</span>
            </div>
            <div class="profile-info-row">
              <span class="info-label">권한</span>
              <span class="info-value">${esc(me.role)}</span>
            </div>
          </div>
        </div>

        <div class="dash-alert-card">
          <div class="section-header dash-alert-head">
            <span>🔔 미확인 알림</span>
            <span class="count ${unreadNotifications.length ? "is-hot" : ""}">${unreadNotifications.length}</span>
          </div>
          <div class="dash-alert-body">
            ${unreadNotificationHtml}
          </div>
          <div class="dash-alert-foot">
            <button type="button" id="dashMarkAllReadBtn" class="btn btn-small" ${unreadNotifications.length ? "" : "disabled"}>전체 읽음 처리</button>
          </div>
        </div>

      </div>

      <!-- Main Content -->
      <div class="dash-main">
        <div class="dash-summary-wrap">
          <div class="dash-summary-toolbar">
            <div class="dash-summary-toolbar-title">월간 요약</div>
            <div class="dash-summary-toolbar-nav">
              <button type="button" class="btn btn-small" onclick="dashboardSummaryShiftMonth(-1)">이전</button>
              <strong>${esc(summaryMonthLabel)}</strong>
              <button type="button" class="btn btn-small" onclick="dashboardSummaryShiftMonth(1)">다음</button>
              <button type="button" class="btn btn-small" onclick="dashboardSummarySetCurrentMonth()">이번달</button>
            </div>
          </div>
          <div class="dash-summary-grid">
            <button type="button" class="dash-summary-card dash-summary-card-btn summary-leave" onclick="openDashboardLeaveSummary()">
              <div class="dash-summary-head">
                <span>🏖 휴가요약</span>
                <span class="dash-summary-period">${esc(summaryMonthLabel)}</span>
              </div>
              ${leaveSummaryLoaded
      ? `<div class="dash-summary-stats">
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(leaveMonthUsed)}</div>
                        <div class="lbl">선택월 사용(일)</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(leaveMonthCount)}</div>
                        <div class="lbl">선택월 건수</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num is-success">${fmtNum(remainingLeave)}</div>
                        <div class="lbl">현재 잔여</div>
                      </div>
                    </div>
                    <div class="dash-summary-meta">
                      <span>누적사용 ${fmtNum(me.used_leave || 0)}일</span>
                      <span>총부여 ${fmtNum(me.total_leave || 0)}일</span>
                    </div>`
      : `<div class="dash-summary-empty">휴가 요약을 불러오지 못했습니다.</div>`
    }
            </button>
            <button type="button" class="dash-summary-card dash-summary-card-btn summary-overtime ${overtimeExceeded ? "is-warn" : ""}" onclick="openDashboardOvertimeSummary()">
              <div class="dash-summary-head">
                <span>⏱ 연장근로 요약</span>
                <span class="dash-summary-period">${esc(overtimeMonthLabel)}</span>
              </div>
              ${overtimeSummary
      ? `<div class="dash-summary-stats">
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(overtimeSummary.monthly_cap_hours || 15)}</div>
                        <div class="lbl">월 기준(시간)</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num ${overtimeExceeded ? "is-danger" : ""}">${fmtNum(overtimeSummary.month_used_hours || 0)}</div>
                        <div class="lbl">사용</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num ${overtimeExceeded ? "is-danger" : "is-success"}">${fmtNum(overtimeSummary.month_remaining_hours || 0)}</div>
                        <div class="lbl">잔여</div>
                      </div>
                    </div>
                    <div class="dash-summary-meta">
                      <span>100% ${fmtNum(overtimeSummary.month_used_100_hours || 0)}h</span>
                      <span>150% ${fmtNum(overtimeSummary.month_used_150_hours || 0)}h</span>
                      <span>환산 ${fmtNum(overtimeSummary.month_weighted_hours || 0)}h</span>
                    </div>`
      : `<div class="dash-summary-empty">연장근로 요약을 불러오지 못했습니다.</div>`
    }
            </button>
            <button type="button" class="dash-summary-card dash-summary-card-btn summary-trip" onclick="openDashboardTripSummary()">
              <div class="dash-summary-head">
                <span>🚗 출장요약</span>
                <span class="dash-summary-period">${esc(tripMonthLabel)}</span>
              </div>
              ${tripSummaryLoaded
      ? `<div class="dash-summary-stats">
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(tripSummary?.month_count || 0)}</div>
                        <div class="lbl">출장 건수</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(tripSummary?.month_total_hours || 0)}</div>
                        <div class="lbl">총 시간</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num">${Number(tripSummary?.month_total_expense || 0) > 0 ? formatWonDisplay(tripSummary.month_total_expense).replace("원", "") : "0"}</div>
                        <div class="lbl">총 출장비(원)</div>
                      </div>
                    </div>
                    <div class="dash-summary-meta">
                      <span>결과문서 ${fmtNum(tripResultDocCount)}건</span>
                    </div>`
      : `<div class="dash-summary-empty">출장 요약을 불러오지 못했습니다.</div>`
    }
            </button>
            <button type="button" class="dash-summary-card dash-summary-card-btn summary-education" onclick="openDashboardEducationSummary()">
              <div class="dash-summary-head">
                <span>📘 교육요약</span>
                <span class="dash-summary-period">${esc(educationMonthLabel)}</span>
              </div>
              ${educationSummaryLoaded
      ? `<div class="dash-summary-stats">
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(educationSummary?.month_count || 0)}</div>
                        <div class="lbl">교육 건수</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num">${fmtNum(educationSummary?.month_total_hours || 0)}</div>
                        <div class="lbl">총 시간</div>
                      </div>
                      <div class="dash-summary-stat">
                        <div class="num">${Number(educationSummary?.month_total_amount || 0) > 0 ? formatWonDisplay(educationSummary.month_total_amount).replace("원", "") : "0"}</div>
                        <div class="lbl">총 교육비(원)</div>
                      </div>
                    </div>
                    <div class="dash-summary-meta">
                      <span>결과문서 ${fmtNum(educationResultDocCount)}건</span>
                    </div>`
      : `<div class="dash-summary-empty">교육 요약을 불러오지 못했습니다.</div>`
    }
            </button>
          </div>
        </div>

        <!-- Draft Documents -->
        <div class="dash-list-card">
          <div class="section-header">
            <span>📝 임시저장 문서</span>
            <span class="count">${rawDrafts.length}</span>
          </div>
          ${renderList(rawDrafts, "draft", "draft")}
        </div>

        <!-- In Review -->
        <div class="dash-list-card">
          <div class="section-header">
            <span>⏳ 진행중 문서</span>
            <span class="count">${inReviewMerged.length}</span>
          </div>
          ${renderList(inReviewMerged, "in_review", "in_review")}
        </div>

        <!-- To Approve -->
        <div class="dash-list-card">
          <div class="section-header">
            <span>📥 결재할 문서</span>
            <span class="count">${rawPendingApprovals.length}</span>
          </div>
          ${renderList(rawPendingApprovals, "to_approve", "to_approve")}
        </div>

        <!-- Completed -->
        <div class="dash-list-card">
          <div class="section-header">
            <span>✅ 처리완료 문서</span>
            <span class="count">${completedApprovedOnly.length}</span>
          </div>
          ${renderList(completedApprovedOnly, "completed", "completed")}
        </div>
      </div>
    </div>
  `;

  const markNotificationRead = async (ids) => {
    const normalized = (ids || [])
      .map((x) => Number(x || 0))
      .filter((x) => x > 0);
    if (!normalized.length) return;
    await api("/api/notifications/read", {
      method: "POST",
      body: JSON.stringify({ ids: normalized }),
    });
  };

  $("#dashMarkAllReadBtn")?.addEventListener("click", async () => {
    try {
      await api("/api/notifications/read", { method: "POST", body: "{}" });
      await loadDashboard();
    } catch (err) {
      alert(err.message || "전체 읽음 처리에 실패했습니다.");
    }
  });

  $("#dashAlertList")?.addEventListener("click", async (e) => {
    const readBtn = e.target.closest("button[data-alert-read-id]");
    if (readBtn) {
      const id = Number(readBtn.dataset.alertReadId || 0);
      if (!id) return;
      try {
        await markNotificationRead([id]);
        await loadDashboard();
      } catch (err) {
        alert(err.message || "읽음 처리에 실패했습니다.");
      }
      return;
    }

    const openBtn = e.target.closest("button[data-alert-open-id]");
    if (!openBtn) return;
    const li = openBtn.closest(".dash-alert-item");
    const id = Number(li?.dataset.notificationId || 0);
    const link = String(li?.dataset.notificationLink || "").trim();
    try {
      if (id > 0) {
        await markNotificationRead([id]);
      }
      await loadDashboard();
      const m = /^\/documents\/(\d+)$/.exec(link);
      if (m) {
        await openDocument(Number(m[1]));
      }
    } catch (err) {
      alert(err.message || "알림 처리 중 오류가 발생했습니다.");
    }
  });
}

async function dashboardSetPage(pageKey, page) {
  if (!state.ui.dashboardPages) state.ui.dashboardPages = {};
  state.ui.dashboardPages[pageKey] = Math.max(1, Number(page) || 1);
  await loadDashboard();
}

async function dashboardSummaryShiftMonth(delta) {
  const now = new Date();
  const s = state.ui.dashboardSummaryMonth || (state.ui.dashboardSummaryMonth = {
    year: now.getFullYear(),
    month: now.getMonth(),
  });
  const base = new Date(Number(s.year) || now.getFullYear(), Number(s.month) || 0, 1);
  base.setMonth(base.getMonth() + Number(delta || 0));
  s.year = base.getFullYear();
  s.month = base.getMonth();
  await loadDashboard();
}

async function dashboardSummarySetCurrentMonth() {
  const now = new Date();
  state.ui.dashboardSummaryMonth = { year: now.getFullYear(), month: now.getMonth() };
  await loadDashboard();
}

function openDashboardLeaveSummary() {
  const dash = state.ui.dashboardSummaryMonth || { year: new Date().getFullYear(), month: new Date().getMonth() };
  state.ui.leaveCalendar = state.ui.leaveCalendar || {};
  state.ui.leaveCalendar.year = Number(dash.year) || new Date().getFullYear();
  state.ui.leaveCalendar.month = Number(dash.month) || 0;
  showView("leaveMgmt");
  loadLeaveMgmt().catch((err) => alert(err.message));
}

function openDashboardOvertimeSummary() {
  const dash = state.ui.dashboardSummaryMonth || { year: new Date().getFullYear(), month: new Date().getMonth() };
  state.ui.overtimeDashboard = state.ui.overtimeDashboard || {};
  state.ui.overtimeDashboard.year = Number(dash.year) || new Date().getFullYear();
  state.ui.overtimeDashboard.month = Number(dash.month) || 0;
  showView("overtimeMgmt");
  loadOvertimeMgmt().catch((err) => alert(err.message));
}

function openDashboardTripSummary() {
  const dash = state.ui.dashboardSummaryMonth || { year: new Date().getFullYear(), month: new Date().getMonth() };
  state.ui.businessTripDashboard = state.ui.businessTripDashboard || {};
  state.ui.businessTripDashboard.year = Number(dash.year) || new Date().getFullYear();
  state.ui.businessTripDashboard.month = Number(dash.month) || 0;
  showView("tripMgmt");
  loadTripMgmt().catch((err) => alert(err.message));
}

function openDashboardEducationSummary() {
  const dash = state.ui.dashboardSummaryMonth || { year: new Date().getFullYear(), month: new Date().getMonth() };
  state.ui.educationDashboard = state.ui.educationDashboard || {};
  state.ui.educationDashboard.year = Number(dash.year) || new Date().getFullYear();
  state.ui.educationDashboard.month = Number(dash.month) || 0;
  showView("educationMgmt");
  loadEducationMgmt().catch((err) => alert(err.message));
}

async function loadLeaveMgmt() {
  const host = $("#leaveMgmt");
  if (!host || !state.me) return;
  const data = await api("/api/leaves");
  const summary = data.summary || { total_leave: 0, used_leave: 0 };
  const remaining = Math.max(0, (summary.total_leave || 0) - (summary.used_leave || 0));

  const renderLeaveRows = (rows, emptyText) => {
    if (!rows || !rows.length) return `<tr><td colspan="7">${emptyText}</td></tr>`;
    return rows.map((r) => `
      <tr data-id="${r.document_id}">
        <td>${esc(r.user?.name || "-")}</td>
        <td>${esc(r.user?.department || "-")}</td>
        <td>${esc(r.start_date || "-")}</td>
        <td>${esc(r.end_date || "-")}</td>
        <td>${r.leave_days ?? "-"}</td>
        <td>${statusLabel[r.document_status] || esc(r.document_status || "-")}</td>
        <td>${esc(r.document_title || "-")}</td>
      </tr>
    `).join("");
  };

  host.innerHTML = `
    <h3>휴가관리</h3>
    <div class="leave-mgmt-top">
      <div class="leave-card">
        <div class="leave-title">내 휴가 요약</div>
        <div class="leave-stats">
          <div class="leave-stat"><div class="num">${summary.total_leave || 0}</div><div class="lbl">부여</div></div>
          <div class="leave-stat"><div class="num">${summary.used_leave || 0}</div><div class="lbl">사용</div></div>
          <div class="leave-stat"><div class="num" style="color:var(--success)">${remaining}</div><div class="lbl">잔여</div></div>
        </div>
      </div>
      ${renderLeaveCalendarPanel(data)}
    </div>
    <div class="table-wrap" style="margin-top:12px;">
      <h4>개인 휴가 사용 내역</h4>
      <table>
        <thead>
          <tr><th>이름</th><th>부서</th><th>시작일</th><th>종료일</th><th>일수</th><th>문서상태</th><th>문서제목</th></tr>
        </thead>
        <tbody id="myLeaveBody">${renderLeaveRows(data.my_records || [], "내 휴가 사용 내역이 없습니다.")}</tbody>
      </table>
    </div>
  `;

  host.querySelectorAll("[data-leave-cal-nav]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const delta = Number(btn.dataset.leaveCalNav || 0);
      if (!state.ui.leaveCalendar) return;
      if (delta === 0) {
        const now = new Date();
        state.ui.leaveCalendar.year = now.getFullYear();
        state.ui.leaveCalendar.month = now.getMonth();
      } else {
        const next = new Date(state.ui.leaveCalendar.year, state.ui.leaveCalendar.month + delta, 1);
        state.ui.leaveCalendar.year = next.getFullYear();
        state.ui.leaveCalendar.month = next.getMonth();
      }
      loadLeaveMgmt().catch((err) => alert(err.message));
    });
  });

  host.querySelectorAll("[data-leave-cal-filter]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const nextFilter = String(btn.dataset.leaveCalFilter || "all");
      if (!state.ui.leaveCalendar) return;
      if (state.ui.leaveCalendar.filter === nextFilter) return;
      state.ui.leaveCalendar.filter = nextFilter;
      loadLeaveMgmt().catch((err) => alert(err.message));
    });
  });

  host.querySelector("#leaveCalDeptFilter")?.addEventListener("change", (e) => {
    if (!state.ui.leaveCalendar) return;
    state.ui.leaveCalendar.department = String(e.target.value || "all");
    loadLeaveMgmt().catch((err) => alert(err.message));
  });

  host.querySelectorAll("[data-leave-doc-id]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const docId = Number(btn.dataset.leaveDocId || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });
}

function getOvertimeMonthLabel(year, month) {
  return `${year}년 ${month + 1}월`;
}

function formatHoursDisplay(value) {
  const n = Number(value || 0);
  if (!Number.isFinite(n)) return "0";
  return Number.isInteger(n) ? String(n) : n.toFixed(2).replace(/\.?0+$/, "");
}

function normalizeOvertimeTypeToken(value) {
  const raw = String(value || "").trim();
  if (!raw) return "unknown";
  if (raw.includes("평일")) return "weekday";
  if (raw.includes("휴일")) return "holiday";
  if (raw.includes("야간")) return "night";
  return "other";
}

function overtimeTypeLabel(value) {
  return String(value || "").trim() || "기타";
}

async function loadOvertimeMgmt() {
  const host = $("#overtimeMgmt");
  if (!host || !state.me) return;
  const dash = state.ui.overtimeDashboard || (state.ui.overtimeDashboard = {
    year: new Date().getFullYear(),
    month: new Date().getMonth(),
  });
  const query = new URLSearchParams({
    year: String(dash.year),
    month: String(dash.month + 1),
  });
  const data = await api(`/api/overtimes?${query.toString()}`);
  const selectedYear = Number(data.selection?.year || dash.year);
  const selectedMonth = Math.max(1, Number(data.selection?.month || (dash.month + 1)));
  dash.year = selectedYear;
  dash.month = selectedMonth - 1;
  const summary = data.summary || {};
  const monthlyCap = Number(summary.monthly_cap_hours || 15);
  const monthUsed = Number(summary.month_used_hours || 0);
  const monthRemaining = Number(summary.month_remaining_hours ?? Math.max(0, monthlyCap - monthUsed));
  const monthUsed100 = Number(summary.month_used_100_hours || 0);
  const monthUsed150 = Number(summary.month_used_150_hours || 0);
  const monthWeighted = Number(summary.month_weighted_hours || 0);
  const yearTotal = Number(summary.year_total_hours || 0);
  const monthOverLimit = monthUsed > monthlyCap;
  const monthRecords = Array.isArray(data.month_records) ? data.month_records : [];
  const monthDaily = Array.isArray(data.month_daily) ? data.month_daily : [];
  const yearMonthly = Array.isArray(data.year_monthly) ? data.year_monthly : [];
  const availableYears = Array.isArray(data.available_years) && data.available_years.length
    ? data.available_years.map((y) => Number(y)).filter((y) => Number.isFinite(y))
    : [selectedYear];
  if (!availableYears.includes(selectedYear)) availableYears.unshift(selectedYear);
  const yearsSorted = Array.from(new Set(availableYears)).sort((a, b) => b - a);

  const usedDayRows = monthDaily.filter((d) => Number(d.hours || 0) > 0);
  const dayChartMax = Math.max(...usedDayRows.map((d) => Number(d.hours || 0)), 1);
  const dayChartHtml = usedDayRows.map((row) => {
    const dayNo = Number(row.day || 0);
    const used = Number(row.hours || 0);
    const used100 = Number(row.hours_100 || 0);
    const used150 = Number(row.hours_150 || 0);
    const safeCap = Math.max(1, dayChartMax);
    const h100 = Math.max(4, Math.round((used100 / safeCap) * 100));
    const h150 = Math.max(0, Math.round((used150 / safeCap) * 100));
    const overClass = used > monthlyCap ? "is-over" : "";
    return `
      <button type="button" class="overtime-day-bar-card ${overClass}" data-ot-day="${dayNo}">
        <div class="day-head">
          <span class="d">${dayNo}일</span>
          <span class="t">${formatHoursDisplay(used)}h</span>
        </div>
        <div class="overtime-day-bar-wrap" title="${dayNo}일 · 100% ${formatHoursDisplay(used100)}h · 150% ${formatHoursDisplay(used150)}h">
          <div class="overtime-day-bar-stack">
            <div class="overtime-day-bar-seg is-150" style="height:${h150}%"></div>
            <div class="overtime-day-bar-seg is-100" style="height:${h100}%"></div>
          </div>
        </div>
        <div class="day-meta">100% ${formatHoursDisplay(used100)}h · 150% ${formatHoursDisplay(used150)}h</div>
      </button>
    `;
  }).join("");

  const overtimeTypeLegendItems = [
    ["weekday", "평일연장"],
    ["holiday", "휴일근로"],
    ["night", "야간근로"],
    ["other", "기타"],
  ];

  const rowsHtml = monthRecords.length ? monthRecords.map((r) => `
      <tr data-id="${r.document_id}">
        <td><span class="ot-type-badge type-${normalizeOvertimeTypeToken(r.overtime_type)}">${esc(overtimeTypeLabel(r.overtime_type))}</span></td>
        <td>${esc(fmt(r.start_date))}</td>
        <td>${esc(fmt(r.end_date))}</td>
        <td>${formatHoursDisplay(r.overtime_hours)}</td>
        <td>${formatHoursDisplay(r.overtime_100_hours || 0)}</td>
        <td>${formatHoursDisplay(r.overtime_150_hours || 0)}</td>
        <td>${statusLabel[r.document_status] || esc(r.document_status || "-")}</td>
        <td>${esc(r.document_title || "-")}</td>
      </tr>
  `).join("") : `<tr><td colspan="8">선택한 월의 개인 연장근로 내역이 없습니다.</td></tr>`;

  host.innerHTML = `
    <h3>연장근로</h3>
    <div class="leave-mgmt-top overtime-mgmt-top">
      <div class="overtime-left-stack">
        <div class="leave-card overtime-summary-card ${monthOverLimit ? "is-over" : ""}">
          <div class="leave-title">월 연장근로 요약 (${selectedYear}.${String(selectedMonth).padStart(2, "0")})</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthlyCap)}</div><div class="lbl">월 기준(시간)</div></div>
            <div class="leave-stat"><div class="num ${monthOverLimit ? "danger" : ""}">${formatHoursDisplay(monthUsed)}</div><div class="lbl">사용</div></div>
            <div class="leave-stat"><div class="num" style="color:var(--success)">${formatHoursDisplay(monthRemaining)}</div><div class="lbl">잔여</div></div>
          </div>
          <div class="overtime-summary-foot">${monthOverLimit ? `<span class="overtime-over-warn">월 15시간 초과</span> · ` : ""}연간 누적 사용: <strong>${formatHoursDisplay(yearTotal)}시간</strong></div>
        </div>
        <div class="leave-card overtime-rate-card">
          <div class="leave-title">월 100% / 150% 구분</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthUsed100)}</div><div class="lbl">100%</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthUsed150)}</div><div class="lbl">150%</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthWeighted)}</div><div class="lbl">환산(1.5x)</div></div>
          </div>
        </div>
      </div>
      <div class="leave-calendar-card overtime-dashboard-card">
        <div class="overtime-dashboard-head">
          <div class="leave-calendar-nav">
            <button type="button" class="btn" data-ot-nav="-1">이전</button>
            <strong>${getOvertimeMonthLabel(selectedYear, selectedMonth - 1)}</strong>
            <button type="button" class="btn" data-ot-nav="1">다음</button>
            <button type="button" class="btn" data-ot-nav="0">이번달</button>
          </div>
          <div class="overtime-dashboard-toolbar">
            <button type="button" id="overtimeExportBtn" class="btn">엑셀 다운로드</button>
          </div>
        </div>
        <div class="overtime-type-legend">
          ${overtimeTypeLegendItems.map(([key, label]) => `<span class="ot-type-badge type-${key}">${label}</span>`).join("")}
          <span class="ot-graph-legend"><i class="seg100"></i>100%</span>
          <span class="ot-graph-legend"><i class="seg150"></i>150%</span>
          <span class="ot-graph-legend"><i class="overflow"></i>15시간 초과</span>
        </div>
        <div class="overtime-year-chart">
          <div class="overtime-year-chart-head">
            <span>${selectedYear}년 ${selectedMonth}월 연장근로 사용 그래프</span>
          </div>
          ${dayChartHtml ? `<div class="overtime-day-bars">${dayChartHtml}</div>` : `<div class="overtime-chart-empty">선택한 월의 연장근로 사용일이 없습니다.</div>`}
        </div>
      </div>
    </div>
    <div class="table-wrap" style="margin-top:12px;">
      <h4>개인 연장근로 사용 내역 (${selectedYear}년 ${selectedMonth}월)</h4>
      <table>
        <thead>
          <tr><th>형태</th><th>시작일시</th><th>종료일시</th><th>시간</th><th>100%</th><th>150%</th><th>문서상태</th><th>문서제목</th></tr>
        </thead>
        <tbody id="myOvertimeBody">${rowsHtml}</tbody>
      </table>
    </div>
  `;

  host.querySelectorAll("[data-ot-nav]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const delta = Number(btn.dataset.otNav || 0);
      const current = new Date(dash.year, dash.month, 1);
      if (delta === 0) {
        const now = new Date();
        dash.year = now.getFullYear();
        dash.month = now.getMonth();
      } else {
        const next = new Date(current.getFullYear(), current.getMonth() + delta, 1);
        dash.year = next.getFullYear();
        dash.month = next.getMonth();
      }
      loadOvertimeMgmt().catch((err) => alert(err.message));
    });
  });

  host.querySelector("#overtimeExportBtn")?.addEventListener("click", () => {
    const qs = new URLSearchParams({
      year: String(dash.year),
      month: String(dash.month + 1),
    });
    window.open(`${API_BASE}/api/overtimes/export?${qs.toString()}`, "_blank");
  });
  host.querySelectorAll("[data-ot-day]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const day = Number(btn.dataset.otDay || 0);
      if (!day) return;
      const firstRow = monthRecords.find((r) => {
        const start = String(r.start_date || "");
        return start.slice(0, 10) === `${selectedYear}-${String(selectedMonth).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      });
      if (firstRow?.document_id) {
        openDocument(Number(firstRow.document_id)).catch((err) => alert(err.message));
      }
    });
  });
  host.querySelectorAll("#myOvertimeBody tr[data-id]").forEach((tr) => {
    tr.addEventListener("click", () => {
      const docId = Number(tr.dataset.id || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });
}

function formatTripDurationText(value) {
  const n = Number(value || 0);
  if (!Number.isFinite(n) || n <= 0) return "-";
  const hours = Math.floor(n);
  const minutes = Math.round((n - hours) * 60);
  if (hours > 0 && minutes > 0) return `${hours}시간 ${minutes}분`;
  if (hours > 0) return `${hours}시간`;
  return `${minutes}분`;
}

function formatWonDisplay(value) {
  const n = Number(value || 0);
  if (!Number.isFinite(n) || n <= 0) return "-";
  const rounded = Math.round(n * 100) / 100;
  if (Number.isInteger(rounded)) return `${rounded.toLocaleString("ko-KR")}원`;
  return `${rounded.toLocaleString("ko-KR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}원`;
}

function getTripSummaryCount(items, key) {
  const found = (Array.isArray(items) ? items : []).find((x) => String(x?.key || "") === String(key));
  return Number(found?.count || 0);
}

function normalizeTripTypeToken(value) {
  const text = String(value || "").trim();
  if (!text) return "other";
  if (text.includes("교육")) return "education";
  if (text.includes("회의")) return "meeting";
  if (text.includes("점검") || text.includes("방문") || text.includes("현장") || text.includes("외근")) return "field";
  if (text.includes("행정")) return "admin";
  if (text.includes("기타")) return "other";
  return "other";
}

async function loadTripMgmt() {
  const host = $("#tripMgmt");
  if (!host) return;
  if (!state.me) {
    host.innerHTML = `<h3>출장현황</h3><p class="hint">로그인 후 조회할 수 있습니다.</p>`;
    return;
  }

  const dash = state.ui.businessTripDashboard || (state.ui.businessTripDashboard = {
    year: new Date().getFullYear(),
    month: new Date().getMonth(),
  });
  const query = new URLSearchParams({
    year: String(dash.year),
    month: String(dash.month + 1),
  });
  let data;
  try {
    data = await api(`/api/business-trips?${query.toString()}`);
  } catch (err) {
    host.innerHTML = `
      <h3>출장현황</h3>
      <div class="card" style="padding:14px;">
        <p class="hint">출장현황을 불러오지 못했습니다.</p>
        <p class="hint">${esc(err.message || "알 수 없는 오류")}</p>
        <div class="row-gap" style="margin-top:8px;">
          <button type="button" class="btn" id="retryTripMgmtBtn">다시 시도</button>
        </div>
      </div>
    `;
    host.querySelector("#retryTripMgmtBtn")?.addEventListener("click", () => {
      loadTripMgmt().catch((e) => alert(e.message));
    });
    return;
  }
  const selectedYear = Number(data.selection?.year || dash.year);
  const selectedMonth = Math.max(1, Number(data.selection?.month || (dash.month + 1)));
  dash.year = selectedYear;
  dash.month = selectedMonth - 1;

  const summary = data.summary || {};
  const monthCount = Number(summary.month_count || 0);
  const monthHours = Number(summary.month_total_hours || 0);
  const avgHours = Number(summary.month_avg_hours || (monthCount > 0 ? (monthHours / monthCount) : 0));
  const monthTotalExpense = Number(summary.month_total_expense || 0);
  const monthTotalExpenseText = monthTotalExpense > 0 ? formatWonDisplay(monthTotalExpense).replace("원", "") : "0";

  const progressSummary = Array.isArray(data.progress_summary) ? data.progress_summary : [];
  const typeSummary = data.type_summary || [];
  const rows = Array.isArray(data.month_records) ? data.month_records : [];
  const resultHistory = Array.isArray(data.result_history) ? data.result_history : [];
  const resultDocTotal = resultHistory.reduce((sum, item) => sum + Number(item?.result_doc_count || 0), 0);
  const docStatusTotal = rows.length;
  const apiWarning = String(data.warning || "").trim();

  const scheduledCount = getTripSummaryCount(progressSummary, "scheduled");
  const ongoingCount = getTripSummaryCount(progressSummary, "ongoing");
  const finishedCount = getTripSummaryCount(progressSummary, "finished");
  const progressRejectedCount = getTripSummaryCount(progressSummary, "rejected");
  const progressTotal = Math.max(0, scheduledCount + ongoingCount + finishedCount + progressRejectedCount);
  const progressPct = (count) => progressTotal ? Math.max(0, Math.round((count / progressTotal) * 1000) / 10) : 0;

  const typeBadges = typeSummary.length
    ? typeSummary
      .map((t) => `<span class="pill">${esc(t.type || "기타")} ${Number(t.count || 0)}건</span>`)
      .join("")
    : `<span class="hint">선택한 월의 출장유형 데이터가 없습니다.</span>`;

  const tableRows = rows.length
    ? rows.map((r) => `
      <tr data-id="${r.document_id}">
        <td><span class="trip-type-badge type-${normalizeTripTypeToken(r.trip_type)}">${esc(r.trip_type || "기타")}</span></td>
        <td>${esc(r.trip_destination || "-")}</td>
        <td>${esc(fmt(r.trip_start_date))}</td>
        <td>${esc(fmt(r.trip_end_date))}</td>
        <td>${esc(r.trip_transportation || "-")}</td>
        <td>${formatWonDisplay(r.trip_expense_amount)}</td>
        <td>${formatTripDurationText(r.trip_hours)}</td>
        <td><span class="trip-progress-chip progress-${esc(r.progress_token || "finished")}">${esc(r.progress_label || "종료")}</span></td>
        <td><span class="trip-status-chip status-${esc(r.document_status || "in_review")}">${statusLabel[r.document_status] || esc(r.document_status || "-")}</span></td>
        <td>${esc(r.document_title || "-")}</td>
        <td>
          <button type="button" class="btn btn-small" data-trip-result="${r.document_id}">
            출장결과${Number(r.result_doc_count || 0) > 0 ? ` ${Number(r.result_doc_count || 0)}건` : ""}
          </button>
        </td>
      </tr>
    `).join("")
    : `<tr><td colspan="11">선택한 월의 출장 내역이 없습니다.</td></tr>`;

  const resultHistoryRows = resultHistory.length
    ? resultHistory.map((item) => {
      const latestStatus = String(item.latest_result_doc_status || "").trim();
      const latestStatusLabel = statusLabel[latestStatus] || (latestStatus || "-");
      const latestDocId = Number(item.latest_result_doc_id || 0);
      const count = Number(item.result_doc_count || 0);
      return `
        <tr>
          <td>${esc(item.source_document_title || "-")}</td>
          <td>${esc(fmt(item.trip_start_date))} ~ ${esc(fmt(item.trip_end_date))}</td>
          <td>${count}건</td>
          <td><span class="trip-status-chip status-${esc(latestStatus || "in_review")}">${esc(latestStatusLabel)}</span></td>
          <td>${esc(item.latest_result_doc_title || "-")}</td>
          <td>${esc(fmt(item.latest_result_doc_time))}</td>
          <td>
            ${latestDocId
          ? `<button type="button" class="btn btn-small" data-trip-result-doc-id="${latestDocId}">열기</button>`
          : `<span class="hint">-</span>`}
          </td>
        </tr>
      `;
    }).join("")
    : `<tr><td colspan="7">선택한 월의 출장결과 내역이 없습니다.</td></tr>`;

  host.innerHTML = `
    <h3>출장현황</h3>
    ${apiWarning ? `<p class="hint" style="color:#b45309;">${esc(apiWarning)}</p>` : ""}
    <div class="leave-mgmt-top overtime-mgmt-top trip-mgmt-top">
      <div class="overtime-left-stack trip-left-stack">
        <div class="leave-card overtime-summary-card">
          <div class="leave-title">월 출장 요약 (${selectedYear}.${String(selectedMonth).padStart(2, "0")})</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${monthCount}</div><div class="lbl">출장 건수</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthHours)}</div><div class="lbl">총 시간</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(avgHours)}</div><div class="lbl">평균 시간</div></div>
          </div>
        </div>
        <div class="leave-card overtime-rate-card trip-expense-card">
          <div class="leave-title">출장비 요약</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${monthTotalExpenseText}</div><div class="lbl">총 출장비(원)</div></div>
            <div class="leave-stat"><div class="num">${docStatusTotal}</div><div class="lbl">상신 문서 수</div></div>
            <div class="leave-stat"><div class="num">${resultDocTotal}</div><div class="lbl">결과 문서 수</div></div>
          </div>
        </div>
      </div>
      <div class="leave-calendar-card overtime-dashboard-card trip-dashboard-card">
        <div class="overtime-dashboard-head">
          <div class="leave-calendar-nav">
            <button type="button" class="btn" data-trip-nav="-1">이전</button>
            <strong>${selectedYear}년 ${selectedMonth}월</strong>
            <button type="button" class="btn" data-trip-nav="1">다음</button>
            <button type="button" class="btn" data-trip-nav="0">이번달</button>
          </div>
        </div>
        <div class="trip-summary-grid">
          <div class="trip-summary-panel">
            <div class="trip-summary-title">출장 일정 상태</div>
            <div class="trip-seg-bar">
              <i class="seg seg-scheduled" style="width:${progressPct(scheduledCount)}%"></i>
              <i class="seg seg-ongoing" style="width:${progressPct(ongoingCount)}%"></i>
              <i class="seg seg-finished" style="width:${progressPct(finishedCount)}%"></i>
              <i class="seg seg-rejected" style="width:${progressPct(progressRejectedCount)}%"></i>
            </div>
            <div class="pill-list">
              <span class="pill">예정 ${scheduledCount}건</span>
              <span class="pill">진행중 ${ongoingCount}건</span>
              <span class="pill">종료 ${finishedCount}건</span>
              <span class="pill">반려 ${progressRejectedCount}건</span>
            </div>
          </div>
          <div class="trip-summary-panel">
            <div class="trip-summary-title">출장유형 분포</div>
            <div class="pill-list">${typeBadges}</div>
          </div>
        </div>
      </div>
    </div>
    <div class="table-wrap trip-result-history-wrap" style="margin-top:12px;">
      <h4>출장결과 내역 (${selectedYear}년 ${selectedMonth}월)</h4>
      <table>
        <thead>
          <tr>
            <th>원 출장문서</th>
            <th>출장기간</th>
            <th>결과문서 수</th>
            <th>최신 상태</th>
            <th>최신 결과문서</th>
            <th>최신 업데이트</th>
            <th>열기</th>
          </tr>
        </thead>
        <tbody id="tripResultHistoryBody">${resultHistoryRows}</tbody>
      </table>
    </div>
    <div class="table-wrap" style="margin-top:12px;">
      <h4>개인 출장 사용 내역 (${selectedYear}년 ${selectedMonth}월)</h4>
      <table>
        <thead>
          <tr><th>출장종류</th><th>출장지</th><th>시작일시</th><th>종료일시</th><th>교통수단</th><th>출장비</th><th>시간</th><th>일정상태</th><th>문서상태</th><th>문서제목</th><th>출장결과</th></tr>
        </thead>
        <tbody id="myTripBody">${tableRows}</tbody>
      </table>
    </div>
  `;

  host.querySelectorAll("[data-trip-nav]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const delta = Number(btn.dataset.tripNav || 0);
      const current = new Date(dash.year, dash.month, 1);
      if (delta === 0) {
        const now = new Date();
        dash.year = now.getFullYear();
        dash.month = now.getMonth();
      } else {
        const next = new Date(current.getFullYear(), current.getMonth() + delta, 1);
        dash.year = next.getFullYear();
        dash.month = next.getMonth();
      }
      loadTripMgmt().catch((err) => alert(err.message));
    });
  });

  const tripRowMap = new Map(rows.map((r) => [Number(r.document_id || 0), r]));
  host.querySelectorAll("button[data-trip-result]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const docId = Number(btn.dataset.tripResult || 0);
      if (!docId) return;
      const rowData = tripRowMap.get(docId);
      if (!rowData) return;
      openTripResultModal(rowData).catch((err) => {
        setTripResultModalMessage(err?.message || "출장결과 팝업을 열 수 없습니다.", true);
      });
    });
  });

  host.querySelectorAll("button[data-trip-result-doc-id]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const docId = Number(btn.dataset.tripResultDocId || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });

  host.querySelectorAll("#myTripBody tr[data-id]").forEach((tr) => {
    tr.addEventListener("click", (e) => {
      if (e.target.closest("button[data-trip-result]")) return;
      const docId = Number(tr.dataset.id || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });
}

async function loadEducationMgmt() {
  const host = $("#educationMgmt");
  if (!host) return;
  if (!state.me) {
    host.innerHTML = `<h3>교육현황</h3><p class="hint">로그인 후 조회할 수 있습니다.</p>`;
    return;
  }

  const dash = state.ui.educationDashboard || (state.ui.educationDashboard = {
    year: new Date().getFullYear(),
    month: new Date().getMonth(),
  });
  const query = new URLSearchParams({
    year: String(dash.year),
    month: String(dash.month + 1),
  });
  let data;
  try {
    data = await api(`/api/educations?${query.toString()}`);
  } catch (err) {
    host.innerHTML = `
      <h3>교육현황</h3>
      <div class="card" style="padding:14px;">
        <p class="hint">교육현황을 불러오지 못했습니다.</p>
        <p class="hint">${esc(err.message || "알 수 없는 오류")}</p>
        <div class="row-gap" style="margin-top:8px;">
          <button type="button" class="btn" id="retryEducationMgmtBtn">다시 시도</button>
        </div>
      </div>
    `;
    host.querySelector("#retryEducationMgmtBtn")?.addEventListener("click", () => {
      loadEducationMgmt().catch((e) => alert(e.message));
    });
    return;
  }

  const selectedYear = Number(data.selection?.year || dash.year);
  const selectedMonth = Math.max(1, Number(data.selection?.month || (dash.month + 1)));
  dash.year = selectedYear;
  dash.month = selectedMonth - 1;

  const summary = data.summary || {};
  const monthCount = Number(summary.month_count || 0);
  const monthTotalHours = Number(summary.month_total_hours || 0);
  const monthAvgHours = Number(summary.month_avg_hours || (monthCount > 0 ? (monthTotalHours / monthCount) : 0));
  const monthTotalAmount = Number(summary.month_total_amount || 0);

  const statusSummary = Array.isArray(data.status_summary) ? data.status_summary : [];
  const statusInReview = getTripSummaryCount(statusSummary, "in_review");
  const statusApproved = getTripSummaryCount(statusSummary, "approved");
  const statusRejected = getTripSummaryCount(statusSummary, "rejected");
  const statusTotal = Math.max(0, statusInReview + statusApproved + statusRejected);
  const statusPct = (count) => statusTotal ? Math.max(0, Math.round((count / statusTotal) * 1000) / 10) : 0;

  const categorySummary = Array.isArray(data.category_summary) ? data.category_summary : [];
  const categoryBadges = categorySummary.length
    ? categorySummary.map((item) => `<span class="pill">${esc(item.category || "기타")} ${Number(item.count || 0)}건</span>`).join("")
    : `<span class="hint">선택한 월의 교육분류 데이터가 없습니다.</span>`;

  const rows = Array.isArray(data.month_records) ? data.month_records : [];
  const rowsHtml = rows.length ? rows.map((r) => `
      <tr data-id="${Number(r.document_id || 0)}">
        <td>${esc(r.education_title || "-")}</td>
        <td>${esc(r.education_category || "-")}</td>
        <td>${esc(r.education_provider || "-")}</td>
        <td>${esc(r.education_period || `${fmt(r.education_start_date)} ~ ${fmt(r.education_end_date)}`)}</td>
        <td>${formatHoursDisplay(r.education_hours || 0)}h</td>
        <td>${formatWonDisplay(r.education_total_amount || 0)}</td>
        <td>${statusLabel[r.document_status] || esc(r.document_status || "-")}</td>
        <td>${esc(r.document_title || "-")}</td>
        <td><button type="button" class="btn ghost" data-education-result="${Number(r.document_id || 0)}">교육결과</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="9">선택한 월의 교육 내역이 없습니다.</td></tr>`;
  const resultHistory = Array.isArray(data.result_history) ? data.result_history : [];
  const resultHistoryRows = resultHistory.length
    ? resultHistory.map((item) => {
      const latestStatus = String(item.latest_result_doc_status || "").trim();
      return `
        <tr>
          <td>${esc(item.source_document_title || "-")}</td>
          <td>${fmt(item.education_start_date)} ~ ${fmt(item.education_end_date)}</td>
          <td>${Number(item.result_doc_count || 0)}건</td>
          <td>${statusLabel[latestStatus] || esc(latestStatus || "-")}</td>
          <td>${esc(item.latest_result_doc_title || "-")}</td>
          <td>${fmt(item.latest_result_doc_time || "-")}</td>
          <td>
            ${item.latest_result_doc_id ? `<button type="button" class="btn ghost" data-education-result-doc-id="${Number(item.latest_result_doc_id || 0)}">열기</button>` : "-"}
          </td>
        </tr>
      `;
    }).join("")
    : `<tr><td colspan="7">선택한 월의 교육결과 내역이 없습니다.</td></tr>`;

  host.innerHTML = `
    <h3>교육현황</h3>
    <div class="leave-mgmt-top overtime-mgmt-top trip-mgmt-top">
      <div class="overtime-left-stack trip-left-stack">
        <div class="leave-card overtime-summary-card">
          <div class="leave-title">월 교육 요약 (${selectedYear}.${String(selectedMonth).padStart(2, "0")})</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${monthCount}</div><div class="lbl">교육 건수</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthTotalHours)}</div><div class="lbl">총 시간</div></div>
            <div class="leave-stat"><div class="num">${formatHoursDisplay(monthAvgHours)}</div><div class="lbl">평균 시간</div></div>
          </div>
        </div>
        <div class="leave-card overtime-rate-card trip-expense-card">
          <div class="leave-title">교육비 요약</div>
          <div class="leave-stats">
            <div class="leave-stat"><div class="num">${monthTotalAmount > 0 ? formatWonDisplay(monthTotalAmount).replace("원", "") : "0"}</div><div class="lbl">총 교육비(원)</div></div>
            <div class="leave-stat"><div class="num">${rows.length}</div><div class="lbl">상신 문서 수</div></div>
          </div>
        </div>
      </div>
      <div class="leave-calendar-card overtime-dashboard-card trip-dashboard-card">
        <div class="overtime-dashboard-head">
          <div class="leave-calendar-nav">
            <button type="button" class="btn" data-edu-nav="-1">이전</button>
            <strong>${selectedYear}년 ${selectedMonth}월</strong>
            <button type="button" class="btn" data-edu-nav="1">다음</button>
            <button type="button" class="btn" data-edu-nav="0">이번달</button>
          </div>
        </div>
        <div class="trip-summary-grid">
          <div class="trip-summary-panel">
            <div class="trip-summary-title">문서 상태</div>
            <div class="trip-seg-bar">
              <i class="seg seg-ongoing" style="width:${statusPct(statusInReview)}%"></i>
              <i class="seg seg-finished" style="width:${statusPct(statusApproved)}%"></i>
              <i class="seg seg-rejected" style="width:${statusPct(statusRejected)}%"></i>
            </div>
            <div class="pill-list">
              <span class="pill">진행중 ${statusInReview}건</span>
              <span class="pill">승인 ${statusApproved}건</span>
              <span class="pill">반려 ${statusRejected}건</span>
            </div>
          </div>
          <div class="trip-summary-panel">
            <div class="trip-summary-title">교육분류 분포</div>
            <div class="pill-list">${categoryBadges}</div>
          </div>
        </div>
      </div>
    </div>
    <div class="table-wrap trip-result-history-wrap" style="margin-top:12px;">
      <h4>교육결과 내역 (${selectedYear}년 ${selectedMonth}월)</h4>
      <table>
        <thead>
          <tr>
            <th>원 교육문서</th>
            <th>교육기간</th>
            <th>결과문서 수</th>
            <th>최신 상태</th>
            <th>최신 결과문서</th>
            <th>최신 업데이트</th>
            <th>열기</th>
          </tr>
        </thead>
        <tbody id="educationResultHistoryBody">${resultHistoryRows}</tbody>
      </table>
    </div>
    <div class="table-wrap" style="margin-top:12px;">
      <h4>개인 교육 사용 내역 (${selectedYear}년 ${selectedMonth}월)</h4>
      <table>
        <thead>
          <tr><th>교육명</th><th>교육분류</th><th>교육기관</th><th>교육기간</th><th>시간</th><th>총 교육비</th><th>문서상태</th><th>문서제목</th><th>교육결과</th></tr>
        </thead>
        <tbody id="myEducationBody">${rowsHtml}</tbody>
      </table>
    </div>
  `;

  host.querySelectorAll("[data-edu-nav]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const delta = Number(btn.dataset.eduNav || 0);
      const current = new Date(dash.year, dash.month, 1);
      if (delta === 0) {
        const now = new Date();
        dash.year = now.getFullYear();
        dash.month = now.getMonth();
      } else {
        const next = new Date(current.getFullYear(), current.getMonth() + delta, 1);
        dash.year = next.getFullYear();
        dash.month = next.getMonth();
      }
      loadEducationMgmt().catch((err) => alert(err.message));
    });
  });

  const educationRowMap = new Map(rows.map((r) => [Number(r.document_id || 0), r]));
  host.querySelectorAll("button[data-education-result]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const docId = Number(btn.dataset.educationResult || 0);
      if (!docId) return;
      const rowData = educationRowMap.get(docId);
      if (!rowData) return;
      openEducationResultModal(rowData).catch((err) => {
        setEducationResultModalMessage(err?.message || "교육결과 팝업을 열 수 없습니다.", true);
      });
    });
  });

  host.querySelectorAll("button[data-education-result-doc-id]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const docId = Number(btn.dataset.educationResultDocId || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });

  host.querySelectorAll("#myEducationBody tr[data-id]").forEach((tr) => {
    tr.addEventListener("click", (e) => {
      if (e.target.closest("button[data-education-result]")) return;
      const docId = Number(tr.dataset.id || 0);
      if (!docId) return;
      openDocument(docId).catch((err) => alert(err.message));
    });
  });
}


async function handleDashboardUserSubmit() {
  const username = $("#dashNewUserUsername")?.value.trim() || "";
  const password = $("#dashNewUserPassword")?.value || "";
  const fullName = $("#dashNewUserFullName")?.value.trim() || "";
  const department = $("#dashNewUserDepartment")?.value.trim() || "";
  const role = $("#dashNewUserRole")?.value || "employee";
  const msg = $("#dashUserAdminMsg");

  if (!msg) return;
  msg.textContent = "";

  const payload = {
    username,
    password,
    full_name: fullName,
    department,
    role,
  };

  try {
    if (dashEditingUserId) {
      await api(`/api/users/${dashEditingUserId}/update`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      msg.textContent = "사용자 정보 수정 완료";
      cancelDashboardEdit();
    } else {
      await api("/api/users", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      msg.textContent = "사용자 등록 완료";
      $("#dashNewUserUsername").value = "";
      $("#dashNewUserPassword").value = "";
      $("#dashNewUserFullName").value = "";
      $("#dashNewUserDepartment").value = "";
      $("#dashNewUserRole").value = "employee";
    }
    msg.style.color = "var(--accent-2)";
    await loadUsers();
  } catch (err) {
    msg.textContent = err.message;
    msg.style.color = "var(--danger)";
  }
}
async function deleteUserFromDashboard(userId) {
  const msg = $("#dashUserAdminMsg");
  if (!window.confirm("정말 이 사용자를 삭제하시겠습니까? 삭제 후 복구할 수 없습니다.")) return;
  try {
    await api(`/api/users/${userId}/delete`, { method: "POST", body: "{}" });
    if (msg) {
      msg.textContent = "사용자 삭제 완료";
      msg.style.color = "var(--accent-2)";
    }
    await loadUsers();
  } catch (err) {
    alert(`삭제 실패: ${err.message}`);
    if (msg) {
      msg.textContent = `삭제 실패: ${err.message}`;
      msg.style.color = "var(--danger)";
    }
  }
}

function renderDocRows(hostId, docs) {
  const host = $(hostId);
  host.innerHTML = "";
  if (!docs.length) {
    host.innerHTML = `<tr><td colspan="6">문서가 없습니다.</td></tr>`;
    return;
  }
  for (const doc of docs) {
    const tr = document.createElement("tr");
    tr.dataset.id = String(doc.id);
    tr.innerHTML = `
      <td>${doc.id}</td>
      <td>${esc(doc.title)}</td>
      <td>${esc(doc.drafter.name)}</td>
      <td>${statusLabel[doc.status] || doc.status}</td>
      <td>${doc.current_step ? esc(doc.current_step.approver_name) : "-"}</td>
      <td>${fmt(doc.created_at)}</td>
    `;
    host.appendChild(tr);
  }
}

function getCurrentDocFilter() {
  return $("#docFilterButtons .doc-filter-btn.is-active")?.dataset.docFilter || "all";
}

function setCurrentDocFilter(filterValue) {
  $$("#docFilterButtons .doc-filter-btn").forEach((btn) => {
    const active = btn.dataset.docFilter === filterValue;
    btn.classList.toggle("is-active", active);
    btn.setAttribute("aria-pressed", active ? "true" : "false");
  });
}

async function loadPending() {
  if (!$("#pendingBody")) return;
  const data = await api("/api/approvals/pending");
  renderDocRows("#pendingBody", data.documents);
}

async function loadMyDocs() {
  const filter = getCurrentDocFilter();
  const search = $("#docSearch").value.trim();
  const params = new URLSearchParams();
  if (filter === "archived") {
    params.set("archived", "1");
  } else if (filter === "all_completed") {
    params.set("all_completed", "1");
  } else if (filter === "pending_me") {
    params.set("pending_me", "1");
  } else {
    params.set("mine", "1");
    if (filter !== "all") params.set("status", filter);
  }
  if (search) params.set("search", search);
  const data = await api(`/api/documents?${params.toString()}`);

  const host = $("#myDocsBody");
  host.innerHTML = "";
  if (!data.documents.length) {
    host.innerHTML = `<tr><td colspan="6">문서가 없습니다.</td></tr>`;
    return;
  }
  for (const doc of data.documents) {
    const tr = document.createElement("tr");
    tr.dataset.id = String(doc.id);
    tr.dataset.canOpen = doc.can_open === false ? "false" : "true";
    const titleText = `${doc.title}${doc.can_open === false ? " [열람제한]" : ""}`;
    const displayStatus = (doc.returned_for_resubmit && doc.status === "in_review")
      ? "반려(재상신 필요)"
      : (statusLabel[doc.status] || doc.status);
    tr.innerHTML = `
      <td>${doc.id}</td>
      <td>${esc(titleText)}</td>
      <td>${esc(templateTypeLabel[doc.template_type] || doc.template_type)}</td>
      <td>${esc(displayStatus)}</td>
      <td>${esc(doc.drafter.name)}</td>
      <td>${fmt(doc.updated_at)}</td>
    `;
    host.appendChild(tr);
  }
}

async function loadNotices() {
  const list = $("#noticeList");
  const noticeForm = $("#noticeForm");
  if (!list || !noticeForm) return;
  const data = await api("/api/notices");
  list.innerHTML = "";

  noticeForm.classList.toggle("hidden", state.me.role !== "admin");

  for (const n of data.notices) {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="row-gap">
        <strong>${esc(n.title)}</strong>
        ${n.pinned ? '<span class="badge pinned">고정</span>' : ""}
      </div>
      <div>${esc(n.content)}</div>
      <div class="meta">${esc(n.author_name)} | ${fmt(n.created_at)}</div>
    `;
    list.appendChild(li);
  }
}

async function loadSchedules() {
  const list = $("#scheduleList");
  if (!list) return;
  const data = await api("/api/schedules");
  list.innerHTML = "";
  for (const s of data.schedules) {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="row-gap">
        <strong>${esc(s.title)}</strong>
        <span class="badge">${esc(s.event_type)}</span>
      </div>
      <div>${esc(s.start_date)} ~ ${esc(s.end_date)}</div>
      <div class="meta">담당: ${esc(s.owner.name)}${s.resource_name ? ` | 자원: ${esc(s.resource_name)}` : ""}</div>
    `;
    list.appendChild(li);
  }
}

async function loadNotifications() {
  const list = $("#notificationList");
  if (!list) return;
  const data = await api("/api/notifications");
  list.innerHTML = "";
  for (const n of data.notifications) {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="row-gap">
        <strong>${esc(n.message)}</strong>
        ${n.is_read ? "" : '<span class="badge unread">미확인</span>'}
      </div>
      <div class="meta">${fmt(n.created_at)}</div>
    `;
    list.appendChild(li);
  }
}

function setEditorFormMode() {
  const providerEl = $("#editorProvider");
  if (providerEl && providerEl.value !== "google_docs") {
    providerEl.value = "google_docs";
  }
  const isGoogle = true;
  $("#googleDocUrlWrap").classList.toggle("hidden", !isGoogle);
  $("#googleDocPanel").classList.toggle("hidden", !isGoogle);
  $("#contentWrap").classList.toggle("hidden", isGoogle);
  if (isGoogle) {
    updateGoogleDocComposePanel(false);
    setGoogleControlState("");

    // Note: Auto-popup is blocked by browser's trusted event requirement.
    // User must click the "Google 연결" button manually.
  }
}

function formatBytes(size) {
  const n = Number(size || 0);
  if (!n) return "0B";
  if (n < 1024) return `${n}B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)}KB`;
  return `${(n / (1024 * 1024)).toFixed(1)}MB`;
}

function renderDraftAttachmentSelection() {
  const list = $("#draftAttachmentList");
  const input = $("#draftAttachments");
  if (!list || !input) return;
  const files = Array.from(input.files || []);
  if (!files.length) {
    list.innerHTML = "";
    return;
  }
  list.innerHTML = files
    .map((f) => `<li>${esc(f.name)} <span class="meta">(${formatBytes(f.size)})</span></li>`)
    .join("");
}

async function fileToBase64Data(file) {
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("파일 읽기 실패"));
    reader.readAsDataURL(file);
  });
  const idx = dataUrl.indexOf(",");
  if (idx < 0) throw new Error("파일 인코딩 실패");
  return dataUrl.slice(idx + 1);
}

async function uploadDraftAttachments(files) {
  if (!files || files.length === 0) return [];
  const payloadFiles = [];
  for (const file of files) {
    payloadFiles.push({
      name: file.name,
      type: file.type || "application/octet-stream",
      size: file.size || 0,
      data: await fileToBase64Data(file),
      folder_id: GOOGLE_ATTACHMENT_FOLDER_ID,
    });
  }
  const data = await api("/api/integrations/google/upload-attachments", {
    method: "POST",
    body: JSON.stringify({ files: payloadFiles }),
  });
  return Array.isArray(data.attachments) ? data.attachments : [];
}

async function saveDraft(submitNow) {
  const form = $("#draftForm");
  const editorProvider = "google_docs";
  const externalDocUrl = form.external_doc_url.value.trim();

  const msg = $("#draftMsg");
  msg.textContent = "";
  const maxApprovers = getDraftMaxApproverSteps();
  const totalSlots = getDraftApprovalTemplateTotalSlots();
  if (state.approverIds.length > maxApprovers) {
    msg.textContent = `현재 템플릿은 결재자 최대 ${maxApprovers}명까지 지원합니다. (기안자 포함 ${totalSlots}칸)`;
    return;
  }
  if (!externalDocUrl) {
    msg.textContent = "Google Docs 링크 또는 문서 ID를 입력해 주세요.";
    return;
  }
  if ((form.template_type.value || "") === "leave") {
    syncLeaveDaysAutoCalc();
    if (!form.leave_start_date.value || !form.leave_end_date.value) {
      msg.textContent = "휴가계 문서는 휴가 시작일/종료일을 입력해 주세요.";
      return;
    }
    if (!String(form.leave_type?.value || "").trim()) {
      msg.textContent = "휴가계 문서는 휴가형태를 선택해 주세요.";
      return;
    }
  }
  if ((form.template_type.value || "") === "overtime") {
    syncOvertimeHoursAutoCalc();
    if (!form.overtime_start_date.value || !form.overtime_end_date.value) {
      msg.textContent = "연장근로 문서는 연장근로 시작일시/종료일시를 입력해 주세요.";
      return;
    }
    if (!String(form.overtime_type?.value || "").trim()) {
      msg.textContent = "연장근로 문서는 형태를 선택해 주세요.";
      return;
    }
    if (!String(form.overtime_content?.value || "").trim()) {
      msg.textContent = "연장근로 문서는 연장근로 내용을 입력해 주세요.";
      return;
    }
  }
  if ((form.template_type.value || "") === "business_trip") {
    if (!form.trip_start_date.value || !form.trip_end_date.value) {
      msg.textContent = "출장신청서 문서는 출장 시작일시/종료일시를 입력해 주세요.";
      return;
    }
    const tripStart = parseLocalDateTimeInput(form.trip_start_date.value);
    const tripEnd = parseLocalDateTimeInput(form.trip_end_date.value);
    if (!tripStart || !tripEnd || tripEnd <= tripStart) {
      msg.textContent = "출장 종료일시는 시작일시보다 뒤여야 합니다.";
      return;
    }
    if (!String(form.trip_type?.value || "").trim()) {
      msg.textContent = "출장신청서 문서는 출장종류를 선택해 주세요.";
      return;
    }
    if (!String(form.trip_destination?.value || "").trim()) {
      msg.textContent = "출장신청서 문서는 출장지를 입력해 주세요.";
      return;
    }
    if (!String(form.trip_purpose?.value || "").trim()) {
      msg.textContent = "출장신청서 문서는 출장목적을 입력해 주세요.";
      return;
    }
  }
  if ((form.template_type.value || "") === "education") {
    if (!form.education_start_date.value || !form.education_end_date.value) {
      msg.textContent = "교육신청서 문서는 교육 시작일시/종료일시를 입력해 주세요.";
      return;
    }
    const eduStart = parseLocalDateTimeInput(form.education_start_date.value);
    const eduEnd = parseLocalDateTimeInput(form.education_end_date.value);
    if (!eduStart || !eduEnd || eduEnd <= eduStart) {
      msg.textContent = "교육 종료일시는 시작일시보다 뒤여야 합니다.";
      return;
    }
    if (!String(form.education_title?.value || "").trim()) {
      msg.textContent = "교육신청서 문서는 교육명을 입력해 주세요.";
      return;
    }
    if (!String(form.education_category?.value || "").trim()) {
      msg.textContent = "교육신청서 문서는 교육분류를 선택해 주세요.";
      return;
    }
    if (!String(form.education_provider?.value || "").trim()) {
      msg.textContent = "교육신청서 문서는 교육기관을 입력해 주세요.";
      return;
    }
    if (!String(form.education_location?.value || "").trim()) {
      msg.textContent = "교육신청서 문서는 교육장소를 입력해 주세요.";
      return;
    }
    if (!String(form.education_purpose?.value || "").trim()) {
      msg.textContent = "교육신청서 문서는 교육목적을 입력해 주세요.";
      return;
    }
  }
  const attachmentFiles = Array.from(form.querySelector("#draftAttachments")?.files || []);
  if (attachmentFiles.length > 10) {
    msg.textContent = "붙임 파일은 최대 10개까지 업로드할 수 있습니다.";
    return;
  }

  const payload = {
    title: form.title.value.trim(),
    template_type: form.template_type.value,
    editor_provider: editorProvider,
    external_doc_url: externalDocUrl || null,
    priority: form.priority.value,
    recipient_text: form.recipient_text.value.trim(),
    issue_department: form.issue_department?.value || "",
    issue_year: String(form.issue_year?.value || "").trim(),
    leave_type: String(form.leave_type?.value || "").trim() || null,
    leave_start_date: form.leave_start_date?.value || null,
    leave_end_date: form.leave_end_date?.value || null,
    leave_days: String(form.leave_days?.value || "").trim() || null,
    leave_reason: String(form.leave_reason?.value || "").trim() || null,
    leave_substitute_name: String(form.leave_substitute_name?.value || "").trim() || null,
    leave_substitute_work: String(form.leave_substitute_work?.value || "").trim() || null,
    overtime_type: String(form.overtime_type?.value || "").trim() || null,
    overtime_start_date: form.overtime_start_date?.value || null,
    overtime_end_date: form.overtime_end_date?.value || null,
    overtime_hours: String(form.overtime_hours?.value || "").trim() || null,
    overtime_content: String(form.overtime_content?.value || "").trim() || null,
    overtime_etc: String(form.overtime_etc?.value || "").trim() || null,
    trip_department: String(form.trip_department?.value || "").trim() || null,
    trip_job_title: String(form.trip_job_title?.value || "").trim() || null,
    trip_name: String(form.trip_name?.value || "").trim() || null,
    trip_type: String(form.trip_type?.value || "").trim() || null,
    trip_destination: String(form.trip_destination?.value || "").trim() || null,
    trip_start_date: form.trip_start_date?.value || null,
    trip_end_date: form.trip_end_date?.value || null,
    trip_transportation: String(form.trip_transportation?.value || "").trim() || null,
    trip_expense: String(form.trip_expense?.value || "").trim() || null,
    trip_purpose: String(form.trip_purpose?.value || "").trim() || null,
    education_department: String(form.education_department?.value || "").trim() || null,
    education_job_title: String(form.education_job_title?.value || "").trim() || null,
    education_name: String(form.education_name?.value || "").trim() || null,
    education_title: String(form.education_title?.value || "").trim() || null,
    education_category: String(form.education_category?.value || "").trim() || null,
    education_provider: String(form.education_provider?.value || "").trim() || null,
    education_location: String(form.education_location?.value || "").trim() || null,
    education_start_date: form.education_start_date?.value || null,
    education_end_date: form.education_end_date?.value || null,
    education_purpose: String(form.education_purpose?.value || "").trim() || null,
    education_tuition_detail: String(form.education_tuition_detail?.value || "").trim() || null,
    education_tuition_amount: String(form.education_tuition_amount?.value || "").trim() || null,
    education_material_detail: String(form.education_material_detail?.value || "").trim() || null,
    education_material_amount: String(form.education_material_amount?.value || "").trim() || null,
    education_transport_detail: String(form.education_transport_detail?.value || "").trim() || null,
    education_transport_amount: String(form.education_transport_amount?.value || "").trim() || null,
    education_other_detail: String(form.education_other_detail?.value || "").trim() || null,
    education_other_amount: String(form.education_other_amount?.value || "").trim() || null,
    education_budget_subject: String(form.education_budget_subject?.value || "").trim() || null,
    education_funding_source: String(form.education_funding_source?.value || "").trim() || null,
    education_payment_method: String(form.education_payment_method?.value || "").trim() || null,
    education_support_budget: String(form.education_support_budget?.value || "").trim() || null,
    education_used_budget: String(form.education_used_budget?.value || "").trim() || null,
    education_remain_budget: String(form.education_remain_budget?.value || "").trim() || null,
    education_companion: String(form.education_companion?.value || "").trim() || null,
    education_ordered: String(form.education_ordered?.value || "").trim() || null,
    education_suggestion: String(form.education_suggestion?.value || "").trim() || null,
    visibility_scope: form.visibility_scope.value || "private",
    due_date: form.due_date.value || null,
    content: "",
    approver_ids: [...state.approverIds],
    reference_ids: [...state.referenceIds],
    attachments: [],
    submit: submitNow,
  };

  try {
    if (attachmentFiles.length) {
      msg.textContent = `붙임 파일 업로드 중... (${attachmentFiles.length}개)`;
      payload.attachments = await uploadDraftAttachments(attachmentFiles);
    }
    const data = await api("/api/documents", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    const warnings = Array.isArray(data.warnings) ? data.warnings.filter((w) => String(w || "").trim()) : [];
    msg.textContent = `문서 #${data.document.id} ${submitNow ? "상신 완료" : "저장 완료"}${warnings.length ? `\n경고: ${warnings.join(" | ")}` : ""}`;
    form.reset();
    state.approverIds = [];
    state.referenceIds = [];
    renderAssigneeLists();
    setEditorFormMode();
    syncDraftTemplateTypeUI();
    closeLeaveInfoModal();
    closeOvertimeInfoModal();
    closeBusinessTripInfoModal();
    closeEducationInfoModal();
    updateGoogleDocComposePanel(true);
    renderDraftAttachmentSelection();
    applyDraftIssueDefaults(true);
    await refreshAll();
    await openDocument(data.document.id);
  } catch (err) {
    msg.textContent = err.message;
  }
}

function userHasPendingStep(doc) {
  return (doc.approval_steps || []).some((s) => s.approver.id === state.me.id && s.status === "pending");
}

function closeDetailPanel() {
  const panel = $("#detailPanel");
  if (!panel) return;
  panel.classList.add("hidden");
  const printBtn = $("#detailPrintBtn");
  if (printBtn) {
    printBtn.classList.add("hidden");
    printBtn.onclick = null;
  }
  syncBodyModalOpenClass();
  const last = state.ui?.lastFocusedBeforeDetail;
  state.ui.lastFocusedBeforeDetail = null;
  if (last && typeof last.focus === "function" && document.contains(last)) {
    try { last.focus(); } catch (_) { }
  }
}

function openDetailPanel() {
  const panel = $("#detailPanel");
  if (!panel) return;
  state.ui.lastFocusedBeforeDetail = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  panel.classList.remove("hidden");
  syncBodyModalOpenClass();
  const focusTarget = $("#closeDetail") || $("#detailModalCard");
  if (focusTarget && typeof focusTarget.focus === "function") {
    setTimeout(() => {
      try { focusTarget.focus(); } catch (_) { }
    }, 0);
  }
}

function trapDetailPanelTabKey(e) {
  if (e.key !== "Tab") return;
  const panel = $("#detailPanel");
  if (!panel || panel.classList.contains("hidden")) return;

  const focusables = Array.from(panel.querySelectorAll(
    'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
  )).filter((el) => el.offsetParent !== null || el === document.activeElement);

  if (!focusables.length) return;
  const first = focusables[0];
  const last = focusables[focusables.length - 1];

  if (e.shiftKey && document.activeElement === first) {
    e.preventDefault();
    last.focus();
  } else if (!e.shiftKey && document.activeElement === last) {
    e.preventDefault();
    first.focus();
  }
}

async function openPrintForDocument(doc) {
  const isCompleted = ["approved", "rejected"].includes(String(doc?.status || ""));
  if (doc?.external_doc?.doc_id && isCompleted) {
    const popup = window.open("about:blank", "_blank");
    if (!popup) {
      alert("팝업이 차단되어 인쇄용 문서를 열지 못했습니다. 팝업 차단을 해제해 주세요.");
      return;
    }
    try {
      popup.document.write("<p style='font-family:sans-serif;padding:16px'>인쇄용 파일 준비 중...</p>");
      const headers = {};
      if (state.token) headers.Authorization = `Bearer ${state.token}`;
      const resp = await fetch(`${API_BASE}/api/documents/${doc.id}/print-binary`, {
        method: "POST",
        headers: { "Content-Type": "application/json", ...headers },
        body: JSON.stringify({ force: true }),
      });
      if (!resp.ok) {
        const errData = await resp.json().catch(() => ({}));
        throw new Error(errData.error || `인쇄용 파일 준비 실패 (${resp.status})`);
      }
      const pdfBlob = await resp.blob();
      const blobUrl = URL.createObjectURL(pdfBlob);
      popup.document.open();
      popup.document.write(`<!doctype html>
<html><head><meta charset="utf-8"><title>인쇄</title>
<style>
html,body{margin:0;height:100%;background:#111}
.bar{display:flex;gap:8px;align-items:center;padding:8px 10px;background:#1f2937;color:#fff;font:14px sans-serif}
.bar button{padding:6px 10px;border:1px solid #4b5563;background:#374151;color:#fff;border-radius:6px;cursor:pointer}
#pdfFrame{width:100%;height:calc(100% - 46px);border:0;background:#222}
</style></head>
<body>
<div class="bar">
  <span>인쇄용 PDF 준비 완료</span>
  <button id="printBtn" type="button">인쇄</button>
  <span style="opacity:.75">자동으로 인쇄창을 시도합니다.</span>
</div>
<iframe id="pdfFrame" src="${blobUrl}"></iframe>
<script>
  const frame = document.getElementById('pdfFrame');
  const doPrint = () => {
    try { frame.contentWindow && frame.contentWindow.focus && frame.contentWindow.focus(); } catch(e) {}
    try { frame.contentWindow && frame.contentWindow.print && frame.contentWindow.print(); return; } catch(e) {}
    try { window.print(); } catch(e) {}
  };
  document.getElementById('printBtn').addEventListener('click', doPrint);
  frame.addEventListener('load', () => setTimeout(doPrint, 250));
  setTimeout(doPrint, 900);
  window.addEventListener('beforeunload', () => { try { URL.revokeObjectURL('${blobUrl}'); } catch(e) {} });
</script>
</body></html>`);
      popup.document.close();
      return;
    } catch (err) {
      try {
        popup.document.open();
        popup.document.write(`<p style="font-family:sans-serif;padding:16px">인쇄용 파일 자동 인쇄 준비에 실패했습니다: ${esc(err.message || "unknown error")}</p>`);
        popup.document.close();
      } catch (_) {
        try { popup.close(); } catch (_) { }
      }
      alert(err.message || "인쇄용 파일을 생성하지 못했습니다.");
      return;
    }
  }
  if (doc?.external_doc?.doc_id) {
    const popup = window.open("about:blank", "_blank");
    if (!popup) {
      alert("팝업이 차단되어 Google Docs 문서를 열지 못했습니다. 팝업 차단을 해제해 주세요.");
      return;
    }
    popup.location.href = doc.external_doc.edit_url;
    return;
  }
  window.print();
}

function buildDetailActions(doc) {
  const host = $("#detailActions");
  host.innerHTML = "";
  const topActions = [];
  const bottomActions = [];
  const pushTop = (el) => { if (el) topActions.push(el); };
  const pushBottom = (el) => { if (el) bottomActions.push(el); };
  const flushActions = () => {
    let currentRow = null;
    for (const el of [...topActions, ...bottomActions]) {
      const isButton = el instanceof HTMLButtonElement;
      if (isButton) {
        if (!currentRow) {
          currentRow = document.createElement("div");
          currentRow.className = "detail-action-row";
          host.appendChild(currentRow);
        }
        currentRow.appendChild(el);
        continue;
      }
      currentRow = null;
      host.appendChild(el);
    }
  };

  const latestRejectedStep = doc.latest_rejection || (doc.approval_steps || [])
    .filter((s) => s && s.status === "rejected")
    .sort((a, b) => String(b.acted_at || "").localeCompare(String(a.acted_at || "")))[0] || null;
  const isRejectedReturnDoc = doc.status === "in_review" && !!latestRejectedStep && !userHasPendingStep(doc) && doc.drafter.id === state.me.id;

  if (doc.drafter.id === state.me.id && doc.status === "draft") {
    const btn = document.createElement("button");
    btn.className = "btn primary";
    btn.textContent = "상신";
    btn.onclick = async () => {
      try {
        await api(`/api/documents/${doc.id}/submit`, { method: "POST", body: "{}" });
        await openDocument(doc.id);
        await refreshAll();
      } catch (err) {
        alert(err.message);
      }
    };
    pushTop(btn);
  }

  if (doc.drafter.id === state.me.id && (doc.status === "rejected" || isRejectedReturnDoc)) {
    const resubmitBtn = document.createElement("button");
    resubmitBtn.className = "btn primary";
    resubmitBtn.textContent = "재기안 상신";
    resubmitBtn.onclick = async () => {
      try {
        await api(`/api/documents/${doc.id}/resubmit`, { method: "POST", body: "{}" });
        await openDocument(doc.id);
        await refreshAll();
      } catch (err) {
        alert(err.message);
      }
    };
    pushTop(resubmitBtn);
  }

  if (doc.is_deleted) {
    if (state.me?.role === "admin") {
      const restoreBtn = document.createElement("button");
      restoreBtn.className = "btn primary";
      restoreBtn.textContent = "복원";
      restoreBtn.onclick = async () => {
        if (!confirm("이 문서를 삭제보관함에서 복원하시겠습니까?")) return;
        try {
          await api(`/api/documents/${doc.id}/restore`, { method: "POST", body: "{}" });
          closeDetailPanel();
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(restoreBtn);

      const purgeArchivedBtn = document.createElement("button");
      purgeArchivedBtn.className = "btn danger";
      purgeArchivedBtn.textContent = "완전삭제";
      purgeArchivedBtn.onclick = async () => {
        if (!confirm("보관삭제 문서를 완전삭제하시겠습니까? 복구할 수 없습니다.")) return;
        try {
          await api(`/api/documents/${doc.id}/delete`, {
            method: "POST",
            body: JSON.stringify({ mode: "purge" }),
          });
          closeDetailPanel();
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(purgeArchivedBtn);
    }
    flushActions();
    return;
  }

  const isCompletedDoc = ["approved", "rejected"].includes(doc.status);
  const req = doc.edit_request || { status: "none" };
  const reqStatus = String(req.status || "none");
  const delReq = doc.delete_request || { status: "none" };
  const delReqStatus = String(delReq.status || "none");
  const isDecisionMaker = !!state.me && (state.me.role === "admin" || Number(req.reviewer_id || 0) === Number(state.me.id || 0));
  const isRequester = !!state.me && Number(req.requested_by || 0) === Number(state.me.id || 0);
  const isDeleteDecisionMaker = !!state.me && (state.me.role === "admin" || Number(delReq.reviewer_id || 0) === Number(state.me.id || 0));
  const isDeleteRequester = !!state.me && Number(delReq.requested_by || 0) === Number(state.me.id || 0);
  const hasEditRequestFlow = isCompletedDoc && !!doc.external_doc && !doc.is_deleted;
  const hasDeleteRequestFlow = isCompletedDoc && !doc.is_deleted;
  const hasPendingEditRequest = reqStatus === "pending";
  const hasPendingDeleteRequest = delReqStatus === "pending";
  const canRequestEdit = hasEditRequestFlow && !hasPendingDeleteRequest && !(isDecisionMaker && hasPendingEditRequest);
  const canRequestDelete = hasDeleteRequestFlow && !hasPendingEditRequest && !(isDeleteDecisionMaker && hasPendingDeleteRequest);
  const isRequesterEditInputContext = canRequestEdit
    && (doc.drafter.id === state.me.id || isRequester || ["none", "rejected", "closed"].includes(reqStatus));
  const isDecisionMakerEditInputContext = hasEditRequestFlow && hasPendingEditRequest && isDecisionMaker;
  const isRequesterDeleteInputContext = canRequestDelete
    && (doc.drafter.id === state.me.id || isDeleteRequester || ["none", "rejected"].includes(delReqStatus));
  const isDecisionMakerDeleteInputContext = hasDeleteRequestFlow && hasPendingDeleteRequest && isDeleteDecisionMaker;
  let editReviewerSelect = null;

  const hasPendingApproval = userHasPendingStep(doc);
  const hasRequestContext = isDecisionMakerEditInputContext || isRequesterEditInputContext || isDecisionMakerDeleteInputContext || isRequesterDeleteInputContext;
  const showCommentButton = !hasRequestContext && doc.status !== "rejected" && !isRejectedReturnDoc;
  const needsTextInput = hasPendingApproval || hasRequestContext || showCommentButton;

  let ta = null;
  if (needsTextInput) {
    const taWrap = document.createElement("div");
    taWrap.className = "row-gap";
    const taLabel = document.createElement("span");
    taLabel.className = "hint";
    taLabel.textContent = "의견 / 반려 사유";
    ta = document.createElement("textarea");
    ta.placeholder = "의견 또는 반려 사유 입력";
    ta.rows = 3;

    if (isDecisionMakerDeleteInputContext) {
      taLabel.textContent = "삭제요청 검토 메모 / 거절 사유";
      ta.placeholder = "삭제요청 거절 사유(또는 검토 메모) 입력";
    } else if (isRequesterDeleteInputContext) {
      taLabel.textContent = "삭제 사유";
      ta.placeholder = delReqStatus === "rejected" ? "삭제 재요청 사유 입력" : "삭제요청 사유 입력";
    } else if (isDecisionMakerEditInputContext) {
      taLabel.textContent = "수정요청 검토 메모 / 거절 사유";
      ta.placeholder = "수정요청 거절 사유(또는 검토 메모) 입력";
    } else if (isRequesterEditInputContext) {
      taLabel.textContent = "수정 사유";
      ta.placeholder = reqStatus === "closed" ? "수정 재요청 사유 입력" : "수정요청 사유 입력";
    }
    taWrap.appendChild(taLabel);
    taWrap.appendChild(ta);
    pushTop(taWrap);
  }

  if (showCommentButton) {
    const addCommentBtn = document.createElement("button");
    addCommentBtn.className = "btn";
    addCommentBtn.textContent = "코멘트 등록";
    addCommentBtn.onclick = async () => {
      try {
        await api(`/api/documents/${doc.id}/actions`, {
          method: "POST",
          body: JSON.stringify({ action: "comment", comment: ta?.value || "" }),
        });
        if (ta) ta.value = "";
        await openDocument(doc.id);
        await Promise.all([loadNotifications(), loadDashboard()]);
      } catch (err) {
        alert(err.message);
      }
    };
    pushTop(addCommentBtn);
  }

  if (isCompletedDoc && reqStatus === "pending" && isDecisionMaker) {
    const reasonWrap = document.createElement("div");
    reasonWrap.className = "card";
    reasonWrap.style.padding = "10px 12px";
    const title = document.createElement("div");
    title.className = "hint";
    title.style.marginBottom = "6px";
    title.textContent = "수정요청 사유";
    const reason = document.createElement("pre");
    reason.textContent = String(req.reason || "사유가 입력되지 않았습니다.");
    reason.style.margin = "0";
    reasonWrap.appendChild(title);
    reasonWrap.appendChild(reason);
    pushTop(reasonWrap);
    if (ta) ta.placeholder = "수정요청 거절 사유(또는 검토 메모) 입력";
  }

  if (isCompletedDoc && delReqStatus === "pending" && isDeleteDecisionMaker) {
    const reasonWrap = document.createElement("div");
    reasonWrap.className = "card";
    reasonWrap.style.padding = "10px 12px";
    const title = document.createElement("div");
    title.className = "hint";
    title.style.marginBottom = "6px";
    title.textContent = "삭제요청 사유";
    const reason = document.createElement("pre");
    reason.textContent = String(delReq.reason || "사유가 입력되지 않았습니다.");
    reason.style.margin = "0";
    reasonWrap.appendChild(title);
    reasonWrap.appendChild(reason);
    pushTop(reasonWrap);
    if (ta) ta.placeholder = "삭제요청 거절 사유(또는 검토 메모) 입력";
  }

  if (isCompletedDoc && doc.external_doc && !doc.is_deleted) {
    const reviewerCandidates = (doc.approval_steps || [])
      .map((s) => s?.approver)
      .filter(Boolean)
      .filter((u, i, arr) => arr.findIndex((x) => Number(x.id) === Number(u.id)) === i)
      .filter((u) => Number(u.id) !== Number(state.me?.id || 0));
    if (reviewerCandidates.length) {
      const reviewerWrap = document.createElement("div");
      reviewerWrap.className = "row-gap";
      const reviewerLabel = document.createElement("span");
      reviewerLabel.className = "hint";
      reviewerLabel.textContent = "요청 결재자";
      editReviewerSelect = document.createElement("select");
      editReviewerSelect.className = "btn";
      const placeholderOpt = document.createElement("option");
      placeholderOpt.value = "";
      placeholderOpt.textContent = "선택";
      editReviewerSelect.appendChild(placeholderOpt);
      for (const cand of reviewerCandidates) {
        const opt = document.createElement("option");
        opt.value = String(cand.id);
        opt.textContent = `${cand.name} (${cand.department || "-"})`;
        if (req.reviewer_id && Number(req.reviewer_id) === Number(cand.id)) opt.selected = true;
        editReviewerSelect.appendChild(opt);
      }
      reviewerWrap.appendChild(reviewerLabel);
      reviewerWrap.appendChild(editReviewerSelect);
      pushTop(reviewerWrap);
    }

    const showRequestEditButton = !(isDecisionMaker && reqStatus === "pending");
    if (showRequestEditButton) {
      const requestEditBtn = document.createElement("button");
      requestEditBtn.className = "btn";
      if (reqStatus === "approved" && isRequester) {
        requestEditBtn.textContent = "수정완료 저장";
        requestEditBtn.className = "btn primary";
        requestEditBtn.disabled = false;
      } else {
        requestEditBtn.textContent =
          reqStatus === "pending" && isRequester ? "수정요청 대기중" :
            reqStatus === "closed" && isRequester ? "수정요청(재요청 가능)" :
              "수정요청";
        requestEditBtn.disabled = (reqStatus === "pending" && isRequester) || (!editReviewerSelect && (reqStatus === "none" || reqStatus === "rejected" || reqStatus === "closed"));
      }
      requestEditBtn.onclick = async () => {
        try {
          if (reqStatus === "approved" && isRequester) {
            await api(`/api/documents/${doc.id}/edit-request/complete`, {
              method: "POST",
              body: "{}",
            });
          } else {
            if (!String(ta?.value || "").trim()) {
              alert(reqStatus === "closed" ? "수정 재요청 사유를 입력해 주세요." : "수정요청 사유를 입력해 주세요.");
              ta?.focus();
              return;
            }
            const reviewerId = Number(editReviewerSelect?.value || 0);
            await api(`/api/documents/${doc.id}/edit-request`, {
              method: "POST",
              body: JSON.stringify({ reason: ta?.value || "", reviewer_id: reviewerId }),
            });
          }
          await openDocument(doc.id);
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(requestEditBtn);
    }

    const showRequestDeleteButton = !(isDeleteDecisionMaker && delReqStatus === "pending");
    if (showRequestDeleteButton) {
      const requestDeleteBtn = document.createElement("button");
      requestDeleteBtn.className = "btn danger";
      requestDeleteBtn.textContent =
        delReqStatus === "pending" && isDeleteRequester ? "삭제요청 대기중" :
          delReqStatus === "rejected" && isDeleteRequester ? "삭제요청(재요청 가능)" :
            "삭제요청";
      requestDeleteBtn.disabled =
        (delReqStatus === "pending" && isDeleteRequester)
        || (!editReviewerSelect && (delReqStatus === "none" || delReqStatus === "rejected"));
      requestDeleteBtn.onclick = async () => {
        try {
          if (!String(ta?.value || "").trim()) {
            alert(delReqStatus === "rejected" ? "삭제 재요청 사유를 입력해 주세요." : "삭제요청 사유를 입력해 주세요.");
            ta?.focus();
            return;
          }
          const reviewerId = Number(editReviewerSelect?.value || 0);
          await api(`/api/documents/${doc.id}/delete-request`, {
            method: "POST",
            body: JSON.stringify({ reason: ta?.value || "", reviewer_id: reviewerId }),
          });
          await openDocument(doc.id);
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(requestDeleteBtn);
    }

    if (isDecisionMaker && reqStatus === "pending") {
      const approveEditBtn = document.createElement("button");
      approveEditBtn.className = "btn primary";
      approveEditBtn.textContent = "수정요청 수락";
      approveEditBtn.onclick = async () => {
        try {
          await api(`/api/documents/${doc.id}/edit-request/decision`, {
            method: "POST",
            body: JSON.stringify({ approve: true }),
          });
          await openDocument(doc.id);
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(approveEditBtn);

      const rejectEditBtn = document.createElement("button");
      rejectEditBtn.className = "btn danger";
      rejectEditBtn.textContent = "수정요청 거절";
      rejectEditBtn.onclick = async () => {
        try {
          await api(`/api/documents/${doc.id}/edit-request/decision`, {
            method: "POST",
            body: JSON.stringify({ approve: false }),
          });
          await openDocument(doc.id);
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(rejectEditBtn);
    }

    if (isDeleteDecisionMaker && delReqStatus === "pending") {
      const approveDeleteBtn = document.createElement("button");
      approveDeleteBtn.className = "btn danger";
      approveDeleteBtn.textContent = "삭제요청 수락";
      approveDeleteBtn.onclick = async () => {
        try {
          await api(`/api/documents/${doc.id}/delete-request/decision`, {
            method: "POST",
            body: JSON.stringify({ approve: true }),
          });
          closeDetailPanel();
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(approveDeleteBtn);

      const rejectDeleteBtn = document.createElement("button");
      rejectDeleteBtn.className = "btn";
      rejectDeleteBtn.textContent = "삭제요청 거절";
      rejectDeleteBtn.onclick = async () => {
        try {
          await api(`/api/documents/${doc.id}/delete-request/decision`, {
            method: "POST",
            body: JSON.stringify({ approve: false }),
          });
          await openDocument(doc.id);
          await refreshAll();
        } catch (err) {
          alert(err.message);
        }
      };
      pushTop(rejectDeleteBtn);
    }
  }

  if (hasPendingApproval) {
    const approveBtn = document.createElement("button");
    approveBtn.className = "btn primary";
    approveBtn.textContent = "승인";
    approveBtn.onclick = async () => {
      try {
        await api(`/api/documents/${doc.id}/actions`, {
          method: "POST",
          body: JSON.stringify({ action: "approve", comment: ta?.value || "" }),
        });
        if (ta) ta.value = "";
        await openDocument(doc.id);
        await refreshAll();
      } catch (err) {
        alert(err.message);
      }
    };

    const rejectBtn = document.createElement("button");
    rejectBtn.className = "btn danger";
    rejectBtn.textContent = "반려";
    rejectBtn.onclick = async () => {
      try {
        await api(`/api/documents/${doc.id}/actions`, {
          method: "POST",
          body: JSON.stringify({ action: "reject", comment: ta?.value || "" }),
        });
        if (ta) ta.value = "";
        await openDocument(doc.id);
        await Promise.all([loadDashboard(), loadPending(), loadMyDocs(), loadNotifications()]);
      } catch (err) {
        alert(err.message);
      }
    };

    pushTop(approveBtn);
    pushTop(rejectBtn);
  }

  if (state.me?.role === "admin") {
    const archiveBtn = document.createElement("button");
    archiveBtn.className = "btn";
    archiveBtn.textContent = "보관삭제";
    archiveBtn.onclick = async () => {
      if (!confirm("이 문서를 보관삭제하시겠습니까? 기본 목록에서 숨겨집니다.")) return;
      try {
        await api(`/api/documents/${doc.id}/delete`, {
          method: "POST",
          body: JSON.stringify({ mode: "archive" }),
        });
        closeDetailPanel();
        await refreshAll();
      } catch (err) {
        alert(err.message);
      }
    };
    pushBottom(archiveBtn);

    const purgeBtn = document.createElement("button");
    purgeBtn.className = "btn danger";
    purgeBtn.textContent = "완전삭제";
    purgeBtn.onclick = async () => {
      if (!confirm("이 문서를 완전삭제하시겠습니까? 복구할 수 없습니다.")) return;
      try {
        await api(`/api/documents/${doc.id}/delete`, {
          method: "POST",
          body: JSON.stringify({ mode: "purge" }),
        });
        closeDetailPanel();
        await refreshAll();
      } catch (err) {
        alert(err.message);
      }
    };
    pushBottom(purgeBtn);
  }

  flushActions();
}

async function openDocument(docId) {
  const data = await api(`/api/documents/${docId}`);
  const doc = data.document;
  state.selectedDocId = doc.id;
  openDetailPanel();

  const rejectedSteps = (doc.approval_steps || [])
    .filter((s) => s && s.status === "rejected")
    .sort((a, b) => String(b.acted_at || "").localeCompare(String(a.acted_at || "")));
  const latestRejectedStep = doc.latest_rejection || rejectedSteps[0] || null;

  const lineItems = (doc.approval_steps || [])
    .map(
      (s) => `<li>
        <div><strong>${s.step_order}. ${esc(s.approver.name)}</strong> (${esc(s.approver.department)})</div>
        <div class="meta">상태: ${statusLabel[s.status] || s.status} | 처리시각: ${fmt(s.acted_at)}${s.comment ? ` | 코멘트: ${esc(s.comment)}` : ""}</div>
      </li>`
    )
    .join("");

  const comments = (doc.comments || [])
    .map((c) => `<li><strong>${esc(c.user.name)}</strong>: ${esc(c.comment)}<div class="meta">${fmt(c.created_at)}</div></li>`)
    .join("");
  const attachments = (doc.attachments || [])
    .map((a) => {
      const href = a.web_view_url || (a.file_id ? `https://drive.google.com/file/d/${encodeURIComponent(a.file_id)}/view` : "#");
      return `<li><a href="${esc(href)}" target="_blank" rel="noopener noreferrer">${esc(a.name || a.file_id)}</a>${a.size ? ` <span class="meta">(${formatBytes(a.size)})</span>` : ""}</li>`;
    })
    .join("");

  const canRefreshRejectedPreview = !!doc.external_doc
    && !!state.me
    && Number(doc.drafter?.id || 0) === Number(state.me.id || 0)
    && (doc.status === "rejected" || doc.returned_for_resubmit);

  const externalDoc = doc.external_doc
    ? `
      <h4>Google Docs</h4>
      <p class="hint">
        ${(["approved", "rejected"].includes(doc.status) && doc.external_doc.can_open_original !== true)
      ? `<span>원문 문서 열기 (수정요청 승인 후 가능)</span>`
      : `<a href="${esc(doc.external_doc.edit_url)}" target="_blank" rel="noopener noreferrer">원문 문서 열기</a>`}
        ${canRefreshRejectedPreview ? `<button type="button" class="btn" id="detailDocRefreshBtn" style="margin-left:8px;">문서 새로고침</button>` : ""}
      </p>
      <iframe id="detailGoogleDocFrame" class="gdoc-frame" src="${esc(doc.external_doc.preview_url)}" loading="lazy" referrerpolicy="no-referrer"></iframe>
    `
    : "";

  const editRequestStatusLabel = {
    none: "없음",
    pending: "대기중",
    approved: "수락",
    rejected: "거절",
    closed: "수정완료",
  };
  const deleteRequestStatusLabel = {
    none: "없음",
    pending: "대기중",
    approved: "수락(보관삭제)",
    rejected: "거절",
  };
  const editReq = doc.edit_request || { status: "none" };
  const deleteReq = doc.delete_request || { status: "none" };
  const editReqLine = (["approved", "rejected"].includes(doc.status) && doc.external_doc)
    ? `<p class="hint">수정요청: ${esc(editRequestStatusLabel[editReq.status] || editReq.status || "없음")}`
    + `${editReq.requested_by_name ? ` | 요청자: ${esc(editReq.requested_by_name)}` : ""}`
    + `${editReq.reviewer_name ? ` | 결재자: ${esc(editReq.reviewer_name)}` : ""}`
    + `${editReq.requested_at ? ` | 요청시각: ${fmt(editReq.requested_at)}` : ""}`
    + `${editReq.decided_by_name ? ` | 처리자: ${esc(editReq.decided_by_name)}` : ""}`
    + `${editReq.decided_at ? ` | 처리시각: ${fmt(editReq.decided_at)}` : ""}`
    + `</p>`
    : "";
  const deleteReqLine = (["approved", "rejected"].includes(doc.status) && !doc.is_deleted)
    ? `<p class="hint">삭제요청: ${esc(deleteRequestStatusLabel[deleteReq.status] || deleteReq.status || "없음")}`
    + `${deleteReq.requested_by_name ? ` | 요청자: ${esc(deleteReq.requested_by_name)}` : ""}`
    + `${deleteReq.reviewer_name ? ` | 결재자: ${esc(deleteReq.reviewer_name)}` : ""}`
    + `${deleteReq.requested_at ? ` | 요청시각: ${fmt(deleteReq.requested_at)}` : ""}`
    + `${deleteReq.decided_by_name ? ` | 처리자: ${esc(deleteReq.decided_by_name)}` : ""}`
    + `${deleteReq.decided_at ? ` | 처리시각: ${fmt(deleteReq.decided_at)}` : ""}`
    + `</p>`
    : "";

  const rejectSummaryHtml = latestRejectedStep
    ? `
      <div class="detail-reject-summary">
        <h4>반려 정보</h4>
        <div class="detail-reject-grid">
          <div><span class="hint">반려자</span><strong>${esc(latestRejectedStep.approver?.name || "-")}</strong>${latestRejectedStep.approver?.department ? ` <span class="hint">(${esc(latestRejectedStep.approver.department)})</span>` : ""}</div>
          <div><span class="hint">반려시각</span><strong>${fmt(latestRejectedStep.acted_at)}</strong></div>
        </div>
        <div class="detail-reject-reason">
          <span class="hint">반려 사유</span>
          <pre>${esc(latestRejectedStep.comment || "반려 사유가 입력되지 않았습니다.")}</pre>
        </div>
      </div>
    `
    : "";

  $("#detailBody").innerHTML = `
    <p><strong>#${doc.id}</strong> ${esc(doc.title)}</p>
    <p class="hint">상태: ${((doc.status === "rejected" || doc.returned_for_resubmit) ? "반려(재상신 필요)" : (statusLabel[doc.status] || doc.status))}${doc.is_deleted ? " (보관삭제)" : ""} | 양식: ${esc(templateTypeLabel[doc.template_type] || doc.template_type)} | 공개: ${esc(visibilityLabel[doc.visibility_scope] || doc.visibility_scope || "비공개")} | 작성도구: ${esc(editorLabel[doc.editor_provider] || doc.editor_provider)} | 기안: ${esc(doc.drafter.name)}</p>
    <p class="hint">생성: ${fmt(doc.created_at)} | 상신: ${fmt(doc.submitted_at)} | 완료: ${fmt(doc.completed_at)}</p>
    ${doc.issue_code ? `<p class="hint">시행번호: ${esc(doc.issue_code)}${doc.recipient_text ? ` | 시행문수신: ${esc(doc.recipient_text)}` : ""}</p>` : (doc.recipient_text ? `<p class="hint">시행문수신: ${esc(doc.recipient_text)}</p>` : "")}
    ${editReqLine}
    ${deleteReqLine}
    ${rejectSummaryHtml}
    <hr />
    <h4>내용</h4>
    <pre>${esc(doc.content || "")}</pre>
    <h4>붙임 파일</h4>
    <ul class="list">${attachments || "<li>없음</li>"}</ul>
    ${externalDoc}
    <h4>결재선</h4>
    <ul class="approval-line">${lineItems || "<li>없음</li>"}</ul>
    <h4>코멘트</h4>
    <ul class="list">${comments || "<li>등록된 코멘트가 없습니다.</li>"}</ul>
  `;

  const topPrintBtn = $("#detailPrintBtn");
  if (topPrintBtn) {
    if (["in_review", "approved", "rejected"].includes(doc.status)) {
      topPrintBtn.classList.remove("hidden");
      topPrintBtn.onclick = () => { void openPrintForDocument(doc); };
    } else {
      topPrintBtn.classList.add("hidden");
      topPrintBtn.onclick = null;
    }
  }

  const detailDocRefreshBtn = $("#detailDocRefreshBtn");
  if (detailDocRefreshBtn) {
    detailDocRefreshBtn.onclick = () => {
      const frame = $("#detailGoogleDocFrame");
      if (!frame) return;
      const currentSrc = frame.getAttribute("src") || "";
      frame.setAttribute("src", "");
      // Force iframe reload after Google Docs edit.
      setTimeout(() => frame.setAttribute("src", currentSrc), 30);
    };
  }

  buildDetailActions(doc);
}

async function loadAdminUsersOnly() {
  if (state.me.role !== "admin") return;
  const data = await api("/api/users");
  state.users = data.users;
  renderAdminUserTable();
}

let adminEditingUserId = null;

function renderAdminUserTable() {
  const tbody = $("#adminUserBody");
  if (!tbody) return;
  tbody.innerHTML = "";
  const sorted = [...state.users].sort((a, b) => a.username.localeCompare(b.username));
  for (const user of sorted) {
    const lockSelf = state.me && user.id === state.me.id;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${user.id}</td>
      <td>${esc(user.username)}</td>
      <td>${esc(user.full_name)}</td>
      <td>${esc(user.department)}</td>
      <td>${esc(user.job_title || "")}</td>
      <td>${esc(user.role)}</td>
      <td>
        <button type="button" class="btn" data-edit-user-id="${user.id}">수정</button>
        <button type="button" class="btn danger" data-delete-user-id="${user.id}" ${lockSelf ? "disabled" : ""}>
          ${lockSelf ? "본인계정" : "삭제"}
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  }
}

function editAdminUser(userId) {
  const user = getUserById(userId);
  if (!user) return;

  adminEditingUserId = userId;
  const form = $("#adminUserForm");
  form.username.value = user.username;
  form.username.disabled = true;
  form.password.value = "";
  form.password.placeholder = "변경시에만 입력";
  form.full_name.value = user.full_name;
  form.department.value = user.department;
  form.job_title.value = user.job_title || "";
  form.role.value = user.role;
  form.total_leave.value = user.total_leave || 0;
  form.used_leave.value = user.used_leave || 0;

  $("#adminUserSubmitBtn").textContent = "정보 수정";
  $("#adminUserCancelBtn").classList.remove("hidden");
  $("#adminUserTitle").textContent = "사용자 정보 수정";
}

function cancelEditMode() {
  adminEditingUserId = null;
  const form = $("#adminUserForm");
  form.username.value = "";
  form.username.disabled = false;
  form.password.value = "";
  form.password.placeholder = "6자 이상";
  form.full_name.value = "";
  form.department.value = "";
  form.job_title.value = "";
  form.role.value = "employee";
  form.total_leave.value = 15;
  form.used_leave.value = 0;
  form.profile_image.value = "";
  if (form.approval_stamp_image) {
    form.approval_stamp_image.value = "";
  }

  $("#adminUserSubmitBtn").textContent = "사용자 등록";
  $("#adminUserCancelBtn").classList.add("hidden");
  $("#adminUserTitle").textContent = "새 사용자 등록";
}

const readFileAsBase64 = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({
      name: file.name,
      type: file.type,
      data: reader.result.split(',')[1]
    });
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
};

async function handleAdminUserSubmit() {
  const form = $("#adminUserForm");
  const username = form.username.value.trim();
  const password = form.password.value;
  const fullName = form.full_name.value.trim();
  const department = form.department.value.trim();
  const jobTitle = form.job_title.value.trim();
  const role = form.role.value;
  const totalLeave = parseFloat(form.total_leave.value) || 0;
  const usedLeave = parseFloat(form.used_leave.value) || 0;
  const fileInput = form.profile_image;
  const stampFileInput = form.approval_stamp_image;
  const msg = $("#adminUserMsg");

  if (!msg) return;
  msg.textContent = "";

  let imagePayload = null;
  if (fileInput.files && fileInput.files[0]) {
    try {
      imagePayload = await readFileAsBase64(fileInput.files[0]);
    } catch (e) {
      alert("이미지 읽기 실패");
      return;
    }
  }

  let approvalStampPayload = null;
  if (stampFileInput && stampFileInput.files && stampFileInput.files[0]) {
    try {
      approvalStampPayload = await readFileAsBase64(stampFileInput.files[0]);
    } catch (e) {
      alert("결재도장 이미지 읽기 실패");
      return;
    }
  }

  const payload = {
    username,
    password,
    full_name: fullName,
    department,
    role,
    job_title: jobTitle,
    total_leave: totalLeave,
    used_leave: usedLeave,
    profile_image: imagePayload, // {name, type, data}
    approval_stamp_image: approvalStampPayload,
  };

  try {
    if (adminEditingUserId) {
      const data = await api(`/api/users/${adminEditingUserId}/update`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      if (state.me && state.me.id === adminEditingUserId) {
        state.me = data.user;
        $("#whoami").textContent = `${state.me.full_name} (${state.me.department}, ${state.me.role})`;
        await loadDashboard();
      }
      msg.textContent = "사용자 정보 수정 완료";
      cancelEditMode();
    } else {
      await api("/api/users", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      msg.textContent = "사용자 등록 완료";
      cancelEditMode(); // clear form
    }
    msg.style.color = "var(--accent-2)";
    await loadUsers();
    renderAdminUserTable(); // update table
  } catch (err) {
    msg.textContent = err.message;
    msg.style.color = "var(--danger)";
  }
}

async function deleteAdminUser(userId) {
  if (!confirm("정말 이 사용자를 삭제하시겠습니까?")) return;
  try {
    await api(`/api/users/${userId}/delete`, { method: "POST", body: "{}" });
    await loadUsers();
    renderAdminUserTable();
  } catch (err) {
    alert(err.message);
  }
}

function bindEvents() {
  ensureTripResultModalMountedToBody();
  ensureEducationResultModalMountedToBody();
  ensureUserPickerModalMountedToBody();

  $("#loginForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    $("#loginError").textContent = "";
    try {
      const data = await api("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({
          username: $("#username").value.trim(),
          password: $("#password").value,
        }),
      });
      await completeLogin(data);
    } catch (err) {
      $("#loginError").textContent = err.message;
    }
  });

  $("#logoutBtn").addEventListener("click", async () => {
    try {
      await api("/api/auth/logout", { method: "POST", body: "{}" });
    } catch (_) { }
    state.token = "";
    state.me = null;
    localStorage.removeItem("approval_token");
    resetGoogleSession();
    closeDetailPanel();
    setTabEditMode(false);
    syncTabEditControlsByRole();
    setLoggedIn(false);
    setCurrentDocFilter("all");
    $("#docFilterArchivedBtn")?.classList.add("hidden");
    initLoginGoogleAuth(true).catch(() => { });
  });

  $("#refreshBtn").addEventListener("click", () => refreshAll().catch((err) => alert(err.message)));

  $("#tabEditBtn")?.addEventListener("click", async () => {
    if (state.me?.role !== "admin") return;
    const editing = !!state.ui?.tabEdit?.editing;
    if (!editing) {
      state.ui.tabEdit.originalOrder = getCurrentTabOrder();
      setTabEditMode(true);
      return;
    }
    const btn = $("#tabEditBtn");
    if (btn) btn.disabled = true;
    try {
      await saveCurrentTabOrder();
      setTabEditMode(false);
    } catch (err) {
      alert(err.message || "탭 순서 저장에 실패했습니다.");
    } finally {
      if (btn) btn.disabled = false;
    }
  });

  $("#tabEditCancelBtn")?.addEventListener("click", () => {
    if (state.me?.role !== "admin") return;
    setTabEditMode(false, { restore: true });
  });

  $("#tabs").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-view]");
    if (!btn) return;
    if (state.ui?.tabEdit?.editing) {
      e.preventDefault();
      return;
    }
    const viewId = btn.dataset.view;
    showView(viewId);
    if (viewId === "tripMgmt") {
      loadTripMgmt().catch((err) => alert(err.message));
      return;
    }
    if (viewId === "educationMgmt") {
      loadEducationMgmt().catch((err) => alert(err.message));
    }
  });

  $("#tabs").addEventListener("dragstart", (e) => {
    if (!state.ui?.tabEdit?.editing) return;
    const btn = e.target.closest("button[data-view]");
    if (!btn || btn.classList.contains("hidden")) return;
    state.ui.tabEdit.draggingView = String(btn.dataset.view || "");
    btn.classList.add("is-dragging");
    try {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", state.ui.tabEdit.draggingView);
    } catch (_) { }
  });

  $("#tabs").addEventListener("dragover", (e) => {
    if (!state.ui?.tabEdit?.editing) return;
    const tabs = $("#tabs");
    if (!tabs) return;
    const draggingView = String(state.ui?.tabEdit?.draggingView || "");
    if (!draggingView) return;
    const draggingBtn = tabs.querySelector(`button[data-view="${draggingView}"]`);
    const targetBtn = e.target.closest("button[data-view]");
    if (!draggingBtn || !targetBtn || draggingBtn === targetBtn || targetBtn.classList.contains("hidden")) return;
    e.preventDefault();
    const rect = targetBtn.getBoundingClientRect();
    const insertBefore = e.clientX < (rect.left + rect.width / 2);
    if (insertBefore) {
      tabs.insertBefore(draggingBtn, targetBtn);
    } else {
      tabs.insertBefore(draggingBtn, targetBtn.nextSibling);
    }
  });

  $("#tabs").addEventListener("drop", (e) => {
    if (!state.ui?.tabEdit?.editing) return;
    e.preventDefault();
  });

  $("#tabs").addEventListener("dragend", () => {
    state.ui.tabEdit.draggingView = "";
    $$("#tabs button[data-view].is-dragging").forEach((btn) => btn.classList.remove("is-dragging"));
  });

  $("#saveDraftBtn").addEventListener("click", () => saveDraft(false));
  $("#submitDraftBtn").addEventListener("click", () => saveDraft(true));
  $("#draftForm [name='template_type']")?.addEventListener("change", handleDraftTemplateTypeChange);
  $("#draftForm [name='leave_start_date']")?.addEventListener("change", syncLeaveDaysAutoCalc);
  $("#draftForm [name='leave_end_date']")?.addEventListener("change", syncLeaveDaysAutoCalc);
  $("#draftForm [name='leave_start_date']")?.addEventListener("input", syncLeaveDaysAutoCalc);
  $("#draftForm [name='leave_end_date']")?.addEventListener("input", syncLeaveDaysAutoCalc);
  $("#draftForm [name='leave_type']")?.addEventListener("change", renderLeaveInfoSummary);
  $("#draftForm [name='leave_reason']")?.addEventListener("input", renderLeaveInfoSummary);
  $("#draftForm [name='leave_substitute_name']")?.addEventListener("input", renderLeaveInfoSummary);
  $("#draftForm [name='leave_substitute_work']")?.addEventListener("input", renderLeaveInfoSummary);
  $("#draftForm [name='overtime_start_date']")?.addEventListener("change", syncOvertimeHoursAutoCalc);
  $("#draftForm [name='overtime_end_date']")?.addEventListener("change", syncOvertimeHoursAutoCalc);
  $("#draftForm [name='overtime_start_date']")?.addEventListener("input", syncOvertimeHoursAutoCalc);
  $("#draftForm [name='overtime_end_date']")?.addEventListener("input", syncOvertimeHoursAutoCalc);
  $("#draftForm [name='overtime_type']")?.addEventListener("change", renderOvertimeInfoSummary);
  $("#draftForm [name='overtime_content']")?.addEventListener("input", renderOvertimeInfoSummary);
  $("#draftForm [name='overtime_etc']")?.addEventListener("input", renderOvertimeInfoSummary);
  $("#draftForm [name='trip_type']")?.addEventListener("change", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_department']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_job_title']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_name']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_destination']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_start_date']")?.addEventListener("change", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_end_date']")?.addEventListener("change", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_start_date']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_end_date']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_transportation']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_expense']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='trip_purpose']")?.addEventListener("input", renderBusinessTripInfoSummary);
  $("#draftForm [name='education_department']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_job_title']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_name']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_title']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_category']")?.addEventListener("change", renderEducationInfoSummary);
  $("#draftForm [name='education_provider']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_location']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_start_date']")?.addEventListener("change", renderEducationInfoSummary);
  $("#draftForm [name='education_end_date']")?.addEventListener("change", renderEducationInfoSummary);
  $("#draftForm [name='education_start_date']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_end_date']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_purpose']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_tuition_detail']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_tuition_amount']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_material_detail']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_material_amount']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_transport_detail']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_transport_amount']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_other_detail']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_other_amount']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_budget_subject']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_funding_source']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_payment_method']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_support_budget']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_used_budget']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_remain_budget']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_companion']")?.addEventListener("input", renderEducationInfoSummary);
  $("#draftForm [name='education_ordered']")?.addEventListener("change", renderEducationInfoSummary);
  $("#draftForm [name='education_suggestion']")?.addEventListener("input", renderEducationInfoSummary);
  $("#openLeaveInfoModalBtn")?.addEventListener("click", openLeaveInfoModal);
  $("#closeLeaveInfoModalBtn")?.addEventListener("click", closeLeaveInfoModal);
  $("#saveLeaveInfoModalBtn")?.addEventListener("click", () => {
    syncLeaveDaysAutoCalc();
    closeLeaveInfoModal();
  });
  $("#resetLeaveInfoBtn")?.addEventListener("click", () => {
    if (!confirm("입력한 휴가정보를 초기화하시겠습니까?")) return;
    resetLeaveInfoInputs();
  });
  $("#leavePresetFullDayBtn")?.addEventListener("click", () => applyLeaveTimePreset(9, 18, "연차"));
  $("#leavePresetMorningHalfBtn")?.addEventListener("click", () => applyLeaveTimePreset(9, 14, "오전반차"));
  $("#leavePresetAfternoonHalfBtn")?.addEventListener("click", () => applyLeaveTimePreset(13, 18, "오후반차"));
  $("#openOvertimeInfoModalBtn")?.addEventListener("click", openOvertimeInfoModal);
  $("#closeOvertimeInfoModalBtn")?.addEventListener("click", closeOvertimeInfoModal);
  $("#saveOvertimeInfoModalBtn")?.addEventListener("click", () => {
    syncOvertimeHoursAutoCalc();
    closeOvertimeInfoModal();
  });
  $("#resetOvertimeInfoBtn")?.addEventListener("click", () => {
    if (!confirm("입력한 연장근로 정보를 초기화하시겠습니까?")) return;
    resetOvertimeInfoInputs();
  });
  $("#openBusinessTripInfoModalBtn")?.addEventListener("click", openBusinessTripInfoModal);
  $("#closeBusinessTripInfoModalBtn")?.addEventListener("click", closeBusinessTripInfoModal);
  $("#saveBusinessTripInfoModalBtn")?.addEventListener("click", () => {
    renderBusinessTripInfoSummary();
    closeBusinessTripInfoModal();
  });
  $("#resetBusinessTripInfoBtn")?.addEventListener("click", () => {
    if (!confirm("입력한 출장 정보를 초기화하시겠습니까?")) return;
    resetBusinessTripInfoInputs();
  });
  $("#openEducationInfoModalBtn")?.addEventListener("click", openEducationInfoModal);
  $("#closeEducationInfoModalBtn")?.addEventListener("click", closeEducationInfoModal);
  $("#saveEducationInfoModalBtn")?.addEventListener("click", () => {
    renderEducationInfoSummary();
    closeEducationInfoModal();
  });
  $("#resetEducationInfoBtn")?.addEventListener("click", () => {
    if (!confirm("입력한 교육 정보를 초기화하시겠습니까?")) return;
    resetEducationInfoInputs();
  });
  $("#closeTripResultModalBtn")?.addEventListener("click", closeTripResultModal);
  $("#cancelTripResultModalBtn")?.addEventListener("click", closeTripResultModal);
  $("#openTripResultApproverPickerBtn")?.addEventListener("click", () => {
    openUserPicker("trip_result_approver").catch((err) => alert(err.message));
  });
  $("#openTripResultRefPickerBtn")?.addEventListener("click", () => {
    openUserPicker("trip_result_ref").catch((err) => alert(err.message));
  });
  $("#submitTripResultBtn")?.addEventListener("click", () => {
    submitTripResultModal().catch((err) => {
      setTripResultModalMessage(err?.message || "출장결과 상신 처리 중 오류가 발생했습니다.", true);
    });
  });
  $("#closeEducationResultModalBtn")?.addEventListener("click", closeEducationResultModal);
  $("#cancelEducationResultModalBtn")?.addEventListener("click", closeEducationResultModal);
  $("#openEducationResultApproverPickerBtn")?.addEventListener("click", () => {
    openUserPicker("education_result_approver").catch((err) => alert(err.message));
  });
  $("#openEducationResultRefPickerBtn")?.addEventListener("click", () => {
    openUserPicker("education_result_ref").catch((err) => alert(err.message));
  });
  $("#submitEducationResultBtn")?.addEventListener("click", () => {
    submitEducationResultModal().catch((err) => {
      setEducationResultModalMessage(err?.message || "교육결과 상신 처리 중 오류가 발생했습니다.", true);
    });
  });
  $("#searchBtn").addEventListener("click", () => loadMyDocs().catch((err) => alert(err.message)));
  $("#openApproverPickerBtn").addEventListener("click", () => openUserPicker("approver"));
  $("#openRefPickerBtn").addEventListener("click", () => openUserPicker("ref"));
  $("#closePickerBtn").addEventListener("click", closeUserPicker);
  $("#applyPickerBtn").addEventListener("click", applyPickerSelection);
  $("#pickerSearchInput").addEventListener("input", (e) => {
    state.picker.query = e.target.value || "";
    renderPickerRows();
  });
  $("#pickerBody").addEventListener("change", (e) => {
    const checkbox = e.target.closest("input[data-picker-user-id]");
    if (!checkbox) return;
    const userId = Number(checkbox.dataset.pickerUserId);
    if (checkbox.checked) {
      if (!state.picker.selectedIds.includes(userId)) {
        state.picker.selectedIds.push(userId);
      }
    } else {
      state.picker.selectedIds = state.picker.selectedIds.filter((id) => id !== userId);
    }
  });
  $("#userPickerModal").addEventListener("click", (e) => {
    if (e.target.id === "userPickerModal") closeUserPicker();
  });
  $("#leaveInfoModal")?.addEventListener("click", (e) => {
    if (e.target.id === "leaveInfoModal") closeLeaveInfoModal();
  });
  $("#overtimeInfoModal")?.addEventListener("click", (e) => {
    if (e.target.id === "overtimeInfoModal") closeOvertimeInfoModal();
  });
  $("#businessTripInfoModal")?.addEventListener("click", (e) => {
    if (e.target.id === "businessTripInfoModal") closeBusinessTripInfoModal();
  });
  $("#educationInfoModal")?.addEventListener("click", (e) => {
    if (e.target.id === "educationInfoModal") closeEducationInfoModal();
  });
  $("#tripResultModal")?.addEventListener("click", (e) => {
    if (e.target.id === "tripResultModal") closeTripResultModal();
  });
  $("#educationResultModal")?.addEventListener("click", (e) => {
    if (e.target.id === "educationResultModal") closeEducationResultModal();
  });
  $("#approverSelectedList").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-approver-action]");
    if (!btn) return;
    const action = btn.dataset.approverAction;
    const userId = Number(btn.dataset.userId);
    if (action === "remove") {
      state.approverIds = state.approverIds.filter((id) => id !== userId);
      renderAssigneeLists();
      return;
    }
    if (action === "up") moveApprover(userId, -1);
    if (action === "down") moveApprover(userId, 1);
  });
  $("#refSelectedList").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-ref-action='remove']");
    if (!btn) return;
    const userId = Number(btn.dataset.userId);
    state.referenceIds = state.referenceIds.filter((id) => id !== userId);
    renderAssigneeLists();
  });
  $("#tripResultApproverSelectedList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-trip-result-approver-action]");
    if (!btn) return;
    const action = btn.dataset.tripResultApproverAction;
    const userId = Number(btn.dataset.userId || 0);
    if (!userId) return;
    if (action === "remove") {
      state.ui.tripResultModal.approverIds = (state.ui.tripResultModal.approverIds || []).filter((id) => id !== userId);
      renderTripResultAssigneeLists();
      return;
    }
    if (action === "up") moveTripResultApprover(userId, -1);
    if (action === "down") moveTripResultApprover(userId, 1);
  });
  $("#tripResultRefSelectedList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-trip-result-ref-action='remove']");
    if (!btn) return;
    const userId = Number(btn.dataset.userId || 0);
    if (!userId) return;
    state.ui.tripResultModal.referenceIds = (state.ui.tripResultModal.referenceIds || []).filter((id) => id !== userId);
    renderTripResultAssigneeLists();
  });
  $("#educationResultApproverSelectedList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-education-result-approver-action]");
    if (!btn) return;
    const action = btn.dataset.educationResultApproverAction;
    const userId = Number(btn.dataset.userId || 0);
    if (!userId) return;
    if (action === "remove") {
      state.ui.educationResultModal.approverIds = (state.ui.educationResultModal.approverIds || []).filter((id) => id !== userId);
      renderEducationResultAssigneeLists();
      return;
    }
    if (action === "up") moveEducationResultApprover(userId, -1);
    if (action === "down") moveEducationResultApprover(userId, 1);
  });
  $("#educationResultRefSelectedList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-education-result-ref-action='remove']");
    if (!btn) return;
    const userId = Number(btn.dataset.userId || 0);
    if (!userId) return;
    state.ui.educationResultModal.referenceIds = (state.ui.educationResultModal.referenceIds || []).filter((id) => id !== userId);
    renderEducationResultAssigneeLists();
  });



  // Admin View Events
  $("#adminUserForm").addEventListener("submit", async (e) => {
    e.preventDefault();
    await handleAdminUserSubmit();
  });

  $("#adminUserCancelBtn")?.addEventListener("click", cancelEditMode);

  $("#adminUserBody").addEventListener("click", async (e) => {
    const editBtn = e.target.closest("button[data-edit-user-id]");
    if (editBtn) {
      editAdminUser(Number(editBtn.dataset.editUserId));
      return;
    }
    const deleteBtn = e.target.closest("button[data-delete-user-id]");
    if (deleteBtn) {
      await deleteAdminUser(Number(deleteBtn.dataset.deleteUserId));
    }
  });

  $("#docFilterButtons")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-doc-filter]");
    if (!btn) return;
    const nextFilter = btn.dataset.docFilter || "all";
    if (nextFilter === getCurrentDocFilter()) return;
    setCurrentDocFilter(nextFilter);
    loadMyDocs().catch((err) => alert(err.message));
  });

  $("#editorProvider").addEventListener("change", setEditorFormMode);
  $("#googleConnectBtn").addEventListener("click", async () => {
    try {
      // 1. Try silent token acquisition first
      await requestGoogleAccessToken("");
    } catch (err) {
      // 2. Fallback to consent popup if silent fails
      try {
        await requestGoogleAccessToken("consent");
      } catch (consentErr) {
        setGoogleControlState(consentErr.message);
      }
    }
  });
  $("#googlePickBtn").addEventListener("click", async () => {
    try {
      await openDrivePicker();
    } catch (err) {
      setGoogleControlState(err.message);
    }
  });

  $("#googleDocUrl").addEventListener("input", () => updateGoogleDocComposePanel(false));
  $("#googleDocUrl").addEventListener("blur", () => updateGoogleDocComposePanel(true));
  $("#refreshGoogleDocBtn").addEventListener("click", () => updateGoogleDocComposePanel(true));
  $("#openGoogleDocBtn").addEventListener("click", () => {
    const href = $("#openGoogleDocBtn").dataset.href;
    if (!href) {
      alert("먼저 Google Docs 링크 또는 문서 ID를 입력해 주세요.");
      return;
    }
    window.open(href, "_blank", "noopener,noreferrer");
  });

  $("#noticeForm")?.addEventListener("submit", async (e) => {
    e.preventDefault();
    const form = e.currentTarget;
    try {
      await api("/api/notices", {
        method: "POST",
        body: JSON.stringify({
          title: form.title.value,
          content: form.content.value,
          pinned: form.pinned.checked,
        }),
      });
      form.reset();
      await loadNotices();
    } catch (err) {
      alert(err.message);
    }
  });

  $("#scheduleForm")?.addEventListener("submit", async (e) => {
    e.preventDefault();
    const form = e.currentTarget;
    try {
      await api("/api/schedules", {
        method: "POST",
        body: JSON.stringify({
          title: form.title.value,
          event_type: form.event_type.value,
          start_date: form.start_date.value,
          end_date: form.end_date.value,
          resource_name: form.resource_name.value,
        }),
      });
      form.reset();
      await Promise.all([loadSchedules(), loadDashboard()]);
    } catch (err) {
      alert(err.message);
    }
  });

  $("#markReadBtn")?.addEventListener("click", async () => {
    try {
      await api("/api/notifications/read", { method: "POST", body: "{}" });
      await Promise.all([loadNotifications(), loadDashboard()]);
    } catch (err) {
      alert(err.message);
    }
  });

  $("#closeDetail").addEventListener("click", closeDetailPanel);
  $("#detailPanel")?.addEventListener("click", (e) => {
    if (e.target === $("#detailPanel")) closeDetailPanel();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    const leaveModal = $("#leaveInfoModal");
    if (leaveModal && !leaveModal.classList.contains("hidden")) {
      closeLeaveInfoModal();
      return;
    }
    const overtimeModal = $("#overtimeInfoModal");
    if (overtimeModal && !overtimeModal.classList.contains("hidden")) {
      closeOvertimeInfoModal();
      return;
    }
    const businessTripModal = $("#businessTripInfoModal");
    if (businessTripModal && !businessTripModal.classList.contains("hidden")) {
      closeBusinessTripInfoModal();
      return;
    }
    const educationModal = $("#educationInfoModal");
    if (educationModal && !educationModal.classList.contains("hidden")) {
      closeEducationInfoModal();
      return;
    }
    const tripResultModal = $("#tripResultModal");
    if (tripResultModal && !tripResultModal.classList.contains("hidden")) {
      closeTripResultModal();
      return;
    }
    const educationResultModal = $("#educationResultModal");
    if (educationResultModal && !educationResultModal.classList.contains("hidden")) {
      closeEducationResultModal();
      return;
    }
    const panel = $("#detailPanel");
    if (panel && !panel.classList.contains("hidden")) {
      closeDetailPanel();
    }
  });
  document.addEventListener("keydown", trapDetailPanelTabKey);
  $("#draftAttachments")?.addEventListener("change", renderDraftAttachmentSelection);

  $("#pendingBody")?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-id]");
    if (!tr) return;
    openDocument(Number(tr.dataset.id)).catch((err) => alert(err.message));
  });

  $("#myDocsBody").addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-id]");
    if (!tr) return;
    if (tr.dataset.canOpen === "false") {
      alert("열람 권한이 제한된 문서입니다. 문서 공개 범위를 확인해 주세요.");
      return;
    }
    openDocument(Number(tr.dataset.id)).catch((err) => alert(err.message));
  });

  $("#leaveMgmt")?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr[data-id]");
    if (!tr) return;
    openDocument(Number(tr.dataset.id)).catch((err) => alert(err.message));
  });
}

ensureTripResultModalMountedToBody();
ensureEducationResultModalMountedToBody();
ensureUserPickerModalMountedToBody();
bindEvents();
setEditorFormMode();
syncDraftTemplateTypeUI();

// 랜딩 페이지 버튼 이벤트
document.getElementById("landingLoginBtn")?.addEventListener("click", () => { hideLandingShowApp(); });
document.getElementById("landingStartBtn")?.addEventListener("click", () => { hideLandingShowApp(); });

bootstrap();
