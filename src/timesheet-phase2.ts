/**
 * FILE: timesheet-phase2.ts
 * Time Keeper — Phase 2 SPA
 * Multi-category time sheet application (TypeScript, no framework, DOM API only).
 *
 * Categories: Customer Deliverable, Business Development, Internal,
 *             Finance, Human Resources, Product / System
 *
 * Customer Deliverable & Business Development load from API (dynamic).
 * Internal, Finance, Human Resources, Product / System use static config.
 *
 * Summary Dashboard shows daily totals, per-category totals, and overall total.
 * Scroll synchronisation keeps header, body, and summary aligned horizontally.
 */

const APP_VERSION = "v2.0.0";

// ================================================================
//#region MSAL / AUTH GLOBALS
// ================================================================
declare const msal: any;

const _entraTenantId: string = "9d97e7fb-d431-4e72-a93c-c5452889c391";
const _entraClientId: string = "e43b1f85-24b8-41bd-b1a7-8c0248a1f153";
const _entraPaSharedId: string = "483a35f3-b479-4461-81f2-ee40ceddcda2";

let _UserEmail: string = "";
let _currentUser: User | null = null;
let _msalApp: any | null = null;
let _isDev: boolean = false;

// ================================================================
//#endregion
//#region MUTEX — prevents concurrent token requests
// ================================================================
class Mutex {
  private _locked = false;
  private _waiting: (() => void)[] = [];

  async lock(): Promise<void> {
    if (this._locked) {
      await new Promise<void>(resolve => this._waiting.push(resolve));
    }
    this._locked = true;
  }

  unlock(): void {
    if (this._waiting.length > 0) {
      const next = this._waiting.shift();
      if (next) next();
    } else {
      this._locked = false;
    }
  }
}

const _lockTokenMutex = new Mutex();

// ================================================================
//#endregion
//#region TYPES & INTERFACES
// ================================================================
type Guid = string | null;

type CategoryName = string;

/** Derive ordered category list from fetched TimeSheetCategories */
function getAllCategories(): CategoryName[] {
  return TimeSheetCategories.map(c => c.keyName);
}

/** Categories whose rows come from the API (accounts/prospects) */
const DYNAMIC_CATEGORIES: CategoryName[] = [
  "Customer Deliverable",
  "Business Development",
];

/** Stage map — numeric codes used by the backend */
const TimeSheetStageMap = {
  Pending: 596610000,
  Submitted: 596610001,
  Finalised: 596610002,
} as const;

type StageText = "Pending" | "Submitted" | "Finalised";

interface Account {
  accountId: Guid;
  accountName: string;
  catNumber?: string;
}

interface User {
  userId: Guid;
  fullName: string;
  emailAddress: string;
}

const TimeSheetCategories: KeyNameValuePair[] = [];
const SubCategories: SubCategory[] = [];

interface SubCategory {
    timeSheetItem: string;
    startDate: string | null;
    endDate: string | null;
    categoryId: string;
    timeSheetCategoryId: Guid;
    timeSheetCategory: number;
    status: number;
}

interface KeyNameValuePair {
  keyName: string;
  keyValue: string;
}

interface ChoiceOptionsRequest {
    entityLogicalName: string;
    choiceFieldLogicalName: string;
}

interface LookupValuesRequest {
    logicalName: string;
    xmlQuery: string;
    filters: string | null;
}

/** Summary model computed on the client */
interface SummaryModel {
  dailyTotals: number[];
  categoryTotals: Record<CategoryName, number>;
  categoryDailyTotals: Record<CategoryName, number[]>;
  subcategoryTotals: Record<CategoryName, Record<string, number>>;
  overall: number;
  expandedCategories: Set<CategoryName>;
  dashboardCollapsed: boolean;
}

interface TimesheetRecord {
  timeSheetId: Guid;
  timeSheetEntryId: string | null;
  yearMonth: string;
  category: CategoryName;
  subCategory: string;
  timeSheetCategory: number;
  timeSheetSubCategoryId: Guid;
  accountId: Guid;
  userId: string;
  timeSheetStage: number;
  totalHours: number;
  timeSheetStartDate: string;
  timeSheetEndDate: string;
  [k: `day${string}`]: number | undefined;
}

interface DataProvider {
  listAccounts(year: number, month: number): Promise<Account[]>;
  listTimesheets(year: number, month: number): Promise<TimesheetRecord[]>;
  saveTimesheet(record: TimesheetRecord): Promise<TimesheetRecord>;
  updateTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]>;
  createTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]>;
  deleteTimeKeepEntry(timeSheetId: Guid): Promise<void>;
  deleteTimeKeepEntries(timeSheetIds: Guid[]): Promise<void>;
  getUser(): Promise<User>;
  getChoiceOptions(request: ChoiceOptionsRequest): Promise<KeyNameValuePair[]>;
  getLookupValues(request: LookupValuesRequest): Promise<Array<Array<{Key: string; Value: any}>>>;
}

// ================================================================
//#endregion
//#region SUB-CATEGORY HELPERS
// ================================================================

/** Extract a field value from a raw lookup key-value pair array */
function lookupField(item: Array<{Key: string; Value: any}>, key: string): any {
  const pair = item.find(p => p.Key === key);
  return pair ? pair.Value : null;
}

/** Parse raw API response rows into SubCategory objects */
function parseLookupToSubCategories(raw: Array<Array<{Key: string; Value: any}>>): SubCategory[] {
  return raw.map(item => ({
    status: lookupField(item, "statecode")?.Value ?? 0,
    categoryId: lookupField(item, "cr91f_categoryid") ?? "",
    timeSheetCategory: lookupField(item, "cr91f_timesheetcategory")?.Value ?? 0,
    startDate: lookupField(item, "cr91f_startdate") ?? null,
    endDate: lookupField(item, "cr91f_enddate") ?? null,
    timeSheetItem: lookupField(item, "cr91f_timesheetitem") ?? "",
    timeSheetCategoryId: lookupField(item, "cr91f_timesheetcategoryid") ?? null,
  }));
}

/** Look up the category name for a SubCategory's timeSheetCategory number */
function getCategoryNameForSubCategory(sub: SubCategory): CategoryName {
  const entry = TimeSheetCategories.find(c => parseInt(c.keyValue, 10) === sub.timeSheetCategory);
  return entry ? entry.keyName : "";
}

/** Get sub-category items for a given category name from the dynamic SubCategories list */
function getSubcategoriesForCategory(cat: CategoryName): SubCategory[] {
  return SubCategories
    .filter(sub => getCategoryNameForSubCategory(sub) === cat)
    .sort((a, b) => a.timeSheetItem.localeCompare(b.timeSheetItem));
}

// ================================================================
//#endregion
//#region MSAL / AUTH HELPERS (reused from Phase 1)
// ================================================================

/** Load MSAL browser library from CDN */
async function loadMsalFromCdn(): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://alcdn.msauth.net/browser/2.35.0/js/msal-browser.min.js";
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Failed to load MSAL CDN"));
    document.head.appendChild(script);
  });
}

/** Create or reuse singleton MSAL instance */
async function getMsalApp(): Promise<any> {
  if (_msalApp) return _msalApp;
  const msalConfig = {
    auth: {
      clientId: _entraClientId,
      authority: `https://login.microsoftonline.com/${_entraTenantId}`,
      redirectUri: `${window.location.origin}/authTokenRedirect`,
    },
    cache: { cacheLocation: "sessionStorage" },
  };
  _msalApp = new msal.PublicClientApplication(msalConfig);
  return _msalApp;
}

/** Silent-first token acquisition for Azure Function API */
async function ensureTokenForFunction(): Promise<string> {
  await _lockTokenMutex.lock();
  try {
    const existingToken = sessionStorage.getItem("msalToken");
    if (existingToken) {
      const tokenParts = existingToken.split(".");
      const payload = JSON.parse(atob(tokenParts[1]));
      _UserEmail = payload.preferred_username ?? (() => { throw new Error("Unable to determine user email from token payload."); })();
      const exp = payload.exp as number;
      const currentTime = Math.floor(Date.now() / 1000);
      if (exp - currentTime > 300) return existingToken;
    }

    const scopes = [`api://${_entraPaSharedId}/user_impersonation`];
    const app = await getMsalApp();
    let account = app.getAllAccounts()[0];
    let msalToken: string = "";

    try {
      const res = await app.acquireTokenSilent({ account, scopes });
      msalToken = res.accessToken;
    } catch {
      try {
        const login = await app.loginPopup({ scopes });
        const res = await app.acquireTokenSilent({ account: login.account, scopes });
        msalToken = res.accessToken;
      } catch (err: any) {
        throw new Error(err?.message || "Authentication failed. Please ensure popups are allowed and try again.");
      }
    }

    sessionStorage.setItem("msalToken", msalToken);
    const payload = JSON.parse(atob(msalToken.split(".")[1]));
    _UserEmail = payload.preferred_username ?? (() => { throw new Error("Unable to determine user email from token payload."); })();
    return msalToken;
  } finally {
    _lockTokenMutex.unlock();
  }
}

// ================================================================
//#endregion
//#region OIDC / GRAPH PROFILE HELPERS (reused from Phase 1)
// ================================================================

type OidcProfile = {
  name?: string;
  preferred_username?: string;
  email?: string;
  sub?: string;
  tenantId?: string;
};

async function getOidcProfile(): Promise<OidcProfile | null> {
  const app = await getMsalApp();
  const account = app.getAllAccounts()[0];
  if (!account) return null;
  const claims = account.idTokenClaims as Record<string, unknown> | undefined;
  if (!claims) return null;
  return {
    name: claims["name"] as string | undefined,
    preferred_username: claims["preferred_username"] as string | undefined,
    email: claims["email"] as string | undefined,
    sub: claims["sub"] as string | undefined,
    tenantId: (claims["tid"] as string | undefined) ?? (claims["tenant"] as string | undefined),
  };
}

async function ensureOidcLogin(): Promise<void> {
  const app = await getMsalApp();
  if (app.getAllAccounts().length > 0) return;
  await app.loginPopup({ scopes: ["openid", "profile", "email"] });
}

async function acquireGraphAccessToken(): Promise<string> {
  const app = await getMsalApp();
  const account = app.getAllAccounts()[0];
  const graphScopes = ["User.Read"];
  try {
    const res = await app.acquireTokenSilent({ account, scopes: graphScopes });
    return res.accessToken;
  } catch {
    const login = await app.loginPopup({ scopes: graphScopes });
    const res = await app.acquireTokenSilent({ account: login.account, scopes: graphScopes });
    return res.accessToken;
  }
}

type GraphMe = {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
};

async function getGraphMe(): Promise<GraphMe> {
  const token = await acquireGraphAccessToken();
  const resp = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) throw new Error(`Graph /me failed: ${resp.status} ${resp.statusText}`);
  return (await resp.json()) as GraphMe;
}

async function getGraphPhotoUrl(): Promise<string | null> {
  const token = await acquireGraphAccessToken();
  const resp = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) return null;
  const blob = await resp.blob();
  return URL.createObjectURL(blob);
}

async function getCurrentUserUiProfile(): Promise<{
  name?: string;
  email?: string;
  photoUrl?: string | null;
}> {
  const oidc = await getOidcProfile();
  const me = await getGraphMe();
  const photoUrl = await getGraphPhotoUrl();
  const email = me.mail ?? me.userPrincipalName ?? oidc?.email ?? oidc?.preferred_username;
  return { name: oidc?.name ?? me.displayName, email, photoUrl };
}

// ================================================================
//#endregion
//#region API PROVIDER (reused from Phase 1, same endpoints)
// ================================================================

class ApiProvider implements DataProvider {
  private baseUrl(): string {
    return _isDev
      ? "http://localhost:7071/api"
      : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api";
  }

  private suffix(): string {
    return _isDev ? "" : "withaadauth";
  }

  private endpoint(name: string): string {
    return `${this.baseUrl()}/${name}${this.suffix()}`;
  }

  private async authHeaders(): Promise<Record<string, string>> {
    return {
      "Content-Type": "application/json",
      Authorization: `Bearer ${await ensureTokenForFunction()}`,
    };
  }

  async listAccounts(_year: number, _month: number): Promise<Account[]> {
    const response = await fetch(this.endpoint("getaccounts"), {
      method: "GET",
      headers: await this.authHeaders(),
    });
    if (!response.ok) throw new Error("Failed to fetch accounts");
    return await response.json();
  }

  async listTimesheets(year: number, month: number): Promise<TimesheetRecord[]> {
    const response = await fetch(this.endpoint("gettimekeepentries"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify({
        userId: _currentUser?.userId,
        yearMonth: `${year}-${String(month).padStart(2, "0")}`,
      }),
    });
    if (!response.ok) throw new Error("Failed to fetch timesheets");
    return await response.json();
  }

  async saveTimesheet(record: TimesheetRecord): Promise<TimesheetRecord> {
    const action = record.timeSheetId ? "updatetimekeepentry" : "createtimekeepentry";
    const response = await fetch(this.endpoint(action), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify(record),
    });
    if (!response.ok) throw new Error("Failed to save timesheet");
    return await response.json();
  }

  async updateTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]> {
    const response = await fetch(this.endpoint("updatetimekeepentries"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify(records),
    });
    if (!response.ok) throw new Error("Failed to update timesheets");
    return await response.json();
  }

  async createTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]> {
    const response = await fetch(this.endpoint("createtimekeepentries"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify(records),
    });
    if (!response.ok) throw new Error("Failed to create timesheets");
    return await response.json();
  }

  async deleteTimeKeepEntry(timeSheetId: Guid): Promise<void> {
    const response = await fetch(this.endpoint("deletetimekeepentry"), {
      method: "DELETE",
      headers: await this.authHeaders(),
      body: JSON.stringify({ timeSheetId }),
    });
    if (!response.ok) throw new Error("Failed to delete timesheet entry");
  }

  async deleteTimeKeepEntries(timeSheetIds: Guid[]): Promise<void> {
    const response = await fetch(this.endpoint("deletetimekeepentries"), {
      method: "DELETE",
      headers: await this.authHeaders(),
      body: JSON.stringify({ timeSheetIds }),
    });
    if (!response.ok) throw new Error("Failed to delete timesheet entries");
  }

  async getUser(): Promise<User> {
    const response = await fetch(this.endpoint("getuser"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify({ emailAddress: _UserEmail }),
    });
    if (!response.ok) throw new Error("Failed to fetch user");
    const user: User = await response.json();
    _currentUser = user;
    return user;
  }

  async getChoiceOptions(request: ChoiceOptionsRequest): Promise<KeyNameValuePair[]> {
    const response = await fetch(this.endpoint("GetChoiceOptions"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify(request),
    });
    if (!response.ok) throw new Error("Failed to fetch choice options");
    return await response.json();
  }

  async getLookupValues(request: LookupValuesRequest): Promise<Array<Array<{Key: string; Value: any}>>> {
    const response = await fetch(this.endpoint("GetLookupValuesDynamic"), {
      method: "POST",
      headers: await this.authHeaders(),
      body: JSON.stringify(request),
    });
    if (!response.ok) throw new Error("Failed to fetch lookup values");
    return await response.json();
  }
}

// ================================================================
//#endregion
//#region UTILITY HELPERS
// ================================================================

/** Pad day number to field key, e.g. dayKey(1) → "day01" */
const dayKey = (d: number): `day${string}` => `day${String(d).padStart(2, "0")}` as `day${string}`;

/** Sum all day fields on a row model */
function sumRowHours(row: TimesheetRecord, daysInMonth: number): number {
  let total = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const v = row[dayKey(d)];
    const n = typeof v === "number" ? v : Number(v) || 0;
    if (Number.isFinite(n)) total += n;
  }
  return Number(total.toFixed(2));
}


function getDaysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate();
}

function isoStartOfMonth(year: number, month: number): string {
  return new Date(Date.UTC(year, month - 1, 1, 0, 0, 0)).toISOString();
}

function isoEndOfMonth(year: number, month: number): string {
  const days = getDaysInMonth(year, month);
  return new Date(Date.UTC(year, month - 1, days, 0, 0, 0)).toISOString();
}

function dayLabel(year: number, month: number, day: number): string {
  const dt = new Date(Date.UTC(year, month - 1, day));
  const weekday = dt.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" });
  const dayNum = dt.getUTCDate();
  return `${weekday} ${String(dayNum).padStart(2, "0")}`;
}

/** Check if a given day column (1-based) is a weekend (Sat or Sun) */
function isWeekend(year: number, month: number, day: number): boolean {
  const d = new Date(Date.UTC(year, month - 1, day));
  const dow = d.getUTCDay();
  return dow === 0 || dow === 6;
}

/** Check if a given day is today (UTC comparison) */
function isToday(year: number, month: number, day: number): boolean {
  const now = new Date();
  return year === now.getUTCFullYear() && month === now.getUTCMonth() + 1 && day === now.getUTCDate();
}

/** Check if a given day is in the future (UTC comparison) */
function isFutureDay(year: number, month: number, day: number): boolean {
  const today = new Date();
  const cellDate = new Date(Date.UTC(year, month - 1, day));
  const todayUtc = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
  return cellDate > todayUtc;
}

function getCategoryValue(cat: CategoryName): number {
  const entry = TimeSheetCategories.find(c => c.keyName === cat);
  return entry ? parseInt(entry.keyValue, 10) : 0;
}

/** Does a category use dynamic (API) rows? */
function isDynamicCategory(cat: CategoryName): boolean {
  return DYNAMIC_CATEGORIES.includes(cat);
}

/** Build a yearMonth string */
function yearMonthStr(year: number, month: number): string {
  return `${year}-${String(month).padStart(2, "0")}`;
}


/** Check if a row has any hours > 0 */
function rowHasHours(row: TimesheetRecord, daysInMonth: number): boolean {
  for (let d = 1; d <= daysInMonth; d++) {
    const v = row[dayKey(d)];
    const n = typeof v === "number" ? v : Number(v) || 0;
    if (n > 0) return true;
  }
  return false;
}

// ================================================================
//#endregion
//#region LOADER & ERROR HELPERS (reused from Phase 1)
// ================================================================

function getTsApp(): HTMLElement | null {
  return document.querySelector<HTMLElement>(".ts-app");
}

function ensureLoader(): HTMLElement | null {
  const app = getTsApp();
  if (!app) return null;
  let loader = app.querySelector<HTMLElement>("#ts-loader");
  if (!loader) {
    loader = document.createElement("div");
    loader.id = "ts-loader";
    loader.setAttribute("role", "status");
    loader.setAttribute("aria-live", "polite");
    loader.setAttribute("aria-label", "Loading");
    loader.innerHTML = `
      <div class="ts-loader-content">
        <div class="ts-spinner" aria-hidden="true"></div>
        <div class="ts-loader-text">Loading\u2026</div>
      </div>
    `;
    app.appendChild(loader);
  }
  return loader;
}

function showLoader(statusText?: string): void {
  const app = getTsApp();
  const loader = ensureLoader();
  if (!app || !loader) return;
  loader.style.display = "flex";
  app.classList.add("is-loading");
  if (typeof statusText === "string") {
    const textDiv = loader.querySelector<HTMLElement>(".ts-loader-text");
    if (textDiv) textDiv.textContent = statusText;
  }
  loader.setAttribute("aria-hidden", "false");
}

function hideLoader(): void {
  const app = getTsApp();
  const loader = app?.querySelector<HTMLElement>("#ts-loader");
  if (!app || !loader) return;
  loader.style.display = "none";
  app.classList.remove("is-loading");
  loader.setAttribute("aria-hidden", "true");
}

function showErrorScreen(message: string): void {
  let errorOverlay = document.getElementById("ts-error-overlay") as HTMLDivElement | null;
  if (!errorOverlay) {
    errorOverlay = document.createElement("div");
    errorOverlay.id = "ts-error-overlay";
    errorOverlay.innerHTML = `
      <div class="ts-error-content">
        <h2>Something went wrong</h2>
        <div id="ts-error-message"></div>
        <button id="ts-error-reload" type="button">Reload</button>
      </div>
    `;
    const app = document.querySelector(".ts-app");
    if (app) app.appendChild(errorOverlay);
    else document.body.appendChild(errorOverlay);
    document.getElementById("ts-error-reload")?.addEventListener("click", () => window.location.reload());
  }
  const msgDiv = document.getElementById("ts-error-message");
  if (msgDiv) {
    if (typeof message === "string" && /popup|blocked/i.test(message)) {
      msgDiv.textContent =
        "Please allow popup windows for this site in your browser settings. " +
        "You can check the top right of your browser address bar for any blocked popup icons." +
        "\n\nAfter allowing popups, please refresh the page.";
    } else {
      msgDiv.textContent = message;
    }
  }
  errorOverlay.style.display = "flex";
}

function hideErrorScreen(): void {
  const overlay = document.getElementById("ts-error-overlay");
  if (overlay) overlay.style.display = "none";
}

/** Delete confirmation modal */
function showDeleteConfirmModal(callback: (confirmed: boolean) => void): void {
  let modal = document.getElementById("ts-delete-confirm-modal") as HTMLDivElement | null;
  if (!modal) {
    modal = document.createElement("div");
    modal.id = "ts-delete-confirm-modal";
    modal.innerHTML = `
      <div class="ts-error-content">
        <h2>Delete Row?</h2>
        <div id="ts-delete-confirm-message">
          Are you sure you want to delete this row?<br><br>
          There are hours logged for this entry.<br><br>
          <strong>This action is <span style='color:#ffd6d6'>irreversible</span>.</strong>
        </div>
        <div style="margin-top: 24px; display: flex; gap: 16px; justify-content: center;">
          <button id="ts-delete-confirm-yes" class="ts-btn ts-btn-danger">Delete</button>
          <button id="ts-delete-confirm-no" class="ts-btn">Cancel</button>
        </div>
      </div>
    `;
    const app = document.querySelector(".ts-app");
    if (app) app.appendChild(modal);
    else document.body.appendChild(modal);
  }
  modal.classList.add("ts-delete-confirm-modal-visible");
  const yesBtn = document.getElementById("ts-delete-confirm-yes");
  const noBtn = document.getElementById("ts-delete-confirm-no");
  const cleanup = () => { modal!.classList.remove("ts-delete-confirm-modal-visible"); };
  yesBtn?.addEventListener("click", () => { cleanup(); callback(true); }, { once: true });
  noBtn?.addEventListener("click", () => { cleanup(); callback(false); }, { once: true });
}

// ================================================================
//#endregion
//#region MAIN APPLICATION CLASS — TimesheetAppV2
// ================================================================

class TimesheetAppV2 {
  private container: HTMLDivElement;
  private provider: DataProvider;

  /** State */
  private year: number;
  private month: number;
  private daysInMonth: number;
  private selectedTab: CategoryName = "Customer Deliverable";
  private rows: TimesheetRecord[] = [];
  private accountsCache: Account[] = [];
  private summary: SummaryModel;
  private stageText: StageText = "Pending";

  /** UI element references */
  private ui = {
    toolbar: null as HTMLDivElement | null,
    tabBar: null as HTMLDivElement | null,
    summaryPanel: null as HTMLDivElement | null,
    summaryScroll: null as HTMLDivElement | null,
    summaryGrid: null as HTMLDivElement | null,
    grid: null as HTMLDivElement | null,
    headerScroll: null as HTMLDivElement | null,
    bodyScroll: null as HTMLDivElement | null,
    headerRow: null as HTMLDivElement | null,
    gridBody: null as HTMLDivElement | null,
    statusText: null as HTMLDivElement | null,
    addRowBtn: null as HTMLButtonElement | null,
    prevMonthBtn: null as HTMLButtonElement | null,
    userPill: null as HTMLDivElement | null,
    stageSelect: null as HTMLSelectElement | null,
    monthInput: null as HTMLInputElement | null,
    footerSection: null as HTMLDivElement | null,
  };

  /** Debounce timers per row */
  private saveTimers = new Map<TimesheetRecord, number>();
  private saveQueues = new Map<TimesheetRecord, (() => void)[]>();

  constructor(container: HTMLDivElement, provider: DataProvider) {
    this.container = container;
    this.provider = provider;
    const today = new Date();
    this.year = today.getFullYear();
    this.month = today.getMonth() + 1;
    this.daysInMonth = getDaysInMonth(this.year, this.month);
    this.summary = this.emptySummary();
  }

  /** Create an empty summary model */
  private emptySummary(): SummaryModel {
    const catTotals: Record<CategoryName, number> = {} as any;
    const catDaily: Record<CategoryName, number[]> = {} as any;
    const subcatTotals: Record<CategoryName, Record<string, number>> = {} as any;
    for (const c of getAllCategories()) {
      catTotals[c] = 0;
      catDaily[c] = new Array(this.daysInMonth).fill(0);
      subcatTotals[c] = {};
    }
    return {
      dailyTotals: new Array(this.daysInMonth).fill(0),
      categoryTotals: catTotals,
      categoryDailyTotals: catDaily,
      subcategoryTotals: subcatTotals,
      overall: 0,
      expandedCategories: new Set<CategoryName>(),
      dashboardCollapsed: true,
    };
  }

// ----------------------------------------------------------------
//#endregion
//#region MOUNT — Entry point
// ---------------------------------------------------------------- */

  /** Mount the application into the container */
  async mount(): Promise<void> {
    this.container.classList.add("ts-app");
    this.container.innerHTML = "";
    ensureLoader();
    showLoader("Logging in\u2026");

    try {
      await ensureTokenForFunction();
    } catch (err: any) {
      hideLoader();
      showErrorScreen(err?.message || "Failed to authenticate.");
      return;
    }

    showLoader("Loading user and data\u2026");
    this.buildShell();

    try {
      const uiProfile = await getCurrentUserUiProfile();
      await this.provider.getUser();
      this.updateUserDisplay(uiProfile);
      await this.loadCategories();
      await this.loadSubCategories();
      await this.loadData();
      this.renderAll();
      hideLoader();
    } catch (err: any) {
      hideLoader();
      showErrorScreen(err?.message || "Failed to load user or data.");
    }
  }

// ----------------------------------------------------------------
//#endregion  
//#region BUILD SHELL — toolbar, tabs, summary, grid structure
// ----------------------------------------------------------------

  private buildShell(): void {
    /* --- Toolbar --- */
    const toolbar = document.createElement("div");
    toolbar.className = "ts-toolbar";
    this.ui.toolbar = toolbar;

    const left = document.createElement("div");
    left.className = "ts-toolbar-left";

    const logo = document.createElement("img");
    logo.src = "Catalyst-Logo-1200-Dark.png";
    logo.alt = "Catalyst Logo";
    logo.className = "ts-logo";

    const monthLabel = document.createElement("label");
    monthLabel.textContent = "Select month:";
    monthLabel.className = "ts-label";
    monthLabel.setAttribute("for", "ts-month-picker");

    const monthInput = document.createElement("input");
    monthInput.type = "month";
    monthInput.id = "ts-month-picker";
    monthInput.value = yearMonthStr(this.year, this.month);
    const now = new Date();
    monthInput.min = "2025-01";
    monthInput.max = yearMonthStr(now.getFullYear(), now.getMonth() + 1);
    monthInput.className = "ts-month-input";
    monthInput.setAttribute("aria-label", "Select month");
    monthInput.addEventListener("change", () => this.handleMonthChange(monthInput.value));
    monthInput.addEventListener("click", () => { if (monthInput.showPicker) monthInput.showPicker(); });
    this.ui.monthInput = monthInput;

    const stageLabel = document.createElement("label");
    stageLabel.textContent = "Stage:";
    stageLabel.className = "ts-label";
    stageLabel.setAttribute("for", "ts-stage-select");

    const stageSelect = document.createElement("select");
    stageSelect.id = "ts-stage-select";
    stageSelect.className = "ts-status-select";
    stageSelect.setAttribute("aria-label", "Timesheet stage");
    for (const s of ["Pending", "Submitted"] as StageText[]) {
      const opt = document.createElement("option");
      opt.value = s;
      opt.textContent = s;
      stageSelect.appendChild(opt);
    }
    stageSelect.value = this.stageText;
    stageSelect.addEventListener("change", () => this.handleStageChange(stageSelect.value as StageText));
    this.ui.stageSelect = stageSelect;

    const userPill = document.createElement("div");
    userPill.className = "ts-user-pill";
    this.ui.userPill = userPill;

    left.append(logo, monthLabel, monthInput, stageLabel, stageSelect);
    toolbar.append(left, userPill);
    this.container.appendChild(toolbar);

    /* --- Summary Panel --- */
    const summaryPanel = document.createElement("div");
    summaryPanel.className = "ts-summary";
    this.ui.summaryPanel = summaryPanel;
    this.container.appendChild(summaryPanel);

    /* --- Tab Bar (above Timesheet Entry) --- */
    const tabBar = document.createElement("div");
    tabBar.className = "ts-tabs";
    tabBar.setAttribute("role", "tablist");
    this.ui.tabBar = tabBar;
    this.renderTabs();
    this.container.appendChild(tabBar);

    /* --- Grid (Dashboard-style layout) --- */
    const grid = document.createElement("div");
    grid.className = "ts-grid";
    this.ui.grid = grid;

    const gridHeader = document.createElement("div");
    gridHeader.className = "ts-grid-header";
    const gridHeaderLeft = document.createElement("div");
    gridHeaderLeft.className = "ts-grid-header-left";
    const gridTitle = document.createElement("span");
    gridTitle.className = "ts-grid-title";
    gridTitle.textContent = "Timesheet Entry";
    gridHeaderLeft.appendChild(gridTitle);
    gridHeader.appendChild(gridHeaderLeft);
    grid.appendChild(gridHeader);

    const gridInner = document.createElement("div");
    gridInner.className = "ts-grid-inner";

    const gridContent = document.createElement("div");
    gridContent.className = "ts-grid-content";

    const headerRow = document.createElement("div");
    headerRow.className = "ts-grid-header-row";
    this.ui.headerRow = headerRow;

    const gridBody = document.createElement("div");
    gridBody.className = "ts-grid-body";
    this.ui.gridBody = gridBody;

    gridContent.append(headerRow, gridBody);
    gridInner.appendChild(gridContent);
    this.ui.bodyScroll = gridInner;
    this.ui.headerScroll = gridInner;

    const checkHScroll = () => {
      const hasH = gridInner.scrollWidth > gridInner.clientWidth;
      gridInner.classList.toggle("has-hscroll", hasH);
    };
    new ResizeObserver(checkHScroll).observe(gridInner);
    new MutationObserver(checkHScroll).observe(gridBody, { childList: true, subtree: true });

    grid.appendChild(gridInner);
    this.container.appendChild(grid);

    /* --- Footer (Add Customer / Copy / Status) --- */
    const footer = document.createElement("div");
    footer.className = "ts-app-footer-section";
    this.ui.footerSection = footer;

    const footerLeft = document.createElement("div");
    footerLeft.className = "ts-app-footer-left";

    const addRowBtn = document.createElement("button");
    addRowBtn.className = "ts-btn";
    addRowBtn.textContent = "Add Customer";
    addRowBtn.setAttribute("aria-label", "Add a new customer row");
    addRowBtn.addEventListener("click", () => this.addCustomerRow());
    this.ui.addRowBtn = addRowBtn;

    const copyLabel = document.createElement("span");
    copyLabel.className = "ts-copy-prev-label";
    copyLabel.textContent = "Copy from previous month:";

    const prevMonthBtn = document.createElement("button");
    prevMonthBtn.className = "ts-btn ts-btn-prev-month";
    prevMonthBtn.title = "Copy from previous month";
    prevMonthBtn.setAttribute("aria-label", "Copy rows from previous month");
    prevMonthBtn.innerHTML = '<span class="ts-icon-copy"></span>';
    prevMonthBtn.addEventListener("click", () => this.copyFromPreviousMonth());
    this.ui.prevMonthBtn = prevMonthBtn;

    footerLeft.append(addRowBtn, copyLabel, prevMonthBtn);

    const statusText = document.createElement("div");
    statusText.className = "ts-app-footer-right";
    statusText.textContent = "Current activity: Idle";
    this.ui.statusText = statusText;

    footer.append(footerLeft, statusText);
    this.container.appendChild(footer);

    /* --- Scroll synchronisation --- */
    this.setupScrollSync();
  }

  /** Wire up horizontal scroll synchronisation between header, body, and summary */
  private setupScrollSync(): void {
    const { headerScroll, bodyScroll } = this.ui;
    let syncing = false;

    const sync = (source: HTMLElement, ...targets: (HTMLElement | null)[]) => {
      if (syncing) return;
      syncing = true;
      for (const t of targets) {
        if (t && t !== source) t.scrollLeft = source.scrollLeft;
      }
      syncing = false;
    };

    bodyScroll?.addEventListener("scroll", () => {
      sync(bodyScroll!, headerScroll, this.ui.summaryScroll);
    });

    headerScroll?.addEventListener("scroll", () => {
      sync(headerScroll!, bodyScroll, this.ui.summaryScroll);
    });
  }

  /** Sync summary scroll when it's scrolled by the user */
  private attachSummaryScrollSync(): void {
    const { summaryScroll, headerScroll, bodyScroll } = this.ui;
    if (!summaryScroll) return;
    let syncing = false;
    summaryScroll.addEventListener("scroll", () => {
      if (syncing) return;
      syncing = true;
      if (headerScroll) headerScroll.scrollLeft = summaryScroll.scrollLeft;
      if (bodyScroll) bodyScroll.scrollLeft = summaryScroll.scrollLeft;
      syncing = false;
    });
  }

// ----------------------------------------------------------------
//#endregion
//#region RENDER METHODS
// ----------------------------------------------------------------

  /** Re-render everything */
  private renderAll(): void {
    this.renderTabs();
    this.renderHeader();
    this.renderBody();
    this.computeSummary();
    this.renderSummary();
    this.applyStageToUi();
  }

  /** Get visible categories (dynamic categories always visible; others only if they have subcategories) */
  private getVisibleCategories(): CategoryName[] {
    return getAllCategories().filter(
      cat => isDynamicCategory(cat) || getSubcategoriesForCategory(cat).length > 0
    );
  }

  /** Render category tabs */
  private renderTabs(): void {
    if (!this.ui.tabBar) return;
    this.ui.tabBar.innerHTML = "";
    const visible = this.getVisibleCategories();
    if (visible.length > 0 && !visible.includes(this.selectedTab)) {
      this.selectedTab = visible[0];
    }
    for (const cat of visible) {
      const btn = document.createElement("button");
      btn.className = `ts-tab${cat === this.selectedTab ? " ts-tab-active" : ""}`;
      btn.textContent = cat;
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", cat === this.selectedTab ? "true" : "false");
      btn.setAttribute("aria-label", `${cat} tab`);
      btn.addEventListener("click", () => this.setTab(cat));
      this.ui.tabBar!.appendChild(btn);
    }
  }

  /** Render the grid header row (sub-category label + day columns + total) */
  private renderHeader(): void {
    if (!this.ui.headerRow) return;
    this.ui.headerRow.innerHTML = "";

    const labelCell = document.createElement("div");
    labelCell.className = "ts-header-cell ts-header-customer";
    labelCell.textContent = isDynamicCategory(this.selectedTab) ? "Customer" : "Item";
    this.ui.headerRow.appendChild(labelCell);

    for (let d = 1; d <= this.daysInMonth; d++) {
      const cell = document.createElement("div");
      cell.className = "ts-header-cell";
      if (isToday(this.year, this.month, d)) cell.classList.add("ts-today");
      if (isWeekend(this.year, this.month, d)) cell.classList.add("ts-weekend");
      const dt = new Date(Date.UTC(this.year, this.month - 1, d));
      const dayName = dt.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" });
      const dayNum = String(dt.getUTCDate()).padStart(2, "0");
      cell.innerHTML = `<span class="ts-header-day-name">${dayName}</span><br>${dayNum}`;
      this.ui.headerRow.appendChild(cell);
    }

    const totalCell = document.createElement("div");
    totalCell.className = "ts-header-cell ts-header-total";
    totalCell.textContent = "Total";
    this.ui.headerRow.appendChild(totalCell);
  }

  /** Render body rows for the currently selected tab */
  private renderBody(): void {
    if (!this.ui.gridBody) return;
    this.ui.gridBody.innerHTML = "";

    const tabRows = this.getRowsForTab(this.selectedTab);

    if (tabRows.length === 0) {
      const emptyRow = document.createElement("div");
      emptyRow.className = "ts-row";
      emptyRow.style.justifyContent = "center";
      emptyRow.style.padding = "24px";
      emptyRow.style.color = "#888";
      emptyRow.textContent = isDynamicCategory(this.selectedTab)
        ? "No entries yet. Click 'Add Customer' to begin."
        : "No entries yet. Hours will auto-save when you enter them.";
      this.ui.gridBody.appendChild(emptyRow);
      return;
    }

    const fragment = document.createDocumentFragment();
    for (const row of tabRows) {
      fragment.appendChild(this.buildRowElement(row));
    }
    this.ui.gridBody.appendChild(fragment);
  }

  /** Build a single row DOM element */
  private buildRowElement(row: TimesheetRecord): HTMLDivElement {
    const rowEl = document.createElement("div");
    rowEl.className = "ts-row";

    const custCell = document.createElement("div");
    custCell.className = "ts-cell ts-cell-customer";

    const isEditing = (row as any)._editing === true;
    const isDynamic = isDynamicCategory(row.category);

    if (isDynamic && (!row.accountId || isEditing)) {
      custCell.appendChild(this.buildAccountSelect(row));
    } else if (isDynamic && row.accountId) {
      const labelWrap = document.createElement("div");
      labelWrap.className = "ts-account-label-wrap";
      const nameSpan = document.createElement("span");
      nameSpan.className = "ts-subcategory-label";
      nameSpan.textContent = row.subCategory || "(No account)";
      nameSpan.title = row.subCategory || "";
      labelWrap.appendChild(nameSpan);

      const btnWrap = document.createElement("div");
      btnWrap.className = "ts-account-btn-wrap";

      const editBtn = document.createElement("button");
      editBtn.className = "ts-btn ts-btn-small";
      editBtn.title = "Edit account";
      editBtn.innerHTML = '<span class="ts-icon-edit"></span>';
      editBtn.setAttribute("aria-label", `Edit ${row.subCategory}`);
      if (this.stageText !== "Pending") editBtn.disabled = true;
      editBtn.addEventListener("click", () => this.editCustomerRow(row));
      btnWrap.appendChild(editBtn);

      const deleteBtn = document.createElement("button");
      deleteBtn.className = "ts-btn ts-btn-small ts-delete-account-btn";
      deleteBtn.title = "Delete row";
      deleteBtn.innerHTML = '<span class="ts-icon-delete"></span>';
      deleteBtn.setAttribute("aria-label", `Delete ${row.subCategory}`);
      if (this.stageText !== "Pending") deleteBtn.disabled = true;
      deleteBtn.addEventListener("click", () => this.deleteCustomerRow(row));
      btnWrap.appendChild(deleteBtn);

      labelWrap.appendChild(btnWrap);
      custCell.appendChild(labelWrap);
    } else {
      const label = document.createElement("span");
      label.className = "ts-subcategory-label";
      label.textContent = row.subCategory;
      label.title = row.subCategory;
      custCell.appendChild(label);
    }

    rowEl.appendChild(custCell);

    const isLocked =
      row.timeSheetStage === TimeSheetStageMap.Submitted ||
      row.timeSheetStage === TimeSheetStageMap.Finalised;
    const noAccount = isDynamic && !row.accountId;

    const totalCell = document.createElement("div");
    totalCell.className = "ts-cell ts-cell-total";
    totalCell.textContent = sumRowHours(row, this.daysInMonth).toFixed(2);

    for (let d = 1; d <= this.daysInMonth; d++) {
      const cell = document.createElement("div");
      cell.className = "ts-cell";
      if (isWeekend(this.year, this.month, d)) cell.classList.add("ts-weekend");

      const input = document.createElement("input");
      input.type = "number";
      input.step = "0.25";
      input.min = "0";
      input.max = "24";
      input.placeholder = "0";
      input.className = "ts-hour-input";
      input.setAttribute("aria-label", `Hours for day ${d}`);

      const k = dayKey(d);
      const currentVal = (row[k] as number) ?? 0;
      input.value = String(currentVal);

      const future = isFutureDay(this.year, this.month, d);
      input.disabled = noAccount || future || isLocked;

      const commitValue = () => {
        const normalized = input.value.replace(",", ".");
        const val = Number(normalized);
        const prev = (row[k] as number) ?? 0;
        const newVal = Number.isFinite(val) ? Math.max(0, Math.min(24, val)) : 0;
        if (prev === newVal) return;
        this.scheduleSave(row, () => {
          (row as any)[k] = newVal;
          input.value = String(newVal);
          row.totalHours = sumRowHours(row, this.daysInMonth);
          totalCell.textContent = row.totalHours.toFixed(2);
          this.computeSummary();
          this.renderSummary();
        });
      };

      input.addEventListener("blur", commitValue);
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { commitValue(); input.blur(); }
      });

      cell.appendChild(input);
      rowEl.appendChild(cell);
    }

    rowEl.appendChild(totalCell);
    return rowEl;
  }

  /** Build account select dropdown for dynamic category rows */
  private buildAccountSelect(row: TimesheetRecord): HTMLDivElement {
    const selectWrap = document.createElement("div");
    selectWrap.className = "ts-account-select-wrap";

    const select = document.createElement("select");
    select.className = "ts-account-select";
    select.setAttribute("aria-label", "Select account");

    const prompt = document.createElement("option");
    prompt.value = "";
    prompt.textContent = "Select account";
    select.appendChild(prompt);

    const usedIds = new Set(
      this.rows
        .filter(r => isDynamicCategory(r.category) && r.accountId && r !== row)
        .map(r => r.accountId!)
    );
    const prevAccountId = (row as any)._prevAccountId;
    const available = this.accountsCache.filter(
      a => a.accountId != null && (!usedIds.has(a.accountId) || a.accountId === prevAccountId)
    );

    // Sort available accounts by accountName ascending
    available.sort((a, b) => a.accountName.localeCompare(b.accountName));

    for (const a of available) {
      const opt = document.createElement("option");
      opt.value = a.accountId!;
      opt.textContent = a.accountName;
      select.appendChild(opt);
    }

    if (prevAccountId) select.value = prevAccountId;

    const okBtn = document.createElement("button");
    okBtn.className = "ts-btn ts-btn-small ts-btn-customer";
    okBtn.title = "Confirm";
    okBtn.setAttribute("aria-label", "Confirm account selection");
    okBtn.innerHTML = '<span class="ts-icon-check"></span>';
    okBtn.addEventListener("click", () => {
      const chosen = select.value;
      if (!chosen) return;
      if (prevAccountId === chosen) {
        row.accountId = prevAccountId;
        row.subCategory = (row as any)._prevSubCategory ?? "";
        delete (row as any)._prevAccountId;
        delete (row as any)._prevSubCategory;
        delete (row as any)._editing;
        this.renderBody();
        return;
      }
      const account = this.accountsCache.find(a => a.accountId === chosen);
      this.scheduleSave(row, () => {
        row.accountId = account?.accountId ?? null;
        row.subCategory = account?.accountName ?? "";
        const ym = yearMonthStr(this.year, this.month);
        row.yearMonth = ym;
        row.timeSheetStage = TimeSheetStageMap.Pending;
        row.timeSheetEntryId = `${ym.replace("-", "")}_${account?.catNumber ?? "000"}_${_currentUser!.fullName[0]}.${_currentUser!.fullName.trim().split(" ").slice(-1)[0]}`;
        row.timeSheetStartDate = isoStartOfMonth(this.year, this.month);
        row.timeSheetEndDate = isoEndOfMonth(this.year, this.month);
        row.userId = _currentUser?.userId ?? "";
        delete (row as any)._prevAccountId;
        delete (row as any)._prevSubCategory;
        delete (row as any)._editing;
        this.renderBody();
      });
    });

    const cancelBtn = document.createElement("button");
    cancelBtn.className = "ts-btn ts-btn-small ts-btn-customer";
    cancelBtn.title = "Cancel";
    cancelBtn.setAttribute("aria-label", "Cancel account selection");
    cancelBtn.innerHTML = '<span class="ts-icon-cancel"></span>';
    cancelBtn.addEventListener("click", () => {
      if ((row as any)._prevAccountId === undefined && !row.accountId) {
        const idx = this.rows.indexOf(row);
        if (idx !== -1) this.rows.splice(idx, 1);
      } else if ((row as any)._prevAccountId !== undefined) {
        row.accountId = (row as any)._prevAccountId;
        row.subCategory = (row as any)._prevSubCategory ?? "";
        delete (row as any)._prevAccountId;
        delete (row as any)._prevSubCategory;
      }
      delete (row as any)._editing;
      this.renderBody();
    });

    selectWrap.append(select, okBtn, cancelBtn);
    return selectWrap;
  }

  /** Render the summary dashboard */
  private renderSummary(): void {
    if (!this.ui.summaryPanel) return;
    this.ui.summaryPanel.innerHTML = "";

    const collapsed = this.summary.dashboardCollapsed;

    const header = document.createElement("div");
    header.className = "ts-summary-header ts-summary-header-toggle";
    header.setAttribute("role", "button");
    header.setAttribute("aria-expanded", String(!collapsed));
    header.style.cursor = "pointer";

    const chevron = document.createElement("span");
    chevron.className = "ts-summary-chevron" + (collapsed ? "" : " ts-summary-chevron-open");
    chevron.textContent = "\u25B6";

    const title = document.createElement("span");
    title.className = "ts-summary-title";
    title.textContent = "Summary Dashboard";

    const overallSpan = document.createElement("span");
    overallSpan.className = "ts-summary-overall";
    overallSpan.textContent = `Total: ${this.summary.overall.toFixed(2)} hrs`;

    const headerLeft = document.createElement("div");
    headerLeft.className = "ts-summary-header-left";
    headerLeft.append(chevron, title);

    header.addEventListener("click", () => {
      this.summary.dashboardCollapsed = !this.summary.dashboardCollapsed;
      this.renderSummary();
    });

    const headerRight = document.createElement("div");
    headerRight.className = "ts-summary-header-right";
    headerRight.append(overallSpan);

    header.append(headerLeft, headerRight);
    this.ui.summaryPanel.appendChild(header);

    const scrollWrap = document.createElement("div");
    scrollWrap.className = "ts-summary-scroll";
    this.ui.summaryScroll = scrollWrap;

    const grid = document.createElement("div");
    grid.className = "ts-summary-grid";
    this.ui.summaryGrid = grid;

    /* Header row for Summary Dashboard */
    const summaryHeaderRow = document.createElement("div");
    summaryHeaderRow.className = "ts-summary-row ts-summary-row-header";

    const headerLabel = document.createElement("div");
    headerLabel.className = "ts-summary-label ts-summary-label-header";
    headerLabel.textContent = "Category";
    summaryHeaderRow.appendChild(headerLabel);

    for (let d = 1; d <= this.daysInMonth; d++) {
      const cell = document.createElement("div");
      cell.className = "ts-summary-cell ts-summary-cell-header";
      if (isToday(this.year, this.month, d)) cell.classList.add("ts-summary-today");
      if (isWeekend(this.year, this.month, d)) cell.classList.add("ts-weekend-summary");
      const dayName = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][new Date(this.year, this.month - 1, d).getDay()];
      cell.innerHTML = `<span class="ts-summary-day-name">${dayName}</span><span>${String(d).padStart(2, "0")}</span>`;
      summaryHeaderRow.appendChild(cell);
    }
    const headerTotalCell = document.createElement("div");
    headerTotalCell.className = "ts-summary-cell ts-summary-cell-header ts-summary-cell-total";
    headerTotalCell.textContent = "Total";
    summaryHeaderRow.appendChild(headerTotalCell);
    grid.appendChild(summaryHeaderRow);

    if (collapsed) {
      const dailyRow = document.createElement("div");
      dailyRow.className = "ts-summary-row ts-summary-row-daily";
      const dailyLabel = document.createElement("div");
      dailyLabel.className = "ts-summary-label ts-summary-label-daily";
      dailyLabel.textContent = "Daily Total";
      dailyRow.appendChild(dailyLabel);
      for (let d = 0; d < this.daysInMonth; d++) {
        const cell = document.createElement("div");
        cell.className = "ts-summary-cell";
        if (isWeekend(this.year, this.month, d + 1)) cell.classList.add("ts-weekend-summary");
        const val = this.summary.dailyTotals[d];
        cell.textContent = val > 0 ? val.toFixed(2) : "";
        dailyRow.appendChild(cell);
      }
      const dailyTotalCell = document.createElement("div");
      dailyTotalCell.className = "ts-summary-cell ts-summary-cell-total";
      dailyTotalCell.textContent = this.summary.overall.toFixed(2);
      dailyRow.appendChild(dailyTotalCell);
      grid.appendChild(dailyRow);
      scrollWrap.appendChild(grid);
      this.ui.summaryPanel.appendChild(scrollWrap);
      this.attachSummaryScrollSync();
      if (this.ui.bodyScroll) {
        scrollWrap.scrollLeft = this.ui.bodyScroll.scrollLeft;
      }
      return;
    }

    /* Per-category rows */
    let rowIdx = 0;
    for (const cat of this.getVisibleCategories()) {
      const catTotal = this.summary.categoryTotals[cat];

      const catRow = document.createElement("div");
      catRow.className = "ts-summary-row ts-summary-category";
      if (rowIdx % 2 === 1) catRow.classList.add("ts-summary-row-alt");
      rowIdx++;

      const catLabel = document.createElement("div");
      catLabel.className = "ts-summary-label";
      catLabel.textContent = cat;
      catRow.appendChild(catLabel);

      const catDailyTotals = this.summary.categoryDailyTotals[cat];
      for (let d = 0; d < this.daysInMonth; d++) {
        const cell = document.createElement("div");
        cell.className = "ts-summary-cell";
        if (isWeekend(this.year, this.month, d + 1)) cell.classList.add("ts-weekend-summary");
        const val = catDailyTotals[d];
        cell.textContent = val > 0 ? val.toFixed(2) : "";
        catRow.appendChild(cell);
      }

      const totalCell = document.createElement("div");
      totalCell.className = "ts-summary-cell ts-summary-cell-total";
      totalCell.textContent = catTotal.toFixed(2);
      catRow.appendChild(totalCell);

      grid.appendChild(catRow);
    }

    /* Daily totals row — at the bottom with dividing line */
    const dailyRow = document.createElement("div");
    dailyRow.className = "ts-summary-row ts-summary-row-daily";

    const dailyLabel = document.createElement("div");
    dailyLabel.className = "ts-summary-label ts-summary-label-daily";
    dailyLabel.textContent = "Daily Total";
    dailyRow.appendChild(dailyLabel);

    for (let d = 0; d < this.daysInMonth; d++) {
      const cell = document.createElement("div");
      cell.className = "ts-summary-cell";
      if (isWeekend(this.year, this.month, d + 1)) cell.classList.add("ts-weekend-summary");
      const val = this.summary.dailyTotals[d];
      cell.textContent = val > 0 ? val.toFixed(2) : "";
      dailyRow.appendChild(cell);
    }

    const dailyTotalCell = document.createElement("div");
    dailyTotalCell.className = "ts-summary-cell ts-summary-cell-total";
    dailyTotalCell.textContent = this.summary.overall.toFixed(2);
    dailyRow.appendChild(dailyTotalCell);
    grid.appendChild(dailyRow);

    scrollWrap.appendChild(grid);
    this.ui.summaryPanel.appendChild(scrollWrap);

    this.attachSummaryScrollSync();

    if (this.ui.bodyScroll) {
      scrollWrap.scrollLeft = this.ui.bodyScroll.scrollLeft;
    }
  }

// ----------------------------------------------------------------
//#endregion
//#region DATA LOADING
// ----------------------------------------------------------------

  private async loadCategories(): Promise<void> {
    const options = await this.provider.getChoiceOptions({
      entityLogicalName: "cr91f_timesheet",
      choiceFieldLogicalName: "cr91f_timesheetcategory",
    });
    TimeSheetCategories.length = 0;
    TimeSheetCategories.push(...options);
  }

  private async loadSubCategories(): Promise<void> {
    const monthStart = `${this.year}-${String(this.month).padStart(2, "0")}-01 00:00:00`;
    const xmlQuery = `<fetch><entity name="cr91f_timesheetcategory"><attribute name="cr91f_categoryid"/><attribute name="cr91f_enddate"/><attribute name="cr91f_startdate"/><attribute name="cr91f_timesheetcategory"/><attribute name="cr91f_timesheetcategoryid"/><attribute name="cr91f_timesheetitem"/><attribute name="statecode"/><filter><condition attribute="cr91f_startdate" operator="le" value="${monthStart}"/><filter type="or"><condition attribute="cr91f_enddate" operator="null"/><condition attribute="cr91f_enddate" operator="ge" value="${monthStart}"/></filter></filter></entity></fetch>`;
    const raw = await this.provider.getLookupValues({
      logicalName: "cr91f_timesheetcategory",
      xmlQuery,
      filters: null,
    });
    SubCategories.length = 0;
    SubCategories.push(...parseLookupToSubCategories(raw));
  }

  /** Load all data for the current month */
  private async loadData(): Promise<void> {
    const [accounts, records] = await Promise.all([
      this.provider.listAccounts(this.year, this.month),
      this.provider.listTimesheets(this.year, this.month),
    ]);

    this.accountsCache = accounts;
    this.daysInMonth = getDaysInMonth(this.year, this.month);
    this.rows = [];

    const ym = yearMonthStr(this.year, this.month);

    /* Build rows from saved records */
    for (const rec of records) {
      const catEntry = rec.timeSheetCategory
        ? TimeSheetCategories.find(c => parseInt(c.keyValue, 10) === rec.timeSheetCategory)
        : null;
      const cat = catEntry?.keyName || "Customer Deliverable";
      let sub = rec.subCategory || "";
      if (!isDynamicCategory(cat) && rec.timeSheetSubCategoryId) {
        const subEntry = SubCategories.find(s => s.timeSheetCategoryId === rec.timeSheetSubCategoryId);
        if (subEntry) sub = subEntry.timeSheetItem;
      }
      const row = this.enrichRecord(rec, cat, sub);
      this.rows.push(row);
    }

    /* Detect stage from existing rows */
    this.detectStage();

    /* For non-dynamic categories, ensure all sub-categories from API have rows (even if unsaved) */
    for (const cat of getAllCategories()) {
      if (isDynamicCategory(cat)) continue;
      const subs = getSubcategoriesForCategory(cat);
      for (const sub of subs) {
        const exists = this.rows.some(r => r.category === cat && r.subCategory === sub.timeSheetItem);
        if (!exists) {
          this.rows.push(this.createEmptyRow(cat, sub.timeSheetItem, ym));
        }
      }
    }

    this.computeSummary();
  }

  private enrichRecord(rec: TimesheetRecord, category: CategoryName, subCategory: string): TimesheetRecord {
    rec.category = category;
    rec.subCategory = subCategory;
    rec.timeSheetCategory = getCategoryValue(category);
    if (isDynamicCategory(category)) {
      rec.timeSheetSubCategoryId = null;
      rec.accountId = rec.accountId || null;
      if (rec.accountId) {
        const acct = this.accountsCache.find(a => a.accountId === rec.accountId);
        if (acct && !rec.subCategory) rec.subCategory = acct.accountName;
      }
    } else {
      rec.accountId = null;
      const sub = SubCategories.find(s => s.timeSheetItem === subCategory && getCategoryNameForSubCategory(s) === category);
      rec.timeSheetSubCategoryId = sub?.timeSheetCategoryId ?? null;
    }
    return rec;
  }

  private createEmptyRow(category: CategoryName, subCategory: string, ym: string): TimesheetRecord {
    const sub = isDynamicCategory(category)
      ? null
      : SubCategories.find(s => s.timeSheetItem === subCategory && getCategoryNameForSubCategory(s) === category) ?? null;
    const row: any = {
      timeSheetId: null,
      timeSheetEntryId: null,
      yearMonth: ym,
      category,
      subCategory,
      timeSheetCategory: getCategoryValue(category),
      timeSheetSubCategoryId: sub?.timeSheetCategoryId ?? null,
      accountId: null,
      userId: _currentUser?.userId ?? "",
      timeSheetStage: TimeSheetStageMap[this.stageText],
      totalHours: 0,
      timeSheetStartDate: isoStartOfMonth(this.year, this.month),
      timeSheetEndDate: isoEndOfMonth(this.year, this.month),
    };
    for (let d = 1; d <= 31; d++) {
      row[dayKey(d)] = 0;
    }
    return row as TimesheetRecord;
  }

  /** Detect stage from existing rows (first row's stage wins) */
  private detectStage(): void {
    if (this.rows.length > 0) {
      const stage = this.rows[0].timeSheetStage;
      if (stage === TimeSheetStageMap.Submitted) this.stageText = "Submitted";
      else if (stage === TimeSheetStageMap.Finalised) this.stageText = "Finalised";
      else this.stageText = "Pending";
    } else {
      this.stageText = "Pending";
    }
    if (this.ui.stageSelect) this.ui.stageSelect.value = this.stageText;
  }

  /** Get rows filtered for the selected tab */
  private getRowsForTab(tab: CategoryName): TimesheetRecord[] {
    const filtered = this.rows.filter(r => r.category === tab);
    filtered.sort((a, b) => a.subCategory.localeCompare(b.subCategory));
    return filtered;
  }

// ----------------------------------------------------------------
//#endregion
//#region SUMMARY COMPUTATION
// ----------------------------------------------------------------

  /** Recompute summary from in-memory rows */
  private computeSummary(): void {
    const expanded = this.summary.expandedCategories;
    this.summary = this.emptySummary();
    this.summary.expandedCategories = expanded;

    for (const row of this.rows) {
      const cat = row.category;
      const sub = row.subCategory;
      for (let d = 1; d <= this.daysInMonth; d++) {
        const v = row[dayKey(d)];
        const n = typeof v === "number" ? v : Number(v) || 0;
        if (!Number.isFinite(n)) continue;
        this.summary.dailyTotals[d - 1] += n;
        this.summary.categoryTotals[cat] += n;
        this.summary.categoryDailyTotals[cat][d - 1] += n;
        if (!this.summary.subcategoryTotals[cat][sub]) {
          this.summary.subcategoryTotals[cat][sub] = 0;
        }
        this.summary.subcategoryTotals[cat][sub] += n;
        this.summary.overall += n;
      }
    }

    this.summary.overall = Number(this.summary.overall.toFixed(2));
    for (const cat of getAllCategories()) {
      this.summary.categoryTotals[cat] = Number(this.summary.categoryTotals[cat].toFixed(2));
    }
  }

// ----------------------------------------------------------------
//#endregion
//#region STAGE & UI GATING
// ----------------------------------------------------------------

  /** Apply stage to UI — enable/disable buttons and inputs */
  private applyStageToUi(): void {
    const isPending = this.stageText === "Pending";
    const isDynamic = isDynamicCategory(this.selectedTab);

    if (this.ui.addRowBtn) {
      this.ui.addRowBtn.disabled = !isPending || !isDynamic;
      this.ui.addRowBtn.style.display = isDynamic ? "" : "none";
    }
    if (this.ui.prevMonthBtn) {
      this.ui.prevMonthBtn.disabled = !isPending || !isDynamic;
      this.ui.prevMonthBtn.style.display = isDynamic ? "" : "none";
    }

    const copyLabel = this.ui.footerSection?.querySelector(".ts-copy-prev-label") as HTMLElement | null;
    if (copyLabel) copyLabel.style.display = isDynamic ? "" : "none";
  }

// ----------------------------------------------------------------
//#endregion
//#region MONTH & STAGE HANDLERS
// ----------------------------------------------------------------

  /** Handle month picker change */
  private async handleMonthChange(value: string): Promise<void> {
    const [yStr, mStr] = value.split("-");
    this.year = Number(yStr);
    this.month = Number(mStr);
    this.daysInMonth = getDaysInMonth(this.year, this.month);
    this.updateStatus("Current activity: Idle");
    showLoader("Loading data\u2026");
    try {
      await this.loadSubCategories();
      await this.loadData();
      this.renderAll();
    } catch (err: any) {
      showErrorScreen(err?.message || "Failed to load data.");
    }
    hideLoader();
  }

  /** Handle stage dropdown change */
  private async handleStageChange(newStage: StageText): Promise<void> {
    this.stageText = newStage;
    const stageValue = TimeSheetStageMap[newStage];

    for (const row of this.rows) {
      row.timeSheetStage = stageValue;
    }

    const savedRows = this.rows.filter(r => r.timeSheetId !== null);
    if (savedRows.length > 0) {
      showLoader("Updating stage\u2026");
      try {
        await this.provider.updateTimesheets(savedRows);
        this.updateStatus(`Stage updated to ${newStage}.`);
      } catch (err: any) {
        showErrorScreen(err?.message || "Failed to update stage.");
      }
      hideLoader();
    }

    this.renderBody();
    this.applyStageToUi();
  }

// ----------------------------------------------------------------
//#endregion
//#region TAB SWITCHING
// ----------------------------------------------------------------

  /** Switch to a different category tab */
  private setTab(category: CategoryName): void {
    if (this.selectedTab === category) return;
    this.selectedTab = category;
    this.renderTabs();
    this.renderHeader();
    this.renderBody();
    this.applyStageToUi();
  }

// ----------------------------------------------------------------
//#endregion
//#region ADD / EDIT / DELETE / COPY (dynamic categories only)
// ----------------------------------------------------------------

  /** Add a new customer/prospect row */
  private addCustomerRow(): void {
    if (this.stageText !== "Pending") return;
    if (!isDynamicCategory(this.selectedTab)) return;

    const ym = yearMonthStr(this.year, this.month);
    const row = this.createEmptyRow(this.selectedTab, "", ym);
    (row as any)._editing = true;
    this.rows.push(row);
    this.renderBody();
  }

  /** Edit an existing customer row (swap account) */
  private editCustomerRow(row: TimesheetRecord): void {
    if (this.stageText !== "Pending") return;
    (row as any)._prevAccountId = row.accountId;
    (row as any)._prevSubCategory = row.subCategory;
    (row as any)._editing = true;
    this.renderBody();
  }

  /** Delete a customer row */
  private deleteCustomerRow(row: TimesheetRecord): void {
    if (this.stageText !== "Pending") return;

    const doDelete = async () => {
      if (row.timeSheetId) {
        showLoader("Deleting row\u2026");
        try {
          await this.provider.deleteTimeKeepEntry(row.timeSheetId);
        } catch (err: any) {
          showErrorScreen(err?.message || "Failed to delete entry.");
          hideLoader();
          return;
        }
      }
      const idx = this.rows.indexOf(row);
      if (idx !== -1) this.rows.splice(idx, 1);
      this.computeSummary();
      this.renderBody();
      this.renderSummary();
      this.updateStatus("Row deleted successfully.");
      hideLoader();
    };

    if (rowHasHours(row, this.daysInMonth)) {
      showDeleteConfirmModal(async (confirmed) => {
        if (confirmed) await doDelete();
      });
    } else {
      void doDelete();
    }
  }

  /** Copy rows from previous month (dynamic categories only) */
  private async copyFromPreviousMonth(): Promise<void> {
    if (this.stageText !== "Pending") return;
    if (!isDynamicCategory(this.selectedTab)) return;

    let prevYear = this.year;
    let prevMonth = this.month - 1;
    if (prevMonth < 1) { prevMonth = 12; prevYear -= 1; }

    showLoader("Copying from previous month\u2026");
    try {
      const prevRecords = await this.provider.listTimesheets(prevYear, prevMonth);
      const currentAccountIds = new Set(
        this.getRowsForTab(this.selectedTab).filter(r => r.accountId).map(r => r.accountId!)
      );

      const tabCategoryFilter = this.selectedTab;
      const toCopy = prevRecords.filter(rec => {
        const cat = rec.category || "Customer Deliverable";
        return cat === tabCategoryFilter && rec.accountId && !currentAccountIds.has(rec.accountId);
      });

      if (toCopy.length === 0) {
        this.updateStatus("No new entries to copy from previous month.");
        hideLoader();
        return;
      }

      const ym = yearMonthStr(this.year, this.month);
      const newRecords: TimesheetRecord[] = toCopy.map(rec => {
        const newRec: any = {
          ...rec,
          timeSheetId: null,
          yearMonth: ym,
          timeSheetStage: TimeSheetStageMap.Pending,
          totalHours: 0,
          timeSheetStartDate: isoStartOfMonth(this.year, this.month),
          timeSheetEndDate: isoEndOfMonth(this.year, this.month),
          timeSheetEntryId: `${ym.replace("-", "")}_${this.accountsCache.find(a => a.accountId === rec.accountId)?.catNumber ?? "000"}_${_currentUser!.fullName[0]}.${_currentUser!.fullName.trim().split(" ").slice(-1)[0]}`,
        };
        for (let d = 1; d <= 31; d++) newRec[dayKey(d)] = 0;
        return newRec as TimesheetRecord;
      });

      const created = await this.provider.createTimesheets(newRecords);

      for (const rec of created) {
        const cat = rec.category || this.selectedTab;
        const sub = rec.subCategory || this.accountsCache.find(a => a.accountId === rec.accountId)?.accountName || "";
        this.rows.push(this.enrichRecord(rec, cat, sub));
      }

      this.computeSummary();
      this.renderBody();
      this.renderSummary();
      this.updateStatus(`Copied ${created.length} entries from previous month.`);
    } catch (err: any) {
      showErrorScreen(err?.message || "Failed to copy from previous month.");
    }
    hideLoader();
  }

// ----------------------------------------------------------------
//#endregion
//#region SAVE LOGIC (debounced, per-row)
// ----------------------------------------------------------------

  /**
   * Queue an update for a row and debounce the save (500 ms).
   * Does not save if stage is Submitted/Finalised.
   * Does not POST/PUT unless at least one day has hours > 0 (except dynamic categories which always save).
   */
  private scheduleSave(row: TimesheetRecord, updateFn?: () => void): void {
    if (
      row.timeSheetStage === TimeSheetStageMap.Submitted ||
      row.timeSheetStage === TimeSheetStageMap.Finalised
    ) return;

    if (updateFn) {
      if (!this.saveQueues.has(row)) this.saveQueues.set(row, []);
      this.saveQueues.get(row)!.push(updateFn);
    }

    const existing = this.saveTimers.get(row);
    if (existing) window.clearTimeout(existing);

    this.updateStatus("Auto saving\u2026");

    const handle = window.setTimeout(() => {
      const queue = this.saveQueues.get(row);
      if (queue && queue.length > 0) {
        while (queue.length > 0) {
          const fn = queue.shift();
          if (fn) fn();
        }
      }
      this.saveRow(row);
    }, 500);
    this.saveTimers.set(row, handle);
  }

  /** Persist a row to the backend */
  private async saveRow(row: TimesheetRecord): Promise<void> {
    if (!isDynamicCategory(row.category) && !rowHasHours(row, this.daysInMonth)) {
      this.updateStatus("No hours to save for this row.");
      return;
    }

    row.totalHours = sumRowHours(row, this.daysInMonth);

    if (!row.timeSheetEntryId && !isDynamicCategory(row.category)) {
      const ym = yearMonthStr(this.year, this.month);
      const subEntry = SubCategories.find(s => s.timeSheetCategoryId === row.timeSheetSubCategoryId);
      const catId = subEntry?.categoryId ?? row.category.replace(/[^a-zA-Z]/g, "").substring(0, 5);
      row.timeSheetEntryId = `${ym.replace("-", "")}_${catId}_${_currentUser!.fullName[0]}.${_currentUser!.fullName.trim().split(" ").slice(-1)[0]}`;
      row.yearMonth = ym;
      row.timeSheetStartDate = isoStartOfMonth(this.year, this.month);
      row.timeSheetEndDate = isoEndOfMonth(this.year, this.month);
      row.userId = _currentUser?.userId ?? "";
      row.timeSheetStage = TimeSheetStageMap[this.stageText];
    }

    try {
      const saved = await this.provider.saveTimesheet(row as TimesheetRecord);
      row.timeSheetId = saved.timeSheetId;
      row.timeSheetEntryId = saved.timeSheetEntryId;
      for (let d = 1; d <= 31; d++) {
        const k = dayKey(d);
        row[k] = saved[k] ?? row[k];
      }
      this.updateStatus("Last modified: " + new Date().toLocaleString());
    } catch (err: any) {
      console.error(err);
      this.updateStatus("Error while saving. Retrying on next change.");
      showErrorScreen(err?.message || "Error while saving timesheet.");
    }
  }

// ----------------------------------------------------------------
//#endregion
//#region USER DISPLAY
// ----------------------------------------------------------------

  private updateUserDisplay(uiProfile?: { name?: string; email?: string; photoUrl?: string | null }): void {
    if (!this.ui.userPill) return;
    this.ui.userPill.innerHTML = "";
    const name = _currentUser?.fullName ?? uiProfile?.name ?? "User";
    const wrapper = document.createElement("div");
    wrapper.className = "ts-user-avatar-wrapper";

    if (uiProfile?.photoUrl) {
      const img = document.createElement("img");
      img.src = uiProfile.photoUrl;
      img.alt = name;
      img.className = "ts-user-avatar-pic";
      img.title = name;
      wrapper.appendChild(img);
    } else {
      const initials = name
        ? name.split(" ").map((n: string) => n[0]).join("").slice(0, 2).toUpperCase()
        : "?";
      const span = document.createElement("span");
      span.className = "ts-user-avatar-initials";
      span.textContent = initials;
      wrapper.appendChild(span);
    }

    const nameSpan = document.createElement("span");
    nameSpan.className = "ts-user-fullname";
    nameSpan.textContent = name;
    wrapper.appendChild(nameSpan);

    this.ui.userPill.appendChild(wrapper);
  }

  private updateStatus(text: string): void {
    if (this.ui.statusText) this.ui.statusText.textContent = text;
  }
}

// ================================================================
//#endregion
//#region INITIALISATION
// ================================================================

/** Initialise the Phase 2 timesheet app */
async function initTimesheetPhase2(container: HTMLDivElement): Promise<TimesheetAppV2> {
  try {
    console.log(`Time Keep ${APP_VERSION}`);
    await loadMsalFromCdn();
    const currentUrl = window.location.href;
    if (currentUrl.includes("localhost:8080") || currentUrl.includes("127.0.0.1:8080")) {
      _isDev = true;
    }
    const provider = new ApiProvider();
    const app = new TimesheetAppV2(container, provider);
    void app.mount();
    return app;
  } catch (err: any) {
    showErrorScreen(err?.message || "Unknown error");
    throw err;
  }
}

/** Auto-init when DOM is ready if element with id="app" exists */
document.addEventListener("DOMContentLoaded", async (): Promise<void> => {
  const container = document.getElementById("app") as HTMLDivElement | null;
  if (container) {
    showLoader("Loading\u2026");
    await initTimesheetPhase2(container);
  }
});

// ================================================================
//#endregion
// ================================================================