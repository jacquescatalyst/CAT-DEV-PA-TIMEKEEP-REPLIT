/* CatTimeSheet.ts
 * SPA to capture time per user, per account, per day — in a month view.
 * Starting point: initTimesheet(container: HTMLDivElement)
 */

// App version
const APP_VERSION = "v1.1.0";

/** Simple async mutex implementation */
class Mutex {
  private _locked = false;
  private _waiting: (() => void)[] = [];

  async lock() {
    if (this._locked) {
      await new Promise<void>(resolve => this._waiting.push(resolve));
    }
    this._locked = true;
  }

  unlock() {
    if (this._waiting.length > 0) {
      const next = this._waiting.shift();
      if (next) next();
    } else {
      this._locked = false;
    }
  }
}

/** No imports: use the global injected by the CDN */
declare const msal: any;

const _entraTenantId: string = "9d97e7fb-d431-4e72-a93c-c5452889c391";
const _entraClientId: string = "e43b1f85-24b8-41bd-b1a7-8c0248a1f153";
const _entraPaSharedId: string = "483a35f3-b479-4461-81f2-ee40ceddcda2";
const _lockTokenMutex = new Mutex();

// Global variable to hold the fetched user
let _UserEmail: string = "";
let _currentUser: User | null = null;
let _userPill: HTMLDivElement | null = null;
let _msalApp: any | null = null;
let _isDev: boolean = false;
let _stageText: string = "Pending";

/** UI row model used by the app */
interface RowModel {
  accountId: Guid | null;
  accountName: string | null;
  record: TimesheetRecord;
  lastSaved?: Date;
  saving?: boolean;
}

type Guid = string;

interface TimesheetRecord {
  // Day fields aligned to PowerApps table
  [key: `Day${string}`]: number; // e.g., day01..31
  timeSheetStartDate: string; // ISO e.g. "2025-12-01T00:00:00Z"
  timeSheetEndDate: string;   // ISO e.g. "2025-12-31T00:00:00Z"
  timeSheetId: Guid | null;
  timeSheetEntryId: string;
  totalHours: number;
  timeSheetStage: number; // Pending(596610000), Submitted(596610001), Finalised(596610002)
  yearMonth: string; // e.g. "2025-12"
  accountId: Guid;     // Account id
  userId: Guid;    // lookup to systemuser, not used here
}

/** Mapping of timesheet stage text to number value */
const TimeSheetStageMap = {
  Pending: 596610000,
  Submitted: 596610001,
  Finalised: 596610002
} as const;

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

  interface DataProvider {
  /** List of accounts available to the user for the month. */    
  listAccounts(year: number, month: number): Promise<Account[]>;

  /** Existing timesheet rows for the month (per account). */
  listTimesheets(year: number, month: number): Promise<TimesheetRecord[]>;

  /** Create or update a timesheet row. */
  saveTimesheet(record: TimesheetRecord): Promise<TimesheetRecord>;

  /** Batch update timesheet stages for multiple records. */
  updateTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]>;

  /** Batch create timesheet records. */
  createTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]>;

  /** Delete timesheet record. */
  deleteTimeKeepEntry(timeSheetId: Guid): Promise<void>;
  
  /** Batch delete timesheet records. */
  deleteTimeKeepEntries(timeSheetIds: Guid[]): Promise<void>;

  /** Fetch the current user info. */
  getUser(): Promise<User>;
}

/** Load MSAL from CDN once, return when window.msal is available 
 * @returns Promise that resolves when MSAL is loaded.
*/
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

/** Create/reuse a singleton MSAL instance 
 * @returns The MSAL PublicClientApplication instance.
*/
async function getMsalApp(): Promise<any> {
  if (_msalApp) return _msalApp;

  const msalConfig = {
    auth: {
      clientId: _entraClientId, // Your Azure AD app client ID
      authority: `https://login.microsoftonline.com/${_entraTenantId}`, // Your Azure AD tenant ID
      // Dataverse web resource you added (see notes below)
      redirectUri: `${window.location.origin}/authTokenRedirect`
    },
    cache: { cacheLocation: "sessionStorage" }
  };
  _msalApp = new msal.PublicClientApplication(msalConfig);
  return _msalApp;
}

/** Silent-first token acquisition 
 * @returns Promise that resolves when token is acquired.
*/
async function ensureTokenForFunction(): Promise<string> {
  // Check local sessionStorage for existing token
  await _lockTokenMutex.lock();
  try {
    const existingToken = sessionStorage.getItem("msalToken");
    if (existingToken) {
      // Validate token expiration, if less than 5 min to expire, get a new one
      const tokenParts = existingToken.split('.');
      const payload = JSON.parse(atob(tokenParts[1]));
      _UserEmail = payload.preferred_username ?? (() => { throw new Error("Unable to determine user email from token payload."); })();
      const exp = payload.exp;
      const iat = payload.iat;
      const currentTime = Math.floor(Date.now() / 1000); // Current time in seconds
      if (exp - currentTime > 300) { // 300 seconds = 5 minutes
        return existingToken;
      }
    }

    // API ID and scopes
    const scopes = [`api://${_entraPaSharedId}/user_impersonation`];
    const app = await getMsalApp();
    let account = app.getAllAccounts()[0];
    let msalToken: string = "";

    try {
      // Try silent first
      const res = await app.acquireTokenSilent({ account, scopes });
      msalToken = res.accessToken;
    } catch {
      // Use popup if silent fails
      try {
        const login = await app.loginPopup({ scopes });
        const res = await app.acquireTokenSilent({ account: login.account, scopes });
        msalToken = res.accessToken;
      } catch (err) {
        // Handle popup blocked or other errors here
        throw new Error(err || "Authentication failed. Please ensure popups are allowed and try again.");
      }
    }

    sessionStorage.setItem("msalToken", msalToken);
    const payload = JSON.parse(atob(msalToken.split('.')[1]));
    _UserEmail = payload.preferred_username ?? (() => { throw new Error("Unable to determine user email from token payload."); })();

    return msalToken;
  } finally {
    _lockTokenMutex.unlock();
  }
}

/** OIDC profile info */
type OidcProfile = {
  name?: string;
  preferred_username?: string;
  email?: string;
  sub?: string;
  tenantId?: string;
};

/** Fetch OIDC profile from ID token claims */
async function getOidcProfile(): Promise<OidcProfile | null> {
  const app = await getMsalApp();

  // Ensure we have an account (if you only ever login with your API scopes, consider a light login with OIDC scopes)
  const account = app.getAllAccounts()[0];
  if (!account) return null;

  // MSAL stores ID token claims on the account
  const claims = account.idTokenClaims as Record<string, unknown> | undefined;
  if (!claims) return null;

  const profile: OidcProfile = {
    name: claims["name"] as string | undefined,
    preferred_username: claims["preferred_username"] as string | undefined,
    email: claims["email"] as string | undefined,
    sub: claims["sub"] as string | undefined,
    tenantId: (claims["tid"] as string | undefined) ?? (claims["tenant"] as string | undefined),
  };

  return profile;
}

/** Ensure user is logged in with OIDC scopes */
async function ensureOidcLogin(): Promise<void> {
  const app = await getMsalApp();
  if (app.getAllAccounts().length > 0) return;

  await app.loginPopup({ scopes: ["openid", "profile", "email"] });
}

/** Acquire Microsoft Graph access token */
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

/** Microsoft Graph /me response type */
type GraphMe = {
  id: string;
  displayName: string;
  mail?: string;                 // Preferred mailbox address
  userPrincipalName: string;     // Fallback if mail is null
  givenName?: string;
  surname?: string;
  jobTitle?: string;
};

/** Fetch /me from Microsoft Graph */
async function getGraphMe(): Promise<GraphMe> {
  const token = await acquireGraphAccessToken();

  const resp = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!resp.ok) {
    throw new Error(`Graph /me failed: ${resp.status} ${resp.statusText}`);
  }

  const me = await resp.json();
  return me as GraphMe;
}

/** Fetch user's photo from Microsoft Graph */
async function getGraphPhotoUrl(): Promise<string | null> {
  const token = await acquireGraphAccessToken();

  const resp = await fetch("https://graph.microsoft.com/v1.0/me/photo/$value", {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!resp.ok) {
    // No photo or permission issue; return null and let UI show a placeholder
    return null;
  }

  const blob = await resp.blob();
  // Create a local object URL (remember to revoke when you unmount)
  return URL.createObjectURL(blob);
}

/** Get current user's UI profile (name, email, photo) */
async function getCurrentUserUiProfile(): Promise<{
  name?: string;
  email?: string;
  photoUrl?: string | null;
}> {
  // OIDC claims (fast, cached)
  const oidc = await getOidcProfile();

  // Graph (authoritative + photo)
  const me = await getGraphMe();
  const photoUrl = await getGraphPhotoUrl();

  // Prefer Graph email; fall back to OIDC claim or UPN
  const email = me.mail ?? me.userPrincipalName ?? oidc?.email ?? oidc?.preferred_username;

  return {
    name: oidc?.name ?? me.displayName,
    email,
    photoUrl
  };
}


/** API-based data provider implementation */
class ApiProvider implements DataProvider {
  async listAccounts(_year: number, _month: number): Promise<Account[]> {
    let url = _isDev ? "http://localhost:7071/api/getaccounts" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/getaccountswithaadauth"
    const response = await fetch(url, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
    });
    if (!response.ok) throw new Error("Failed to fetch accounts");
    return await response.json();
  }

  async listTimesheets(year: number, month: number): Promise<TimesheetRecord[]> {
    let url = _isDev ? "http://localhost:7071/api/gettimekeepentries" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/gettimekeepentrieswithaadauth"
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify({ userId: _currentUser?.userId, yearMonth: `${year}-${String(month).padStart(2, "0")}` }),
    });
    if (!response.ok) throw new Error("Failed to fetch timesheets");
    return await response.json();
  }

  async saveTimesheet(record: TimesheetRecord): Promise<TimesheetRecord> {
    const url = record.timeSheetId
      ? _isDev ? "http://localhost:7071/api/updatetimekeepentry" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/updatetimekeepentrywithaadauth"
      : _isDev ? "http://localhost:7071/api/createtimekeepentry" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/createtimekeepentrywithaadauth";
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify(record),
    });
    if (!response.ok) throw new Error("Failed to save timesheet");
    return await response.json();
  }

  async updateTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]> {
    let url = _isDev ? "http://localhost:7071/api/updatetimekeepentries" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/updatetimekeepentrieswithaadauth"
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify(records),
    });
    if (!response.ok) throw new Error("Failed to save timesheet");
    return await response.json();
  }

  async createTimesheets(records: TimesheetRecord[]): Promise<TimesheetRecord[]> {
    let url = _isDev ? "http://localhost:7071/api/createtimekeepentries" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/createtimekeepentrieswithaadauth";
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify(records),
    });
    if (!response.ok) throw new Error("Failed to save timesheet");
    return await response.json();
  }

  async deleteTimeKeepEntry(timeSheetId: Guid): Promise<void> {
    let url = _isDev ? "http://localhost:7071/api/deletetimekeepentry" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/deletetimekeepentrywithaadauth";
    const response = await fetch(url, {
      method: "DELETE",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify({ timeSheetId }),
    });
    if (!response.ok) throw new Error("Failed to delete timesheet entry");
  }

  async deleteTimeKeepEntries(timeSheetIds: Guid[]): Promise<void> {
    let url = _isDev ? "http://localhost:7071/api/deletetimekeepentries" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/deletetimekeepentrieswithaadauth";
    const response = await fetch(url, {
      method: "DELETE",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify({ timeSheetIds }),
    });
    if (!response.ok) throw new Error("Failed to delete timesheet entries");
  }

  async getUser(): Promise<User> {
    let url = _isDev ? "http://localhost:7071/api/getuser" : "https://cat-dev-pa-dataverse-b6fmgzfgd4f9fkfd.canadacentral-01.azurewebsites.net/api/getuserwithaadauth";
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${await ensureTokenForFunction()}`
      },
      body: JSON.stringify({ emailAddress: _UserEmail }),
    });
    if (!response.ok) throw new Error("Failed to fetch user");
    const user: User = await response.json();
    _currentUser = user;
    return user;
  }
}

/** Utility: pad day number to PowerApps field key */
const dayKey = (d: number) => `day${String(d).padStart(2, "0")}` as const;

const sumHours = (rec: TimesheetRecord) => {
  let total = 0;
  for (let d = 1; d <= 31; d++) {
    const k = dayKey(d);
    const val = (rec[k] ?? 0) as number;
    total += Number.isFinite(val) ? val : 0;
  }
  return Number(total.toFixed(2));
};

/** Month helpers */
function getDaysInMonth(year: number, month: number): number {
  return new Date(year, month, 0).getDate(); // month is 1-based here
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
  // Use UTC weekday by formatting with a locale and timeZone: 'UTC'
  const weekday = dt.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" });
  const dayNum = dt.getUTCDate();
  return `${weekday} ${String(dayNum).padStart(2, "0")}`; // e.g. "Mon 01"
}

/** Main App */
class TimesheetApp {
  private container: HTMLDivElement;
  private provider: DataProvider;

  // State
  private year: number;
  private month: number; // 1..12
  private daysInMonth: number;
  private rows: RowModel[] = [];
  private Accounts: Account[] = [];

  // Elements
  private toolbar!: HTMLDivElement;
  private gridScrollWrap!: HTMLDivElement;
  private bodyScroll!: HTMLDivElement;
  private headerRow!: HTMLDivElement;
  private gridBody!: HTMLDivElement;
  private addRowBtn!: HTMLButtonElement;
  private statusText!: HTMLDivElement;
  private prevMonthBtn!: HTMLButtonElement;

  // Debounce buffers and update queues per row
  private saveTimers = new Map<RowModel, number>();
  private saveQueues = new Map<RowModel, (() => void)[]>();

  constructor(container: HTMLDivElement, provider: DataProvider) {
    this.container = container;
    this.provider = provider;

    const today = new Date();
    this.year = today.getFullYear();
    this.month = today.getMonth() + 1;
    this.daysInMonth = getDaysInMonth(this.year, this.month);
  }

  async mount() {
    // Ensure .ts-app exists before any async calls that may error
    this.container.classList.add("ts-app");
    this.container.innerHTML = ""; // clear
    ensureLoader();
      showLoader("Logging in...");
    try {
      await ensureTokenForFunction(); // kick off token fetch
    } catch (err: any) {
      hideLoader();
      showErrorScreen(err?.message || "Failed to load user or data.");
      // Ensure error overlay is visible and on top
      const errorOverlay = document.getElementById("ts-error-overlay");
      if (errorOverlay) {
        errorOverlay.style.display = "flex";
        errorOverlay.style.zIndex = "9999";
      }
      return;
    }
    showLoader("Loading user and data...");
    this.buildShell();
    try {
      const uiProfile = await getCurrentUserUiProfile();
      await this.provider.getUser();
      this.UpdateUserDisplay(uiProfile);
      await this.loadInitialData();
      this.renderAll();
      hideLoader();
    } catch (err: any) {
      hideLoader();
      showErrorScreen(err?.message || "Failed to load user or data.");
      const errorOverlay = document.getElementById("ts-error-overlay");
      if (errorOverlay) {
        errorOverlay.style.display = "flex";
        errorOverlay.style.zIndex = "9999";
      }
    }
  }

  private buildShell() {
    // Top toolbar: select month + user pill + add customer
    this.toolbar = document.createElement("div");
    this.toolbar.className = "ts-toolbar";

    const left = document.createElement("div");
    left.className = "ts-toolbar-left";

    // Add logo image
    const logo = document.createElement("img");
    logo.src = "Catalyst-Logo-1200-Dark.png";
    logo.alt = "Catalyst Logo";
    logo.className = "ts-logo";

    const monthLabel = document.createElement("label");
    monthLabel.textContent = "Select month:";
    monthLabel.className = "ts-label";

    const monthInput = document.createElement("input");
    monthInput.type = "month";
    monthInput.value = `${this.year}-${String(this.month).padStart(2, "0")}`;
    // Restrict to 2025-01 or later, up to current month
    const now = new Date();
    const maxMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
    monthInput.min = "2025-01";
    monthInput.max = maxMonth;
    monthInput.className = "ts-month-input";
    monthInput.addEventListener("change", async () => {
      this.updateStatus("Current activity: Idle");
      showLoader("Loading data...");
      const [yStr, mStr] = monthInput.value.split("-");
      this.year = Number(yStr);
      this.month = Number(mStr);
      this.daysInMonth = getDaysInMonth(this.year, this.month);
      await this.loadInitialData();
      this.renderAll();
      hideLoader();
    });
    monthInput.addEventListener("click", () => {
      if (monthInput.showPicker) monthInput.showPicker();
    });


    // Add status dropdown
    const statusLabel = document.createElement("label");
    statusLabel.textContent = "Stage:";
    statusLabel.className = "ts-label";

    const statusSelect = document.createElement("select");
    statusSelect.className = "ts-status-select";
    const pendingOption = document.createElement("option");
    pendingOption.value = "Pending";
    pendingOption.textContent = "Pending";
    const submittedOption = document.createElement("option");
    submittedOption.value = "Submitted";
    submittedOption.textContent = "Submitted";
    statusSelect.append(pendingOption, submittedOption);

    // Determine default stage: Pending unless any row has a different stage
    this.setDefaultStageValue([]);

    // Add event to update all rows' stage and call batch API
    statusSelect.addEventListener("change", async () => {
      _stageText = statusSelect.value;
      const stageValue = TimeSheetStageMap[_stageText as keyof typeof TimeSheetStageMap];
      // Update all rows locally
      this.rows.forEach(row => {
        row.record.timeSheetStage = stageValue;
      });
      this.renderRows();
      // Enable/disable Add Customer button
      if (this.addRowBtn) {
        this.addRowBtn.disabled = _stageText !== "Pending";
      }
      // Enable/disable Copy from Prev Month button
      if (this.prevMonthBtn) {
        console.log('Stage changed:', _stageText, 'prevMonthBtn:', this.prevMonthBtn, 'disabled should be', _stageText !== 'Pending');
        this.prevMonthBtn.disabled = _stageText !== "Pending";
      } else {
        console.log('prevMonthBtn not found on stage change');
      }
      // Enable/disable edit/delete buttons in the DOM
      const editBtns = this.container.querySelectorAll<HTMLButtonElement>(".ts-edit-account-btn");
      const deleteBtns = this.container.querySelectorAll<HTMLButtonElement>(".ts-delete-account-btn");
      editBtns.forEach(btn => { btn.disabled = _stageText !== "Pending"; });
      deleteBtns.forEach(btn => { btn.disabled = _stageText !== "Pending"; });
      // Indicate saving status
      this.updateStatus("Saving stage...");
      try {
        await this.provider.updateTimesheets(this.rows.map(r => r.record));
        this.updateStatus(`All rows updated to stage: ${_stageText}`);
      } catch (err) {
        this.updateStatus("Error while saving stage. Retrying on next change.");
        showErrorScreen(err?.message || "Failed to update all stages.");
      }
    });

    // Add label and button for copying previous month's customers
    const copyPrevMonthWrap = document.createElement("div");
    copyPrevMonthWrap.style.display = "flex";
    copyPrevMonthWrap.style.alignItems = "center";
    copyPrevMonthWrap.style.gap = "8px";

    const copyPrevMonthLabel = document.createElement("span");
    copyPrevMonthLabel.textContent = "Copy Customers from Previous Month:";
    copyPrevMonthLabel.className = "ts-copy-prev-label";

    this.prevMonthBtn = document.createElement("button");
    this.prevMonthBtn.className = "ts-btn ts-btn-prev-month";
    this.prevMonthBtn.title = "Retrieve previous month's timesheet records";
    const copyIcon = document.createElement("span");
    copyIcon.className = "ts-icon-copy";
    this.prevMonthBtn.appendChild(copyIcon);
    this.prevMonthBtn.addEventListener("click", async () => {
      // Guard: Only allow if stage is Pending
      if (_stageText !== "Pending") {
        this.updateStatus("Cannot copy from previous month unless stage is Pending.");
        return;
      }
      this.updateStatus("Fetching previous month's records...");
      showLoader("Fetching previous month's records...");
      try {
        const prevRecords = await this.getPreviousMonthTimesheetRecords();
        // Map current records by accountId for quick lookup
        const currentByAccount = new Map<string, RowModel>();
        this.rows.forEach(row => {
          if (row.accountId) currentByAccount.set(row.accountId, row);
        });

        // Prepare new records to add (clear hours, set new meta)
        const ym = `${this.year}-${String(this.month).padStart(2, "0")}`;
        const newRows: RowModel[] = [];
        for (const prev of prevRecords) {
          // If already present, skip
          if (currentByAccount.has(prev.accountId)) continue;
          // Clone and clear hours
          // Build timeSheetEntryId using the same pattern as addRow/buildaccountSelect
          const accountEntry = this.Accounts.find(a => a.accountId === prev.accountId);
          const catNumber = accountEntry?.catNumber ?? "";
          const userInitial = _currentUser?.fullName[0] ?? "";
          const userLast = _currentUser?.fullName.trim().split(" ").slice(-1)[0] ?? "";
          const entryId = `${ym.replace("-", "")}_${catNumber}_${userInitial}.${userLast}`;

          const cleared: TimesheetRecord = {
            ...prev,
            yearMonth: ym,
            timeSheetId: null,
            timeSheetEntryId: entryId,
            timeSheetStartDate: isoStartOfMonth(this.year, this.month),
            timeSheetEndDate: isoEndOfMonth(this.year, this.month),
            timeSheetStage: TimeSheetStageMap.Pending,
            totalHours: 0,
          };
          // Clear all day fields
          for (let d = 1; d <= 31; d++) {
            cleared[dayKey(d)] = 0;
          }
          // Set userId to current user
          cleared.userId = _currentUser?.userId ?? "";
          // Find account name
          const account = this.Accounts.find(a => a.accountId === cleared.accountId);
          newRows.push({
            accountId: cleared.accountId,
            accountName: account?.accountName ?? "(Unknown)",
            record: cleared,
          });
        }

        // Add new rows to the current list
        this.rows.push(...newRows);
        this.renderRows();

        // Save new timesheet records if any
        if (newRows.length > 0) {
          this.updateStatus("Saving imported records...");
          try {
            const saved = await this.provider.createTimesheets(newRows.map(r => r.record));
            // Update the newRows with saved data
            for (let i = 0; i < newRows.length; i++) {
              newRows[i].record = saved[i];
              newRows[i].lastSaved = new Date();
            }
            this.updateStatus("Imported records saved.");
            this.renderRows();
          } catch (err) {
            this.updateStatus("Failed to save imported records.");
            showErrorScreen(err?.message || "Failed to save imported records.");
          }
        } else {
          this.updateStatus("No new records to import from previous month.");
        }
      } catch (err) {
        this.updateStatus("Failed to fetch previous month's records.");
        showErrorScreen(err?.message || "Failed to fetch previous month's records.");
      } finally {
        hideLoader();
      }
    });

    //left.append(logo, monthLabel, monthInput, statusLabel, statusSelect);
    copyPrevMonthWrap.append(copyPrevMonthLabel, this.prevMonthBtn);
    left.append(monthLabel, monthInput, statusLabel, statusSelect, copyPrevMonthWrap);

    const right = document.createElement("div");
    right.className = "ts-toolbar-right";

    _userPill = document.createElement("div");
    _userPill.className = "ts-user-pill";
    this.UpdateUserDisplay();

    right.appendChild(_userPill);

    this.toolbar.append(left, right);

    // Header + body scroll containers
    const grid = document.createElement("div");
    grid.className = "ts-grid";

    // Create scroll wrap for header and body
    this.gridScrollWrap = document.createElement("div");
    this.gridScrollWrap.className = "ts-grid-scroll-wrap";

    this.headerRow = document.createElement("div");
    this.headerRow.className = "ts-grid-header-row";
    this.gridScrollWrap.appendChild(this.headerRow);

    this.bodyScroll = document.createElement("div");
    this.bodyScroll.className = "ts-grid-body-scroll";
    this.gridBody = document.createElement("div");
    this.gridBody.className = "ts-grid-body";
    this.bodyScroll.appendChild(this.gridBody);
    this.gridScrollWrap.appendChild(this.bodyScroll);

    // Sync header scroll with body and vice versa
    this.bodyScroll.addEventListener("scroll", () => {
      this.headerRow.scrollLeft = this.bodyScroll.scrollLeft;
    });
    this.headerRow.addEventListener("scroll", () => {
      this.bodyScroll.scrollLeft = this.headerRow.scrollLeft;
    });

    // Add row + status into toolbar left
    this.addRowBtn = document.createElement("button");
    this.addRowBtn.className = "ts-btn";
    this.addRowBtn.textContent = "Add Customer";
    if (_stageText !== "Pending") {
      this.addRowBtn.disabled = true;
    }
    this.addRowBtn.addEventListener("click", () => {
      if (_stageText !== "Pending") return;
      this.addRow();
    });

    this.statusText = document.createElement("div");
    this.statusText.className = "ts-status";
    this.statusText.textContent = "Current activity: Idle";

    // Add new parent div for Add Customer and status with class names
    const appFooterSection = document.createElement("div");
    appFooterSection.className = "ts-app-footer-section";

    const appFooterLeft = document.createElement("div");
    appFooterLeft.className = "ts-app-footer-left";
    appFooterLeft.append(this.addRowBtn, this.statusText);

    const appFooterRight = document.createElement("div");
    appFooterRight.className = "ts-app-footer-right";
    appFooterRight.textContent = APP_VERSION;

    appFooterSection.append(appFooterLeft, appFooterRight);

    // Assemble
    this.container.append(this.toolbar, grid, appFooterSection);
    grid.append(this.gridScrollWrap);
  }

  private setDefaultStageValue(records: TimesheetRecord[]): string {
    // Directly set the stage value on the provided select element
    // Usage: this.setDefaultStageValue(records, statusSelect)
    // statusSelect: HTMLSelectElement
    if (!Array.isArray(records) || records.length === 0) return;
    const nonPendingRow = records.find(r => r.timeSheetStage !== TimeSheetStageMap.Pending);
    let defaultStageText: string = "Pending";
    if (nonPendingRow) {
      const stageEntry = Object.entries(TimeSheetStageMap).find(([k, v]) => v === nonPendingRow.timeSheetStage);
      if (stageEntry) {
        defaultStageText = stageEntry[0];
      }
    }
    // Find the status select element in the toolbar
    const statusSelect = this.toolbar?.querySelector<HTMLSelectElement>(".ts-status-select");
    if (statusSelect) {
      // Remove any existing Finalised option
      const finalisedOpt = Array.from(statusSelect.options).find(opt => opt.value === "Finalised");
      if (finalisedOpt) statusSelect.removeChild(finalisedOpt);

      // Remove Pending and Submitted options if present
      const pendingOpt = Array.from(statusSelect.options).find(opt => opt.value === "Pending");
      const submittedOpt = Array.from(statusSelect.options).find(opt => opt.value === "Submitted");

      // If default is Finalised, only show Finalised and disable
      if (defaultStageText === "Finalised") {
        if (pendingOpt) statusSelect.removeChild(pendingOpt);
        if (submittedOpt) statusSelect.removeChild(submittedOpt);
        const tempFinalised = document.createElement("option");
        tempFinalised.value = "Finalised";
        tempFinalised.textContent = "Finalised";
        statusSelect.appendChild(tempFinalised);
        statusSelect.value = "Finalised";
        statusSelect.disabled = true;
        _stageText = "Finalised";
      } else {
        // If Pending or Submitted, ensure those options exist and Finalised is not present
        if (!pendingOpt) {
          const opt = document.createElement("option");
          opt.value = "Pending";
          opt.textContent = "Pending";
          statusSelect.insertBefore(opt, statusSelect.firstChild);
        }
        if (!submittedOpt) {
          const opt = document.createElement("option");
          opt.value = "Submitted";
          opt.textContent = "Submitted";
          statusSelect.appendChild(opt);
        }
        _stageText = defaultStageText;
        statusSelect.value = defaultStageText;
        statusSelect.disabled = false;
      }
    }
    // Also update prevMonthBtn state on initial load
    if (this.prevMonthBtn) {
      this.prevMonthBtn.disabled = defaultStageText !== "Pending";
    }
  }

  private async loadInitialData() {
    const [accounts, existing] = await Promise.all([
      this.provider.listAccounts(this.year, this.month),
      this.provider.listTimesheets(this.year, this.month),
    ]);
    this.Accounts = accounts.slice().sort((a, b) => a.accountName.localeCompare(b.accountName));
    // convert to RowModels
    const ym = `${this.year}-${String(this.month).padStart(2, "0")}`;
    this.rows = existing.map(rec => {
      const account = this.Accounts.find(c => c.accountId === rec.accountId);
      return {
        accountId: rec.accountId,
        accountName: account?.accountName ?? "(Unknown)",
        record: rec,
        lastSaved: new Date(),
      };
    }).sort((a, b) => a.accountName.localeCompare(b.accountName));
    this.daysInMonth = getDaysInMonth(this.year, this.month);

    // Set stage value based on existing rows
    this.setDefaultStageValue(this.rows.map(r => r.record));
    // Ensure Add Customer button is enabled/disabled according to stage after stage is set
    if (this.addRowBtn) {
      this.addRowBtn.disabled = _stageText !== "Pending";
    }
    // Ensure prevMonthBtn is enabled/disabled according to stage after stage is set
    if (this.prevMonthBtn) {
      this.prevMonthBtn.disabled = _stageText !== "Pending";
    }
  }

  private renderAll() {
    this.renderHeader();
    this.renderRows();
    // this.updateStatus("Last modified: " + (new Date()).toLocaleString());
  }

  private renderHeader() {
    this.headerRow.innerHTML = "";
    const stickyCustomer = document.createElement("div");
    stickyCustomer.className = "ts-header-cell ts-header-customer";
    stickyCustomer.textContent = "Customer";
    this.headerRow.appendChild(stickyCustomer);

    for (let d = 1; d <= this.daysInMonth; d++) {
      const hc = document.createElement("div");
      hc.className = "ts-header-cell";
      hc.textContent = dayLabel(this.year, this.month, d);
      // highlight today
      const today = new Date();
      if (
        today.getUTCFullYear() === this.year &&
        today.getUTCMonth() + 1 === this.month &&
        today.getUTCDate() === d
      ) {
        hc.classList.add("ts-today");
      }
      this.headerRow.appendChild(hc);
    }

    const totalHeader = document.createElement("div");
    totalHeader.className = "ts-header-cell ts-header-total";
    totalHeader.textContent = "Total";
    this.headerRow.appendChild(totalHeader);
  }

  private renderRows() {
    this.gridBody.innerHTML = "";

    this.rows.forEach(row => {
      const rowEl = document.createElement("div");
      rowEl.className = "ts-row";

      // Customer cell (sticky)
      const custCell = document.createElement("div");
      custCell.className = "ts-cell ts-cell-customer";

      // If no account yet, show dropdown
      if (!row.accountId) {
        const select = this.buildaccountSelect(row);
        custCell.appendChild(select);
      } else {
        const labelWrap = document.createElement("div");
        labelWrap.className = "ts-account-label-wrap";

        // Customer name (left)
        const name = document.createElement("span");
        name.className = "ts-account-name";
        name.textContent = row.accountName ?? "account";
        labelWrap.appendChild(name);

        // Button container (right)
        const btnWrap = document.createElement("div");
        btnWrap.className = "ts-account-btn-wrap";

        // Edit button
        const editBtn = document.createElement("button");
        editBtn.className = "ts-btn ts-btn-small ts-edit-account-btn";
        editBtn.title = "Change customer";
        editBtn.innerHTML = '<span class="ts-icon-edit"></span>';
        // Disable edit button if _stageText is not 'Pending'
        if (_stageText !== "Pending") {
          editBtn.disabled = true;
        }
        editBtn.addEventListener("click", () => {
          if (_stageText !== "Pending") return;
          (row as any)._prevAccountId = row.accountId;
          (row as any)._prevAccountName = row.accountName;
          row.accountId = null;
          row.accountName = null;
          this.renderRows();
        });
        btnWrap.appendChild(editBtn);

        // Delete button
        const deleteBtn = document.createElement("button");
        deleteBtn.className = "ts-btn ts-btn-small ts-delete-account-btn";
        deleteBtn.title = "Delete customer row";
        deleteBtn.innerHTML = '<span class="ts-icon-delete"></span>';
        // Disable delete button if _stageText is not 'Pending'
        if (_stageText !== "Pending") {
          deleteBtn.disabled = true;
        }
        deleteBtn.addEventListener("click", async () => {
          if (_stageText !== "Pending") return;
          // Only show confirmation modal if total hours > 0
          const total = sumHours(row.record);
          if (total > 0) {
            showDeleteConfirmModal(async (confirmed) => {
              if (!confirmed) return;
              await handleDeleteRow();
            });
          } else {
            await handleDeleteRow();
          }
        });
        btnWrap.appendChild(deleteBtn);

        // Delete row handler
        const handleDeleteRow = async () => {
          if (row.record.timeSheetId) {
            try {
              showLoader("Deleting row...");
              await this.provider.deleteTimeKeepEntry(row.record.timeSheetId);
            } catch (err) {
              showErrorScreen(err?.message || "Failed to delete timesheet entry.");
              hideLoader();
              return;
            }
          }
          const idx = this.rows.indexOf(row);
          if (idx !== -1) this.rows.splice(idx, 1);
          this.renderRows();
          this.updateStatus("Row deleted successfully.");
          hideLoader();
        };
/**
 * Show a modal confirmation window for deleting a row.
 * Calls callback(true) if confirmed, callback(false) if cancelled.
 */
function showDeleteConfirmModal(callback: (confirmed: boolean) => void) {
  let modal = document.getElementById("ts-delete-confirm-modal") as HTMLDivElement | null;
  if (!modal) {
    modal = document.createElement("div");
    modal.id = "ts-delete-confirm-modal";
    modal.innerHTML = `
      <div class="ts-error-content">
        <h2>Delete Row?</h2>
        <div id="ts-delete-confirm-message">
          Are you sure you want to delete this row?<br><br>  
          There are hours logged for this customer.<br><br>
          <strong>This action is <span style='color:#ffd6d6'>irreversible</span>.</strong>
        </div>
        <div style="margin-top: 24px; display: flex; gap: 16px; justify-content: center;">
          <button id="ts-delete-confirm-yes" class="ts-btn ts-btn-danger">Delete</button>
          <button id="ts-delete-confirm-no" class="ts-btn">Cancel</button>
        </div>
      </div>
    `;
    // Append to .ts-app if possible
    const app = document.querySelector('.ts-app');
    if (app) {
      app.appendChild(modal);
    } else {
      document.body.appendChild(modal);
    }
  }
  modal.classList.add("ts-delete-confirm-modal-visible");

  const yesBtn = document.getElementById("ts-delete-confirm-yes");
  const noBtn = document.getElementById("ts-delete-confirm-no");
  const cleanup = () => { modal!.classList.remove("ts-delete-confirm-modal-visible"); };
  yesBtn?.addEventListener("click", () => { cleanup(); callback(true); }, { once: true });
  noBtn?.addEventListener("click", () => { cleanup(); callback(false); }, { once: true });
}

        labelWrap.appendChild(btnWrap);
        custCell.appendChild(labelWrap);
      }

      rowEl.appendChild(custCell);

      // Determine if row should be editable
      const isStageFinalisedOrSubmitted =
        row.record.timeSheetStage === TimeSheetStageMap.Finalised ||
        row.record.timeSheetStage === TimeSheetStageMap.Submitted;

      // Day inputs
      for (let d = 1; d <= this.daysInMonth; d++) {
        const cell = document.createElement("div");
        cell.className = "ts-cell";
        const input = document.createElement("input");
        input.type = "number";
        input.step = "0.25";
        input.min = "0";
        input.max = "24";
        input.placeholder = "0";
        input.className = "ts-hour-input";

        const key = dayKey(d);
        input.value = String(row.record[key] ?? 0);

        // Only allow editing for today and past dates, and not if stage is Finalised or Submitted
        const today = new Date();
        const cellDate = new Date(Date.UTC(this.year, this.month - 1, d));
        const isFuture = cellDate > new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
        input.disabled = !row.accountId || isFuture || isStageFinalisedOrSubmitted;

        const commitValue = () => {
          // Support both comma and period as decimal separator
          const normalized = input.value.replace(',', '.');
          const val = Number(normalized);
          const prev = row.record[key];
          const newVal = Number.isFinite(val) ? Math.max(0, Math.min(24, val)) : 0;
          if (prev === newVal) return; // Only save if changed
          // Queue the update and batch save
          this.scheduleSave(row, () => {
            row.record[key] = newVal;
            // reflect corrected value
            input.value = String(row.record[key]);
            // update total cell
            totalCell.textContent = sumHours(row.record).toFixed(2);
          });
        };
        input.addEventListener("blur", commitValue);
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            commitValue();
            input.blur();
          }
        });

        cell.appendChild(input);
        rowEl.appendChild(cell);
      }

      // Total
      const totalCell = document.createElement("div");
      totalCell.className = "ts-cell ts-cell-total";
      totalCell.textContent = sumHours(row.record).toFixed(2);
      rowEl.appendChild(totalCell);

      this.gridBody.appendChild(rowEl);
    });
  }

  private buildaccountSelect(row: RowModel) {
    const selectWrap = document.createElement("div");
    selectWrap.className = "ts-account-select-wrap";

    const select = document.createElement("select");
    select.className = "ts-account-select";

    const prompt = document.createElement("option");
    prompt.value = "";
    prompt.textContent = "Select account";
    select.appendChild(prompt);

    const usedaccountIds = new Set(this.rows.filter(r => r.accountId).map(r => r.accountId!));
    // Always include the previous accountId (if editing), so user can see their current selection
    const prevAccountId = (row as any)._prevAccountId;
    const availableAccounts = this.Accounts.filter(
      c => !usedaccountIds.has(c.accountId) || c.accountId === prevAccountId
    );
    availableAccounts.forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.accountId;
      opt.textContent = c.accountName;
      select.appendChild(opt);
    });

    // If editing, set the previous accountId as the selected option
    if (prevAccountId) {
      select.value = prevAccountId;
    }

    // OK button
    const okBtn = document.createElement("button");
    okBtn.className = "ts-btn ts-btn-small ts-btn-customer";
    okBtn.title = "Confirm customer";
    okBtn.textContent = "";
    const checkIcon = document.createElement("span");
    checkIcon.className = "ts-icon-check";
    okBtn.appendChild(checkIcon);
    okBtn.addEventListener("click", () => {
      const chosen = select.value;
      if (!chosen) return;
      const prevAccountId = (row as any)._prevAccountId;
      if (prevAccountId === chosen) {
        // No change: restore previous and just close dropdown
        row.accountId = prevAccountId;
        row.accountName = (row as any)._prevAccountName;
        delete (row as any)._prevAccountId;
        delete (row as any)._prevAccountName;
        this.renderRows();
        return;
      }
      const account = this.Accounts.find(c => c.accountId === chosen);
      // Batch all updates and save
      this.scheduleSave(row, () => {
        row.accountId = account?.accountId ?? null;
        row.accountName = account?.accountName ?? null;
        // Hydrate record meta for this account & month
        const ym = `${this.year}-${String(this.month).padStart(2, "0")}`;
        row.record.yearMonth = ym;
        row.record.timeSheetStage = TimeSheetStageMap.Pending;
        row.record.timeSheetEntryId = `${ym.replace("-", "")}_${account?.catNumber}_${_currentUser.fullName[0]}.${_currentUser.fullName.trim().split(" ").slice(-1)[0]}`; // new entry      
        row.record.timeSheetId = row.record.timeSheetId ?? null;
        row.record.timeSheetStartDate = isoStartOfMonth(this.year, this.month);
        row.record.timeSheetEndDate = isoEndOfMonth(this.year, this.month);
        row.record.accountId = row.accountId ?? "";
        row.record["userId"] = _currentUser?.userId ?? "";
        // Remove previous accountId tracking after selection
        delete (row as any)._prevAccountId;
        delete (row as any)._prevAccountName;
        this.renderRows(); // switch to label + enable inputs
      });
    });

    // Cancel button
    const cancelBtn = document.createElement("button");
    cancelBtn.className = "ts-btn ts-btn-small ts-btn-customer";
    cancelBtn.title = "Cancel customer selection";
    cancelBtn.textContent = "";
    const cancelIcon = document.createElement("span");
    cancelIcon.className = "ts-icon-cancel";
    cancelBtn.appendChild(cancelIcon);
    cancelBtn.addEventListener("click", () => {
      // If this is a new, unsaved row (no previous accountId and no accountId set), remove it
      if ((row as any)._prevAccountId === undefined && !row.accountId) {
        const idx = this.rows.indexOf(row);
        if (idx !== -1) {
          this.rows.splice(idx, 1);
        }
        this.renderRows();
        return;
      }
      // If editing, revert to previous accountId/accountName
      if ((row as any)._prevAccountId !== undefined) {
        row.accountId = (row as any)._prevAccountId;
        row.accountName = (row as any)._prevAccountName;
        delete (row as any)._prevAccountId;
        delete (row as any)._prevAccountName;
      }
      // Close the dropdown (re-render rows)
      this.renderRows();
    });

    selectWrap.append(select, okBtn, cancelBtn);
    return selectWrap;
  }

  private addRow() {
    const ym = `${this.year}-${String(this.month).padStart(2, "0")}`;
    const empty: TimesheetRecord = {
      timeSheetStartDate: isoStartOfMonth(this.year, this.month),
      timeSheetEndDate: isoEndOfMonth(this.year, this.month),
      timeSheetId: null,
      timeSheetEntryId: null,
      totalHours: 0,
      yearMonth: ym,
      accountId: null,
      userId: _currentUser?.userId ?? "",
      timeSheetStage: 596610000, // Pending
      // Initialize day fields to 0
      ...(() => {
        const o: Partial<TimesheetRecord> = {};
        for (let d = 1; d <= 31; d++) (o as any)[dayKey(d)] = 0.0;
        return o;
      })(),
    };

    const model: RowModel = {
      accountId: null,
      accountName: null,
      record: empty,
    };

    this.rows.push(model);
    this.renderRows();
  }

  /**
   * Queue an update for a row and debounce the save.
   * If updateFn is provided, it will be executed before saving.
   */
  private scheduleSave(row: RowModel, updateFn?: () => void) {
    // Disable auto saving if stage is Submitted or Finalised
    if (
      row.record.timeSheetStage === TimeSheetStageMap.Submitted ||
      row.record.timeSheetStage === TimeSheetStageMap.Finalised
    ) {
      return;
    }

    // Queue the update function if provided
    if (updateFn) {
      if (!this.saveQueues.has(row)) this.saveQueues.set(row, []);
      this.saveQueues.get(row)!.push(updateFn);
    }

    // Clear any existing timer
    const existing = this.saveTimers.get(row);
    if (existing) window.clearTimeout(existing);

    // Mark as saving
    row.saving = true;
    this.updateStatus("Auto saving...");

    // Debounce and batch updates
    const handle = window.setTimeout(() => {
      // Flush all queued updates for this row
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

  private async saveRow(row: RowModel): Promise<void> {
    try {
      const saved = await this.provider.saveTimesheet({
        ...row.record,
        totalHours: sumHours(row.record),
      });

      row.record = saved;
      row.lastSaved = new Date();
      row.saving = false;
      this.updateStatus("Last modified: " + row.lastSaved.toLocaleString());
      // Re-render total to reflect canonical saved value if needed
      this.renderRows();
    } catch (err) {
      console.error(err);
      row.saving = false;
      this.updateStatus("Error while saving. Retrying on next change.");
      showErrorScreen(err?.message || "Error while saving timesheet.");
    }
  }

  private updateStatus(text: string) {
    this.statusText.textContent = text;
  }

  private UpdateUserDisplay(uiProfile?: { name?: string; email?: string; photoUrl?: string | null }): void {
    // Clear previous content
    if (_userPill) _userPill.innerHTML = "";
    let avatar: HTMLElement;
    let name: string = _currentUser?.fullName ?? uiProfile?.name ?? "User";
    if (uiProfile) {
      // Show photo and full name
      const wrapper = document.createElement("div");
      wrapper.className = ".ts-user-avatar-wrapper";
      const img = document.createElement("img");
      img.src = (uiProfile.photoUrl && uiProfile.photoUrl.trim() !== "") ? uiProfile.photoUrl : window.location.origin + "/default-user.png";
      img.alt = name;
      img.className = "ts-user-avatar-pic";
      img.title = name;
      const nameSpan = document.createElement("span");
      nameSpan.className = "ts-user-fullname";
      nameSpan.textContent = name;
      wrapper.appendChild(img);
      wrapper.appendChild(nameSpan);
      avatar = wrapper;
    } else {
      // Show initials and full name
      const wrapper = document.createElement("div");
      wrapper.className = ".ts-user-avatar-wrapper";
      const initials = name
        ? name.split(" ").map(n => n[0]).join("").slice(0, 2).toUpperCase()
        : "?";
      const span = document.createElement("span");
      span.className = "ts-user-avatar-initials";
      span.textContent = initials;
      const nameSpan = document.createElement("span");
      nameSpan.className = "ts-user-fullname";
      nameSpan.textContent = name;
      wrapper.appendChild(span);
      wrapper.appendChild(nameSpan);
      avatar = wrapper;
    }
    if (_userPill) _userPill.appendChild(avatar);
  }
  
  /**
   * Retrieve TimesheetRecord rows for the previous month relative to the current selection.
   * @returns Promise<TimesheetRecord[]> for the previous month
   */
  async getPreviousMonthTimesheetRecords(): Promise<TimesheetRecord[]> {
    let prevYear = this.year;
    let prevMonth = this.month - 1;
    if (prevMonth < 1) {
      prevMonth = 12;
      prevYear -= 1;
    }
    return this.provider.listTimesheets(prevYear, prevMonth);
  }
}

/**
 * Utility: find the .ts-app container (positioning context for the loader)
 */
function getTsApp(): HTMLElement | null {
  return document.querySelector<HTMLElement>(".ts-app");
}

/**
 * Ensure the loader exists inside .ts-app with the right structure.
 * Creates it if missing.
 */
function ensureLoader(): HTMLElement | null {
  const app = getTsApp();
  if (!app) {
    console.warn("[ts-loader] .ts-app container not found.");
    return null;
  }

  let loader = app.querySelector<HTMLElement>("#ts-loader");
  if (!loader) {
    loader = document.createElement("div");
    loader.id = "ts-loader";
    loader.setAttribute("role", "status");
    loader.setAttribute("aria-live", "polite");
    loader.setAttribute("aria-label", "Loading");

    // Structure matches the CSS provided earlier
    loader.innerHTML = `
      <div class="ts-loader-content">
        <div class="ts-spinner" aria-hidden="true"></div>
        <div class="ts-loader-text">Loading…</div>
      </div>
    `;

    app.appendChild(loader);
  }

  return loader;
}

/** Loader screen helpers */
function showLoader(statusText?: string): void {
  const app = getTsApp();
  const loader = ensureLoader();
  if (!app || !loader) return;

  // Display the overlay and mark the app as loading
  loader.style.display = "flex";
  app.classList.add("is-loading");

  // Set loader status text if provided
  if (typeof statusText === "string") {
    const textDiv = loader.querySelector<HTMLElement>(".ts-loader-text");
    if (textDiv) textDiv.textContent = statusText;
  }

  // Accessibility: reveal status
  loader.setAttribute("aria-hidden", "false");
}

/**
 * Hide the loader overlay and clear loading state.
 */
function hideLoader(): void {
  const app = getTsApp();
  const loader = app?.querySelector<HTMLElement>("#ts-loader");
  if (!app || !loader) return;

  loader.style.display = "none";
  app.classList.remove("is-loading");

  // Accessibility: hide status
  loader.setAttribute("aria-hidden", "true");
}

/** Error screen helpers */
function showErrorScreen(message: string): void {
  let errorOverlay = document.getElementById("ts-error-overlay") as HTMLDivElement | null;
  if (!errorOverlay) {
    // Check for popup blocked error
    errorOverlay = document.createElement("div");
    errorOverlay.id = "ts-error-overlay";
    errorOverlay.innerHTML = `
      <div class="ts-error-content">
        <h2>Something went wrong</h2>
        <div id="ts-error-message"></div>
        <button id="ts-error-reload" type="button">Reload</button>
      </div>
    `;
    // Append to .ts-app instead of body
    const app = document.querySelector('.ts-app');
    if (app) {
      app.appendChild(errorOverlay);
    } else {
      document.body.appendChild(errorOverlay);
    }
    document.getElementById("ts-error-reload")?.addEventListener("click", () => window.location.reload());
  }
  const msgDiv = document.getElementById("ts-error-message");
  if (msgDiv) {
    if (typeof message === "string" && /popup|blocked/i.test(message)) {
      msgDiv.textContent =
        "Please allow popup windows for this site in your browser settings. " +
        "You can check the top right of your browser address bar for any blocked popup icons." +
        "\n\n" +
        "After allowing popups, please refresh the page.";
      return;
    }
    else
      msgDiv.textContent = message;
  }
  errorOverlay.style.display = "flex";
}

function hideErrorScreen(): void {
  const errorOverlay = document.getElementById("ts-error-overlay");
  if (errorOverlay) errorOverlay.style.display = "none";
}

// Example usage: wrap async entrypoints and API calls
async function initTimesheet(container: HTMLDivElement): Promise<TimesheetApp> {
  try {
    console.log(`Time Keep ${APP_VERSION}`);
    await loadMsalFromCdn();
    // Check for Dev
    const currentUrl: string = window.location.href;
    if (currentUrl == "http://localhost:8080/"
      || currentUrl == "http://127.0.0.1:8080/") {
      _isDev = true;
    }
    const provider = new ApiProvider();
    const app = new TimesheetApp(container, provider);
    void app.mount();
    return app;
  } catch (err: any) {
    showErrorScreen(err?.message || "Unknown error");
    throw err;
  }
}

// Auto-init on DOM ready
document.addEventListener("DOMContentLoaded", async (): Promise<void> => {
  const container = document.getElementById('app') as HTMLDivElement;
    showLoader("Loading...");
    await initTimesheet(container);
});