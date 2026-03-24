import {
  SYSTEM_ROLE_LABELS,
  type AppPermissionKey,
  type ManageableSystemRoleKey,
} from "./shared/authAccess";
import type {
  AdminPanelData,
  AdminUserSummary,
  AuthActionResult,
  AuthSessionState,
  IssuedOneTimeToken,
  PermissionGroupSummary,
} from "./shared/authTypes";

type InitOptions = {
  setBusy: (isBusy: boolean) => void;
  onSessionChanged?: (session: AuthSessionState) => void;
};

function byId<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function escapeHtml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatDateTime(value: string | null | undefined): string {
  if (!value) return "brak";
  try {
    return new Date(value).toLocaleString("pl-PL");
  } catch {
    return String(value);
  }
}

function authErrorMessage<T>(result: AuthActionResult<T>): string {
  return "error" in result ? String(result.error ?? "unknown") : "";
}

function authFieldErrors<T>(
  result: AuthActionResult<T>,
): Record<string, string> | undefined {
  return "fieldErrors" in result ? result.fieldErrors : undefined;
}

const els = {
  authSummaryText: byId<HTMLElement>("auth-summary-text"),
  btnAuthLogout: byId<HTMLButtonElement>("btn-auth-logout"),
  btnAdminSessionRefresh: byId<HTMLButtonElement>("btn-admin-session-refresh"),
  btnAdminCopyToken: byId<HTMLButtonElement>("btn-admin-copy-token"),
  btnAdminLogoutInline: byId<HTMLButtonElement>("btn-admin-logout-inline"),

  adminAuthStatus: byId<HTMLElement>("admin-auth-status"),
  adminNoAccessCard: byId<HTMLElement>("admin-no-access-card"),
  adminPanelWrap: byId<HTMLElement>("admin-panel-wrap"),
  adminIssuedTokenCard: byId<HTMLElement>("admin-issued-token-card"),
  adminIssuedTokenMeta: byId<HTMLElement>("admin-issued-token-meta"),
  adminIssuedTokenValue: byId<HTMLElement>("admin-issued-token-value"),

  adminBootstrapCard: byId<HTMLElement>("admin-bootstrap-card"),
  adminBootstrapForm: byId<HTMLFormElement>("admin-bootstrap-form"),
  adminBootstrapLogin: byId<HTMLInputElement>("admin-bootstrap-login"),
  adminBootstrapFullName: byId<HTMLInputElement>("admin-bootstrap-full-name"),
  adminBootstrapLoginError: byId<HTMLElement>("admin-bootstrap-login-error"),
  adminBootstrapFullNameError: byId<HTMLElement>(
    "admin-bootstrap-full-name-error",
  ),

  adminLoginCard: byId<HTMLElement>("admin-login-card"),
  adminPasswordLoginForm: byId<HTMLFormElement>("admin-password-login-form"),
  adminPasswordLoginLogin: byId<HTMLInputElement>("admin-password-login-login"),
  adminPasswordLoginPassword: byId<HTMLInputElement>(
    "admin-password-login-password",
  ),
  adminPasswordLoginLoginError: byId<HTMLElement>(
    "admin-password-login-login-error",
  ),
  adminPasswordLoginPasswordError: byId<HTMLElement>(
    "admin-password-login-password-error",
  ),

  adminTokenLoginForm: byId<HTMLFormElement>("admin-token-login-form"),
  adminTokenLoginLogin: byId<HTMLInputElement>("admin-token-login-login"),
  adminTokenLoginToken: byId<HTMLInputElement>("admin-token-login-token"),
  adminTokenLoginLoginError: byId<HTMLElement>("admin-token-login-login-error"),
  adminTokenLoginTokenError: byId<HTMLElement>("admin-token-login-token-error"),

  adminPasswordSetupCard: byId<HTMLElement>("admin-password-setup-card"),
  adminPasswordSetupForm: byId<HTMLFormElement>("admin-password-setup-form"),
  adminPasswordSetupPassword: byId<HTMLInputElement>(
    "admin-password-setup-password",
  ),
  adminPasswordSetupPasswordConfirm: byId<HTMLInputElement>(
    "admin-password-setup-password-confirm",
  ),
  adminPasswordSetupPasswordError: byId<HTMLElement>(
    "admin-password-setup-password-error",
  ),
  adminPasswordSetupPasswordConfirmError: byId<HTMLElement>(
    "admin-password-setup-password-confirm-error",
  ),

  adminSessionCard: byId<HTMLElement>("admin-session-card"),
  adminSessionMeta: byId<HTMLElement>("admin-session-meta"),
  adminSessionGroups: byId<HTMLElement>("admin-session-groups"),
  adminSessionPermissions: byId<HTMLElement>("admin-session-permissions"),

  btnAdminUserResetForm: byId<HTMLButtonElement>("btn-admin-user-reset-form"),
  adminUserForm: byId<HTMLFormElement>("admin-user-form"),
  adminUserId: byId<HTMLInputElement>("admin-user-id"),
  adminUserLogin: byId<HTMLInputElement>("admin-user-login"),
  adminUserFullName: byId<HTMLInputElement>("admin-user-full-name"),
  adminUserSystemRole: byId<HTMLSelectElement>("admin-user-system-role"),
  adminUserIsActive: byId<HTMLInputElement>("admin-user-is-active"),
  adminUserGroupList: byId<HTMLElement>("admin-user-group-list"),
  adminUserLoginError: byId<HTMLElement>("admin-user-login-error"),
  adminUserFullNameError: byId<HTMLElement>("admin-user-full-name-error"),
  adminUserSystemRoleError: byId<HTMLElement>("admin-user-system-role-error"),
  adminUserGroupError: byId<HTMLElement>("admin-user-group-error"),
  adminUserStatus: byId<HTMLElement>("admin-user-status"),
  adminUserList: byId<HTMLElement>("admin-user-list"),

  btnAdminGroupResetForm: byId<HTMLButtonElement>("btn-admin-group-reset-form"),
  adminGroupForm: byId<HTMLFormElement>("admin-group-form"),
  adminGroupId: byId<HTMLInputElement>("admin-group-id"),
  adminGroupName: byId<HTMLInputElement>("admin-group-name"),
  adminGroupDescription: byId<HTMLInputElement>("admin-group-description"),
  adminGroupPermissionList: byId<HTMLElement>("admin-group-permission-list"),
  adminGroupNameError: byId<HTMLElement>("admin-group-name-error"),
  adminGroupDescriptionError: byId<HTMLElement>("admin-group-description-error"),
  adminGroupPermissionsError: byId<HTMLElement>(
    "admin-group-permissions-error",
  ),
  adminGroupStatus: byId<HTMLElement>("admin-group-status"),
  adminGroupList: byId<HTMLElement>("admin-group-list"),
};

let initOptions: InitOptions | null = null;
let initialized = false;
let sessionState: AuthSessionState | null = null;
let adminPanelData: AdminPanelData | null = null;
let lastIssuedToken: IssuedOneTimeToken | null = null;

function setStatus(el: HTMLElement, text: string) {
  el.textContent = text;
}

function setFieldErrors(
  map: Record<string, HTMLElement>,
  fieldErrors?: Record<string, string>,
) {
  for (const el of Object.values(map)) {
    el.textContent = "";
  }
  if (!fieldErrors) return;

  for (const [field, message] of Object.entries(fieldErrors)) {
    const el = map[field];
    if (el) el.textContent = message;
  }
}

function getSelectedCheckboxValues(container: HTMLElement): string[] {
  return Array.from(
    container.querySelectorAll<HTMLInputElement>('input[type="checkbox"]:checked'),
  )
    .map((input) => String(input.value ?? "").trim())
    .filter(Boolean);
}

function renderPills(target: HTMLElement, values: string[], emptyText: string) {
  if (!values.length) {
    target.innerHTML = `<span class="pill">${escapeHtml(emptyText)}</span>`;
    return;
  }

  target.innerHTML = values
    .map((value) => `<span class="pill">${escapeHtml(value)}</span>`)
    .join("");
}

function canAccessAdminPanel(): boolean {
  return Boolean(
    sessionState?.isAuthenticated &&
      sessionState.user?.canAccessAdminPanel &&
      !sessionState.requiresPasswordSetup,
  );
}

function canManageUser(user: AdminUserSummary): boolean {
  const actorRole = sessionState?.user?.systemRole;
  if (!canAccessAdminPanel() || !actorRole) return false;
  if (user.systemRole === "SUPER_ADMIN") return false;
  if (actorRole === "ADMIN" && user.systemRole !== "USER") return false;
  return true;
}

function renderAuthSummary() {
  if (!sessionState) {
    els.authSummaryText.textContent = "Sprawdzanie sesji...";
    els.btnAuthLogout.classList.add("hidden");
    return;
  }

  if (sessionState.bootstrapRequired) {
    els.authSummaryText.textContent = "Bootstrap Super Admin wymagany";
    els.btnAuthLogout.classList.add("hidden");
    return;
  }

  if (!sessionState.isAuthenticated || !sessionState.user) {
    els.authSummaryText.textContent = "Nie zalogowano";
    els.btnAuthLogout.classList.add("hidden");
    return;
  }

  const roleLabel = SYSTEM_ROLE_LABELS[sessionState.user.systemRole];
  const suffix = sessionState.requiresPasswordSetup
    ? " - ustaw nowe haslo"
    : "";
  els.authSummaryText.textContent = `${sessionState.user.fullName} (${sessionState.user.login}) - ${roleLabel}${suffix}`;
  els.btnAuthLogout.classList.remove("hidden");
}

function renderIssuedToken() {
  const token = lastIssuedToken;
  els.adminIssuedTokenCard.classList.toggle("hidden", !token);
  if (!token) return;

  els.adminIssuedTokenMeta.textContent = `${token.fullName} (${token.login}) - wygasa ${formatDateTime(token.expiresAt)}`;
  els.adminIssuedTokenValue.textContent = token.token;
}

function renderSessionBlock() {
  const session = sessionState;
  els.adminBootstrapCard.classList.toggle("hidden", !session?.bootstrapRequired);
  els.adminLoginCard.classList.toggle(
    "hidden",
    session?.bootstrapRequired || session?.isAuthenticated,
  );
  els.adminPasswordSetupCard.classList.toggle(
    "hidden",
    !(session?.isAuthenticated && session?.requiresPasswordSetup),
  );
  els.adminSessionCard.classList.toggle(
    "hidden",
    !(session?.isAuthenticated && session?.user),
  );
  els.adminNoAccessCard.classList.toggle(
    "hidden",
    !(
      session?.isAuthenticated &&
        !session.requiresPasswordSetup &&
        session.user &&
        !session.user.canAccessAdminPanel
    ),
  );
  els.adminPanelWrap.classList.toggle("hidden", !canAccessAdminPanel());

  if (!session?.isAuthenticated || !session.user) {
    els.adminSessionMeta.textContent = "-";
    els.adminSessionGroups.innerHTML = "";
    els.adminSessionPermissions.innerHTML = "";
    return;
  }

  els.adminSessionMeta.textContent = `${session.user.fullName} (${session.user.login}) - ${SYSTEM_ROLE_LABELS[session.user.systemRole]}`;
  renderPills(
    els.adminSessionGroups,
    session.user.groupNames.map((groupName) => `Grupa: ${groupName}`),
    "Brak grup",
  );
  renderPills(
    els.adminSessionPermissions,
    session.user.effectivePermissions,
    "Brak dostepow",
  );
}

function renderRoleOptions() {
  const roles = adminPanelData?.manageableSystemRoles ?? [];
  els.adminUserSystemRole.innerHTML = roles
    .map(
      (role) =>
        `<option value="${escapeHtml(role)}">${escapeHtml(
          SYSTEM_ROLE_LABELS[role],
        )}</option>`,
    )
    .join("");
}

function renderGroupCheckboxes(selectedGroupIds: string[]) {
  const groups = adminPanelData?.groups ?? [];
  if (!groups.length) {
    els.adminUserGroupList.innerHTML =
      '<div class="admin-empty">Najpierw utworz grupe dostepow.</div>';
    return;
  }

  const selected = new Set(selectedGroupIds);
  els.adminUserGroupList.innerHTML = groups
    .map(
      (group) => `
        <label class="admin-checkbox-item">
          <input type="checkbox" value="${escapeHtml(group.id)}" ${
            selected.has(group.id) ? "checked" : ""
          } />
          <div class="admin-checkbox-item-main">
            <div class="admin-checkbox-item-title">${escapeHtml(group.name)}</div>
            <div class="admin-checkbox-item-sub">${escapeHtml(
              group.description || group.permissionKeys.join(", ") || "Bez opisu",
            )}</div>
          </div>
        </label>
      `,
    )
    .join("");
}

function renderPermissionCheckboxes(selectedKeys: AppPermissionKey[]) {
  const selected = new Set(selectedKeys);
  const permissions = adminPanelData?.permissionCatalog ?? [];
  els.adminGroupPermissionList.innerHTML = permissions
    .map(
      (permission) => `
        <label class="admin-checkbox-item">
          <input type="checkbox" value="${escapeHtml(permission.key)}" ${
            selected.has(permission.key) ? "checked" : ""
          } />
          <div class="admin-checkbox-item-main">
            <div class="admin-checkbox-item-title">${escapeHtml(
              permission.label,
            )}</div>
            <div class="admin-checkbox-item-sub">${escapeHtml(
              `${permission.category} - ${permission.description}`,
            )}</div>
          </div>
        </label>
      `,
    )
    .join("");
}

function resetUserForm() {
  els.adminUserId.value = "";
  els.adminUserLogin.value = "";
  els.adminUserFullName.value = "";
  els.adminUserIsActive.checked = true;
  if (els.adminUserSystemRole.options.length > 0) {
    els.adminUserSystemRole.selectedIndex = 0;
  }
  renderGroupCheckboxes([]);
  setFieldErrors(
    {
      login: els.adminUserLoginError,
      fullName: els.adminUserFullNameError,
      systemRole: els.adminUserSystemRoleError,
      groupIds: els.adminUserGroupError,
    },
    {},
  );
  setStatus(els.adminUserStatus, "");
}

function resetGroupForm() {
  els.adminGroupId.value = "";
  els.adminGroupName.value = "";
  els.adminGroupDescription.value = "";
  renderPermissionCheckboxes([]);
  setFieldErrors(
    {
      name: els.adminGroupNameError,
      description: els.adminGroupDescriptionError,
      permissionKeys: els.adminGroupPermissionsError,
    },
    {},
  );
  setStatus(els.adminGroupStatus, "");
}

function renderUserList() {
  const users = adminPanelData?.users ?? [];
  if (!users.length) {
    els.adminUserList.innerHTML =
      '<div class="admin-empty">Brak uzytkownikow do wyswietlenia.</div>';
    return;
  }

  els.adminUserList.innerHTML = users
    .map((user) => {
      const canManage = canManageUser(user);
      const tokenText = user.activeToken.exists
        ? `Aktywny token do ${formatDateTime(user.activeToken.expiresAt)}`
        : user.mustSetPassword
          ? "Brak aktywnego tokena"
          : "Haslo ustawione";

      const meta = [
        `Rola: ${SYSTEM_ROLE_LABELS[user.systemRole]}`,
        user.isActive ? "Aktywne" : "Nieaktywne",
        tokenText,
      ];

      const groups = user.groupNames.length
        ? user.groupNames.map((groupName) => `Grupa: ${groupName}`)
        : ["Brak grup"];
      const permissions = user.effectivePermissions.length
        ? user.effectivePermissions
        : ["Brak dostepow"];

      return `
        <div class="admin-list-card">
          <div class="admin-list-card-top">
            <div>
              <div class="admin-list-title">${escapeHtml(user.fullName)}</div>
              <div class="admin-list-sub">${escapeHtml(
                `${user.login} - ${meta.join(" - ")}`,
              )}</div>
            </div>
            <div class="admin-list-actions">
              ${
                canManage
                  ? `<button class="btn btn-outline-light btn-sm" type="button" data-admin-user-edit="${escapeHtml(
                      user.id,
                    )}">Edytuj</button>`
                  : ""
              }
              ${
                canManage
                  ? `<button class="btn btn-outline-light btn-sm" type="button" data-admin-user-token="${escapeHtml(
                      user.id,
                    )}">Nowy token</button>`
                  : ""
              }
            </div>
          </div>
          <div class="admin-list-meta">
            ${groups
              .map((value) => `<span class="pill">${escapeHtml(value)}</span>`)
              .join("")}
            ${permissions
              .map((value) => `<span class="pill">${escapeHtml(value)}</span>`)
              .join("")}
          </div>
        </div>
      `;
    })
    .join("");
}

function renderGroupList() {
  const groups = adminPanelData?.groups ?? [];
  if (!groups.length) {
    els.adminGroupList.innerHTML =
      '<div class="admin-empty">Brak grup dostepow. Utworz pierwsza grupe po prawej.</div>';
    return;
  }

  els.adminGroupList.innerHTML = groups
    .map(
      (group) => `
        <div class="admin-list-card">
          <div class="admin-list-card-top">
            <div>
              <div class="admin-list-title">${escapeHtml(group.name)}</div>
              <div class="admin-list-sub">${escapeHtml(
                `${group.slug} - ${group.memberCount} uzytk. - ${group.description || "bez opisu"}`,
              )}</div>
            </div>
            <div class="admin-list-actions">
              <button class="btn btn-outline-light btn-sm" type="button" data-admin-group-edit="${escapeHtml(
                group.id,
              )}">Edytuj</button>
            </div>
          </div>
          <div class="admin-list-meta">
            ${group.permissionKeys
              .map(
                (permissionKey) =>
                  `<span class="pill">${escapeHtml(permissionKey)}</span>`,
              )
              .join("")}
          </div>
        </div>
      `,
    )
    .join("");
}

function renderAdminData() {
  renderRoleOptions();

  const currentUserGroupIds =
    els.adminUserId.value && adminPanelData
      ? adminPanelData.users.find((user) => user.id === els.adminUserId.value)
          ?.groupIds ?? getSelectedCheckboxValues(els.adminUserGroupList)
      : getSelectedCheckboxValues(els.adminUserGroupList);
  renderGroupCheckboxes(currentUserGroupIds);

  const currentPermissionKeys = getSelectedCheckboxValues(
    els.adminGroupPermissionList,
  ).filter(Boolean) as AppPermissionKey[];
  renderPermissionCheckboxes(currentPermissionKeys);

  renderUserList();
  renderGroupList();
}

async function runWithBusy<T>(action: () => Promise<T>): Promise<T> {
  initOptions?.setBusy(true);
  try {
    return await action();
  } finally {
    initOptions?.setBusy(false);
  }
}

async function refreshAdminData() {
  if (!canAccessAdminPanel()) {
    adminPanelData = null;
    renderAdminData();
    return;
  }

  const result = await runWithBusy(() => window.api.getAdminPanelData());
  if (!result.ok) {
    adminPanelData = null;
    setStatus(els.adminAuthStatus, authErrorMessage(result));
    renderAdminData();
    return;
  }

  adminPanelData = result.data;
  renderAdminData();
}

async function refreshSession(options?: { refreshAdmin?: boolean }) {
  const session = await runWithBusy(() => window.api.getAuthSession());
  sessionState = session;
  renderAuthSummary();
  renderSessionBlock();
  renderIssuedToken();
  initOptions?.onSessionChanged?.(session);

  if (options?.refreshAdmin) {
    await refreshAdminData();
  }
}

function fillUserForm(user: AdminUserSummary) {
  els.adminUserId.value = user.id;
  els.adminUserLogin.value = user.login;
  els.adminUserFullName.value = user.fullName;
  els.adminUserIsActive.checked = user.isActive;
  els.adminUserSystemRole.value =
    user.systemRole === "SUPER_ADMIN" ? "" : user.systemRole;
  renderGroupCheckboxes(user.groupIds);
  setStatus(els.adminUserStatus, `Edycja: ${user.fullName}`);
}

function fillGroupForm(group: PermissionGroupSummary) {
  els.adminGroupId.value = group.id;
  els.adminGroupName.value = group.name;
  els.adminGroupDescription.value = group.description ?? "";
  renderPermissionCheckboxes(group.permissionKeys);
  setStatus(els.adminGroupStatus, `Edycja grupy: ${group.name}`);
}

async function handleBootstrapSubmit(event: Event) {
  event.preventDefault();
  setFieldErrors(
    {
      login: els.adminBootstrapLoginError,
      fullName: els.adminBootstrapFullNameError,
    },
    {},
  );

  const result = await runWithBusy(() =>
    window.api.bootstrapSuperAdmin({
      login: els.adminBootstrapLogin.value,
      fullName: els.adminBootstrapFullName.value,
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        login: els.adminBootstrapLoginError,
        fullName: els.adminBootstrapFullNameError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminAuthStatus, authErrorMessage(result));
    return;
  }

  lastIssuedToken = result.data;
  els.adminPasswordLoginLogin.value = result.data.login;
  els.adminTokenLoginLogin.value = result.data.login;
  els.adminTokenLoginToken.value = result.data.token;
  els.adminBootstrapForm.reset();
  setStatus(
    els.adminAuthStatus,
    "Super admin utworzony. Zaloguj sie jednorazowym tokenem i ustaw haslo.",
  );
  await refreshSession({ refreshAdmin: true });
}

async function handlePasswordLoginSubmit(event: Event) {
  event.preventDefault();
  setFieldErrors(
    {
      login: els.adminPasswordLoginLoginError,
      password: els.adminPasswordLoginPasswordError,
    },
    {},
  );

  const result = await runWithBusy(() =>
    window.api.loginWithPassword({
      login: els.adminPasswordLoginLogin.value,
      password: els.adminPasswordLoginPassword.value,
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        login: els.adminPasswordLoginLoginError,
        password: els.adminPasswordLoginPasswordError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminAuthStatus, authErrorMessage(result));
    return;
  }

  els.adminPasswordLoginPassword.value = "";
  setStatus(els.adminAuthStatus, "Zalogowano.");
  await refreshSession({ refreshAdmin: true });
}

async function handleTokenLoginSubmit(event: Event) {
  event.preventDefault();
  setFieldErrors(
    {
      login: els.adminTokenLoginLoginError,
      token: els.adminTokenLoginTokenError,
    },
    {},
  );

  const result = await runWithBusy(() =>
    window.api.loginWithToken({
      login: els.adminTokenLoginLogin.value,
      token: els.adminTokenLoginToken.value,
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        login: els.adminTokenLoginLoginError,
        token: els.adminTokenLoginTokenError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminAuthStatus, authErrorMessage(result));
    return;
  }

  els.adminPasswordSetupPassword.focus();
  setStatus(
    els.adminAuthStatus,
    "Token zaakceptowany. Ustaw nowe haslo, aby zakonczyc logowanie.",
  );
  await refreshSession({ refreshAdmin: true });
}

async function handlePasswordSetupSubmit(event: Event) {
  event.preventDefault();
  setFieldErrors(
    {
      password: els.adminPasswordSetupPasswordError,
      confirm: els.adminPasswordSetupPasswordConfirmError,
    },
    {},
  );

  const password = els.adminPasswordSetupPassword.value;
  const confirm = els.adminPasswordSetupPasswordConfirm.value;
  if (password !== confirm) {
    setFieldErrors(
      {
        password: els.adminPasswordSetupPasswordError,
        confirm: els.adminPasswordSetupPasswordConfirmError,
      },
      { confirm: "Hasla musza byc identyczne." },
    );
    return;
  }

  const result = await runWithBusy(() =>
    window.api.completePasswordSetup({
      password,
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        password: els.adminPasswordSetupPasswordError,
        confirm: els.adminPasswordSetupPasswordConfirmError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminAuthStatus, authErrorMessage(result));
    return;
  }

  els.adminPasswordSetupForm.reset();
  setStatus(els.adminAuthStatus, "Haslo zapisane.");
  await refreshSession({ refreshAdmin: true });
}

async function handleLogout() {
  lastIssuedToken = null;
  adminPanelData = null;
  resetUserForm();
  resetGroupForm();
  await runWithBusy(() => window.api.logout());
  setStatus(els.adminAuthStatus, "Wylogowano.");
  await refreshSession({ refreshAdmin: true });
}

async function handleUserSave(event: Event) {
  event.preventDefault();
  if (!adminPanelData) return;

  const result = await runWithBusy(() =>
    window.api.saveAdminUser({
      userId: els.adminUserId.value || undefined,
      login: els.adminUserLogin.value,
      fullName: els.adminUserFullName.value,
      systemRole: els.adminUserSystemRole
        .value as ManageableSystemRoleKey,
      isActive: els.adminUserIsActive.checked,
      groupIds: getSelectedCheckboxValues(els.adminUserGroupList),
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        login: els.adminUserLoginError,
        fullName: els.adminUserFullNameError,
        systemRole: els.adminUserSystemRoleError,
        groupIds: els.adminUserGroupError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminUserStatus, authErrorMessage(result));
    return;
  }

  lastIssuedToken = result.data.issuedToken ?? lastIssuedToken;
  resetUserForm();
  setStatus(
    els.adminUserStatus,
    result.data.issuedToken
      ? "Uzytkownik zapisany i otrzymal nowy token 24h."
      : "Uzytkownik zapisany.",
  );
  renderIssuedToken();
  await refreshAdminData();
}

async function handleGroupSave(event: Event) {
  event.preventDefault();
  if (!adminPanelData) return;

  const result = await runWithBusy(() =>
    window.api.savePermissionGroup({
      groupId: els.adminGroupId.value || undefined,
      name: els.adminGroupName.value,
      description: els.adminGroupDescription.value,
      permissionKeys: getSelectedCheckboxValues(
        els.adminGroupPermissionList,
      ) as AppPermissionKey[],
    }),
  );

  if (!result.ok) {
    setFieldErrors(
      {
        name: els.adminGroupNameError,
        description: els.adminGroupDescriptionError,
        permissionKeys: els.adminGroupPermissionsError,
      },
      authFieldErrors(result),
    );
    setStatus(els.adminGroupStatus, authErrorMessage(result));
    return;
  }

  resetGroupForm();
  setStatus(els.adminGroupStatus, "Grupa zapisana.");
  await refreshAdminData();
}

async function handleRotateToken(userId: string) {
  const user = adminPanelData?.users.find((item) => item.id === userId);
  if (!user) return;

  const confirmed = window.confirm(
    `Wygenerowac nowy token 24h dla ${user.fullName}? To zresetuje haslo i wymusi ustawienie nowego.`,
  );
  if (!confirmed) return;

  const result = await runWithBusy(() =>
    window.api.rotateAdminUserToken({
      userId,
    }),
  );

  if (!result.ok) {
    setStatus(els.adminUserStatus, authErrorMessage(result));
    return;
  }

  lastIssuedToken = result.data;
  renderIssuedToken();
  setStatus(els.adminUserStatus, `Nowy token wydany dla ${user.fullName}.`);
  await refreshAdminData();
}

async function copyLastIssuedToken() {
  if (!lastIssuedToken?.token) return;
  const result = await window.api.writeClipboardText(lastIssuedToken.token);
  setStatus(
    els.adminAuthStatus,
    result.ok ? "Token skopiowany do schowka." : "Nie udalo sie skopiowac tokena.",
  );
}

function bindEvents() {
  if (initialized) return;
  initialized = true;

  els.btnAdminSessionRefresh.addEventListener("click", () => {
    void refreshSession({ refreshAdmin: true });
  });
  els.btnAdminCopyToken.addEventListener("click", () => {
    void copyLastIssuedToken();
  });
  els.btnAuthLogout.addEventListener("click", () => {
    void handleLogout();
  });
  els.btnAdminLogoutInline.addEventListener("click", () => {
    void handleLogout();
  });

  els.adminBootstrapForm.addEventListener("submit", (event) => {
    void handleBootstrapSubmit(event);
  });
  els.adminPasswordLoginForm.addEventListener("submit", (event) => {
    void handlePasswordLoginSubmit(event);
  });
  els.adminTokenLoginForm.addEventListener("submit", (event) => {
    void handleTokenLoginSubmit(event);
  });
  els.adminPasswordSetupForm.addEventListener("submit", (event) => {
    void handlePasswordSetupSubmit(event);
  });
  els.adminUserForm.addEventListener("submit", (event) => {
    void handleUserSave(event);
  });
  els.adminGroupForm.addEventListener("submit", (event) => {
    void handleGroupSave(event);
  });

  els.btnAdminUserResetForm.addEventListener("click", () => resetUserForm());
  els.btnAdminGroupResetForm.addEventListener("click", () => resetGroupForm());

  els.adminUserList.addEventListener("click", (event) => {
    const target = event.target as HTMLElement | null;
    const editButton = target?.closest("[data-admin-user-edit]") as
      | HTMLButtonElement
      | null;
    if (editButton) {
      const userId = String(editButton.dataset.adminUserEdit ?? "").trim();
      const user = adminPanelData?.users.find((item) => item.id === userId);
      if (user) fillUserForm(user);
      return;
    }

    const tokenButton = target?.closest("[data-admin-user-token]") as
      | HTMLButtonElement
      | null;
    if (!tokenButton) return;
    const userId = String(tokenButton.dataset.adminUserToken ?? "").trim();
    if (userId) void handleRotateToken(userId);
  });

  els.adminGroupList.addEventListener("click", (event) => {
    const target = event.target as HTMLElement | null;
    const button = target?.closest("[data-admin-group-edit]") as
      | HTMLButtonElement
      | null;
    if (!button) return;
    const groupId = String(button.dataset.adminGroupEdit ?? "").trim();
    const group = adminPanelData?.groups.find((item) => item.id === groupId);
    if (group) fillGroupForm(group);
  });
}

export async function initAdminUi(options: InitOptions): Promise<void> {
  initOptions = options;
  bindEvents();
  renderIssuedToken();
  await refreshSession({ refreshAdmin: true });
  resetUserForm();
  resetGroupForm();
}

export function getCurrentAuthSession(): AuthSessionState | null {
  return sessionState;
}

export function hasAppPermission(permission: AppPermissionKey): boolean {
  return Boolean(
    sessionState?.isAuthenticated &&
      sessionState.user &&
      !sessionState.requiresPasswordSetup &&
      sessionState.user.effectivePermissions.includes(permission),
  );
}
