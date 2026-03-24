import type {
  AppPermissionDefinition,
  AppPermissionKey,
  ManageableSystemRoleKey,
  SystemRoleKey,
} from "./authAccess";

export type FieldErrorMap = Record<string, string>;

export type AuthActionResult<T> =
  | { ok: true; data: T; message?: string }
  | { ok: false; error: string; fieldErrors?: FieldErrorMap };

export type IssuedOneTimeToken = {
  userId: string;
  login: string;
  fullName: string;
  token: string;
  tokenPreview: string;
  expiresAt: string;
};

export type AuthSessionUser = {
  id: string;
  login: string;
  fullName: string;
  systemRole: SystemRoleKey;
  isActive: boolean;
  groupIds: string[];
  groupNames: string[];
  effectivePermissions: AppPermissionKey[];
  canAccessAdminPanel: boolean;
};

export type AuthSessionState = {
  bootstrapRequired: boolean;
  isAuthenticated: boolean;
  requiresPasswordSetup: boolean;
  user: AuthSessionUser | null;
};

export type AdminUserSummary = {
  id: string;
  login: string;
  fullName: string;
  systemRole: SystemRoleKey;
  isActive: boolean;
  mustSetPassword: boolean;
  passwordSetAt: string | null;
  createdAt: string;
  updatedAt: string;
  groupIds: string[];
  groupNames: string[];
  effectivePermissions: AppPermissionKey[];
  activeToken: {
    exists: boolean;
    expiresAt: string | null;
    tokenPreview: string | null;
  };
};

export type PermissionGroupSummary = {
  id: string;
  name: string;
  slug: string;
  description: string | null;
  permissionKeys: AppPermissionKey[];
  memberCount: number;
  createdAt: string;
  updatedAt: string;
};

export type AdminPanelData = {
  permissionCatalog: AppPermissionDefinition[];
  manageableSystemRoles: ManageableSystemRoleKey[];
  users: AdminUserSummary[];
  groups: PermissionGroupSummary[];
};

export type BootstrapSuperAdminInput = {
  login: string;
  fullName: string;
};

export type PasswordLoginInput = {
  login: string;
  password: string;
};

export type TokenLoginInput = {
  login: string;
  token: string;
};

export type CompletePasswordSetupInput = {
  password: string;
};

export type AdminUpsertUserInput = {
  userId?: string;
  login: string;
  fullName: string;
  systemRole: ManageableSystemRoleKey;
  isActive: boolean;
  groupIds: string[];
};

export type AdminRotateUserTokenInput = {
  userId: string;
};

export type AdminUpsertPermissionGroupInput = {
  groupId?: string;
  name: string;
  description?: string;
  permissionKeys: AppPermissionKey[];
};
