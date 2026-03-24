import {
  createHash,
  randomBytes,
  scryptSync,
  timingSafeEqual,
} from "crypto";
import { getPostgresPrisma } from "./postgresDb";
import {
  APP_PERMISSION_CATALOG,
  APP_PERMISSION_KEYS,
  getBasePermissionsForSystemRole,
  isAppPermissionKey,
  isManageableSystemRoleKey,
  type AppPermissionKey,
  type ManageableSystemRoleKey,
  type SystemRoleKey,
} from "../shared/authAccess";
import type {
  AdminPanelData,
  AdminRotateUserTokenInput,
  AdminUpsertPermissionGroupInput,
  AdminUpsertUserInput,
  AuthActionResult,
  AuthSessionState,
  AuthSessionUser,
  BootstrapSuperAdminInput,
  CompletePasswordSetupInput,
  FieldErrorMap,
  IssuedOneTimeToken,
  PasswordLoginInput,
  TokenLoginInput,
} from "../shared/authTypes";

type SessionRecord = {
  userId: string;
  requiresPasswordSetup: boolean;
};

class ValidationError extends Error {
  fieldErrors: FieldErrorMap;

  constructor(message: string, fieldErrors: FieldErrorMap = {}) {
    super(message);
    this.name = "ValidationError";
    this.fieldErrors = fieldErrors;
  }
}

let currentSession: SessionRecord | null = null;

const TOKEN_TTL_MS = 24 * 60 * 60 * 1000;
const PASSWORD_MIN_LENGTH = 10;
const LOGIN_PATTERN = /^[a-z0-9._-]{3,64}$/;
const PASSWORD_UPPERCASE = /[A-Z]/;
const PASSWORD_LOWERCASE = /[a-z]/;
const PASSWORD_DIGIT = /\d/;

async function wrapAction<T>(
  action: () => Promise<T>,
): Promise<AuthActionResult<T>> {
  try {
    const data = await action();
    return { ok: true, data };
  } catch (error: unknown) {
    if (error instanceof ValidationError) {
      return {
        ok: false,
        error: error.message,
        fieldErrors: error.fieldErrors,
      };
    }
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

function normalizeLogin(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

function normalizeFullName(value: unknown): string {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizeGroupName(value: unknown): string {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function uniqueStrings(values: string[]): string[] {
  return [...new Set(values)];
}

function makeSlug(value: string): string {
  const normalized = value
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return normalized || "group";
}

function hashToken(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

function generateSalt(): string {
  return randomBytes(16).toString("hex");
}

function hashPassword(password: string, salt: string): string {
  return scryptSync(password, salt, 64).toString("hex");
}

function verifyPasswordHash(
  password: string,
  salt: string,
  expectedHash: string,
): boolean {
  const actual = Buffer.from(hashPassword(password, salt), "hex");
  const expected = Buffer.from(expectedHash, "hex");
  if (actual.length !== expected.length) return false;
  return timingSafeEqual(actual, expected);
}

function generatePlainToken(): {
  token: string;
  tokenHash: string;
  tokenPreview: string;
  expiresAt: Date;
} {
  const token = randomBytes(24).toString("base64url");
  return {
    token,
    tokenHash: hashToken(token),
    tokenPreview: `${token.slice(0, 6)}...${token.slice(-4)}`,
    expiresAt: new Date(Date.now() + TOKEN_TTL_MS),
  };
}

function validateLogin(login: string): string {
  if (!login) return "Login jest wymagany.";
  if (!LOGIN_PATTERN.test(login)) {
    return "Login moze zawierac tylko male litery, cyfry oraz . _ -.";
  }
  return "";
}

function validateFullName(fullName: string): string {
  if (!fullName) return "Full Name jest wymagane.";
  if (fullName.length < 3) return "Full Name musi miec minimum 3 znaki.";
  if (fullName.length > 160) return "Full Name jest za dlugie.";
  return "";
}

function validatePassword(password: string): string {
  if (!password) return "Haslo jest wymagane.";
  if (password.length < PASSWORD_MIN_LENGTH) {
    return `Haslo musi miec minimum ${PASSWORD_MIN_LENGTH} znakow.`;
  }
  if (!PASSWORD_UPPERCASE.test(password)) {
    return "Haslo musi zawierac przynajmniej jedna wielka litere.";
  }
  if (!PASSWORD_LOWERCASE.test(password)) {
    return "Haslo musi zawierac przynajmniej jedna mala litere.";
  }
  if (!PASSWORD_DIGIT.test(password)) {
    return "Haslo musi zawierac przynajmniej jedna cyfre.";
  }
  return "";
}

function validatePermissionGroupName(name: string): string {
  if (!name) return "Nazwa grupy jest wymagana.";
  if (name.length < 3) return "Nazwa grupy musi miec minimum 3 znaki.";
  if (name.length > 120) return "Nazwa grupy jest za dluga.";
  return "";
}

function buildValidationError(fieldErrors: FieldErrorMap): ValidationError {
  return new ValidationError("Popraw oznaczone pola.", fieldErrors);
}

function getManageableRolesForActor(
  actorRole: SystemRoleKey,
): ManageableSystemRoleKey[] {
  if (actorRole === "SUPER_ADMIN") return ["ADMIN", "USER"];
  if (actorRole === "ADMIN") return ["USER"];
  return [];
}

function buildEffectivePermissions(params: {
  systemRole: SystemRoleKey;
  groupPermissionKeys: string[];
}): AppPermissionKey[] {
  const set = new Set<AppPermissionKey>(
    getBasePermissionsForSystemRole(params.systemRole),
  );

  for (const key of params.groupPermissionKeys) {
    if (isAppPermissionKey(key)) set.add(key);
  }

  return APP_PERMISSION_KEYS.filter((key) => set.has(key));
}

async function superAdminExists(): Promise<boolean> {
  const prisma = getPostgresPrisma();
  const count = await prisma.authUser.count({
    where: { systemRole: "SUPER_ADMIN" as never },
  });
  return count > 0;
}

async function findUniqueSlug(
  desiredName: string,
  currentGroupId?: string,
): Promise<string> {
  const prisma = getPostgresPrisma();
  const baseSlug = makeSlug(desiredName);

  for (let index = 0; index < 200; index += 1) {
    const candidate = index === 0 ? baseSlug : `${baseSlug}-${index + 1}`;
    const existing = await prisma.authPermissionGroup.findUnique({
      where: { slug: candidate },
      select: { id: true },
    });

    if (!existing || existing.id === currentGroupId) return candidate;
  }

  throw new Error("Nie udalo sie wygenerowac unikalnego slug.");
}

async function fetchSessionUserById(userId: string) {
  const prisma = getPostgresPrisma();
  return prisma.authUser.findUnique({
    where: { id: userId },
    include: {
      groups: {
        include: {
          group: {
            include: {
              permissions: {
                orderBy: { permissionKey: "asc" },
              },
            },
          },
        },
      },
    },
  });
}

async function getSessionStateInternal(): Promise<AuthSessionState> {
  const bootstrapRequired = !(await superAdminExists());
  if (bootstrapRequired) {
    currentSession = null;
    return {
      bootstrapRequired: true,
      isAuthenticated: false,
      requiresPasswordSetup: false,
      user: null,
    };
  }

  if (!currentSession) {
    return {
      bootstrapRequired: false,
      isAuthenticated: false,
      requiresPasswordSetup: false,
      user: null,
    };
  }

  const user = await fetchSessionUserById(currentSession.userId);
  if (!user || !user.isActive) {
    currentSession = null;
    return {
      bootstrapRequired: false,
      isAuthenticated: false,
      requiresPasswordSetup: false,
      user: null,
    };
  }

  const groupPermissionKeys = user.groups.flatMap((membership) =>
    membership.group.permissions.map((permission) => permission.permissionKey),
  );
  const effectivePermissions = buildEffectivePermissions({
    systemRole: user.systemRole as SystemRoleKey,
    groupPermissionKeys,
  });

  const sessionUser: AuthSessionUser = {
    id: user.id,
    login: user.login,
    fullName: user.fullName,
    systemRole: user.systemRole as SystemRoleKey,
    isActive: user.isActive,
    groupIds: user.groups.map((membership) => membership.groupId),
    groupNames: user.groups.map((membership) => membership.group.name),
    effectivePermissions,
    canAccessAdminPanel:
      user.systemRole === "SUPER_ADMIN" || user.systemRole === "ADMIN",
  };

  return {
    bootstrapRequired: false,
    isAuthenticated: true,
    requiresPasswordSetup:
      currentSession.requiresPasswordSetup || Boolean(user.mustSetPassword),
    user: sessionUser,
  };
}

async function requireAuthenticatedContext(options?: {
  allowPasswordSetup?: boolean;
}) {
  const session = await getSessionStateInternal();
  if (session.bootstrapRequired) {
    throw new Error(
      "Najpierw utworz pierwszego super admina w zakladce administracyjnej.",
    );
  }
  if (!session.isAuthenticated || !session.user) {
    throw new Error("Najpierw zaloguj sie do aplikacji.");
  }
  if (!options?.allowPasswordSetup && session.requiresPasswordSetup) {
    throw new Error("Najpierw ustaw nowe haslo po logowaniu tokenem.");
  }
  return session;
}

async function requireAdminActor() {
  const session = await requireAuthenticatedContext();
  if (!session.user?.canAccessAdminPanel) {
    throw new Error("Brak dostepu do panelu administratora.");
  }
  return session;
}

async function assertCanManageTargetUser(
  actorRole: SystemRoleKey,
  targetRole: SystemRoleKey,
) {
  if (targetRole === "SUPER_ADMIN") {
    throw new Error("Panel nie pozwala zarzadzac kontami super admin.");
  }
  if (actorRole === "ADMIN" && targetRole !== "USER") {
    throw new Error("Admin moze zarzadzac tylko kontami user.");
  }
}

async function buildIssuedTokenPayload(
  userId: string,
  login: string,
  fullName: string,
  issuedById?: string,
): Promise<IssuedOneTimeToken> {
  const prisma = getPostgresPrisma();
  const token = generatePlainToken();

  await prisma.$transaction(async (tx) => {
    await tx.authOneTimeToken.updateMany({
      where: {
        userId,
        usedAt: null,
        revokedAt: null,
      },
      data: {
        revokedAt: new Date(),
      },
    });

    await tx.authUser.update({
      where: { id: userId },
      data: {
        passwordHash: null,
        passwordSalt: null,
        passwordSetAt: null,
        mustSetPassword: true,
      },
    });

    await tx.authOneTimeToken.create({
      data: {
        userId,
        issuedById: issuedById ?? null,
        tokenHash: token.tokenHash,
        tokenPreview: token.tokenPreview,
        expiresAt: token.expiresAt,
      },
    });
  });

  return {
    userId,
    login,
    fullName,
    token: token.token,
    tokenPreview: token.tokenPreview,
    expiresAt: token.expiresAt.toISOString(),
  };
}

async function ensureGroupIdsExist(groupIds: string[]): Promise<string[]> {
  const normalized = uniqueStrings(
    groupIds.map((groupId) => String(groupId ?? "").trim()).filter(Boolean),
  );
  if (!normalized.length) return [];

  const prisma = getPostgresPrisma();
  const groups = await prisma.authPermissionGroup.findMany({
    where: {
      id: {
        in: normalized,
      },
    },
    select: { id: true },
  });
  const existingIds = new Set(groups.map((group) => group.id));
  const missing = normalized.filter((groupId) => !existingIds.has(groupId));
  if (missing.length) {
    throw new ValidationError("Wybrano nieistniejaca grupe dostepow.", {
      groupIds: "Jedna lub wiecej grup nie istnieje.",
    });
  }
  return normalized;
}

function mapAdminUserSummary(user: {
  id: string;
  login: string;
  fullName: string;
  systemRole: string;
  isActive: boolean;
  mustSetPassword: boolean;
  passwordSetAt: Date | null;
  createdAt: Date;
  updatedAt: Date;
  groups: Array<{
    groupId: string;
    group: {
      name: string;
      permissions: Array<{ permissionKey: string }>;
    };
  }>;
  tokens: Array<{
    expiresAt: Date;
    tokenPreview: string;
  }>;
}) {
  const groupPermissionKeys = user.groups.flatMap((membership) =>
    membership.group.permissions.map((permission) => permission.permissionKey),
  );
  return {
    id: user.id,
    login: user.login,
    fullName: user.fullName,
    systemRole: user.systemRole as SystemRoleKey,
    isActive: user.isActive,
    mustSetPassword: user.mustSetPassword,
    passwordSetAt: user.passwordSetAt?.toISOString() ?? null,
    createdAt: user.createdAt.toISOString(),
    updatedAt: user.updatedAt.toISOString(),
    groupIds: user.groups.map((membership) => membership.groupId),
    groupNames: user.groups.map((membership) => membership.group.name),
    effectivePermissions: buildEffectivePermissions({
      systemRole: user.systemRole as SystemRoleKey,
      groupPermissionKeys,
    }),
    activeToken: {
      exists: user.tokens.length > 0,
      expiresAt: user.tokens[0]?.expiresAt.toISOString() ?? null,
      tokenPreview: user.tokens[0]?.tokenPreview ?? null,
    },
  };
}

function validateUserInput(
  input: AdminUpsertUserInput,
  actorRole: SystemRoleKey,
): {
  login: string;
  fullName: string;
  systemRole: ManageableSystemRoleKey;
  isActive: boolean;
} {
  const login = normalizeLogin(input.login);
  const fullName = normalizeFullName(input.fullName);
  const systemRole = String(input.systemRole ?? "").trim();
  const fieldErrors: FieldErrorMap = {};

  const loginError = validateLogin(login);
  if (loginError) fieldErrors.login = loginError;

  const fullNameError = validateFullName(fullName);
  if (fullNameError) fieldErrors.fullName = fullNameError;

  if (!isManageableSystemRoleKey(systemRole)) {
    fieldErrors.systemRole = "Nieprawidlowa rola systemowa.";
  } else if (!getManageableRolesForActor(actorRole).includes(systemRole)) {
    fieldErrors.systemRole = "Brak uprawnien do nadania tej roli.";
  }

  if (Object.keys(fieldErrors).length) throw buildValidationError(fieldErrors);

  return {
    login,
    fullName,
    systemRole: systemRole as ManageableSystemRoleKey,
    isActive: Boolean(input.isActive),
  };
}

function validatePermissionGroupInput(input: AdminUpsertPermissionGroupInput): {
  name: string;
  description: string | null;
  permissionKeys: AppPermissionKey[];
} {
  const name = normalizeGroupName(input.name);
  const description = normalizeGroupName(input.description ?? "");
  const rawPermissionKeys = Array.isArray(input.permissionKeys)
    ? input.permissionKeys
    : [];
  const permissionKeys = uniqueStrings(
    rawPermissionKeys
      .map((permissionKey) => String(permissionKey ?? "").trim())
      .filter(Boolean),
  );

  const fieldErrors: FieldErrorMap = {};
  const nameError = validatePermissionGroupName(name);
  if (nameError) fieldErrors.name = nameError;

  const invalidPermission = permissionKeys.find(
    (permissionKey) => !isAppPermissionKey(permissionKey),
  );
  if (invalidPermission) {
    fieldErrors.permissionKeys = "Wybrano nieprawidlowy dostep.";
  }

  if (Object.keys(fieldErrors).length) throw buildValidationError(fieldErrors);

  return {
    name,
    description: description ? description.slice(0, 255) : null,
    permissionKeys: permissionKeys.filter(isAppPermissionKey),
  };
}

export async function getAuthSessionState(): Promise<AuthSessionState> {
  return getSessionStateInternal();
}

export async function bootstrapSuperAdmin(
  input: BootstrapSuperAdminInput,
): Promise<AuthActionResult<IssuedOneTimeToken>> {
  return wrapAction(async () => {
    if (await superAdminExists()) {
      throw new Error("Super admin juz istnieje. Uzyj standardowego logowania.");
    }

    const login = normalizeLogin(input.login);
    const fullName = normalizeFullName(input.fullName);
    const fieldErrors: FieldErrorMap = {};

    const loginError = validateLogin(login);
    if (loginError) fieldErrors.login = loginError;

    const fullNameError = validateFullName(fullName);
    if (fullNameError) fieldErrors.fullName = fullNameError;

    if (Object.keys(fieldErrors).length) throw buildValidationError(fieldErrors);

    const prisma = getPostgresPrisma();
    const existing = await prisma.authUser.findUnique({
      where: { login },
      select: { id: true },
    });
    if (existing) {
      throw new ValidationError("Login jest juz zajety.", {
        login: "Ten login juz istnieje.",
      });
    }

    const user = await prisma.authUser.create({
      data: {
        login,
        fullName,
        systemRole: "SUPER_ADMIN" as never,
        isActive: true,
        mustSetPassword: true,
      },
      select: {
        id: true,
        login: true,
        fullName: true,
      },
    });

    return buildIssuedTokenPayload(user.id, user.login, user.fullName);
  });
}

export async function loginWithPassword(
  input: PasswordLoginInput,
): Promise<AuthActionResult<AuthSessionState>> {
  return wrapAction(async () => {
    const login = normalizeLogin(input.login);
    const password = String(input.password ?? "");
    const fieldErrors: FieldErrorMap = {};

    const loginError = validateLogin(login);
    if (loginError) fieldErrors.login = loginError;
    if (!password) fieldErrors.password = "Haslo jest wymagane.";
    if (Object.keys(fieldErrors).length) throw buildValidationError(fieldErrors);

    if (!(await superAdminExists())) {
      throw new Error("Najpierw utworz pierwszego super admina.");
    }

    const prisma = getPostgresPrisma();
    const user = await prisma.authUser.findUnique({
      where: { login },
      select: {
        id: true,
        isActive: true,
        passwordHash: true,
        passwordSalt: true,
      },
    });

    if (!user || !user.isActive || !user.passwordHash || !user.passwordSalt) {
      throw new ValidationError("Nieprawidlowy login lub haslo.", {
        login: "Sprawdz dane logowania.",
        password: "Sprawdz dane logowania.",
      });
    }

    if (!verifyPasswordHash(password, user.passwordSalt, user.passwordHash)) {
      throw new ValidationError("Nieprawidlowy login lub haslo.", {
        login: "Sprawdz dane logowania.",
        password: "Sprawdz dane logowania.",
      });
    }

    currentSession = {
      userId: user.id,
      requiresPasswordSetup: false,
    };

    return getSessionStateInternal();
  });
}

export async function loginWithToken(
  input: TokenLoginInput,
): Promise<AuthActionResult<AuthSessionState>> {
  return wrapAction(async () => {
    const login = normalizeLogin(input.login);
    const tokenRaw = String(input.token ?? "").trim();
    const fieldErrors: FieldErrorMap = {};

    const loginError = validateLogin(login);
    if (loginError) fieldErrors.login = loginError;
    if (!tokenRaw) fieldErrors.token = "Token jest wymagany.";
    if (Object.keys(fieldErrors).length) throw buildValidationError(fieldErrors);

    if (!(await superAdminExists())) {
      throw new Error("Najpierw utworz pierwszego super admina.");
    }

    const prisma = getPostgresPrisma();
    const now = new Date();
    const tokenHash = hashToken(tokenRaw);
    const token = await prisma.authOneTimeToken.findFirst({
      where: {
        tokenHash,
        usedAt: null,
        revokedAt: null,
        expiresAt: { gt: now },
        user: {
          login,
          isActive: true,
        },
      },
      select: {
        id: true,
        userId: true,
      },
    });

    if (!token) {
      throw new ValidationError("Token jest nieprawidlowy lub wygasl.", {
        token: "Token jest nieprawidlowy lub wygasl.",
      });
    }

    await prisma.authOneTimeToken.update({
      where: { id: token.id },
      data: { usedAt: now },
    });

    currentSession = {
      userId: token.userId,
      requiresPasswordSetup: true,
    };

    return getSessionStateInternal();
  });
}

export async function completePasswordSetup(
  input: CompletePasswordSetupInput,
): Promise<AuthActionResult<AuthSessionState>> {
  return wrapAction(async () => {
    const session = await requireAuthenticatedContext({ allowPasswordSetup: true });
    if (!session.requiresPasswordSetup || !session.user) {
      throw new Error("Zmiana hasla w tym trybie nie jest wymagana.");
    }

    const password = String(input.password ?? "");
    const passwordError = validatePassword(password);
    if (passwordError) {
      throw buildValidationError({ password: passwordError });
    }

    const salt = generateSalt();
    const passwordHash = hashPassword(password, salt);
    const prisma = getPostgresPrisma();
    const now = new Date();

    await prisma.$transaction(async (tx) => {
      await tx.authUser.update({
        where: { id: session.user!.id },
        data: {
          passwordSalt: salt,
          passwordHash,
          passwordSetAt: now,
          mustSetPassword: false,
        },
      });

      await tx.authOneTimeToken.updateMany({
        where: {
          userId: session.user!.id,
          usedAt: null,
          revokedAt: null,
        },
        data: {
          revokedAt: now,
        },
      });
    });

    currentSession = {
      userId: session.user.id,
      requiresPasswordSetup: false,
    };

    return getSessionStateInternal();
  });
}

export async function logout(): Promise<AuthSessionState> {
  currentSession = null;
  return getSessionStateInternal();
}

export async function getAdminPanelData(): Promise<AuthActionResult<AdminPanelData>> {
  return wrapAction(async () => {
    const session = await requireAdminActor();
    const prisma = getPostgresPrisma();
    const now = new Date();

    const [users, groups] = await Promise.all([
      prisma.authUser.findMany({
        orderBy: [{ systemRole: "asc" }, { login: "asc" }],
        include: {
          groups: {
            include: {
              group: {
                include: {
                  permissions: {
                    orderBy: { permissionKey: "asc" },
                  },
                },
              },
            },
          },
          tokens: {
            where: {
              usedAt: null,
              revokedAt: null,
              expiresAt: { gt: now },
            },
            orderBy: { expiresAt: "desc" },
            take: 1,
          },
        },
      }),
      prisma.authPermissionGroup.findMany({
        orderBy: { name: "asc" },
        include: {
          permissions: {
            orderBy: { permissionKey: "asc" },
          },
          _count: {
            select: {
              users: true,
            },
          },
        },
      }),
    ]);

    const actorRole = session.user!.systemRole;

    return {
      permissionCatalog: [...APP_PERMISSION_CATALOG],
      manageableSystemRoles: getManageableRolesForActor(actorRole),
      users: users.map(mapAdminUserSummary),
      groups: groups.map((group) => ({
        id: group.id,
        name: group.name,
        slug: group.slug,
        description: group.description,
        permissionKeys: group.permissions
          .map((permission) => permission.permissionKey)
          .filter(isAppPermissionKey),
        memberCount: group._count.users,
        createdAt: group.createdAt.toISOString(),
        updatedAt: group.updatedAt.toISOString(),
      })),
    };
  });
}

export async function saveAdminUser(
  input: AdminUpsertUserInput,
): Promise<AuthActionResult<{ userId: string; issuedToken?: IssuedOneTimeToken }>> {
  return wrapAction(async () => {
    const session = await requireAdminActor();
    const actor = session.user!;
    const normalized = validateUserInput(input, actor.systemRole);
    const groupIds = await ensureGroupIdsExist(input.groupIds ?? []);
    const prisma = getPostgresPrisma();

    const existingByLogin = await prisma.authUser.findUnique({
      where: { login: normalized.login },
      select: {
        id: true,
      },
    });

    if (existingByLogin && existingByLogin.id !== input.userId) {
      throw new ValidationError("Login jest juz zajety.", {
        login: "Ten login juz istnieje.",
      });
    }

    if (input.userId) {
      const target = await prisma.authUser.findUnique({
        where: { id: String(input.userId) },
        select: {
          id: true,
          systemRole: true,
        },
      });

      if (!target) {
        throw new Error("Nie znaleziono wskazanego uzytkownika.");
      }

      await assertCanManageTargetUser(
        actor.systemRole,
        target.systemRole as SystemRoleKey,
      );

      await prisma.$transaction(async (tx) => {
        await tx.authUser.update({
          where: { id: target.id },
          data: {
            login: normalized.login,
            fullName: normalized.fullName,
            systemRole: normalized.systemRole as never,
            isActive: normalized.isActive,
          },
        });

        await tx.authUserPermissionGroup.deleteMany({
          where: { userId: target.id },
        });

        if (groupIds.length) {
          await tx.authUserPermissionGroup.createMany({
            data: groupIds.map((groupId) => ({
              userId: target.id,
              groupId,
            })),
          });
        }
      });

      return { userId: target.id };
    }

    const createdUser = await prisma.$transaction(async (tx) => {
      const user = await tx.authUser.create({
        data: {
          login: normalized.login,
          fullName: normalized.fullName,
          systemRole: normalized.systemRole as never,
          isActive: normalized.isActive,
          mustSetPassword: true,
        },
        select: {
          id: true,
          login: true,
          fullName: true,
        },
      });

      if (groupIds.length) {
        await tx.authUserPermissionGroup.createMany({
          data: groupIds.map((groupId) => ({
            userId: user.id,
            groupId,
          })),
        });
      }

      return user;
    });

    const issuedToken = await buildIssuedTokenPayload(
      createdUser.id,
      createdUser.login,
      createdUser.fullName,
      actor.id,
    );

    return {
      userId: createdUser.id,
      issuedToken,
    };
  });
}

export async function rotateAdminUserToken(
  input: AdminRotateUserTokenInput,
): Promise<AuthActionResult<IssuedOneTimeToken>> {
  return wrapAction(async () => {
    const session = await requireAdminActor();
    const actor = session.user!;
    const userId = String(input.userId ?? "").trim();
    if (!userId) {
      throw buildValidationError({ userId: "Wskaz uzytkownika." });
    }

    const prisma = getPostgresPrisma();
    const target = await prisma.authUser.findUnique({
      where: { id: userId },
      select: {
        id: true,
        login: true,
        fullName: true,
        systemRole: true,
        isActive: true,
      },
    });

    if (!target) {
      throw new Error("Nie znaleziono wskazanego uzytkownika.");
    }

    await assertCanManageTargetUser(
      actor.systemRole,
      target.systemRole as SystemRoleKey,
    );

    if (!target.isActive) {
      throw new Error("Najpierw aktywuj konto uzytkownika.");
    }

    return buildIssuedTokenPayload(
      target.id,
      target.login,
      target.fullName,
      actor.id,
    );
  });
}

export async function savePermissionGroup(
  input: AdminUpsertPermissionGroupInput,
): Promise<AuthActionResult<{ groupId: string }>> {
  return wrapAction(async () => {
    await requireAdminActor();
    const normalized = validatePermissionGroupInput(input);
    const slug = await findUniqueSlug(normalized.name, input.groupId);
    const prisma = getPostgresPrisma();

    const existingByName = await prisma.authPermissionGroup.findFirst({
      where: {
        name: normalized.name,
      },
      select: {
        id: true,
      },
    });

    if (existingByName && existingByName.id !== input.groupId) {
      throw new ValidationError("Grupa o tej nazwie juz istnieje.", {
        name: "Wybierz inna nazwe grupy.",
      });
    }

    if (input.groupId) {
      const groupId = String(input.groupId).trim();
      const group = await prisma.authPermissionGroup.findUnique({
        where: { id: groupId },
        select: { id: true },
      });
      if (!group) throw new Error("Nie znaleziono wskazanej grupy.");

      await prisma.$transaction(async (tx) => {
        await tx.authPermissionGroup.update({
          where: { id: groupId },
          data: {
            name: normalized.name,
            slug,
            description: normalized.description,
          },
        });

        await tx.authPermissionGroupAccess.deleteMany({
          where: { groupId },
        });

        if (normalized.permissionKeys.length) {
          await tx.authPermissionGroupAccess.createMany({
            data: normalized.permissionKeys.map((permissionKey) => ({
              groupId,
              permissionKey,
            })),
          });
        }
      });

      return { groupId };
    }

    const group = await prisma.authPermissionGroup.create({
      data: {
        name: normalized.name,
        slug,
        description: normalized.description,
        permissions: {
          create: normalized.permissionKeys.map((permissionKey) => ({
            permissionKey,
          })),
        },
      },
      select: {
        id: true,
      },
    });

    return { groupId: group.id };
  });
}

export async function assertFeatureAccess(
  permission: AppPermissionKey | "AUTHENTICATED",
): Promise<void> {
  const session = await requireAuthenticatedContext();
  if (permission === "AUTHENTICATED") return;

  if (!session.user?.effectivePermissions.includes(permission)) {
    throw new Error("Brak uprawnien do wykonania tej akcji.");
  }
}
