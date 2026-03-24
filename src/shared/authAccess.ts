export const SYSTEM_ROLE_KEYS = ["SUPER_ADMIN", "ADMIN", "USER"] as const;
export type SystemRoleKey = (typeof SYSTEM_ROLE_KEYS)[number];

export const MANAGEABLE_SYSTEM_ROLE_KEYS = ["ADMIN", "USER"] as const;
export type ManageableSystemRoleKey = (typeof MANAGEABLE_SYSTEM_ROLE_KEYS)[number];

export const SYSTEM_ROLE_LABELS: Record<SystemRoleKey, string> = {
  SUPER_ADMIN: "Super Admin",
  ADMIN: "Admin",
  USER: "User",
};

type AppPermissionDefinitionShape = {
  key: string;
  label: string;
  description: string;
  category: string;
  parentKey?: string;
};

export const APP_PERMISSION_CATALOG = [
  {
    key: "REPORT_IMPORT",
    label: "Import danych",
    description: "Import plikow Excel do bazy aplikacji.",
    category: "Dane",
  },
  {
    key: "REPORT_VIEW",
    label: "Podglad danych",
    description: "Podglad zaimportowanych wierszy i metadanych raportu.",
    category: "Dane",
  },
  {
    key: "REPORT_DUPLICATES_VIEW",
    label: "Duplikaty MRN",
    description: "Podglad i przebudowa grup duplikatow MRN.",
    category: "Dane",
  },
  {
    key: "VALIDATION_VIEW",
    label: "Walidacja IQR",
    description: "Dostep do ekranu walidacji i wynikow IQR.",
    category: "Walidacja",
  },
  {
    key: "VALIDATION_VERIFY_MANUAL",
    label: "Reczna weryfikacja",
    description: "Oznaczanie pozycji jako zweryfikowanych recznie.",
    category: "Walidacja",
  },
  {
    key: "ATTENTION_VIEW",
    label: "Do mojej uwagi",
    description: "Dostep do listy odchylen wymagajacych uwagi.",
    category: "Walidacja",
  },
  {
    key: "ATTENTION_VIEW_ALL",
    label: "Widok wszystkich agentow",
    description:
      "W zakladce Do mojej uwagi pozwala filtrowac i przegladac wpisy wszystkich agentow.",
    category: "Walidacja",
    parentKey: "ATTENTION_VIEW",
  },
  {
    key: "EXPORT_VIEW",
    label: "Eksport wynikow",
    description: "Podglad i eksport wynikow do Excel.",
    category: "Eksport",
  },
  {
    key: "SETTINGS_VIEW",
    label: "Ustawienia",
    description: "Podglad ustawien lokalnych i zasobow aplikacji.",
    category: "Aplikacja",
  },
] as const satisfies readonly AppPermissionDefinitionShape[];

export type AppPermissionDefinition = (typeof APP_PERMISSION_CATALOG)[number];
export type AppPermissionKey = AppPermissionDefinition["key"];

export const APP_PERMISSION_KEYS: AppPermissionKey[] = APP_PERMISSION_CATALOG.map(
  (permission) => permission.key,
);

const APP_PERMISSION_KEY_SET = new Set<string>(APP_PERMISSION_KEYS);
const APP_PERMISSION_MAP = new Map<string, AppPermissionDefinition>(
  APP_PERMISSION_CATALOG.map((permission) => [permission.key, permission]),
);

export function isAppPermissionKey(value: unknown): value is AppPermissionKey {
  return APP_PERMISSION_KEY_SET.has(String(value ?? "").trim());
}

export function getAppPermissionDefinition(
  key: unknown,
): AppPermissionDefinition | null {
  const normalizedKey = String(key ?? "").trim();
  return APP_PERMISSION_MAP.get(normalizedKey) ?? null;
}

export function expandPermissionDependencies(
  permissionKeys: readonly AppPermissionKey[],
): AppPermissionKey[] {
  const expanded = new Set<AppPermissionKey>(permissionKeys);
  const queue = [...expanded];

  while (queue.length > 0) {
    const key = queue.pop();
    if (!key) continue;
    const definition = getAppPermissionDefinition(key);
    const parentKey =
      definition && "parentKey" in definition ? definition.parentKey : undefined;
    if (!parentKey || !isAppPermissionKey(parentKey)) continue;
    if (expanded.has(parentKey)) continue;
    expanded.add(parentKey);
    queue.push(parentKey);
  }

  return APP_PERMISSION_KEYS.filter((key) => expanded.has(key));
}

export function getBasePermissionsForSystemRole(
  systemRole: SystemRoleKey,
): AppPermissionKey[] {
  if (systemRole === "SUPER_ADMIN" || systemRole === "ADMIN") {
    return [...APP_PERMISSION_KEYS];
  }
  return [];
}

export function isSystemRoleKey(value: unknown): value is SystemRoleKey {
  return SYSTEM_ROLE_KEYS.includes(value as SystemRoleKey);
}

export function isManageableSystemRoleKey(
  value: unknown,
): value is ManageableSystemRoleKey {
  return MANAGEABLE_SYSTEM_ROLE_KEYS.includes(value as ManageableSystemRoleKey);
}
