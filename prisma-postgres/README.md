# Postgres Prisma Flow

This folder contains the Prisma schema and migrations for the auth/RBAC Postgres layer used by the Electron client.

## Environment selection

- `dev` uses `client/.env.test`
- `prod` uses `client/.env`

All helper scripts switch between those files with `--target dev|prod`.

## Recommended flow

### 1. Local development

Use `db push` when you need to sync a disposable dev database quickly:

```bash
npm run prisma:push:postgres
```

If you need the generated client refreshed explicitly:

```bash
npm run prisma:generate:postgres
```

### 2. Create a durable migration

When schema changes should be promoted beyond local dev, create a migration in dev:

```bash
npm run prisma:migrate:create:postgres -- --name add_some_change
```

That command uses `prisma migrate dev --create-only`, so it creates migration files without applying a destructive prod workflow.

Validate that migration files match the current schema:

```bash
npm run prisma:migrate:check:postgres
```

That check is non-destructive: it compares the combined SQL in `prisma-postgres/migrations` with the SQL Prisma currently generates from the schema.

### 3. Deploy schema safely

Preferred deploy path:

```bash
npm run prisma:migrate:deploy:postgres
```

For production, an explicit opt-in is required by the wrapper:

```bash
npm run prisma:migrate:deploy:postgres:prod
```

You can inspect migration status with:

```bash
npm run prisma:migrate:status:postgres
npm run prisma:migrate:status:postgres:prod
```

### 4. Emergency-only `db push` on prod

`db push` on prod is intentionally named unsafe and requires explicit prod allowance:

```bash
npm run prisma:push:postgres:prod:unsafe
```

Use it only for disposable environments or urgent recovery. Normal prod rollout should use `migrate deploy`.

## Seed first Super Admin

After schema is applied, create the first `SUPER_ADMIN` and print a 24-hour one-time token:

```bash
npm run seed:super-admin:postgres -- --login root.admin --full-name "Root Admin"
```

Production variant:

```bash
npm run seed:super-admin:postgres:prod -- --login root.admin --full-name "Root Admin"
```

Behavior:

- refuses to run if auth tables do not exist yet
- refuses to create a second `SUPER_ADMIN`
- stores the seeded account without a password
- creates a one-time token valid for 24 hours
- prints the raw token once to stdout

The seeded user must then log in with that token in the admin tab and set a password immediately.
