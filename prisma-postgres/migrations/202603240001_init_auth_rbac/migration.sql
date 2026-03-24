-- CreateSchema
CREATE SCHEMA IF NOT EXISTS "public";

-- CreateEnum
CREATE TYPE "AuthSystemRole" AS ENUM ('SUPER_ADMIN', 'ADMIN', 'USER');

-- CreateTable
CREATE TABLE "AuthUser" (
    "id" UUID NOT NULL,
    "login" VARCHAR(64) NOT NULL,
    "fullName" VARCHAR(160) NOT NULL,
    "systemRole" "AuthSystemRole" NOT NULL DEFAULT 'USER',
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "passwordHash" VARCHAR(255),
    "passwordSalt" VARCHAR(255),
    "passwordSetAt" TIMESTAMP(3),
    "mustSetPassword" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "AuthUser_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "AuthPermissionGroup" (
    "id" UUID NOT NULL,
    "name" VARCHAR(120) NOT NULL,
    "slug" VARCHAR(120) NOT NULL,
    "description" VARCHAR(255),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "AuthPermissionGroup_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "AuthPermissionGroupAccess" (
    "groupId" UUID NOT NULL,
    "permissionKey" VARCHAR(80) NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "AuthPermissionGroupAccess_pkey" PRIMARY KEY ("groupId","permissionKey")
);

-- CreateTable
CREATE TABLE "AuthUserPermissionGroup" (
    "userId" UUID NOT NULL,
    "groupId" UUID NOT NULL,
    "assignedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "AuthUserPermissionGroup_pkey" PRIMARY KEY ("userId","groupId")
);

-- CreateTable
CREATE TABLE "AuthOneTimeToken" (
    "id" UUID NOT NULL,
    "userId" UUID NOT NULL,
    "issuedById" UUID,
    "tokenHash" VARCHAR(255) NOT NULL,
    "tokenPreview" VARCHAR(32) NOT NULL,
    "expiresAt" TIMESTAMP(3) NOT NULL,
    "usedAt" TIMESTAMP(3),
    "revokedAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "AuthOneTimeToken_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "AuthUser_login_key" ON "AuthUser"("login");

-- CreateIndex
CREATE UNIQUE INDEX "AuthPermissionGroup_slug_key" ON "AuthPermissionGroup"("slug");

-- CreateIndex
CREATE INDEX "AuthPermissionGroupAccess_permissionKey_idx" ON "AuthPermissionGroupAccess"("permissionKey");

-- CreateIndex
CREATE INDEX "AuthUserPermissionGroup_groupId_idx" ON "AuthUserPermissionGroup"("groupId");

-- CreateIndex
CREATE UNIQUE INDEX "AuthOneTimeToken_tokenHash_key" ON "AuthOneTimeToken"("tokenHash");

-- CreateIndex
CREATE INDEX "AuthOneTimeToken_userId_expiresAt_idx" ON "AuthOneTimeToken"("userId", "expiresAt");

-- AddForeignKey
ALTER TABLE "AuthPermissionGroupAccess" ADD CONSTRAINT "AuthPermissionGroupAccess_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "AuthPermissionGroup"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "AuthUserPermissionGroup" ADD CONSTRAINT "AuthUserPermissionGroup_userId_fkey" FOREIGN KEY ("userId") REFERENCES "AuthUser"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "AuthUserPermissionGroup" ADD CONSTRAINT "AuthUserPermissionGroup_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "AuthPermissionGroup"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "AuthOneTimeToken" ADD CONSTRAINT "AuthOneTimeToken_userId_fkey" FOREIGN KEY ("userId") REFERENCES "AuthUser"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "AuthOneTimeToken" ADD CONSTRAINT "AuthOneTimeToken_issuedById_fkey" FOREIGN KEY ("issuedById") REFERENCES "AuthUser"("id") ON DELETE SET NULL ON UPDATE CASCADE;
