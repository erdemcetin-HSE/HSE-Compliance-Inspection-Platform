-- CreateTable
CREATE TABLE "Company" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "legalName" TEXT,
    "taxNumber" TEXT,
    "country" TEXT,
    "city" TEXT,
    "address" TEXT,
    "currency" TEXT NOT NULL DEFAULT 'RUB',
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Company_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ProjectMaster" (
    "id" TEXT NOT NULL,
    "companyId" TEXT,
    "name" TEXT NOT NULL,
    "country" TEXT NOT NULL,
    "city" TEXT NOT NULL,
    "address" TEXT NOT NULL,
    "contractScope" TEXT NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ProjectMaster_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DepartmentMaster" (
    "id" TEXT NOT NULL,
    "companyId" TEXT,
    "name" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "description" TEXT,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "DepartmentMaster_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ContractorMaster" (
    "id" TEXT NOT NULL,
    "companyId" TEXT,
    "projectId" TEXT,
    "companyName" TEXT NOT NULL,
    "country" TEXT NOT NULL,
    "city" TEXT NOT NULL,
    "projectLocation" TEXT NOT NULL,
    "contractScope" TEXT NOT NULL,
    "hseWarningCount" INTEGER NOT NULL DEFAULT 0,
    "hseWarningDate" TIMESTAMP(3),
    "fireWarningCount" INTEGER NOT NULL DEFAULT 0,
    "fireWarningDate" TIMESTAMP(3),
    "environmentWarningCount" INTEGER NOT NULL DEFAULT 0,
    "environmentWarningDate" TIMESTAMP(3),
    "penaltyCount" INTEGER NOT NULL DEFAULT 0,
    "penaltyLegalClause" TEXT,
    "totalPenaltyAmount" INTEGER NOT NULL DEFAULT 0,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ContractorMaster_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "AuditLog" (
    "id" TEXT NOT NULL,
    "companyId" TEXT,
    "entityType" TEXT NOT NULL,
    "entityId" TEXT,
    "action" TEXT NOT NULL,
    "actor" TEXT NOT NULL DEFAULT 'system',
    "beforeJson" JSONB,
    "afterJson" JSONB,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "AuditLog_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "ProjectMaster_companyId_idx" ON "ProjectMaster"("companyId");

-- CreateIndex
CREATE INDEX "DepartmentMaster_companyId_idx" ON "DepartmentMaster"("companyId");

-- CreateIndex
CREATE INDEX "ContractorMaster_companyId_idx" ON "ContractorMaster"("companyId");

-- CreateIndex
CREATE INDEX "ContractorMaster_projectId_idx" ON "ContractorMaster"("projectId");

-- CreateIndex
CREATE INDEX "AuditLog_companyId_createdAt_idx" ON "AuditLog"("companyId", "createdAt");

-- AddForeignKey
ALTER TABLE "ProjectMaster" ADD CONSTRAINT "ProjectMaster_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DepartmentMaster" ADD CONSTRAINT "DepartmentMaster_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ContractorMaster" ADD CONSTRAINT "ContractorMaster_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ContractorMaster" ADD CONSTRAINT "ContractorMaster_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "ProjectMaster"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "AuditLog" ADD CONSTRAINT "AuditLog_companyId_fkey" FOREIGN KEY ("companyId") REFERENCES "Company"("id") ON DELETE SET NULL ON UPDATE CASCADE;
