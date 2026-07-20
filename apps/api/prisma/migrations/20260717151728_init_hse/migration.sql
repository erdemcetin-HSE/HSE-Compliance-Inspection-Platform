-- CreateEnum
CREATE TYPE "OpenClosedStatus" AS ENUM ('OPEN', 'CLOSED');

-- CreateEnum
CREATE TYPE "RiskLevel" AS ENUM ('LOW', 'MEDIUM', 'HIGH', 'CRITICAL');

-- CreateEnum
CREATE TYPE "PtwStatus" AS ENUM ('DRAFT', 'PENDING_APPROVAL', 'ACTIVE', 'SUSPENDED', 'COMPLETED', 'CANCELLED');

-- CreateEnum
CREATE TYPE "PtwControlStatus" AS ENUM ('COMPLIANT', 'NON_COMPLIANT', 'NA');

-- CreateEnum
CREATE TYPE "PtwActionStatus" AS ENUM ('OPEN', 'IN_PROGRESS', 'CLOSED');

-- CreateEnum
CREATE TYPE "PtwAttachmentType" AS ENUM ('PHOTO', 'PDF', 'RISK_ASSESSMENT', 'METHOD_STATEMENT', 'TOOLBOX_TALK', 'CERTIFICATE', 'OTHER');

-- CreateTable
CREATE TABLE "Employee" (
    "id" TEXT NOT NULL,
    "firstName" TEXT NOT NULL,
    "lastName" TEXT NOT NULL,
    "contractor" TEXT,
    "region" TEXT,
    "roleTitle" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Employee_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Inspection" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "contractor" TEXT NOT NULL,
    "region" TEXT NOT NULL,
    "workNature" TEXT NOT NULL,
    "findings" TEXT,
    "riskLevel" "RiskLevel" NOT NULL,
    "correctiveAction" TEXT,
    "dueDate" TIMESTAMP(3),
    "status" "OpenClosedStatus" NOT NULL DEFAULT 'OPEN',
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Inspection_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "InspectionAttachment" (
    "id" TEXT NOT NULL,
    "inspectionId" TEXT NOT NULL,
    "fileName" TEXT NOT NULL,
    "fileType" TEXT NOT NULL,
    "fileUrl" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "InspectionAttachment_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PpeItem" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "region" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PpeItem_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PpeTransaction" (
    "id" TEXT NOT NULL,
    "itemId" TEXT NOT NULL,
    "quantityIn" INTEGER NOT NULL DEFAULT 0,
    "quantityUsed" INTEGER NOT NULL DEFAULT 0,
    "quantityLeft" INTEGER NOT NULL,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PpeTransaction_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwRecord" (
    "id" TEXT NOT NULL,
    "permitNo" TEXT NOT NULL,
    "organization" TEXT NOT NULL,
    "department" TEXT NOT NULL,
    "project" TEXT NOT NULL,
    "location" TEXT NOT NULL,
    "permitType" TEXT NOT NULL,
    "issueDate" TIMESTAMP(3) NOT NULL,
    "validityStart" TIMESTAMP(3) NOT NULL,
    "validityEnd" TIMESTAMP(3) NOT NULL,
    "status" "PtwStatus" NOT NULL DEFAULT 'DRAFT',
    "requesterName" TEXT NOT NULL,
    "issuerName" TEXT NOT NULL,
    "jobResponsibleName" TEXT NOT NULL,
    "siteResponsibleName" TEXT NOT NULL,
    "hseResponsibleName" TEXT NOT NULL,
    "authorizedApprover" TEXT NOT NULL,
    "jobTitle" TEXT NOT NULL,
    "jobDescription" TEXT NOT NULL,
    "workArea" TEXT NOT NULL,
    "plannedWork" TEXT NOT NULL,
    "workingConditions" TEXT NOT NULL,
    "workStartDate" TIMESTAMP(3) NOT NULL,
    "workStartTime" TEXT NOT NULL,
    "workEndDate" TIMESTAMP(3) NOT NULL,
    "workEndTime" TEXT NOT NULL,
    "specialConditions" TEXT,
    "preparedBy" TEXT,
    "hseApprovalBy" TEXT,
    "projectManagerBy" TEXT,
    "employerRepBy" TEXT,
    "digitalSignature" TEXT,
    "approvalDate" TIMESTAMP(3),
    "workCompleted" BOOLEAN NOT NULL DEFAULT false,
    "areaSafe" BOOLEAN NOT NULL DEFAULT false,
    "materialsCollected" BOOLEAN NOT NULL DEFAULT false,
    "ptwClosed" BOOLEAN NOT NULL DEFAULT false,
    "closedBy" TEXT,
    "closedAt" TIMESTAMP(3),
    "closureRemarks" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PtwRecord_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwHazard" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwHazard_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwSafetySystem" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwSafetySystem_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwEquipment" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwEquipment_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwTeamMember" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "fullName" TEXT NOT NULL,
    "duty" TEXT NOT NULL,
    "company" TEXT NOT NULL,
    "trainingStatus" TEXT NOT NULL,
    "signature" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwTeamMember_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwPreWorkCheck" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "checkItem" TEXT NOT NULL,
    "status" "PtwControlStatus" NOT NULL DEFAULT 'COMPLIANT',
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwPreWorkCheck_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwPrecaution" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "measure" TEXT NOT NULL,
    "responsible" TEXT NOT NULL,
    "dueDate" TIMESTAMP(3),
    "status" "PtwActionStatus" NOT NULL DEFAULT 'OPEN',
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwPrecaution_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwDailyLog" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "logDate" TIMESTAMP(3) NOT NULL,
    "logTime" TEXT NOT NULL,
    "workStarted" TEXT,
    "workEnded" TEXT,
    "note" TEXT,
    "responsible" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwDailyLog_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwTeamChange" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "addedPersonnel" TEXT,
    "removedPersonnel" TEXT,
    "changeDate" TIMESTAMP(3) NOT NULL,
    "approvedBy" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwTeamChange_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PtwAttachment" (
    "id" TEXT NOT NULL,
    "ptwId" TEXT NOT NULL,
    "type" "PtwAttachmentType" NOT NULL DEFAULT 'OTHER',
    "fileName" TEXT NOT NULL,
    "fileType" TEXT NOT NULL,
    "fileUrl" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "PtwAttachment_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Observation" (
    "id" TEXT NOT NULL,
    "observationType" TEXT NOT NULL,
    "location" TEXT NOT NULL,
    "details" TEXT NOT NULL,
    "ownerName" TEXT NOT NULL,
    "status" "OpenClosedStatus" NOT NULL DEFAULT 'OPEN',
    "closedAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Observation_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "RiskRecord" (
    "id" TEXT NOT NULL,
    "title" TEXT NOT NULL,
    "probability" INTEGER NOT NULL,
    "impact" INTEGER NOT NULL,
    "score" INTEGER NOT NULL,
    "riskLevel" "RiskLevel" NOT NULL,
    "status" "OpenClosedStatus" NOT NULL DEFAULT 'OPEN',
    "actionOwner" TEXT,
    "dueDate" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "RiskRecord_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Training" (
    "id" TEXT NOT NULL,
    "employeeId" TEXT NOT NULL,
    "trainingName" TEXT NOT NULL,
    "trainingDate" TIMESTAMP(3) NOT NULL,
    "validUntil" TIMESTAMP(3) NOT NULL,
    "certificateNo" TEXT,
    "certificateUrl" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Training_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "PtwRecord_permitNo_key" ON "PtwRecord"("permitNo");

-- AddForeignKey
ALTER TABLE "Inspection" ADD CONSTRAINT "Inspection_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "Employee"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "InspectionAttachment" ADD CONSTRAINT "InspectionAttachment_inspectionId_fkey" FOREIGN KEY ("inspectionId") REFERENCES "Inspection"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PpeTransaction" ADD CONSTRAINT "PpeTransaction_itemId_fkey" FOREIGN KEY ("itemId") REFERENCES "PpeItem"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwHazard" ADD CONSTRAINT "PtwHazard_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwSafetySystem" ADD CONSTRAINT "PtwSafetySystem_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwEquipment" ADD CONSTRAINT "PtwEquipment_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwTeamMember" ADD CONSTRAINT "PtwTeamMember_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwPreWorkCheck" ADD CONSTRAINT "PtwPreWorkCheck_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwPrecaution" ADD CONSTRAINT "PtwPrecaution_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwDailyLog" ADD CONSTRAINT "PtwDailyLog_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwTeamChange" ADD CONSTRAINT "PtwTeamChange_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PtwAttachment" ADD CONSTRAINT "PtwAttachment_ptwId_fkey" FOREIGN KEY ("ptwId") REFERENCES "PtwRecord"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Training" ADD CONSTRAINT "Training_employeeId_fkey" FOREIGN KEY ("employeeId") REFERENCES "Employee"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
