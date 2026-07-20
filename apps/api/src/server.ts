import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import multer from 'multer';
import nodemailer from 'nodemailer';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableCell, TableRow } from 'docx';

import {
  OpenClosedStatus,
  PtwActionStatus,
  PtwAttachmentType,
  PtwControlStatus,
  PtwStatus,
  RiskLevel
} from '@prisma/client';
import { z } from 'zod';
import { prisma } from './prisma.js';
dotenv.config();

type LanguageCode = 'tr' | 'en' | 'ru';

const app = express();
const frontendOrigin = process.env.FRONTEND_ORIGIN;
app.use(
  cors({
    origin: frontendOrigin ? [frontendOrigin] : true,
    credentials: true
  })
);
app.use(express.json());

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const uploadsDir = path.resolve(__dirname, '../uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (_req, file, cb) => {
    const safe = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
    cb(null, `${Date.now()}-${safe}`);
  }
});

const upload = multer({ storage, limits: { fileSize: 8 * 1024 * 1024 } });
app.use('/uploads', express.static(uploadsDir));


const employeeCreateSchema = z.object({
  firstName: z.string().min(1),
  lastName: z.string().min(1),
  contractor: z.string().optional(),
  region: z.string().optional(),
  roleTitle: z.string().optional()
});

const inspectionCreateSchema = z.object({
  employeeId: z.string().min(1),
  contractor: z.string().min(1),
  region: z.string().min(1),
  workNature: z.string().min(1),
  findings: z.string().optional(),
  riskLevel: z.nativeEnum(RiskLevel),
  correctiveAction: z.string().optional(),
  dueDate: z.string().datetime().optional(),
  status: z.nativeEnum(OpenClosedStatus).optional(),
  attachments: z
    .array(
      z.object({
        fileName: z.string().min(1),
        fileType: z.string().min(1),
        fileUrl: z.string().min(1)
      })
    )
    .optional()
});

const ppeItemCreateSchema = z.object({
  name: z.string().min(1),
  region: z.string().optional()
});

const ppeTransactionCreateSchema = z.object({
  itemId: z.string().min(1),
  quantityIn: z.number().int().min(0).default(0),
  quantityUsed: z.number().int().min(0).default(0),
  note: z.string().optional()
});
const companyCreateSchema = z.object({
  name: z.string().min(1),
  legalName: z.string().optional(),
  taxNumber: z.string().optional(),
  country: z.string().optional(),
  city: z.string().optional(),
  address: z.string().optional(),
  currency: z.string().min(1).default('RUB')
});

const companyUpdateSchema = companyCreateSchema.partial();

const projectMasterCreateSchema = z.object({
  companyId: z.string().optional(),
  name: z.string().min(1),
  country: z.string().min(1),
  city: z.string().min(1),
  address: z.string().min(1),
  contractScope: z.string().min(1),
  isActive: z.boolean().default(true)
});

const projectMasterUpdateSchema = projectMasterCreateSchema.partial();

const departmentMasterCreateSchema = z.object({
  companyId: z.string().optional(),
  name: z.string().min(1),
  code: z.string().min(1),
  description: z.string().optional(),
  isActive: z.boolean().default(true)
});

const departmentMasterUpdateSchema = departmentMasterCreateSchema.partial();

const contractorMasterCreateSchema = z.object({
  companyId: z.string().optional(),
  projectId: z.string().optional(),
  companyName: z.string().min(1),
  country: z.string().min(1),
  city: z.string().min(1),
  projectLocation: z.string().min(1),
  contractScope: z.string().min(1),
  hseWarningCount: z.number().int().min(0).default(0),
  hseWarningDate: z.string().datetime().optional(),
  fireWarningCount: z.number().int().min(0).default(0),
  fireWarningDate: z.string().datetime().optional(),
  environmentWarningCount: z.number().int().min(0).default(0),
  environmentWarningDate: z.string().datetime().optional(),
  penaltyCount: z.number().int().min(0).default(0),
  penaltyLegalClause: z.string().optional(),
  totalPenaltyAmount: z.number().int().min(0).default(0),
  isActive: z.boolean().default(true)
});

const contractorMasterUpdateSchema = contractorMasterCreateSchema.partial();

const ptwStatusSchema = z.enum([
  'Taslak',
  'Onay Bekliyor',
  'Aktif',
  'Askıya Alındı',
  'Tamamlandı',
  'İptal'
]);

const ptwTeamMemberSchema = z.object({
  adSoyad: z.string().min(1),
  gorevi: z.string().min(1),
  firma: z.string().min(1),
  egitimDurumu: z.string().min(1),
  imza: z.string().min(1)
});

const ptwPreWorkCheckSchema = z.object({
  kontrol: z.string().min(1),
  durum: z.enum(['Uygun', 'Uygun Değil', 'N/A']),
  aciklama: z.string().optional()
});

const ptwPrecautionSchema = z.object({
  tedbir: z.string().min(1),
  sorumlu: z.string().min(1),
  termin: z.string().datetime().optional(),
  durum: z.enum(['Açık', 'Devam Ediyor', 'Tamamlandı'])
});

const ptwDailyLogSchema = z.object({
  tarih: z.string().datetime(),
  saat: z.string().min(1),
  calismaBasladi: z.string().optional(),
  calismaBitti: z.string().optional(),
  aciklama: z.string().optional(),
  sorumlu: z.string().min(1)
});

const ptwTeamChangeSchema = z.object({
  eklenenPersonel: z.string().optional(),
  ayrilanPersonel: z.string().optional(),
  tarih: z.string().datetime(),
  onaylayan: z.string().min(1)
});

const ptwCreateSchema = z.object({
  permitNo: z.string().min(1),
  organization: z.string().min(1),
  department: z.string().min(1),
  project: z.string().min(1),
  location: z.string().min(1),
  permitType: z.string().min(1),
  issueDate: z.string().datetime(),
  validityStart: z.string().datetime(),
  validityEnd: z.string().datetime(),
  status: ptwStatusSchema.default('Taslak'),
  requesterName: z.string().min(1),
  issuerName: z.string().min(1),
  jobResponsibleName: z.string().min(1),
  siteResponsibleName: z.string().min(1),
  hseResponsibleName: z.string().min(1),
  authorizedApprover: z.string().min(1),
  jobTitle: z.string().min(1),
  jobDescription: z.string().min(1),
  workArea: z.string().min(1),
  plannedWork: z.string().min(1),
  workingConditions: z.string().min(1),
  workStartDate: z.string().datetime(),
  workStartTime: z.string().min(1),
  workEndDate: z.string().datetime(),
  workEndTime: z.string().min(1),
  hazards: z.array(z.string().min(1)).default([]),
  safetySystems: z.array(z.string().min(1)).default([]),
  equipmentUsed: z.array(z.string().min(1)).default([]),
  teamMembers: z.array(ptwTeamMemberSchema).default([]),
  preWorkChecks: z.array(ptwPreWorkCheckSchema).default([]),
  precautions: z.array(ptwPrecautionSchema).default([]),
  specialConditions: z.string().optional(),
  preparedBy: z.string().optional(),
  hseApprovalBy: z.string().optional(),
  projectManagerBy: z.string().optional(),
  employerRepBy: z.string().optional(),
  digitalSignature: z.string().optional(),
  approvalDate: z.string().datetime().optional(),
  dailyLogs: z.array(ptwDailyLogSchema).default([]),
  teamChanges: z.array(ptwTeamChangeSchema).default([]),
  workCompleted: z.boolean().default(false),
  areaSafe: z.boolean().default(false),
  materialsCollected: z.boolean().default(false),
  ptwClosed: z.boolean().default(false),
  closedBy: z.string().optional(),
  closedAt: z.string().datetime().optional(),
  closureRemarks: z.string().optional(),
  attachments: z
    .array(
      z.object({
        type: z.nativeEnum(PtwAttachmentType).default(PtwAttachmentType.OTHER),
        fileName: z.string().min(1),
        fileType: z.string().min(1),
        fileUrl: z.string().min(1)
      })
    )
    .default([])
});

const ptwUpdateSchema = ptwCreateSchema.partial();

const ptwCloseSchema = z.object({
  closedBy: z.string().min(1),
  closureRemarks: z.string().optional(),
  closedAt: z.string().datetime().optional(),
  workCompleted: z.boolean().default(true),
  areaSafe: z.boolean().default(true),
  materialsCollected: z.boolean().default(true),
  ptwClosed: z.boolean().default(true)
});

const ptwEmailSummarySchema = z.object({
  ptwId: z.string().min(1),
  to: z.string().email(),
  subject: z.string().optional(),
  message: z.string().optional()
});

const observationCreateSchema = z.object({
  observationType: z.string().min(1),
  location: z.string().min(1),
  details: z.string().min(1),
  ownerName: z.string().min(1),
  status: z.nativeEnum(OpenClosedStatus).optional(),
  closedAt: z.string().datetime().optional()
});

const riskCreateSchema = z.object({
  title: z.string().min(1),
  probability: z.number().int().min(1).max(5),
  impact: z.number().int().min(1).max(5),
  status: z.nativeEnum(OpenClosedStatus).optional(),
  actionOwner: z.string().optional(),
  dueDate: z.string().datetime().optional()
});

const trainingCreateSchema = z.object({
  employeeId: z.string().min(1),
  trainingName: z.string().min(1),
  trainingDate: z.string().datetime(),
  validUntil: z.string().datetime(),
  certificateNo: z.string().optional(),
  certificateUrl: z.string().optional()
});

const statusPatchSchema = z.object({
  status: z.nativeEnum(OpenClosedStatus)
});

const reportEmailSchema = z.object({
  to: z.string().email(),
  subject: z.string().min(1).optional(),
  message: z.string().optional(),
  format: z.enum(['PDF', 'XLSX']).default('PDF'),
  language: z.enum(['tr', 'en', 'ru']).default('tr')
});

const reportTexts: Record<
  LanguageCode,
  {
    reportTitle: string;
    generatedAt: string;
    latestInspections: string;
    metric: string;
    value: string;
    totalInspections: string;
    ptwOpen: string;
    ptwClosed: string;
    observationOpen: string;
    observationClosed: string;
    highRisk: string;
    criticalRisk: string;
    trainingCompliance: string;
    ppeUsage: string;
    emailSubject: string;
    emailBody: string;
  }
> = {
  tr: {
    reportTitle: 'HSE Haftalik/Aylik Yonetici Raporu',
    generatedAt: 'Uretim Tarihi',
    latestInspections: 'Son Inspection Kayitlari',
    metric: 'Metrik',
    value: 'Deger',
    totalInspections: 'Toplam Inspection',
    ptwOpen: 'PTW Acik',
    ptwClosed: 'PTW Kapali',
    observationOpen: 'Observation Acik',
    observationClosed: 'Observation Kapali',
    highRisk: 'Yuksek Risk Sayisi',
    criticalRisk: 'Kritik Risk Sayisi',
    trainingCompliance: 'Egitim Uyum %',
    ppeUsage: 'PPE Kullanim %',
    emailSubject: 'HSE Raporu',
    emailBody: 'HSE raporu ekte yer almaktadir.'
  },
  en: {
    reportTitle: 'HSE Weekly/Monthly Executive Report',
    generatedAt: 'Generated At',
    latestInspections: 'Latest Inspection Records',
    metric: 'Metric',
    value: 'Value',
    totalInspections: 'Total Inspections',
    ptwOpen: 'PTW Open',
    ptwClosed: 'PTW Closed',
    observationOpen: 'Observation Open',
    observationClosed: 'Observation Closed',
    highRisk: 'High Risk Count',
    criticalRisk: 'Critical Risk Count',
    trainingCompliance: 'Training Compliance %',
    ppeUsage: 'PPE Usage %',
    emailSubject: 'HSE Report',
    emailBody: 'Please find the attached HSE report.'
  },
  ru: {
    reportTitle: 'Еженедельный/ежемесячный отчет HSE для руководства',
    generatedAt: 'Дата формирования',
    latestInspections: 'Последние записи инспекций',
    metric: 'Показатель',
    value: 'Значение',
    totalInspections: 'Всего инспекций',
    ptwOpen: 'Открытые PTW',
    ptwClosed: 'Закрытые PTW',
    observationOpen: 'Открытые наблюдения',
    observationClosed: 'Закрытые наблюдения',
    highRisk: 'Количество высоких рисков',
    criticalRisk: 'Количество критических рисков',
    trainingCompliance: 'Соответствие обучению %',
    ppeUsage: 'Использование PPE %',
    emailSubject: 'Отчет HSE',
    emailBody: 'Отчет HSE приложен к письму.'
  }
};

const parseLanguage = (value: unknown): LanguageCode => {
  if (value === 'en' || value === 'ru' || value === 'tr') {
    return value;
  }
  return 'tr';
};

const mapPtwStatusToDb = (status: string): PtwStatus => {
  switch (status) {
    case 'Taslak':
      return PtwStatus.DRAFT;
    case 'Onay Bekliyor':
      return PtwStatus.PENDING_APPROVAL;
    case 'Aktif':
      return PtwStatus.ACTIVE;
    case 'Askıya Alındı':
      return PtwStatus.SUSPENDED;
    case 'Tamamlandı':
      return PtwStatus.COMPLETED;
    case 'İptal':
      return PtwStatus.CANCELLED;
    default:
      return PtwStatus.DRAFT;
  }
};

const mapPtwStatusFromDb = (status: PtwStatus): string => {
  switch (status) {
    case PtwStatus.DRAFT:
      return 'Taslak';
    case PtwStatus.PENDING_APPROVAL:
      return 'Onay Bekliyor';
    case PtwStatus.ACTIVE:
      return 'Aktif';
    case PtwStatus.SUSPENDED:
      return 'Askıya Alındı';
    case PtwStatus.COMPLETED:
      return 'Tamamlandı';
    case PtwStatus.CANCELLED:
      return 'İptal';
    default:
      return 'Taslak';
  }
};

const mapControlStatusToDb = (status: string): PtwControlStatus => {
  switch (status) {
    case 'Uygun':
      return PtwControlStatus.COMPLIANT;
    case 'Uygun Değil':
      return PtwControlStatus.NON_COMPLIANT;
    default:
      return PtwControlStatus.NA;
  }
};

const mapControlStatusFromDb = (status: PtwControlStatus): string => {
  switch (status) {
    case PtwControlStatus.COMPLIANT:
      return 'Uygun';
    case PtwControlStatus.NON_COMPLIANT:
      return 'Uygun Değil';
    default:
      return 'N/A';
  }
};

const mapActionStatusToDb = (status: string): PtwActionStatus => {
  switch (status) {
    case 'Açık':
      return PtwActionStatus.OPEN;
    case 'Devam Ediyor':
      return PtwActionStatus.IN_PROGRESS;
    default:
      return PtwActionStatus.CLOSED;
  }
};

const mapActionStatusFromDb = (status: PtwActionStatus): string => {
  switch (status) {
    case PtwActionStatus.OPEN:
      return 'Açık';
    case PtwActionStatus.IN_PROGRESS:
      return 'Devam Ediyor';
    default:
      return 'Tamamlandı';
  }
};

const riskLevelByScore = (score: number): RiskLevel => {
  if (score >= 20) return RiskLevel.CRITICAL;
  if (score >= 12) return RiskLevel.HIGH;
  if (score >= 6) return RiskLevel.MEDIUM;
  return RiskLevel.LOW;
};

const parseOrThrow = <T>(schema: z.ZodType<T>, body: unknown): T => {
  const parsed = schema.safeParse(body);
  if (!parsed.success) {
    throw new Error(parsed.error.issues.map((issue) => issue.message).join('; '));
  }
  return parsed.data;
};
const writeAuditLog = async (input: {
  companyId?: string | null;
  entityType: string;
  entityId?: string | null;
  action: string;
  actor?: string;
  beforeJson?: unknown;
  afterJson?: unknown;
}) => {
  await prisma.auditLog.create({
    data: {
      companyId: input.companyId ?? undefined,
      entityType: input.entityType,
      entityId: input.entityId ?? undefined,
      action: input.action,
      actor: input.actor ?? 'system',
      beforeJson: input.beforeJson === undefined ? undefined : JSON.parse(JSON.stringify(input.beforeJson)),
      afterJson: input.afterJson === undefined ? undefined : JSON.parse(JSON.stringify(input.afterJson))
    }
  });
};

const getDashboardSummary = async () => {
  const [
    inspectionTotal,
    ptwOpen,
    ptwClosed,
    observationOpen,
    observationClosed,
    highRiskCount,
    criticalRiskCount,
    trainingTotal,
    trainingValid
  ] = await Promise.all([
    prisma.inspection.count(),
    prisma.ptwRecord.count({
      where: {
        status: { in: [PtwStatus.DRAFT, PtwStatus.PENDING_APPROVAL, PtwStatus.ACTIVE, PtwStatus.SUSPENDED] }
      }
    }),
    prisma.ptwRecord.count({ where: { status: { in: [PtwStatus.COMPLETED, PtwStatus.CANCELLED] } } }),
    prisma.observation.count({ where: { status: OpenClosedStatus.OPEN } }),
    prisma.observation.count({ where: { status: OpenClosedStatus.CLOSED } }),
    prisma.riskRecord.count({ where: { riskLevel: RiskLevel.HIGH } }),
    prisma.riskRecord.count({ where: { riskLevel: RiskLevel.CRITICAL } }),
    prisma.training.count(),
    prisma.training.count({ where: { validUntil: { gte: new Date() } } })
  ]);

  const ppeAgg = await prisma.ppeTransaction.aggregate({
    _sum: {
      quantityUsed: true,
      quantityLeft: true
    }
  });

  const used = ppeAgg._sum.quantityUsed ?? 0;
  const left = ppeAgg._sum.quantityLeft ?? 0;
  const ppeUsageRate = used + left > 0 ? Math.round((used / (used + left)) * 100) : 0;
  const trainingComplianceRate =
    trainingTotal > 0 ? Math.round((trainingValid / trainingTotal) * 100) : 0;

  return {
    inspectionTotal,
    ptwOpen,
    ptwClosed,
    observationOpen,
    observationClosed,
    highRiskCount,
    criticalRiskCount,
    trainingComplianceRate,
    ppeUsageRate
  };
};

const buildOverviewPdf = async (language: LanguageCode) => {
  const text = reportTexts[language];
  const summary = await getDashboardSummary();
  const inspections = await prisma.inspection.findMany({
    include: { employee: true },
    orderBy: { createdAt: 'desc' },
    take: 20
  });

  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([842, 595]);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  page.drawText(text.reportTitle, {
    x: 40,
    y: 560,
    size: 20,
    font: bold,
    color: rgb(0.05, 0.15, 0.35)
  });

  page.drawText(`${text.generatedAt}: ${new Date().toISOString()}`, {
    x: 40,
    y: 540,
    size: 10,
    font,
    color: rgb(0.3, 0.3, 0.3)
  });

  const summaryRows = [
    `${text.totalInspections}: ${summary.inspectionTotal}`,
    `${text.ptwOpen}/${text.ptwClosed}: ${summary.ptwOpen}/${summary.ptwClosed}`,
    `${text.observationOpen}/${text.observationClosed}: ${summary.observationOpen}/${summary.observationClosed}`,
    `${text.highRisk}/${text.criticalRisk}: ${summary.highRiskCount}/${summary.criticalRiskCount}`,
    `${text.trainingCompliance}: ${summary.trainingComplianceRate}%`,
    `${text.ppeUsage}: ${summary.ppeUsageRate}%`
  ];

  let y = 505;
  summaryRows.forEach((row) => {
    page.drawText(row, { x: 40, y, size: 12, font });
    y -= 18;
  });

  page.drawText(text.latestInspections, {
    x: 40,
    y: y - 12,
    size: 13,
    font: bold
  });
  y -= 36;

  const inspectionLines = inspections.map(
    (item) =>
      `${item.createdAt.toISOString().slice(0, 10)} | ${item.employee.firstName} ${item.employee.lastName} | ${item.region} | ${item.workNature} | ${item.riskLevel} | ${item.status}`
  );

  inspectionLines.slice(0, 18).forEach((line) => {
    page.drawText(line, { x: 40, y, size: 9, font });
    y -= 14;
  });

  return Buffer.from(await pdfDoc.save());
};

const buildOverviewWorkbook = async (language: LanguageCode) => {
  const text = reportTexts[language];
  const summary = await getDashboardSummary();
  const inspections = await prisma.inspection.findMany({
    include: { employee: true, attachments: true },
    orderBy: { createdAt: 'desc' }
  });

  const wb = XLSX.utils.book_new();

  const summarySheetData = [
    [text.metric, text.value],
    [text.totalInspections, summary.inspectionTotal],
    [text.ptwOpen, summary.ptwOpen],
    [text.ptwClosed, summary.ptwClosed],
    [text.observationOpen, summary.observationOpen],
    [text.observationClosed, summary.observationClosed],
    [text.highRisk, summary.highRiskCount],
    [text.criticalRisk, summary.criticalRiskCount],
    [text.trainingCompliance, summary.trainingComplianceRate],
    [text.ppeUsage, summary.ppeUsageRate]
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summarySheetData), 'Summary');

  const inspectionRows = inspections.map((item) => ({
    inspectionDate: item.createdAt.toISOString(),
    employeeName: `${item.employee.firstName} ${item.employee.lastName}`,
    contractor: item.contractor,
    region: item.region,
    workNature: item.workNature,
    findings: item.findings ?? '',
    riskLevel: item.riskLevel,
    status: item.status,
    attachmentCount: item.attachments.length
  }));
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(inspectionRows),
    'Inspections'
  );

  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
};

const buildMailTransport = () => {
  const host = process.env.SMTP_HOST;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;
  const port = Number(process.env.SMTP_PORT ?? 587);

  if (host && user && pass) {
    return nodemailer.createTransport({
      host,
      port,
      secure: false,
      auth: { user, pass }
    });
  }

  return nodemailer.createTransport({ jsonTransport: true });
};

const ptwIncludeShape = {
  hazards: true,
  safetySystems: true,
  equipmentUsed: true,
  teamMembers: true,
  preWorkChecks: true,
  precautions: true,
  dailyLogs: true,
  teamChanges: true,
  attachments: true
} as const;

const serializePtw = (ptw: any) => ({
  id: ptw.id,
  permitNo: ptw.permitNo,
  organization: ptw.organization,
  department: ptw.department,
  project: ptw.project,
  location: ptw.location,
  permitType: ptw.permitType,
  issueDate: ptw.issueDate,
  validityStart: ptw.validityStart,
  validityEnd: ptw.validityEnd,
  status: mapPtwStatusFromDb(ptw.status),
  requesterName: ptw.requesterName,
  issuerName: ptw.issuerName,
  jobResponsibleName: ptw.jobResponsibleName,
  siteResponsibleName: ptw.siteResponsibleName,
  hseResponsibleName: ptw.hseResponsibleName,
  authorizedApprover: ptw.authorizedApprover,
  jobTitle: ptw.jobTitle,
  jobDescription: ptw.jobDescription,
  workArea: ptw.workArea,
  plannedWork: ptw.plannedWork,
  workingConditions: ptw.workingConditions,
  workStartDate: ptw.workStartDate,
  workStartTime: ptw.workStartTime,
  workEndDate: ptw.workEndDate,
  workEndTime: ptw.workEndTime,
  hazards: ptw.hazards.map((h: any) => h.name),
  safetySystems: ptw.safetySystems.map((s: any) => s.name),
  equipmentUsed: ptw.equipmentUsed.map((e: any) => e.name),
  teamMembers: ptw.teamMembers.map((m: any) => ({
    id: m.id,
    adSoyad: m.fullName,
    gorevi: m.duty,
    firma: m.company,
    egitimDurumu: m.trainingStatus,
    imza: m.signature
  })),
  preWorkChecks: ptw.preWorkChecks.map((c: any) => ({
    id: c.id,
    kontrol: c.checkItem,
    durum: mapControlStatusFromDb(c.status),
    aciklama: c.note
  })),
  precautions: ptw.precautions.map((p: any) => ({
    id: p.id,
    tedbir: p.measure,
    sorumlu: p.responsible,
    termin: p.dueDate,
    durum: mapActionStatusFromDb(p.status)
  })),
  specialConditions: ptw.specialConditions,
  preparedBy: ptw.preparedBy,
  hseApprovalBy: ptw.hseApprovalBy,
  projectManagerBy: ptw.projectManagerBy,
  employerRepBy: ptw.employerRepBy,
  digitalSignature: ptw.digitalSignature,
  approvalDate: ptw.approvalDate,
  dailyLogs: ptw.dailyLogs.map((d: any) => ({
    id: d.id,
    tarih: d.logDate,
    saat: d.logTime,
    calismaBasladi: d.workStarted,
    calismaBitti: d.workEnded,
    aciklama: d.note,
    sorumlu: d.responsible
  })),
  teamChanges: ptw.teamChanges.map((t: any) => ({
    id: t.id,
    eklenenPersonel: t.addedPersonnel,
    ayrilanPersonel: t.removedPersonnel,
    tarih: t.changeDate,
    onaylayan: t.approvedBy
  })),
  workCompleted: ptw.workCompleted,
  areaSafe: ptw.areaSafe,
  materialsCollected: ptw.materialsCollected,
  ptwClosed: ptw.ptwClosed,
  closedBy: ptw.closedBy,
  closedAt: ptw.closedAt,
  closureRemarks: ptw.closureRemarks,
  attachments: ptw.attachments.map((a: any) => ({
    id: a.id,
    type: a.type,
    fileName: a.fileName,
    fileType: a.fileType,
    fileUrl: a.fileUrl
  })),
  createdAt: ptw.createdAt,
  updatedAt: ptw.updatedAt
});

const buildPtwPdf = async (ptw: any) => {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([842, 595]);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  page.drawText(`PTW Ozet Raporu - ${ptw.permitNo}`, {
    x: 40,
    y: 560,
    size: 18,
    font: bold,
    color: rgb(0.05, 0.18, 0.35)
  });

  const lines = [
    `Organizasyon: ${ptw.organization}`,
    `Departman/Proje: ${ptw.department} / ${ptw.project}`,
    `Lokasyon: ${ptw.location}`,
    `PTW Turu: ${ptw.permitType}`,
    `Durum: ${mapPtwStatusFromDb(ptw.status)}`,
    `Gecerlilik: ${ptw.validityStart.toISOString().slice(0, 10)} - ${ptw.validityEnd.toISOString().slice(0, 10)}`,
    `Sorumlular: Talep Eden ${ptw.requesterName} | Veren ${ptw.issuerName} | HSE ${ptw.hseResponsibleName}`,
    `Is: ${ptw.jobTitle}`,
    `Planlanan Is: ${ptw.plannedWork}`,
    `Riskler: ${ptw.hazards.map((h: any) => h.name).join(', ') || '-'}`,
    `Guvenlik Sistemleri: ${ptw.safetySystems.map((s: any) => s.name).join(', ') || '-'}`,
    `Ekip Sayisi: ${ptw.teamMembers.length}`,
    `On Kontrol Sayisi: ${ptw.preWorkChecks.length}`,
    `Tedbir Sayisi: ${ptw.precautions.length}`,
    `Ek Sayisi: ${ptw.attachments.length}`
  ];

  let y = 528;
  lines.forEach((line) => {
    page.drawText(line, { x: 40, y, size: 10, font });
    y -= 16;
  });

  return Buffer.from(await pdfDoc.save());
};

const buildPtwDocx = async (ptw: any) => {
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({ text: `PTW Ozet Raporu - ${ptw.permitNo}`, heading: HeadingLevel.HEADING_1 }),
          new Paragraph({ children: [new TextRun(`Organizasyon: ${ptw.organization}`)] }),
          new Paragraph({ children: [new TextRun(`Departman: ${ptw.department}`)] }),
          new Paragraph({ children: [new TextRun(`Proje: ${ptw.project}`)] }),
          new Paragraph({ children: [new TextRun(`Lokasyon: ${ptw.location}`)] }),
          new Paragraph({ children: [new TextRun(`PTW Turu: ${ptw.permitType}`)] }),
          new Paragraph({ children: [new TextRun(`Durum: ${mapPtwStatusFromDb(ptw.status)}`)] }),
          new Paragraph({ children: [new TextRun(`Is Basligi: ${ptw.jobTitle}`)] }),
          new Paragraph({ children: [new TextRun(`Is Aciklamasi: ${ptw.jobDescription}`)] }),
          new Paragraph({ text: 'Ekip Bilgileri', heading: HeadingLevel.HEADING_2 }),
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph('Ad Soyad')] }),
                  new TableCell({ children: [new Paragraph('Gorev')] }),
                  new TableCell({ children: [new Paragraph('Firma')] }),
                  new TableCell({ children: [new Paragraph('Egitim Durumu')] })
                ]
              }),
              ...ptw.teamMembers.map(
                (m: any) =>
                  new TableRow({
                    children: [
                      new TableCell({ children: [new Paragraph(m.fullName)] }),
                      new TableCell({ children: [new Paragraph(m.duty)] }),
                      new TableCell({ children: [new Paragraph(m.company)] }),
                      new TableCell({ children: [new Paragraph(m.trainingStatus)] })
                    ]
                  })
              )
            ]
          })
        ]
      }
    ]
  });

  return Buffer.from(await Packer.toBuffer(doc));
};

const buildPtwCsv = (ptw: any, full: boolean) => {
  const wb = XLSX.utils.book_new();
  const summaryRows = [
    { alan: 'PTW No', deger: ptw.permitNo },
    { alan: 'Organizasyon', deger: ptw.organization },
    { alan: 'Departman', deger: ptw.department },
    { alan: 'Proje', deger: ptw.project },
    { alan: 'Lokasyon', deger: ptw.location },
    { alan: 'Durum', deger: mapPtwStatusFromDb(ptw.status) }
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), 'Summary');

  const actionRows = ptw.precautions.map((p: any) => ({
    tedbir: p.measure,
    sorumlu: p.responsible,
    termin: p.dueDate ? p.dueDate.toISOString() : '',
    durum: mapActionStatusFromDb(p.status)
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(actionRows), 'Actions');

  if (full) {
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        ptw.preWorkChecks.map((c: any) => ({
          kontrol: c.checkItem,
          durum: mapControlStatusFromDb(c.status),
          aciklama: c.note ?? ''
        }))
      ),
      'Checks'
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(
        ptw.dailyLogs.map((d: any) => ({
          tarih: d.logDate.toISOString(),
          saat: d.logTime,
          basladi: d.workStarted ?? '',
          bitti: d.workEnded ?? '',
          aciklama: d.note ?? '',
          sorumlu: d.responsible
        }))
      ),
      'DailyLogs'
    );
  }

  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'csv' }));
};
app.get('/health', (_req, res) => {
  res.json({ ok: true, service: 'hse-api' });
});

app.post('/api/uploads/inspection-attachment', upload.single('file'), (req, res) => {
  if (!req.file) {
    res.status(400).json({ message: 'File is required.' });
    return;
  }

  res.status(201).json({
    fileName: req.file.originalname,
    fileType: req.file.mimetype,
    fileUrl: `/uploads/${req.file.filename}`,
    fileSize: req.file.size
  });
});

app.get('/api/companies', async (_req, res) => {
  const rows = await prisma.company.findMany({ orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/companies', async (req, res) => {
  try {
    const input = parseOrThrow(companyCreateSchema, req.body);
    const row = await prisma.company.create({ data: input });
    await writeAuditLog({ companyId: row.id, entityType: 'Company', entityId: row.id, action: 'CREATE', afterJson: row });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.put('/api/companies/:id', async (req, res) => {
  try {
    const input = parseOrThrow(companyUpdateSchema, req.body);
    const before = await prisma.company.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Company not found.' });
      return;
    }
    const row = await prisma.company.update({ where: { id: req.params.id }, data: input });
    await writeAuditLog({ companyId: row.id, entityType: 'Company', entityId: row.id, action: 'UPDATE', beforeJson: before, afterJson: row });
    res.json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.delete('/api/companies/:id', async (req, res) => {
  try {
    const before = await prisma.company.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Company not found.' });
      return;
    }
    await writeAuditLog({ companyId: before.id, entityType: 'Company', entityId: before.id, action: 'DELETE', beforeJson: before });
    await prisma.company.delete({ where: { id: req.params.id } });
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/master/projects', async (_req, res) => {
  const rows = await prisma.projectMaster.findMany({ include: { company: true }, orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/master/projects', async (req, res) => {
  try {
    const input = parseOrThrow(projectMasterCreateSchema, req.body);
    const row = await prisma.projectMaster.create({ data: input, include: { company: true } });
    await writeAuditLog({ companyId: row.companyId, entityType: 'ProjectMaster', entityId: row.id, action: 'CREATE', afterJson: row });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.put('/api/master/projects/:id', async (req, res) => {
  try {
    const input = parseOrThrow(projectMasterUpdateSchema, req.body);
    const before = await prisma.projectMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Project not found.' });
      return;
    }
    const row = await prisma.projectMaster.update({ where: { id: req.params.id }, data: input, include: { company: true } });
    await writeAuditLog({ companyId: row.companyId, entityType: 'ProjectMaster', entityId: row.id, action: 'UPDATE', beforeJson: before, afterJson: row });
    res.json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.delete('/api/master/projects/:id', async (req, res) => {
  try {
    const before = await prisma.projectMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Project not found.' });
      return;
    }
    await prisma.projectMaster.delete({ where: { id: req.params.id } });
    await writeAuditLog({ companyId: before.companyId, entityType: 'ProjectMaster', entityId: before.id, action: 'DELETE', beforeJson: before });
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/master/departments', async (_req, res) => {
  const rows = await prisma.departmentMaster.findMany({ include: { company: true }, orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/master/departments', async (req, res) => {
  try {
    const input = parseOrThrow(departmentMasterCreateSchema, req.body);
    const row = await prisma.departmentMaster.create({ data: input, include: { company: true } });
    await writeAuditLog({ companyId: row.companyId, entityType: 'DepartmentMaster', entityId: row.id, action: 'CREATE', afterJson: row });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.put('/api/master/departments/:id', async (req, res) => {
  try {
    const input = parseOrThrow(departmentMasterUpdateSchema, req.body);
    const before = await prisma.departmentMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Department not found.' });
      return;
    }
    const row = await prisma.departmentMaster.update({ where: { id: req.params.id }, data: input, include: { company: true } });
    await writeAuditLog({ companyId: row.companyId, entityType: 'DepartmentMaster', entityId: row.id, action: 'UPDATE', beforeJson: before, afterJson: row });
    res.json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.delete('/api/master/departments/:id', async (req, res) => {
  try {
    const before = await prisma.departmentMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Department not found.' });
      return;
    }
    await prisma.departmentMaster.delete({ where: { id: req.params.id } });
    await writeAuditLog({ companyId: before.companyId, entityType: 'DepartmentMaster', entityId: before.id, action: 'DELETE', beforeJson: before });
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/master/contractors', async (_req, res) => {
  const rows = await prisma.contractorMaster.findMany({ include: { company: true, project: true }, orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/master/contractors', async (req, res) => {
  try {
    const input = parseOrThrow(contractorMasterCreateSchema, req.body);
    const row = await prisma.contractorMaster.create({
      data: {
        ...input,
        hseWarningDate: input.hseWarningDate ? new Date(input.hseWarningDate) : undefined,
        fireWarningDate: input.fireWarningDate ? new Date(input.fireWarningDate) : undefined,
        environmentWarningDate: input.environmentWarningDate ? new Date(input.environmentWarningDate) : undefined
      },
      include: { company: true, project: true }
    });
    await writeAuditLog({ companyId: row.companyId, entityType: 'ContractorMaster', entityId: row.id, action: 'CREATE', afterJson: row });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.put('/api/master/contractors/:id', async (req, res) => {
  try {
    const input = parseOrThrow(contractorMasterUpdateSchema, req.body);
    const before = await prisma.contractorMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Contractor not found.' });
      return;
    }
    const row = await prisma.contractorMaster.update({
      where: { id: req.params.id },
      data: {
        ...input,
        hseWarningDate: input.hseWarningDate ? new Date(input.hseWarningDate) : undefined,
        fireWarningDate: input.fireWarningDate ? new Date(input.fireWarningDate) : undefined,
        environmentWarningDate: input.environmentWarningDate ? new Date(input.environmentWarningDate) : undefined
      },
      include: { company: true, project: true }
    });
    await writeAuditLog({ companyId: row.companyId, entityType: 'ContractorMaster', entityId: row.id, action: 'UPDATE', beforeJson: before, afterJson: row });
    res.json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.delete('/api/master/contractors/:id', async (req, res) => {
  try {
    const before = await prisma.contractorMaster.findUnique({ where: { id: req.params.id } });
    if (!before) {
      res.status(404).json({ message: 'Contractor not found.' });
      return;
    }
    await prisma.contractorMaster.delete({ where: { id: req.params.id } });
    await writeAuditLog({ companyId: before.companyId, entityType: 'ContractorMaster', entityId: before.id, action: 'DELETE', beforeJson: before });
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/audit-logs', async (_req, res) => {
  const rows = await prisma.auditLog.findMany({ include: { company: true }, orderBy: { createdAt: 'desc' }, take: 500 });
  res.json(rows);
});

app.get('/api/employees', async (_req, res) => {
  const rows = await prisma.employee.findMany({ orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/employees', async (req, res) => {
  try {
    const input = parseOrThrow(employeeCreateSchema, req.body);
    const row = await prisma.employee.create({ data: input });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/inspections', async (_req, res) => {
  const rows = await prisma.inspection.findMany({
    include: { employee: true, attachments: true },
    orderBy: { createdAt: 'desc' }
  });
  res.json(rows);
});

app.post('/api/inspections', async (req, res) => {
  try {
    const input = parseOrThrow(inspectionCreateSchema, req.body);
    const row = await prisma.inspection.create({
      data: {
        employeeId: input.employeeId,
        contractor: input.contractor,
        region: input.region,
        workNature: input.workNature,
        findings: input.findings,
        riskLevel: input.riskLevel,
        correctiveAction: input.correctiveAction,
        dueDate: input.dueDate ? new Date(input.dueDate) : undefined,
        status: input.status ?? OpenClosedStatus.OPEN,
        attachments: input.attachments
          ? {
              createMany: {
                data: input.attachments
              }
            }
          : undefined
      },
      include: { attachments: true }
    });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/ppe/items', async (_req, res) => {
  const rows = await prisma.ppeItem.findMany({
    include: { transactions: { orderBy: { createdAt: 'desc' } } },
    orderBy: { createdAt: 'desc' }
  });
  res.json(rows);
});

app.post('/api/ppe/items', async (req, res) => {
  try {
    const input = parseOrThrow(ppeItemCreateSchema, req.body);
    const row = await prisma.ppeItem.create({ data: input });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.post('/api/ppe/transactions', async (req, res) => {
  try {
    const input = parseOrThrow(ppeTransactionCreateSchema, req.body);
    const quantityIn = input.quantityIn ?? 0;
    const quantityUsed = input.quantityUsed ?? 0;
    const latest = await prisma.ppeTransaction.findFirst({
      where: { itemId: input.itemId },
      orderBy: { createdAt: 'desc' }
    });
    const previousLeft = latest?.quantityLeft ?? 0;
    const quantityLeft = previousLeft + quantityIn - quantityUsed;
    const row = await prisma.ppeTransaction.create({
      data: {
        itemId: input.itemId,
        quantityIn,
        quantityUsed,
        quantityLeft,
        note: input.note
      }
    });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.post('/api/uploads/ptw-attachment', upload.single('file'), (req, res) => {
  if (!req.file) {
    res.status(400).json({ message: 'File is required.' });
    return;
  }

  const typeRaw = String(req.body.type ?? 'OTHER').toUpperCase();
  const type = Object.values(PtwAttachmentType).includes(typeRaw as PtwAttachmentType)
    ? (typeRaw as PtwAttachmentType)
    : PtwAttachmentType.OTHER;

  res.status(201).json({
    type,
    fileName: req.file.originalname,
    fileType: req.file.mimetype,
    fileUrl: `/uploads/${req.file.filename}`,
    fileSize: req.file.size
  });
});

app.get('/api/ptw', async (_req, res) => {
  const rows = await prisma.ptwRecord.findMany({
    include: ptwIncludeShape,
    orderBy: { createdAt: 'desc' }
  });
  res.json(rows.map(serializePtw));
});

app.get('/api/ptw/:id', async (req, res) => {
  const row = await prisma.ptwRecord.findUnique({
    where: { id: req.params.id },
    include: ptwIncludeShape
  });
  if (!row) {
    res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
    return;
  }
  res.json(serializePtw(row));
});

app.post('/api/ptw', async (req, res) => {
  try {
    const input = parseOrThrow(ptwCreateSchema, req.body);
    const row = await prisma.ptwRecord.create({
      data: {
        permitNo: input.permitNo,
        organization: input.organization,
        department: input.department,
        project: input.project,
        location: input.location,
        permitType: input.permitType,
        issueDate: new Date(input.issueDate),
        validityStart: new Date(input.validityStart),
        validityEnd: new Date(input.validityEnd),
        status: mapPtwStatusToDb(input.status ?? 'Taslak'),
        requesterName: input.requesterName,
        issuerName: input.issuerName,
        jobResponsibleName: input.jobResponsibleName,
        siteResponsibleName: input.siteResponsibleName,
        hseResponsibleName: input.hseResponsibleName,
        authorizedApprover: input.authorizedApprover,
        jobTitle: input.jobTitle,
        jobDescription: input.jobDescription,
        workArea: input.workArea,
        plannedWork: input.plannedWork,
        workingConditions: input.workingConditions,
        workStartDate: new Date(input.workStartDate),
        workStartTime: input.workStartTime,
        workEndDate: new Date(input.workEndDate),
        workEndTime: input.workEndTime,
        specialConditions: input.specialConditions,
        preparedBy: input.preparedBy,
        hseApprovalBy: input.hseApprovalBy,
        projectManagerBy: input.projectManagerBy,
        employerRepBy: input.employerRepBy,
        digitalSignature: input.digitalSignature,
        approvalDate: input.approvalDate ? new Date(input.approvalDate) : undefined,
        workCompleted: input.workCompleted,
        areaSafe: input.areaSafe,
        materialsCollected: input.materialsCollected,
        ptwClosed: input.ptwClosed,
        closedBy: input.closedBy,
        closedAt: input.closedAt ? new Date(input.closedAt) : undefined,
        closureRemarks: input.closureRemarks,
        hazards: { createMany: { data: (input.hazards ?? []).map((name) => ({ name })) } },
        safetySystems: { createMany: { data: (input.safetySystems ?? []).map((name) => ({ name })) } },
        equipmentUsed: { createMany: { data: (input.equipmentUsed ?? []).map((name) => ({ name })) } },
        teamMembers: {
          createMany: {
            data: (input.teamMembers ?? []).map((member) => ({
              fullName: member.adSoyad,
              duty: member.gorevi,
              company: member.firma,
              trainingStatus: member.egitimDurumu,
              signature: member.imza
            }))
          }
        },
        preWorkChecks: {
          createMany: {
            data: (input.preWorkChecks ?? []).map((check) => ({
              checkItem: check.kontrol,
              status: mapControlStatusToDb(check.durum),
              note: check.aciklama
            }))
          }
        },
        precautions: {
          createMany: {
            data: (input.precautions ?? []).map((item) => ({
              measure: item.tedbir,
              responsible: item.sorumlu,
              dueDate: item.termin ? new Date(item.termin) : undefined,
              status: mapActionStatusToDb(item.durum)
            }))
          }
        },
        dailyLogs: {
          createMany: {
            data: (input.dailyLogs ?? []).map((log) => ({
              logDate: new Date(log.tarih),
              logTime: log.saat,
              workStarted: log.calismaBasladi,
              workEnded: log.calismaBitti,
              note: log.aciklama,
              responsible: log.sorumlu
            }))
          }
        },
        teamChanges: {
          createMany: {
            data: (input.teamChanges ?? []).map((change) => ({
              addedPersonnel: change.eklenenPersonel,
              removedPersonnel: change.ayrilanPersonel,
              changeDate: new Date(change.tarih),
              approvedBy: change.onaylayan
            }))
          }
        },
        attachments: {
          createMany: {
            data: (input.attachments ?? []).map((file) => ({
              type: file.type,
              fileName: file.fileName,
              fileType: file.fileType,
              fileUrl: file.fileUrl
            }))
          }
        }
      },
      include: ptwIncludeShape
    });
    res.status(201).json(serializePtw(row));
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.put('/api/ptw/:id', async (req, res) => {
  try {
    const input = parseOrThrow(ptwUpdateSchema, req.body);

    const row = await prisma.$transaction(async (tx) => {
      await tx.ptwHazard.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwSafetySystem.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwEquipment.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwTeamMember.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwPreWorkCheck.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwPrecaution.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwDailyLog.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwTeamChange.deleteMany({ where: { ptwId: req.params.id } });
      await tx.ptwAttachment.deleteMany({ where: { ptwId: req.params.id } });

      return tx.ptwRecord.update({
        where: { id: req.params.id },
        data: {
          permitNo: input.permitNo,
          organization: input.organization,
          department: input.department,
          project: input.project,
          location: input.location,
          permitType: input.permitType,
          issueDate: input.issueDate ? new Date(input.issueDate) : undefined,
          validityStart: input.validityStart ? new Date(input.validityStart) : undefined,
          validityEnd: input.validityEnd ? new Date(input.validityEnd) : undefined,
          status: input.status ? mapPtwStatusToDb(input.status) : undefined,
          requesterName: input.requesterName,
          issuerName: input.issuerName,
          jobResponsibleName: input.jobResponsibleName,
          siteResponsibleName: input.siteResponsibleName,
          hseResponsibleName: input.hseResponsibleName,
          authorizedApprover: input.authorizedApprover,
          jobTitle: input.jobTitle,
          jobDescription: input.jobDescription,
          workArea: input.workArea,
          plannedWork: input.plannedWork,
          workingConditions: input.workingConditions,
          workStartDate: input.workStartDate ? new Date(input.workStartDate) : undefined,
          workStartTime: input.workStartTime,
          workEndDate: input.workEndDate ? new Date(input.workEndDate) : undefined,
          workEndTime: input.workEndTime,
          specialConditions: input.specialConditions,
          preparedBy: input.preparedBy,
          hseApprovalBy: input.hseApprovalBy,
          projectManagerBy: input.projectManagerBy,
          employerRepBy: input.employerRepBy,
          digitalSignature: input.digitalSignature,
          approvalDate: input.approvalDate ? new Date(input.approvalDate) : undefined,
          workCompleted: input.workCompleted,
          areaSafe: input.areaSafe,
          materialsCollected: input.materialsCollected,
          ptwClosed: input.ptwClosed,
          closedBy: input.closedBy,
          closedAt: input.closedAt ? new Date(input.closedAt) : undefined,
          closureRemarks: input.closureRemarks,
          hazards: input.hazards ? { createMany: { data: input.hazards.map((name) => ({ name })) } } : undefined,
          safetySystems: input.safetySystems
            ? { createMany: { data: input.safetySystems.map((name) => ({ name })) } }
            : undefined,
          equipmentUsed: input.equipmentUsed
            ? { createMany: { data: input.equipmentUsed.map((name) => ({ name })) } }
            : undefined,
          teamMembers: input.teamMembers
            ? {
                createMany: {
                  data: input.teamMembers.map((member) => ({
                    fullName: member.adSoyad,
                    duty: member.gorevi,
                    company: member.firma,
                    trainingStatus: member.egitimDurumu,
                    signature: member.imza
                  }))
                }
              }
            : undefined,
          preWorkChecks: input.preWorkChecks
            ? {
                createMany: {
                  data: input.preWorkChecks.map((check) => ({
                    checkItem: check.kontrol,
                    status: mapControlStatusToDb(check.durum),
                    note: check.aciklama
                  }))
                }
              }
            : undefined,
          precautions: input.precautions
            ? {
                createMany: {
                  data: input.precautions.map((item) => ({
                    measure: item.tedbir,
                    responsible: item.sorumlu,
                    dueDate: item.termin ? new Date(item.termin) : undefined,
                    status: mapActionStatusToDb(item.durum)
                  }))
                }
              }
            : undefined,
          dailyLogs: input.dailyLogs
            ? {
                createMany: {
                  data: input.dailyLogs.map((log) => ({
                    logDate: new Date(log.tarih),
                    logTime: log.saat,
                    workStarted: log.calismaBasladi,
                    workEnded: log.calismaBitti,
                    note: log.aciklama,
                    responsible: log.sorumlu
                  }))
                }
              }
            : undefined,
          teamChanges: input.teamChanges
            ? {
                createMany: {
                  data: input.teamChanges.map((change) => ({
                    addedPersonnel: change.eklenenPersonel,
                    removedPersonnel: change.ayrilanPersonel,
                    changeDate: new Date(change.tarih),
                    approvedBy: change.onaylayan
                  }))
                }
              }
            : undefined,
          attachments: input.attachments
            ? {
                createMany: {
                  data: input.attachments.map((file) => ({
                    type: file.type,
                    fileName: file.fileName,
                    fileType: file.fileType,
                    fileUrl: file.fileUrl
                  }))
                }
              }
            : undefined
        },
        include: ptwIncludeShape
      });
    });

    res.json(serializePtw(row));
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.patch('/api/ptw/:id/status', async (req, res) => {
  try {
    const payload = parseOrThrow(z.object({ status: ptwStatusSchema }), req.body);
    const row = await prisma.ptwRecord.update({
      where: { id: req.params.id },
      data: { status: mapPtwStatusToDb(payload.status) },
      include: ptwIncludeShape
    });
    res.json(serializePtw(row));
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.post('/api/ptw/:id/close', async (req, res) => {
  try {
    const input = parseOrThrow(ptwCloseSchema, req.body);
    const row = await prisma.ptwRecord.update({
      where: { id: req.params.id },
      data: {
        status: PtwStatus.COMPLETED,
        ptwClosed: input.ptwClosed,
        workCompleted: input.workCompleted,
        areaSafe: input.areaSafe,
        materialsCollected: input.materialsCollected,
        closedBy: input.closedBy,
        closedAt: input.closedAt ? new Date(input.closedAt) : new Date(),
        closureRemarks: input.closureRemarks
      },
      include: ptwIncludeShape
    });
    res.json(serializePtw(row));
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/ptw/:id/export/pdf', async (req, res) => {
  try {
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: req.params.id },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const buffer = await buildPtwPdf(ptw);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${ptw.permitNo}.pdf"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.get('/api/ptw/:id/export/docx', async (req, res) => {
  try {
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: req.params.id },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const buffer = await buildPtwDocx(ptw);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${ptw.permitNo}.docx"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.get('/api/ptw/:id/export/action-csv', async (req, res) => {
  try {
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: req.params.id },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const buffer = buildPtwCsv(ptw, false);
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', `attachment; filename="${ptw.permitNo}-actions.csv"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.get('/api/ptw/:id/export/full-csv', async (req, res) => {
  try {
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: req.params.id },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const buffer = buildPtwCsv(ptw, true);
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', `attachment; filename="${ptw.permitNo}-full.csv"`);
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.get('/api/ptw/:id/print', async (req, res) => {
  try {
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: req.params.id },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const buffer = await buildPtwPdf(ptw);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.post('/api/ptw/email-summary', async (req, res) => {
  try {
    const input = parseOrThrow(ptwEmailSummarySchema, req.body);
    const ptw = await prisma.ptwRecord.findUnique({
      where: { id: input.ptwId },
      include: ptwIncludeShape
    });
    if (!ptw) {
      res.status(404).json({ message: 'PTW kaydi bulunamadi.' });
      return;
    }

    const transporter = buildMailTransport();
    const fromAddress = process.env.REPORT_FROM_EMAIL ?? 'hse-platform@local.test';
    const attachment = await buildPtwPdf(ptw);

    const info = await transporter.sendMail({
      from: fromAddress,
      to: input.to,
      subject: input.subject ?? `PTW Ozet - ${ptw.permitNo}`,
      text:
        input.message ??
        `PTW No: ${ptw.permitNo}\nDurum: ${mapPtwStatusFromDb(ptw.status)}\nProje: ${ptw.project}\nLokasyon: ${ptw.location}`,
      attachments: [
        {
          filename: `${ptw.permitNo}.pdf`,
          content: attachment,
          contentType: 'application/pdf'
        }
      ]
    });

    res.json({ accepted: true, messageId: info.messageId });
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/observations', async (_req, res) => {
  const rows = await prisma.observation.findMany({ orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/observations', async (req, res) => {
  try {
    const input = parseOrThrow(observationCreateSchema, req.body);
    const row = await prisma.observation.create({
      data: {
        observationType: input.observationType,
        location: input.location,
        details: input.details,
        ownerName: input.ownerName,
        status: input.status ?? OpenClosedStatus.OPEN,
        closedAt: input.closedAt ? new Date(input.closedAt) : undefined
      }
    });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.patch('/api/observations/:id/status', async (req, res) => {
  try {
    const input = parseOrThrow(statusPatchSchema, req.body);
    const row = await prisma.observation.update({
      where: { id: req.params.id },
      data: {
        status: input.status,
        closedAt: input.status === OpenClosedStatus.CLOSED ? new Date() : null
      }
    });
    res.json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/risks', async (_req, res) => {
  const rows = await prisma.riskRecord.findMany({ orderBy: { createdAt: 'desc' } });
  res.json(rows);
});

app.post('/api/risks', async (req, res) => {
  try {
    const input = parseOrThrow(riskCreateSchema, req.body);
    const score = input.probability * input.impact;
    const row = await prisma.riskRecord.create({
      data: {
        title: input.title,
        probability: input.probability,
        impact: input.impact,
        score,
        riskLevel: riskLevelByScore(score),
        status: input.status ?? OpenClosedStatus.OPEN,
        actionOwner: input.actionOwner,
        dueDate: input.dueDate ? new Date(input.dueDate) : undefined
      }
    });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/trainings', async (_req, res) => {
  const rows = await prisma.training.findMany({
    include: { employee: true },
    orderBy: { trainingDate: 'desc' }
  });
  res.json(rows);
});

app.post('/api/trainings', async (req, res) => {
  try {
    const input = parseOrThrow(trainingCreateSchema, req.body);
    const row = await prisma.training.create({
      data: {
        employeeId: input.employeeId,
        trainingName: input.trainingName,
        trainingDate: new Date(input.trainingDate),
        validUntil: new Date(input.validUntil),
        certificateNo: input.certificateNo,
        certificateUrl: input.certificateUrl
      }
    });
    res.status(201).json(row);
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/reports/overview.pdf', async (_req, res) => {
  try {
    const language = parseLanguage(_req.query.lang);
    const buffer = await buildOverviewPdf(language);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="hse-overview-report.pdf"');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.get('/api/reports/overview.xlsx', async (_req, res) => {
  try {
    const language = parseLanguage(_req.query.lang);
    const buffer = await buildOverviewWorkbook(language);
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader('Content-Disposition', 'attachment; filename="hse-overview-report.xlsx"');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ message: (error as Error).message });
  }
});

app.post('/api/reports/email', async (req, res) => {
  try {
    const input = parseOrThrow(reportEmailSchema, req.body);
    const language: LanguageCode = input.language ?? 'tr';
    const format = input.format ?? 'PDF';
    const text = reportTexts[language];
    const transporter = buildMailTransport();
    const fromAddress = process.env.REPORT_FROM_EMAIL ?? 'hse-platform@local.test';

    const attachmentBuffer =
      format === 'PDF' ? await buildOverviewPdf(language) : await buildOverviewWorkbook(language);
    const attachmentName =
      format === 'PDF' ? 'hse-overview-report.pdf' : 'hse-overview-report.xlsx';
    const mimeType =
      format === 'PDF'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    const info = await transporter.sendMail({
      from: fromAddress,
      to: input.to,
      subject: input.subject ?? `${text.emailSubject} (${format.toUpperCase()})`,
      text: input.message ?? text.emailBody,
      attachments: [
        {
          filename: attachmentName,
          content: attachmentBuffer,
          contentType: mimeType
        }
      ]
    });

    res.json({
      accepted: true,
      transportMode:
        process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS
          ? 'smtp'
          : 'jsonTransport',
      messageId: info.messageId
    });
  } catch (error) {
    res.status(400).json({ message: (error as Error).message });
  }
});

app.get('/api/dashboard/ceo-summary', async (_req, res) => {
  const summary = await getDashboardSummary();
  res.json(summary);
});

const port = Number(process.env.PORT ?? 4000);
app.listen(port, () => {
  console.log(`HSE API running on :${port}`);
});
