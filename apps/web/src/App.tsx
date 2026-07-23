import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { countryList, getCitiesForCountry } from './data/locations';
import jsPDF from 'jspdf';
import { Document as DocxDocument, Packer, Paragraph, TextRun, HeadingLevel, Table as DocxTable, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } from 'docx';

declare global {
  interface Window {
    __HSE_GUEST_MODE__?: boolean;
  }
}

let hasSeededProjectCatalog = false;

type Status = 'OPEN' | 'IN_PROGRESS' | 'CLOSED';
type Language = 'tr' | 'en' | 'ru';
type DashboardPeriod = 'daily' | 'weekly' | 'monthly' | 'yearly' | 'custom';
type ComparisonMode = 'none' | 'project' | 'department';
type GroupKey = 'operations' | 'workforce' | 'compliance' | 'reports' | 'system';
type ModuleKey =
  | 'dashboard'
  | 'inspections'
  | 'observations'
  | 'risk-assessments'
  | 'permit-to-work'
  | 'incidents'
  | 'audits'
  | 'equipment-management'
  | 'environmental'
  | 'emergency-preparedness'
  | 'employees'
  | 'trainings'
  | 'ppe-stocks'
  | 'occupational-health'
  | 'legal-register'
  | 'documents'
  | 'action-tracker'
  | 'kpis-analytics'
  | 'reports'
  | 'export-center'
  | 'projects'
  | 'departments'
  | 'contractors'
  | 'settings'
type Project = {
  id: string;
  name: string;
  country: string;
  city: string;
  address: string;
  contractScope: string;
};

type ContractorMasterRecord = {
  id: string;
  companyName: string;
  projectId: string;
  projectName: string;
  country: string;
  city: string;
  projectLocation: string;
  contractScope: string;
  hseWarningCount: number;
  hseWarningDate: string;
  fireWarningCount: number;
  fireWarningDate: string;
  environmentWarningCount: number;
  environmentWarningDate: string;
  penaltyCount: number;
  penaltyLegalClause: string;
  totalPenaltyAmount: number;
};

type ContractorMasterForm = Omit<ContractorMasterRecord, 'id'>;

type DepartmentRecord = {
  id: string;
  name: string;
  code: string;
  description: string;
};

type DepartmentForm = Omit<DepartmentRecord, 'id'>;

type SettingsTabKey =
  | 'company-info'
  | 'general-settings'
  | 'roles-permissions'
  | 'notifications'
  | 'risk-matrix'
  | 'hse-categories'
  | 'penalty-settings'
  | 'document-numbering'
  | 'backup-restore'
  | 'audit-log';

type SettingsConfig = {
  companyName: string;
  taxNumber: string;
  headquartersAddress: string;
  timezone: string;
  defaultLanguage: string;
  dateFormat: string;
  roleMatrixSummary: string;
  approvalFlow: string;
  emailNotifications: boolean;
  smsNotifications: boolean;
  escalationHours: number;
  riskLikelihoodScale: string;
  riskSeverityScale: string;
  riskThresholds: string;
  hseCategories: string;
  environmentalCategories: string;
  fireSafetyCategories: string;
  warningPenaltyRule: string;
  currency: string;
  defaultPenaltyAmount: number;
  documentPrefix: string;
  documentSequenceStart: number;
  revisionFormat: string;
  backupFrequency: string;
  retentionDays: number;
  restoreApprovalRole: string;
  auditLogPolicy: string;
  auditLogRetentionMonths: number;
};

type ModuleRecord = {
  projectId: string;
  date: string;
  title: string;
  valueA: number;
  valueB: number;
  status: Status;
};

type CustomFileUploadProps = {
  buttonLabel: string;
  emptyLabel: string;
  singleLabel: string;
  multipleLabel: string;
  accept?: string;
  multiple?: boolean;
  onFilesChange: (files: File[]) => void;
};

function CustomFileUpload({
  buttonLabel,
  emptyLabel,
  singleLabel,
  multipleLabel,
  accept,
  multiple = false,
  onFilesChange
}: CustomFileUploadProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  const handleFiles = (incomingFiles: File[]) => {
    const nextFiles = multiple ? incomingFiles : incomingFiles.slice(0, 1);
    setSelectedFiles(nextFiles);
    onFilesChange(nextFiles);
    if (inputRef.current) {
      inputRef.current.value = '';
    }
  };

  const summaryText =
    selectedFiles.length === 0
      ? emptyLabel
      : selectedFiles.length === 1
        ? `${singleLabel}${selectedFiles[0]?.name ?? ''}`
        : `${multipleLabel}${selectedFiles.length}`;

  return (
    <div className="custom-file-upload">
      <div className="custom-file-upload-summary">{summaryText}</div>
      <button type="button" className="secondary" onClick={() => inputRef.current?.click()}>
        {buttonLabel}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        multiple={multiple}
        onChange={(event) => handleFiles(Array.from(event.target.files ?? []))}
        style={{ display: 'none' }}
      />
    </div>
  );
}

type ObservationPriority = 'DUSUK' | 'ORTA' | 'YUKSEK' | 'KRITIK';
type ObservationWorkflowStatus = 'ACIK' | 'DEVAM_EDIYOR' | 'KAPALI';

type ObservationAttachmentRecord = {
  id: string;
  name: string;
  mimeType: string;
  dataUrl: string;
};

type ObservationRecord = {
  id: string;
  observationNo: string;
  projectId: string;
  country: string;
  city: string;
  projectLocation: string;
  inspectionDate: string;
  inspectionTime: string;
  inspectorName: string;
  contractor: string;
  subcontractor: string;
  responsiblePerson: string;
  category: string;
  subject: string;
  observationLocation: string;
  violatedRequirement: string;
  observationDescription: string;
  correctiveAction: string;
  dueDate: string;
  comment: string;
  priority: ObservationPriority;
  status: ObservationWorkflowStatus;
  attachments: ObservationAttachmentRecord[];
};

type ObservationForm = {
  projectId: string;
  country: string;
  city: string;
  projectLocation: string;
  inspectionDate: string;
  inspectionTime: string;
  inspectorName: string;
  contractor: string;
  subcontractor: string;
  responsiblePerson: string;
  category: string;
  subject: string;
  observationLocation: string;
  violatedRequirement: string;
  observationDescription: string;
  correctiveAction: string;
  dueDate: string;
  comment: string;
  priority: ObservationPriority;
  status: ObservationWorkflowStatus;
};

type RiskRecord = {
  id: string;
  riskId: string;
  projectId: string;
  departmentActivity: string;
  assessmentDate: string;
  hazard: string;
  potentialConsequence: string;
  existingControls: string;
  recommendedControls: string;
  likelihood: number;
  severity: number;
  initialRiskScore: number;
  residualLikelihood: number;
  residualSeverity: number;
  residualRiskScore: number;
  responsiblePerson: string;
  targetCompletionDate: string;
  status: Status;
  attachments: string[];
  photos: string[];
  notes: string;
};

type RiskForm = {
  projectId: string;
  departmentActivity: string;
  assessmentDate: string;
  hazard: string;
  potentialConsequence: string;
  existingControls: string;
  recommendedControls: string;
  likelihood: number;
  severity: number;
  residualLikelihood: number;
  residualSeverity: number;
  responsiblePerson: string;
  targetCompletionDate: string;
  status: Status;
  attachments: string[];
  photos: string[];
  notes: string;
};

type IncidentForm = {
  projectId: string;
  date: string;
  title: string;
  incidentCount: number;
  lostWorkDays: number;
  status: Status;
};

type EmergencyType = 'YANGIN' | 'TAHLIYE' | 'ILK_YARDIM' | 'KIMYASAL_SIZINTI' | 'KURTARMA' | 'DEPREM' | 'DIGER';
type EmergencyDrillResult = 'BASARILI' | 'KISMEN_BASARILI' | 'BASARISIZ';
type EmergencyDrillStatus = 'PLANLANDI' | 'TAMAMLANDI' | 'IPTAL_EDILDI';

type EmergencyDrillRecord = {
  id: string;
  drillId: string;
  projectId: string;
  emergencyType: EmergencyType;
  drillName: string;
  drillDate: string;
  participantCount: number;
  drillResult: EmergencyDrillResult;
  openActions: number;
  closedActions: number;
  responsiblePerson: string;
  nextPlannedDrillDate: string;
  status: EmergencyDrillStatus;
  attachments: string[];
  drillReports: string[];
  photos: string[];
  notes: string;
};

type EmergencyDrillForm = {
  projectId: string;
  emergencyType: EmergencyType;
  drillName: string;
  drillDate: string;
  participantCount: number;
  drillResult: EmergencyDrillResult;
  openActions: number;
  closedActions: number;
  responsiblePerson: string;
  nextPlannedDrillDate: string;
  status: EmergencyDrillStatus;
  attachments: string[];
  drillReports: string[];
  photos: string[];
  notes: string;
};

type WorkforceRecord = {
  id: string;
  projectId: string;
  date: string;
  departmentArea: string;
  contractor: string;
  totalWorkforce: number;
  newEmployees: number;
  maleEmployees: number;
  femaleEmployees: number;
  dayShiftWorkers: number;
  nightShiftWorkers: number;
  overtimeWorkers: number;
  age18_25: number;
  age26_35: number;
  age36_45: number;
  age46_55: number;
  age56Plus: number;
  status: Status;
  notes: string;
};

type WorkforceForm = {
  projectId: string;
  date: string;
  departmentArea: string;
  contractor: string;
  totalWorkforce: number;
  newEmployees: number;
  maleEmployees: number;
  femaleEmployees: number;
  dayShiftWorkers: number;
  nightShiftWorkers: number;
  overtimeWorkers: number;
  age18_25: number;
  age26_35: number;
  age36_45: number;
  age46_55: number;
  age56Plus: number;
  status: Status;
  notes: string;
};

type TrainingStatus = 'PLANLANDI' | 'TAMAMLANDI' | 'DEVAM_EDIYOR' | 'SURESI_DOLDU';

type TrainingRecord = {
  id: string;
  projectId: string;
  trainingId: string;
  trainingType: string;
  trainingTitle: string;
  trainingCategory: string;
  trainingDate: string;
  provider: string;
  department: string;
  position: string;
  projectEmployeeCount: number;
  certifiedEmployeeCount: number;
  certificateRequired: boolean;
  certificateValidityDate: string;
  totalTrainingCost: number;
  costPerEmployee: number;
  status: TrainingStatus;
  attachments: string[];
  participantList: string[];
  certificates: string[];
  notes: string;
};

type TrainingForm = {
  projectId: string;
  trainingType: string;
  trainingTitle: string;
  trainingCategory: string;
  trainingDate: string;
  provider: string;
  department: string;
  position: string;
  projectEmployeeCount: number;
  certifiedEmployeeCount: number;
  certificateRequired: boolean;
  certificateValidityDate: string;
  totalTrainingCost: number;
  status: TrainingStatus;
  attachments: string[];
  participantList: string[];
  certificates: string[];
  notes: string;
};

type EquipmentStatus = 'ACTIVE' | 'OUT_OF_SERVICE' | 'UNDER_MAINTENANCE';
type EquipmentInspectionStatus = 'COMPLIANT' | 'UPCOMING' | 'OVERDUE';
type EquipmentRiskLevel = 'LOW' | 'MEDIUM' | 'HIGH';

type EquipmentRecord = {
  id: string;
  equipmentId: string;
  projectId: string;
  equipmentName: string;
  equipmentType: string;
  brandModel: string;
  serialNumber: string;
  location: string;
  responsiblePerson: string;
  lastInspectionDate: string;
  nextInspectionDate: string;
  inspectionStatus: EquipmentInspectionStatus;
  certificateNumber: string;
  certificateExpiryDate: string;
  equipmentStatus: EquipmentStatus;
  riskLevel: EquipmentRiskLevel;
  attachments: string[];
  inspectionReports: string[];
  equipmentPhotos: string[];
  notes: string;
};

type EquipmentForm = {
  projectId: string;
  equipmentName: string;
  equipmentType: string;
  brandModel: string;
  serialNumber: string;
  location: string;
  responsiblePerson: string;
  lastInspectionDate: string;
  nextInspectionDate: string;
  inspectionStatus: EquipmentInspectionStatus;
  certificateNumber: string;
  certificateExpiryDate: string;
  equipmentStatus: EquipmentStatus;
  riskLevel: EquipmentRiskLevel;
  attachments: string[];
  inspectionReports: string[];
  equipmentPhotos: string[];
  notes: string;
};

type PpeTransactionType = 'STOK_GIRISI' | 'STOK_CIKISI' | 'STOGA_IADE' | 'HASARLI_HURDA' | 'STOK_DUZELTME';
type PpeTransactionLifecycle = 'TAMAMLANDI' | 'ONAY_BEKLIYOR' | 'IPTAL';
type PpeStockStatus = 'YETERLI' | 'DUSUK_STOK' | 'STOKTA_YOK';

type PpeTransactionRecord = {
  id: string;
  transactionId: string;
  transactionType: PpeTransactionType;
  lifecycle: PpeTransactionLifecycle;
  projectId: string;
  warehouse: string;
  date: string;
  category: string;
  itemName: string;
  brandModel: string;
  unit: string;
  quantity: number;
  minimumStockLevel: number;
  responsiblePerson: string;
  supplier: string;
  targetPersonDepartment: string;
  unitPrice: number;
  totalCost: number;
  notes: string;
};

type PpeTransactionForm = {
  transactionType: PpeTransactionType;
  lifecycle: PpeTransactionLifecycle;
  projectId: string;
  warehouse: string;
  date: string;
  category: string;
  itemName: string;
  brandModel: string;
  unit: string;
  quantity: number;
  minimumStockLevel: number;
  responsiblePerson: string;
  supplier: string;
  targetPersonDepartment: string;
  unitPrice: number;
  notes: string;
};

type HealthRecord = {
  employeeName: string;
  employeeId: string;
  department: string;
  position: string;
  bloodGroup: string;
  allergies: string;
  chronicDisease: string;
  communicableDisease: string;
  medication: string;
  disabilityStatus: string;
  fitForWork: 'Yes' | 'No';
  restrictedWork: string;
  medicalExaminationDate: string;
  nextMedicalExamination: string;
  vaccinationStatus: string;
  remarks: string;
};

const initialHealthRecords: HealthRecord[] = [];

const isHealthRecord = (value: unknown): value is HealthRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.employeeName === 'string'
    && typeof candidate.employeeId === 'string'
    && typeof candidate.department === 'string'
    && typeof candidate.position === 'string'
    && typeof candidate.bloodGroup === 'string'
    && typeof candidate.allergies === 'string'
    && typeof candidate.chronicDisease === 'string'
    && typeof candidate.communicableDisease === 'string'
    && typeof candidate.medication === 'string'
    && typeof candidate.disabilityStatus === 'string'
    && (candidate.fitForWork === 'Yes' || candidate.fitForWork === 'No')
    && typeof candidate.restrictedWork === 'string'
    && typeof candidate.medicalExaminationDate === 'string'
    && typeof candidate.nextMedicalExamination === 'string'
    && typeof candidate.vaccinationStatus === 'string'
    && typeof candidate.remarks === 'string'
  );
};

const loadHealthRecords = (): HealthRecord[] => {
  if (typeof window === 'undefined') {
    return initialHealthRecords;
  }

  try {
    const raw = window.localStorage.getItem(HEALTH_RECORDS_STORAGE_KEY);
    if (!raw) {
      return initialHealthRecords;
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isHealthRecord)) {
      return parsed;
    }
  } catch {
    // Fall back to empty health records if local storage is unavailable or corrupted.
  }

  return initialHealthRecords;
};

type LegalComplianceStatus = 'UYUMLU' | 'KISMEN_UYUMLU' | 'UYUMSUZ' | 'UYGULANAMAZ';
type LegalRiskLevel = 'DUSUK' | 'ORTA' | 'YUKSEK' | 'KRITIK';
type LegalDocumentKind = 'MEVZUAT_BELGESI' | 'UYUMLULUK_KANITI' | 'DENETIM_RAPORU' | 'IZIN_SERTIFIKA' | 'DIGER';
type LegalUserRole = 'HSE_MANAGER' | 'CORPORATE_HSE_MANAGER' | 'COMPLIANCE_MANAGER' | 'ISO_REPRESENTATIVE' | 'VIEWER';

type LegalDocumentVersion = {
  version: number;
  fileName: string;
  fileType: string;
  fileUrl: string;
  uploadedAt: string;
  uploadedBy: string;
};

type LegalDocument = {
  id: string;
  kind: LegalDocumentKind;
  fileName: string;
  fileType: string;
  fileUrl: string;
  uploadedAt: string;
  uploadedBy: string;
  versions: LegalDocumentVersion[];
};

type LegalAuditEvent = {
  id: string;
  eventType:
    | 'CREATED'
    | 'UPDATED'
    | 'REVIEWED'
    | 'EXPORTED_PDF_SINGLE'
    | 'EXPORTED_PDF_FULL'
    | 'EXPORTED_EXCEL'
    | 'DOWNLOADED_DOCUMENT'
    | 'REPLACED_DOCUMENT'
    | 'DELETED_DOCUMENT';
  actor: string;
  eventDate: string;
  detail: string;
};

type LegalRecord = {
  id: string;
  projectId: string;
  regulationId: string;
  category: string;
  title: string;
  authority: string;
  department: string;
  legalRequirement: string;
  responsiblePerson: string;
  complianceStatus: LegalComplianceStatus;
  effectiveDate: string;
  lastReviewDate: string;
  nextReviewDate: string;
  openActions: number;
  riskLevel: LegalRiskLevel;
  notes: string;
  documents: LegalDocument[];
  createdBy: string;
  createdAt: string;
  modifiedBy: string;
  modifiedAt: string;
  reviewedBy: string;
  reviewedAt: string;
  exportHistory: Array<{ format: 'PDF_SINGLE' | 'PDF_FULL' | 'EXCEL'; actor: string; date: string }>;
  downloadHistory: Array<{ fileName: string; actor: string; date: string }>;
  auditTrail: LegalAuditEvent[];
};

type LegalRecordForm = {
  projectId: string;
  category: string;
  title: string;
  authority: string;
  department: string;
  legalRequirement: string;
  responsiblePerson: string;
  complianceStatus: LegalComplianceStatus;
  effectiveDate: string;
  lastReviewDate: string;
  nextReviewDate: string;
  openActions: number;
  riskLevel: LegalRiskLevel;
  notes: string;
};

type ControlledDocumentStatus = 'TASLAK' | 'GOZDEN_GECIRMEDE' | 'ONAYLANDI' | 'GECERSIZ';

type ControlledDocumentRevision = {
  id: string;
  revisionNumber: number;
  status: ControlledDocumentStatus;
  fileName: string;
  fileType: string;
  fileUrl: string;
  uploadedAt: string;
  uploadedBy: string;
  note: string;
};

type ControlledDocumentRecord = {
  id: string;
  projectId: string;
  documentId: string;
  title: string;
  category: string;
  documentType: string;
  revisionNumber: number;
  status: ControlledDocumentStatus;
  effectiveDate: string;
  reviewDate: string;
  department: string;
  preparedBy: string;
  reviewedBy: string;
  approvedBy: string;
  notes: string;
  fileName: string;
  fileType: string;
  fileUrl: string;
  uploadedAt: string;
  uploadedBy: string;
  revisions: ControlledDocumentRevision[];
};

type ControlledDocumentForm = {
  projectId: string;
  title: string;
  category: string;
  documentType: string;
  effectiveDate: string;
  reviewDate: string;
  department: string;
  preparedBy: string;
  reviewedBy: string;
  approvedBy: string;
  status: ControlledDocumentStatus;
  notes: string;
};

type CorporateReportFormat = 'PDF' | 'EXCEL';

type CorporateReportKey =
  | 'executive-dashboard'
  | 'inspections'
  | 'observations'
  | 'risk-assessments'
  | 'incidents'
  | 'trainings'
  | 'ppe-stocks'
  | 'equipment-management'
  | 'employees'
  | 'emergency-preparedness'
  | 'capa'
  | 'legal-register'
  | 'documents'
  | 'daily-hse'
  | 'weekly-hse'
  | 'monthly-hse'
  | 'quarterly-hse'
  | 'annual-hse'
  | 'client-hse'
  | 'corporate-hse'
  | 'employer-warning-letters'
  | 'employer-penalty-letters'
  | 'hse-environment-fire-letters'
  | 'state-inspection-warnings'
  | 'state-inspection-letters'
  | 'state-inspection-penalties';

type CorporateReportOption = {
  key: CorporateReportKey;
  label: string;
  group: 'module' | 'management' | 'correspondence';
};

type GeneratedCorporateReport = {
  id: string;
  reportName: string;
  reportTypeKey: CorporateReportKey;
  reportTypeLabel: string;
  projectId: string;
  projectName: string;
  department: string;
  periodStart: string;
  periodEnd: string;
  format: CorporateReportFormat;
  createdAt: string;
  createdBy: string;
  summary: string;
  tableRows: Array<{ label: string; value: string; note?: string }>;
  recommendations: string[];
  html: string;
  csvContent: string;
  downloadName: string;
};

type SummaryKpi = { label: string; value: string; trend?: string };
type IncidentKpiStatus = 'normal' | 'monitor' | 'risk' | 'critical';
type IncidentExecutiveKpi = {
  label: string;
  value: string;
  status: IncidentKpiStatus;
  note?: string;
};
type ExecutiveSection = {
  key: string;
  title: string;
  chartA: { type: 'line' | 'bar' | 'area' | 'donut' | 'gauge' | 'heatmap'; values: number[]; label: string };
  chartB?: { type: 'line' | 'bar' | 'area' | 'donut' | 'gauge' | 'heatmap'; values: number[]; label: string };
  highlights: Array<{ label: string; value: string }>;
  incidentExecutiveKpis?: IncidentExecutiveKpi[];
  details: Array<{ label: string; value: string; note?: string }>;
};
type ChartThemeName = 'operations' | 'safety' | 'risk' | 'compliance';
type ChartTheme = {
  primary: string;
  secondary: string;
  accent: string;
  danger: string;
  neutral: string;
};
type ChecklistAnswer = 'YES' | 'NO' | 'NA' | '';
type ActionStatus = 'OPEN' | 'IN_PROGRESS' | 'CLOSED';

type InspectionSection = {
  id: string;
  title: string;
  description: string;
  questions: string[];
};

type InspectionForm = {
  title: string;
  businessUnit: string;
  siteArea: string;
  locationType: string;
  inspectionDate: string;
  inspectorName: string;
  department: string;
  positiveObservations: string;
  inspectionNotes: string;
};

type CorrectiveAction = {
  text: string;
  status: ActionStatus;
};

type PtwStatus = 'Taslak' | 'Onay Bekliyor' | 'Aktif' | 'Askıya Alındı' | 'Tamamlandı' | 'İptal';
type PtwControlStatus = 'Uygun' | 'Uygun Değil' | 'N/A';
type PtwActionStatus = 'Açık' | 'Devam Ediyor' | 'Tamamlandı';

type PtwTeamMember = {
  adSoyad: string;
  gorevi: string;
  firma: string;
  egitimDurumu: string;
  imza: string;
};

type PtwControlItem = {
  kontrol: string;
  durum: PtwControlStatus;
  aciklama: string;
};

type PtwPrecautionItem = {
  tedbir: string;
  sorumlu: string;
  termin: string;
  durum: PtwActionStatus;
};

type PtwDailyLog = {
  tarih: string;
  saat: string;
  calismaBasladi: string;
  calismaBitti: string;
  aciklama: string;
  sorumlu: string;
};

type PtwTeamChange = {
  eklenenPersonel: string;
  ayrilanPersonel: string;
  tarih: string;
  onaylayan: string;
};

type PtwAttachmentApiType =
  | 'PHOTO'
  | 'PDF'
  | 'RISK_ASSESSMENT'
  | 'METHOD_STATEMENT'
  | 'TOOLBOX_TALK'
  | 'CERTIFICATE'
  | 'OTHER';

type PtwAttachment = {
  id?: string;
  type: PtwAttachmentApiType;
  tip: string;
  dosyaAdi: string;
  fileType: string;
  fileUrl: string;
  previewUrl?: string;
};

type PtwApiRecord = {
  id: string;
  permitNo: string;
  organization: string;
  department: string;
  project: string;
  location: string;
  permitType: string;
  issueDate: string;
  validityStart: string;
  validityEnd: string;
  status: PtwStatus;
  requesterName: string;
  issuerName: string;
  jobResponsibleName: string;
  siteResponsibleName: string;
  hseResponsibleName: string;
  authorizedApprover: string;
  jobTitle: string;
  jobDescription: string;
  workArea: string;
  plannedWork: string;
  workingConditions: string;
  workStartDate: string;
  workStartTime: string;
  workEndDate: string;
  workEndTime: string;
  hazards: string[];
  safetySystems: string[];
  equipmentUsed: string[];
  teamMembers: PtwTeamMember[];
  preWorkChecks: PtwControlItem[];
  precautions: PtwPrecautionItem[];
  specialConditions?: string;
  preparedBy?: string;
  hseApprovalBy?: string;
  projectManagerBy?: string;
  employerRepBy?: string;
  digitalSignature?: string;
  approvalDate?: string;
  dailyLogs: PtwDailyLog[];
  teamChanges: PtwTeamChange[];
  workCompleted: boolean;
  areaSafe: boolean;
  materialsCollected: boolean;
  ptwClosed: boolean;
  closedBy?: string;
  closedAt?: string;
  closureRemarks?: string;
  attachments: Array<{
    id?: string;
    type: PtwAttachmentApiType;
    fileName: string;
    fileType: string;
    fileUrl: string;
  }>;
  createdAt: string;
  updatedAt: string;
};

type PtwLocalRecord = {
  id: string;
  payload: Record<string, unknown>;
  savedAt: string;
};

const isPtwLocalRecord = (value: unknown): value is PtwLocalRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && !!candidate.id
    && typeof candidate.savedAt === 'string'
    && typeof candidate.payload === 'object'
    && candidate.payload !== null
  );
};

type PtwForm = {
  ptwNo: string;
  organizasyon: string;
  departman: string;
  proje: string;
  lokasyon: string;
  ptwTuru: string;
  duzenlenmeTarihi: string;
  gecerlilikBaslangici: string;
  gecerlilikBitisi: string;
  durum: PtwStatus;
  isiTalepEden: string;
  isiVeren: string;
  isSorumlusu: string;
  sahaSorumlusu: string;
  hseSorumlusu: string;
  yetkiliOnaylayan: string;
  isinAdi: string;
  isinAciklamasi: string;
  calismaAlani: string;
  yapilacakIs: string;
  calismaKosullari: string;
  baslangicTarihi: string;
  baslangicSaati: string;
  bitisTarihi: string;
  bitisSaati: string;
  ozelSartlar: string;
  ptwHazirlayan: string;
  hseOnayi: string;
  projeMuduru: string;
  isverenTemsilcisi: string;
  dijitalImza: string;
  onayTarihi: string;
  isTamamlandi: boolean;
  alanGuvenli: boolean;
  malzemelerToplandi: boolean;
  ptwKapatildi: boolean;
  kapatan: string;
  kapanisTarihi: string;
  kapanisAciklama: string;
};

type PtwTabKey =
  | 'kayit'
  | 'sorumlular'
  | 'ekip'
  | 'is-bilgileri'
  | 'riskler'
  | 'guvenlik-sistemleri'
  | 'ekipman'
  | 'on-kontroller'
  | 'tedbirler'
  | 'ozel-sartlar'
  | 'onaylar'
  | 'gunluk-takip'
  | 'ekip-degisiklik'
  | 'kapanis'
  | 'dosya-ekleri';

const initialProjects: Project[] = [];

const PROJECTS_STORAGE_KEY = 'hse-project-catalog-v1';
const DEPARTMENTS_STORAGE_KEY = 'hse-department-records-v1';
const CONTRACTORS_STORAGE_KEY = 'hse-contractor-records-v1';
const PPE_TRANSACTIONS_STORAGE_KEY = 'hse-ppe-transactions-v1';
const HEALTH_RECORDS_STORAGE_KEY = 'hse-health-records-v1';
const OBSERVATION_RECORDS_STORAGE_KEY = 'hse-observation-records-v1';
const PTW_LOCAL_RECORDS_STORAGE_KEY = 'hse-ptw-local-records-v1';

const observationCategoryOptions = [
  'Güvensiz Davranış',
  'Güvensiz Durum',
  'Çevresel Uygunsuzluk',
  'Yangın Güvenliği',
  'KKD Uygunsuzluğu',
  'Housekeeping',
  'Pozitif Gözlem',
  'Diğer'
];

const observationPriorityLabel: Record<ObservationPriority, string> = {
  DUSUK: 'Düşük',
  ORTA: 'Orta',
  YUKSEK: 'Yüksek',
  KRITIK: 'Kritik'
};

const observationStatusLabel: Record<ObservationWorkflowStatus, string> = {
  ACIK: 'Açık',
  DEVAM_EDIYOR: 'Devam Ediyor',
  KAPALI: 'Kapalı'
};

const observationCategoryOptionsRu: Record<string, string> = {
  'Güvensiz Davranış': 'Небезопасные действия',
  'Güvensiz Durum': 'Небезопасные условия',
  'Çevresel Uygunsuzluk': 'Экологическое несоответствие',
  'Yangın Güvenliği': 'Пожарная безопасность',
  'KKD Uygunsuzluğu': 'Несоответствие по СИЗ',
  Housekeeping: 'Содержание рабочих мест',
  'Pozitif Gözlem': 'Позитивное предписание',
  'Diğer': 'Прочее'
};

const observationPriorityLabelRu: Record<ObservationPriority, string> = {
  DUSUK: 'Низкий',
  ORTA: 'Средний',
  YUKSEK: 'Высокий',
  KRITIK: 'Критический'
};

const observationStatusLabelRu: Record<ObservationWorkflowStatus, string> = {
  ACIK: 'Открыто',
  DEVAM_EDIYOR: 'В работе',
  KAPALI: 'Закрыто'
};

const createEmptyObservationForm = (project?: Project): ObservationForm => ({
  projectId: project?.id ?? '',
  country: project?.country ?? '',
  city: project?.city ?? '',
  projectLocation: project?.address ?? '',
  inspectionDate: new Date().toISOString().slice(0, 10),
  inspectionTime: new Date().toISOString().slice(11, 16),
  inspectorName: '',
  contractor: '',
  subcontractor: '',
  responsiblePerson: '',
  category: observationCategoryOptions[0],
  subject: '',
  observationLocation: '',
  violatedRequirement: '',
  observationDescription: '',
  correctiveAction: '',
  dueDate: new Date().toISOString().slice(0, 10),
  comment: '',
  priority: 'ORTA',
  status: 'ACIK'
});

const isObservationAttachmentRecord = (value: unknown): value is ObservationAttachmentRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.name === 'string'
    && typeof candidate.mimeType === 'string'
    && typeof candidate.dataUrl === 'string'
  );
};

const isObservationRecord = (value: unknown): value is ObservationRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.observationNo === 'string'
    && typeof candidate.projectId === 'string'
    && typeof candidate.country === 'string'
    && typeof candidate.city === 'string'
    && typeof candidate.projectLocation === 'string'
    && typeof candidate.inspectionDate === 'string'
    && typeof candidate.inspectionTime === 'string'
    && typeof candidate.inspectorName === 'string'
    && typeof candidate.contractor === 'string'
    && typeof candidate.subcontractor === 'string'
    && typeof candidate.responsiblePerson === 'string'
    && typeof candidate.category === 'string'
    && typeof candidate.subject === 'string'
    && typeof candidate.observationLocation === 'string'
    && typeof candidate.violatedRequirement === 'string'
    && typeof candidate.observationDescription === 'string'
    && typeof candidate.correctiveAction === 'string'
    && typeof candidate.dueDate === 'string'
    && typeof candidate.comment === 'string'
    && (candidate.priority === 'DUSUK' || candidate.priority === 'ORTA' || candidate.priority === 'YUKSEK' || candidate.priority === 'KRITIK')
    && (candidate.status === 'ACIK' || candidate.status === 'DEVAM_EDIYOR' || candidate.status === 'KAPALI')
    && Array.isArray(candidate.attachments)
    && candidate.attachments.every(isObservationAttachmentRecord)
  );
};

const loadObservationRecords = (): ObservationRecord[] => {
  if (typeof window === 'undefined') {
    return [];
  }

  try {
    const raw = window.localStorage.getItem(OBSERVATION_RECORDS_STORAGE_KEY);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isObservationRecord)) {
      return parsed;
    }
  } catch {
    // Keep empty list on parse or storage failures.
  }

  return [];
};

const isProjectRecord = (value: unknown): value is Project => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.name === 'string'
    && typeof candidate.country === 'string'
    && typeof candidate.city === 'string'
    && typeof candidate.address === 'string'
    && typeof candidate.contractScope === 'string'
  );
};

const loadProjectCatalog = (): Project[] => {
  if (typeof window === 'undefined') {
    return initialProjects;
  }

  try {
    const raw = window.localStorage.getItem(PROJECTS_STORAGE_KEY);
    if (!raw) {
      return initialProjects;
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isProjectRecord)) {
      return parsed;
    }
  } catch {
    // Fall back to seeded projects if local storage is unavailable or corrupted.
  }

  return [];
};

const initialContractorRecords: ContractorMasterRecord[] = [];

const isContractorRecord = (value: unknown): value is ContractorMasterRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.companyName === 'string'
    && typeof candidate.projectId === 'string'
    && typeof candidate.projectName === 'string'
    && typeof candidate.country === 'string'
    && typeof candidate.city === 'string'
    && typeof candidate.projectLocation === 'string'
    && typeof candidate.contractScope === 'string'
    && typeof candidate.hseWarningCount === 'number'
    && typeof candidate.hseWarningDate === 'string'
    && typeof candidate.fireWarningCount === 'number'
    && typeof candidate.fireWarningDate === 'string'
    && typeof candidate.environmentWarningCount === 'number'
    && typeof candidate.environmentWarningDate === 'string'
    && typeof candidate.penaltyCount === 'number'
    && typeof candidate.penaltyLegalClause === 'string'
    && typeof candidate.totalPenaltyAmount === 'number'
  );
};

const loadContractorRecords = (): ContractorMasterRecord[] => {
  if (typeof window === 'undefined') {
    return initialContractorRecords;
  }

  try {
    const raw = window.localStorage.getItem(CONTRACTORS_STORAGE_KEY);
    if (!raw) {
      return initialContractorRecords;
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isContractorRecord)) {
      return parsed;
    }
  } catch {
    // Fall back to seeded contractors if local storage is unavailable or corrupted.
  }

  return initialContractorRecords;
};

const settingsTabOptions: Array<{ key: SettingsTabKey; label: string }> = [
  { key: 'company-info', label: 'Şirket Bilgileri' },
  { key: 'general-settings', label: 'Genel Ayarlar' },
  { key: 'roles-permissions', label: 'Kullanıcı Rolleri ve Yetkiler' },
  { key: 'notifications', label: 'Bildirim Ayarları' },
  { key: 'risk-matrix', label: 'Risk Matrisi Ayarları' },
  { key: 'hse-categories', label: 'İSG Kategorileri' },
  { key: 'penalty-settings', label: 'Ceza Ayarları' },
  { key: 'document-numbering', label: 'Belge Numaralandırma' },
  { key: 'backup-restore', label: 'Yedekleme ve Geri Yükleme' },
  { key: 'audit-log', label: 'Denetim Günlüğü' }
];

const initialSettingsConfig: SettingsConfig = {
  companyName: 'HSE Uyum Platformu A.Ş.',
  taxNumber: '1234567890',
  headquartersAddress: 'Maslak Mah. Büyükdere Cad. No: 123, İstanbul',
  timezone: 'Europe/Istanbul',
  defaultLanguage: 'Türkçe',
  dateFormat: 'DD.MM.YYYY',
  roleMatrixSummary: 'HSE Manager: Full, Supervisor: Approve, Viewer: Read-only',
  approvalFlow: 'Kritik değişiklikler için iki aşamalı onay',
  emailNotifications: true,
  smsNotifications: false,
  escalationHours: 24,
  riskLikelihoodScale: '1-5',
  riskSeverityScale: '1-5',
  riskThresholds: 'Düşük: 1-5, Orta: 6-10, Yüksek: 11-15, Kritik: 16-25',
  hseCategories: 'Davranışsal Güvenlik, PPE, Çalışma İzni, Yüksekte Çalışma',
  environmentalCategories: 'Atık Yönetimi, Emisyon, Gürültü, Su Kullanımı',
  fireSafetyCategories: 'Yangın Söndürücüler, Acil Kaçış, Alarm Sistemleri',
  warningPenaltyRule: '3 uyarı sonrası sözleşmesel ceza tetiklenir',
  currency: 'RUB',
  defaultPenaltyAmount: 15000,
  documentPrefix: 'HSE-DOC',
  documentSequenceStart: 1000,
  revisionFormat: 'REV-{NN}',
  backupFrequency: 'Günlük 02:00',
  retentionDays: 365,
  restoreApprovalRole: 'Kurumsal İSG Yöneticisi',
  auditLogPolicy: 'Tüm yapılandırma değişiklikleri zorunlu kayıt altına alınır',
  auditLogRetentionMonths: 24
};

const uiText: Record<Language, Record<string, string>> = {
  tr: {
    dashboard: 'Gösterge Paneli',
    project: 'Proje',
    language: 'Dil',
    allProjects: 'Tüm Projeler',
    overview: 'Proje performansına ve tüm projelerin görünümüne genel bakış.',
    moduleAnalytics: 'veri girişi ve analiz',
    executiveKpi: 'Yönetici Gösterge Paneli / KPI',
    chart: 'Grafik',
    dataEntry: 'Veri Girişi',
    records: 'Kayıtlar',
    save: 'Kaydet',
    noData: 'Seçili kapsam için veri yok.',
    noHealth: 'Henüz iş sağlığı kaydı yok.',
    projectDataEntry: 'Proje Ana Veri Girişi',
    projectName: 'Proje Adı',
    country: 'Ülke',
    city: 'Şehir',
    address: 'Adres',
    whereToEnter: 'Proje adı, ülke, şehir ve adres bilgilerini SİSTEM > Projeler ekranından girin.',
    primaryMetric: 'Ana Metrik',
    secondaryMetric: 'İkincil Metrik'
  },
  en: {
    dashboard: 'Dashboard',
    project: 'Project',
    language: 'Language',
    allProjects: 'All Projects',
    overview: 'Project performance overview and cross-project visibility.',
    moduleAnalytics: 'data entry and analytics',
    executiveKpi: 'Executive Dashboard / KPIs',
    chart: 'Chart',
    dataEntry: 'Data Entry',
    records: 'Records',
    save: 'Save',
    noData: 'No data available for selected scope.',
    noHealth: 'No occupational health record yet.',
    projectDataEntry: 'Project Master Data Entry',
    projectName: 'Project Name',
    country: 'Country',
    city: 'City',
    address: 'Address',
    whereToEnter: 'Enter project name, country, city and address in SYSTEM > Projects.',
    primaryMetric: 'Primary Metric',
    secondaryMetric: 'Secondary Metric'
  },
  ru: {
    dashboard: 'Панель управления',
    project: 'Проект',
    language: 'Язык',
    allProjects: 'Все проекты',
    overview: 'Обзор производительности проекта и сравнение всех проектов.',
    moduleAnalytics: 'ввод данных и аналитика',
    executiveKpi: 'Панель руководителя / KPI',
    chart: 'График',
    dataEntry: 'Ввод данных',
    records: 'Записи',
    save: 'Сохранить',
    noData: 'Нет данных для выбранной области.',
    noHealth: 'Пока нет записей по охране труда.',
    projectDataEntry: 'Ввод основных данных проекта',
    projectName: 'Название проекта',
    country: 'Страна',
    city: 'Город',
    address: 'Адрес',
    whereToEnter: 'Название проекта, страну, город и адрес вводите в разделе SYSTEM > Projects.',
    primaryMetric: 'Основной показатель',
    secondaryMetric: 'Вторичный показатель'
  }
};

const moduleLabelsByLanguage: Record<Language, Record<ModuleKey, string>> = {
  tr: {
    dashboard: 'Gösterge Paneli',
    inspections: 'Denetimler',
    observations: 'Gözlemler',
    'risk-assessments': 'Risk Değerlendirmeleri',
    'permit-to-work': 'Çalışma İzni',
    incidents: 'Olaylar',
    audits: 'Tetkikler',
    'equipment-management': 'Ekipman Yönetimi',
    environmental: 'Çevre',
    'emergency-preparedness': 'Acil Durum Hazırlığı',
    employees: 'Çalışanlar',
    trainings: 'Eğitimler',
    'ppe-stocks': 'KKD Stokları',
    'occupational-health': 'İş Sağlığı',
    'legal-register': 'Yasal Envanter',
    documents: 'Dokümanlar',
    'action-tracker': 'Aksiyon Takibi (DÖF)',
    'kpis-analytics': 'KPI ve Analitik',
    reports: 'Raporlar',
    'export-center': 'Dışa Aktarım Merkezi',
    projects: 'Projeler',
    departments: 'Departmanlar',
    contractors: 'Alt Yükleniciler',
    settings: 'Ayarlar'
  },
  en: {
    dashboard: 'Dashboard',
    inspections: 'Inspections',
    observations: 'Observations',
    'risk-assessments': 'Risk Assessments',
    'permit-to-work': 'Permit to Work',
    incidents: 'Incidents',
    audits: 'Audits',
    'equipment-management': 'Equipment Management',
    environmental: 'Environmental',
    'emergency-preparedness': 'Emergency Preparedness',
    employees: 'Employees',
    trainings: 'Trainings',
    'ppe-stocks': 'PPE Stocks',
    'occupational-health': 'Occupational Health',
    'legal-register': 'Legal Register',
    documents: 'Documents',
    'action-tracker': 'Action Tracker (CAPA)',
    'kpis-analytics': 'KPIs & Analytics',
    reports: 'Reports',
    'export-center': 'Export Center',
    projects: 'Projects',
    departments: 'Departments',
    contractors: 'Contractors',
    settings: 'Settings'
  },
  ru: {
    dashboard: 'Панель',
    inspections: 'Инспекции',
    observations: 'Предписания',
    'risk-assessments': 'Оценка рисков',
    'permit-to-work': 'Наряд-допуск',
    incidents: 'Инциденты',
    audits: 'Аудиты',
    'equipment-management': 'Управление оборудованием',
    environmental: 'Экология',
    'emergency-preparedness': 'Готовность к ЧС',
    employees: 'Сотрудники',
    trainings: 'Обучения',
    'ppe-stocks': 'СИЗ Запасы',
    'occupational-health': 'Охрана труда',
    'legal-register': 'Реестр законодательства',
    documents: 'Документы',
    'action-tracker': 'Трекер действий (CAPA)',
    'kpis-analytics': 'KPI и Аналитика',
    reports: 'Отчеты',
    'export-center': 'Центр экспорта',
    projects: 'Проекты',
    departments: 'Отделы',
    contractors: 'Подрядчики',
    settings: 'Настройки'
  }
};

const metricFieldLabels: Partial<Record<ModuleKey, { primary: string; secondary: string }>> = {
  inspections: { primary: 'Inspected Areas', secondary: 'Findings Count' },
  observations: { primary: 'Unsafe Acts', secondary: 'Positive Observations' },
  'risk-assessments': { primary: 'Risk Score (Probability x Severity)', secondary: 'Control Measures (Count/Effectiveness)' },
  'permit-to-work': { primary: 'Permits Issued', secondary: 'Active Permits' },
  incidents: { primary: 'Incident Count', secondary: 'Lost Work Days' },
  audits: { primary: 'Audit Checks', secondary: 'Non-Conformities' },
  'equipment-management': { primary: 'Equipment Count', secondary: 'Due Inspections' },
  environmental: { primary: 'Waste Generated', secondary: 'Waste Recycled' },
  'emergency-preparedness': { primary: 'Drill Count', secondary: 'Open Actions' },
  employees: { primary: 'Employee Count', secondary: 'Absent Count' },
  trainings: { primary: 'Employees Trained', secondary: 'Pending Trainings' },
  'ppe-stocks': { primary: 'PPE Received', secondary: 'PPE Issued' },
  'legal-register': { primary: 'Legal Items', secondary: 'Pending Actions' },
  documents: { primary: 'Document Count', secondary: 'Revisions' },
  'action-tracker': { primary: 'Actions Opened', secondary: 'Actions Overdue' },
  'kpis-analytics': { primary: 'KPI Score', secondary: 'Target Achievement' },
  reports: { primary: 'Reports Generated', secondary: 'Reports Shared' },
  'export-center': { primary: 'Exports Count', secondary: 'Failed Exports' },
  contractors: { primary: 'Contractor Count', secondary: 'Pending Approvals' },
  settings: { primary: 'Updated Settings', secondary: 'Pending Changes' }
};

const textByLanguage: Partial<Record<Language, Record<string, string>>> = {
  tr: {
    'HSE Compliance Platform': 'HSE Uyum Platformu',
    'Designed and developed by Erdem Cetin': 'Erdem Cetin tarafından tasarlandı ve geliştirildi',
    'Design and Development: Erdem Cetin © 2026 All Rights Reserved.': 'Tasarım ve Geliştirme: Erdem Cetin © 2026 Tüm Hakları Saklıdır.',
    'Project performance and executive summary.': 'Proje performansı ve yönetici özeti.',
    'Corporate report center and automated output generation.': 'Kurumsal raporlama merkezi ve otomatik çıktı üretimi.',
    'Project master data management.': 'Proje ana veri yönetimi.',
    'Subcontractor master data management.': 'Alt yüklenici ana veri yönetimi.',
    'System configuration and governance settings.': 'Sistem yapılandırması ve yönetişim ayarları.',
    'Permit to Work': 'Çalışma İzni',
    'Weather Summary': 'Hava durumu özeti',
    'Wind Summary': 'Rüzgar özeti',
    'Incident Executive KPI Summary': 'Olay Yönetici KPI Özeti',
    'HSE Activity by Project': 'Projelere Göre HSE Aktivitesi',
    'HSE Activity by Department': 'Departmanlara Göre HSE Aktivitesi',
    'Project Distribution': 'Proje Dağılımı',
    'Department Distribution': 'Departman Dağılımı',
    Period: 'Periyot',
    Daily: 'Günlük',
    Weekly: 'Haftalık',
    Monthly: 'Aylık',
    Yearly: 'Yıllık',
    'Custom Date Range': 'Özel Tarih Aralığı',
    Comparison: 'Karşılaştırma',
    None: 'Yok',
    From: 'Başlangıç',
    To: 'Bitiş',
    'Select Projects': 'Projeleri Seç',
    'Select Departments': 'Departmanları Seç',
    'Apply Filters': 'Filtreleri Uygula',
    'Reset Filters': 'Filtreleri Sıfırla',
    'Executive PDF Output': 'Yönetici PDF Çıktısı',
    'Safety Performance': 'Güvenlik Performansı',
    'Safety Score Trend': 'Güvenlik Skoru Trendi',
    'Overall Safety Index': 'Genel Güvenlik Endeksi',
    'Check Details': 'Kontrol Et',
    Increase: 'Artış',
    Decrease: 'Azalış',
    Stable: 'Sabit',
    Average: 'Ortalama',
    Minimum: 'Min',
    Maximum: 'Max',
    Last: 'Son',
    Compliance: 'Uygunluk',
    Strength: 'Güçlü',
    Critical: 'Kritik',
    'Fit for Work Trend': 'İşe Uygunluk Trendi',
    'Health Status Distribution': 'Sağlık Durum Dağılımı',
    'Open Risk Trend': 'Açık Risk Trend',
    'Risk Level Distribution': 'Risk Seviyesi Dağılımı',
    'Waste Generation Trend': 'Atık Üretimi Trendi',
    'Recycling Trend': 'Geri Dönüşüm Trendi',
    'Compliance Score': 'Uyum Skoru',
    'Audit and Non-Conformity View': 'Tetkik ve Uygunsuzluk Görünümü',
    'Training Participation Trend': 'Eğitime Katılım Trendi',
    'Training Completion Rate': 'Eğitim Tamamlama Oranı',
    'PPE Receipt Trend': 'KKD Teslim Alım Trendi',
    'PPE Stock Distribution': 'KKD Stok Dağılımı',
    'Equipment Inventory': 'Ekipman Envanteri',
    'Certification Compliance Rate': 'Sertifikasyon Uygunluk Oranı',
    'Inspection Count Trend': 'Denetim Sayısı Trendi',
    'Finding Status': 'Bulgu Durumu',
    'Incident Trend (Last 7 Periods)': 'Olay Trendi (Son 7 Dönem)',
    'Incident Category Distribution': 'Olay Kategori Dağılımı',
    'Permit to Work Volume Trend': 'Çalışma İzni Hacim Trendi',
    'Permit to Work Status Distribution': 'Çalışma İzni Durum Dağılımı',
    'Permit to Work Violation': 'Çalışma İzni İhlali',
    'Wind Direction': 'Rüzgar Yönü',
    Open: 'Açık',
    'All Projects': 'Tüm Projeler',
    'Workforce Safety': 'İş Güvenliği',
    'Inspections & Observations': 'Denetimler ve Gözlemler',
    Training: 'Eğitim',
    'PPE Management': 'KKD Yönetimi',
    'Equipment Management': 'Ekipman Yönetimi',
    'Occupational Health': 'İş Sağlığı',
    'Risk Management': 'Risk Yönetimi',
    Environmental: 'Çevre',
    'Emergency Preparedness': 'Acil Durum Hazırlığı',
    'Reports & Analytics': 'Raporlar ve Analitik',
    'Total Workforce': 'Toplam İş Gücü',
    'PTW Active': 'Aktif Çalışma İzni',
    'PTW Closed': 'Kapatılan Çalışma İzni',
    'Dashboard': 'Gösterge Paneli',
    OPERATIONS: 'OPERASYONLAR',
    WORKFORCE: 'İŞ GÜCÜ',
    REPORTS: 'RAPORLAR',
    SYSTEM: 'SİSTEM',
    'Daily Workforce': 'Günlük İş Gücü',
    'Safe Man-Hours (Daily / Weekly / Monthly / Total)': 'Güvenli Adam-Saat (Günlük / Haftalık / Aylık / Toplam)',
    LTI: 'LTI',
    MTI: 'MTI',
    'Near Misses': 'Ramak Kala Olay',
    'Unsafe Acts': 'Güvensiz Davranışlar',
    'Unsafe Conditions': 'Güvensiz Koşullar',
    'Lost Time Injury Frequency Rate (LTIFR)': 'Kayıp Zamanlı Yaralanma Sıklık Oranı (LTIFR)',
    'Total Recordable Incident Rate (TRIR)': 'Toplam Kayıt Altına Alınan Olay Oranı (TRIR)',
    'Total Inspections': 'Toplam Denetim',
    'Open Findings': 'Açık Bulgular',
    'Closed Findings': 'Kapanan Bulgular',
    'Safety Observations': 'Güvenlik Gözlemleri',
    'Positive Observations': 'Pozitif Gözlemler',
    'Overdue Actions': 'Geciken Aksiyonlar',
    'Weekly Inspection Performance': 'Haftalık Denetim Performansı',
    'Daily Inductions': 'Günlük Oryantasyon',
    'Weekly Inductions': 'Haftalık Oryantasyon',
    'Monthly Inductions': 'Aylık Oryantasyon',
    'Total Inductions': 'Toplam Oryantasyon',
    'Toolbox Talks (Daily / Weekly / Monthly)': 'Toolbox Konuşmaları (Günlük / Haftalık / Aylık)',
    'Employees Trained (Weekly / Monthly)': 'Eğitilen Çalışanlar (Haftalık / Aylık)',
    'Upcoming Trainings': 'Yaklaşan Eğitimler',
    'Expired Training Certificates': 'Süresi Dolan Eğitim Sertifikaları',
    'PPE Received': 'Teslim Alınan KKD',
    'PPE Issued': 'Dağıtılan KKD',
    'PPE In Stock': 'Stoktaki KKD',
    'Low Stock Items': 'Düşük Stok Kalemleri',
    'PPE Distribution by Department': 'Bölümlere Göre KKD Dağılımı',
    'PPE Consumption Trend': 'KKD Tüketim Trendi',
    'Total Equipment': 'Toplam Ekipman',
    'Mobile Equipment': 'Mobil Ekipman',
    'Heavy Equipment': 'Ağır Ekipman',
    'Lifting Equipment': 'Kaldırma Ekipmanı',
    'Inspection Due': 'Denetimi Yaklaşan',
    'Certified Equipment': 'Sertifikalı Ekipman',
    'Out of Service Equipment': 'Servis Dışı Ekipman',
    'Total Employees': 'Toplam Çalışan',
    'Fit for Work': 'İşe Uygun',
    'Restricted Duty': 'Kısıtlı Görev',
    'Medical Leave': 'Sağlık İzni',
    'Periodic Medical Examination Due': 'Periyodik Muayene Zamanı Gelen',
    'Completed Medical Examinations': 'Tamamlanan Muayeneler',
    'Employees with Chronic Diseases': 'Kronik Hastalığı Olan Çalışanlar',
    'Employees with Communicable Diseases': 'Bulaşıcı Hastalığı Olan Çalışanlar',
    'Fitness Certificates Expiring': 'Süresi Dolacak Uygunluk Belgeleri',
    'Vaccination Status': 'Aşı Durumu',
    'Open Risks': 'Açık Riskler',
    'Closed Risks': 'Kapanan Riskler',
    'Top 10 Critical Risks': 'En Kritik 10 Risk',
    'Risk Trend': 'Risk Trendi',
    'Waste Generated': 'Oluşan Atık',
    'Waste Recycled': 'Geri Dönüştürülen Atık',
    'Environmental Inspections': 'Çevre Denetimleri',
    'Internal Audits': 'İç Tetkikler',
    'External Audits': 'Dış Tetkikler',
    'Open NCRs': 'Açık Uygunsuzluklar',
    'Closed NCRs': 'Kapanan Uygunsuzluklar',
    'CAPA Open': 'Açık DÖF',
    'CAPA Closed': 'Kapanan DÖF',
    'Legal Compliance Status': 'Yasal Uygunluk Durumu',
    'Emergency Drills': 'Acil Durum Tatbikatları',
    'Fire Drills': 'Yangın Tatbikatları',
    'Evacuation Drills': 'Tahliye Tatbikatları',
    'Emergency Equipment Status': 'Acil Ekipman Durumu',
    'Fire Extinguisher Inspections': 'Yangın Söndürücü Denetimleri',
    'First Aid Kit Inspections': 'İlk Yardım Çantası Denetimleri',
    'Daily Report': 'Günlük Rapor',
    'Weekly Report': 'Haftalık Rapor',
    'Monthly Report': 'Aylık Rapor',
    'Yearly Statistics': 'Yıllık İstatistikler',
    'KPI Trends': 'KPI Trendleri',
    'Project Comparison': 'Proje Karşılaştırması',
    'Department Comparison': 'Departman Karşılaştırması',
    'Inspected Areas': 'Denetlenen Alan Sayısı',
    'Findings Count': 'Bulgu Sayısı',
    'Risk Score': 'Risk Skoru',
    'Control Measures': 'Kontrol Tedbirleri',
    'Permits Issued': 'Açılan İzin Sayısı',
    'Active Permits': 'Aktif İzin Sayısı',
    'Incident Count': 'Olay Sayısı',
    'Lost Work Days': 'Kayıp İş Günü',
    'Audit Checks': 'Tetkik Kontrol Sayısı',
    'Non-Conformities': 'Uygunsuzluk Sayısı',
    'Equipment Count': 'Ekipman Sayısı',
    'Due Inspections': 'Muayene Zamanı Gelen',
    'Drill Count': 'Tatbikat Sayısı',
    'Open Actions': 'Açık Aksiyonlar',
    'Employee Count': 'Çalışan Sayısı',
    'Absent Count': 'Eksik Çalışan Sayısı',
    'Employees Trained': 'Eğitilen Çalışan Sayısı',
    'Pending Trainings': 'Bekleyen Eğitimler',
    'Legal Items': 'Yasal Madde Sayısı',
    'Pending Actions': 'Bekleyen Aksiyonlar',
    'Document Count': 'Doküman Sayısı',
    Revisions: 'Revizyon Sayısı',
    'Actions Opened': 'Açılan Aksiyonlar',
    'Actions Overdue': 'Geciken Aksiyonlar',
    'KPI Score': 'KPI Skoru',
    'Target Achievement': 'Hedef Gerçekleşme',
    'Reports Generated': 'Oluşturulan Rapor Sayısı',
    'Reports Shared': 'Paylaşılan Rapor Sayısı',
    'Exports Count': 'Dışa Aktarım Sayısı',
    'Failed Exports': 'Başarısız Dışa Aktarım',
    'Contractor Count': 'Alt Yüklenici Sayısı',
    'Pending Approvals': 'Bekleyen Onaylar',
    'Updated Settings': 'Güncellenen Ayarlar',
    'Pending Changes': 'Bekleyen Değişiklikler',
    'Risk Score (Probability x Severity)': 'Risk Skoru (Olasılık x Şiddet)',
    'Control Measures (Count/Effectiveness)': 'Kontrol Tedbirleri (Adet/Etkinlik)',
    'System policy updates': 'Sistem politika güncellemeleri',
    'Scaffold inspection': 'İskele denetimi',
    'Unsafe stacking': 'Güvensiz istifleme',
    'Electrical risk review': 'Elektrik risk değerlendirmesi',
    'Hot work permit': 'Sıcak çalışma izni',
    'Minor finger injury': 'Hafif parmak yaralanması',
    'Hafif parmak yaralanması': 'Незначительная травма пальца',
    'Weekly internal audit': 'Haftalık iç tetkik',
    'Lifting equipment check': 'Kaldırma ekipmanı kontrolü',
    'Waste segregation tracking': 'Atık ayrıştırma takibi',
    'Evacuation drill exercise': 'Tahliye tatbikatı uygulaması',
    'New employees onboarded': 'Yeni çalışan oryantasyonu',
    'Work-at-height training': 'Yüksekte çalışma eğitimi',
    'Helmet stock update': 'Baret stok güncellemesi',
    'Health screening': 'Sağlık taraması',
    'Regulation update': 'Mevzuat güncellemesi',
    'SOP revision': 'SOP revizyonu',
    'CAPA closure batch': 'DÖF kapatma paketi',
    'KPI summary refresh': 'KPI özet yenilemesi',
    'Weekly report generated': 'Haftalık rapor oluşturuldu',
    'Exports executed': 'Dışa aktarımlar tamamlandı',
    'Active projects': 'Aktif projeler',
    'Approved contractors': 'Onaylı alt yükleniciler'
  },
  ru: {
    Dashboard: 'Панель управления',
    Project: 'Проект',
    'HSE Compliance Platform': 'Платформа соответствия HSE',
    'Designed and developed by Erdem Cetin': 'Разработано и спроектировано Erdem Cetin',
    'Design and Development: Erdem Cetin © 2026 All Rights Reserved.': 'Дизайн и разработка: Erdem Cetin © 2026 Все права защищены.',
    OPERATIONS: 'ОПЕРАЦИИ',
    WORKFORCE: 'ПЕРСОНАЛ',
    REPORTS: 'ОТЧЕТЫ',
    SYSTEM: 'СИСТЕМА',
    'Project performance and executive summary.': 'Производительность проекта и сводка для руководства.',
    'Corporate report center and automated output generation.': 'Корпоративный центр отчетности и автоматическое формирование отчетов.',
    'Project master data management.': 'Управление основными данными проекта.',
    'Subcontractor master data management.': 'Управление основными данными подрядчиков.',
    'System configuration and governance settings.': 'Системные настройки и параметры управления.',
    'Permit to Work': 'Наряд-допуск',
    'Weather Summary': 'Сводка погоды',
    'Wind Summary': 'Сводка ветра',
    'Incident Executive KPI Summary': 'Сводка KPI по инцидентам для руководства',
    'HSE Activity by Project': 'HSE-активность по проектам',
    'HSE Activity by Department': 'HSE-активность по отделам',
    'Project Distribution': 'Распределение по проектам',
    'Department Distribution': 'Распределение по отделам',
    Period: 'Период',
    Daily: 'Ежедневно',
    Weekly: 'Еженедельно',
    Monthly: 'Ежемесячно',
    Yearly: 'Ежегодно',
    'Custom Date Range': 'Произвольный период',
    Comparison: 'Сравнение',
    None: 'Нет',
    From: 'С',
    To: 'По',
    'Select Projects': 'Выберите проекты',
    'Select Departments': 'Выберите отделы',
    'Apply Filters': 'Применить фильтры',
    'Reset Filters': 'Сбросить фильтры',
    'Executive PDF Output': 'PDF-отчет для руководства',
    'Safety Performance': 'Показатели безопасности',
    'Safety Score Trend': 'Тренд показателя безопасности',
    'Overall Safety Index': 'Общий индекс безопасности',
    'Check Details': 'Проверить',
    Increase: 'Рост',
    Decrease: 'Снижение',
    Stable: 'Стабильно',
    Average: 'Среднее',
    Minimum: 'Мин',
    Maximum: 'Макс',
    Last: 'Последнее',
    Compliance: 'Соответствие',
    Strength: 'Сильный',
    Critical: 'Критично',
    'Fit for Work Trend': 'Тренд годности к работе',
    'Health Status Distribution': 'Распределение статусов здоровья',
    'Open Risk Trend': 'Тренд открытых рисков',
    'Risk Level Distribution': 'Распределение уровней риска',
    'Waste Generation Trend': 'Тренд образования отходов',
    'Recycling Trend': 'Тренд переработки',
    'Compliance Score': 'Оценка соответствия',
    'Audit and Non-Conformity View': 'Обзор аудитов и несоответствий',
    'Training Participation Trend': 'Тренд участия в обучении',
    'Training Completion Rate': 'Уровень завершения обучения',
    'PPE Receipt Trend': 'Тренд поступления СИЗ',
    'PPE Stock Distribution': 'Распределение запасов СИЗ',
    'Equipment Inventory': 'Инвентаризация оборудования',
    'Certification Compliance Rate': 'Уровень соответствия сертификации',
    'Inspection Count Trend': 'Тренд количества инспекций',
    'Finding Status': 'Статус замечаний',
    'Incident Trend (Last 7 Periods)': 'Тренд инцидентов (последние 7 периодов)',
    'Incident Category Distribution': 'Распределение категорий инцидентов',
    'PTW Volume Trend': 'Тренд объема нарядов-допусков',
    'PTW Status Distribution': 'Распределение статусов нарядов-допусков',
    'PTW Violation': 'Нарушение наряда-допуска',
    'PTW Active': 'Активные наряды-допуски',
    'PTW Closed': 'Закрытые наряды-допуски',
    'Hot Work': 'Огневые работы',
    'Confined Space': 'Работы в замкнутом пространстве',
    'Work at Height': 'Работы на высоте',
    'Near Miss': 'Почти произошедший инцидент',
    'Safe Man-Hours - Daily': 'Безопасные человеко-часы - день',
    'Safe Man-Hours - Weekly': 'Безопасные человеко-часы - неделя',
    'Safe Man-Hours - Monthly': 'Безопасные человеко-часы - месяц',
    'Safe Man-Hours - Total': 'Безопасные человеко-часы - всего',
    'Daily x weekly x monthly annualized capacity': 'Годовой эквивалент на основе дневной, недельной и месячной мощности',
    'Permit to Work Volume Trend': 'Тренд объема нарядов-допусков',
    'Permit to Work Status Distribution': 'Распределение статусов нарядов-допусков',
    'Permit to Work Violation': 'Нарушение наряда-допуска',
    'Wind Direction': 'Направление ветра',
    Open: 'Открыто',
    'All Projects': 'Все проекты',
    'First Aid Cases': 'Случаи первой помощи',
    'Near Miss Reports': 'Сообщения о почти случившихся инцидентах',
    'Critical Incidents': 'Критические инциденты',
    'Training Completion': 'Завершение обучения',
    'PPE Stock Health': 'Состояние запасов СИЗ',
    'Environmental Incidents': 'Экологические инциденты',
    'High Risks': 'Высокие риски',
    'Medium Risks': 'Средние риски',
    'Low Risks': 'Низкие риски',
    'Inspection Overdue': 'Просроченные инспекции',
    'Out of Service': 'Вне эксплуатации',
    'Water Consumption': 'Потребление воды',
    'Electricity Consumption': 'Потребление электроэнергии',
    'Fuel Consumption': 'Расход топлива',
    'Legal Compliance': 'Юридическое соответствие',
    'Open Non-Conformities': 'Открытые несоответствия',
    'Closed CAPA': 'Закрытые CAPA',
    'Open CAPA': 'Открытые CAPA',
    'Workforce Safety': 'Безопасность персонала',
    'Inspections & Observations': 'Инспекции и предписания',
    Training: 'Обучение',
    'PPE Management': 'Управление СИЗ',
    'Equipment Management': 'Управление оборудованием',
    'Occupational Health': 'Охрана труда',
    'Risk Management': 'Управление рисками',
    Environmental: 'Экология',
    'Emergency Preparedness': 'Готовность к ЧС',
    'Reports & Analytics': 'Отчеты и аналитика',
    'Total Workforce': 'Общая численность',
    'Daily Workforce': 'Ежедневная численность',
    'Safe Man-Hours (Daily / Weekly / Monthly / Total)': 'Безопасные человеко-часы (день / неделя / месяц / всего)',
    'Open Findings': 'Открытые замечания',
    'Closed Findings': 'Закрытые замечания',
    'Safety Observations': 'Предписания по безопасности',
    'Positive Observations': 'Позитивные предписания',
    'Overdue Actions': 'Просроченные действия',
    'Weekly Inspection Performance': 'Недельная эффективность инспекций',
    'Total Employees': 'Всего сотрудников',
    'Fit for Work': 'Годен к работе',
    'Restricted Duty': 'Ограниченный режим труда',
    'Vaccination Status': 'Статус вакцинации',
    'Open Risks': 'Открытые риски',
    'Closed Risks': 'Закрытые риски',
    'Waste Generated': 'Образовано отходов',
    'Waste Recycled': 'Переработано отходов',
    'Internal Audits': 'Внутренние аудиты',
    'External Audits': 'Внешние аудиты',
    'Emergency Drills': 'Аварийные учения',
    'Fire Drills': 'Пожарные учения',
    'Project Comparison': 'Сравнение проектов',
    'Department Comparison': 'Сравнение отделов',
    'Inspected Areas': 'Проверенные зоны',
    'Findings Count': 'Количество замечаний',
    'PPE Received': 'Получено СИЗ',
    'PPE Issued': 'Выдано СИЗ',
    'Employee Count': 'Количество сотрудников',
    'Absent Count': 'Отсутствующие сотрудники',
    'Reports Generated': 'Сформированные отчеты',
    'Reports Shared': 'Отправленные отчеты',
    'Updated Settings': 'Обновленные настройки',
    'Pending Changes': 'Ожидающие изменения'
  }
};

function localizeText(label: string, language: Language): string {
  const trToEnMap = Object.fromEntries(Object.entries(textByLanguage.tr ?? {}).map(([english, translated]) => [translated, english]));

  if (language === 'en') {
    return (trToEnMap[label] as string | undefined) ?? label;
  }

  if (language === 'ru') {
    const ruDirect = textByLanguage.ru?.[label];
    if (ruDirect) {
      return ruDirect;
    }

    const englishLabel = (trToEnMap[label] as string | undefined) ?? label;
    return textByLanguage.ru?.[englishLabel] ?? englishLabel;
  }

  return textByLanguage[language]?.[label] ?? label;
}

const buildTurkishReplacementMap = (language: Exclude<Language, 'tr'>): Map<string, string> => {
  const replacements = new Map<string, string>();
  const trEntries = Object.entries(textByLanguage.tr ?? {});

  trEntries.forEach(([english, turkish]) => {
    if (language === 'en') {
      replacements.set(turkish, english);
      return;
    }

    const russian = textByLanguage.ru?.[english];
    replacements.set(turkish, russian ?? english);

    if (russian) {
      replacements.set(english, russian);
    }
  });

  Object.entries(moduleLabelsByLanguage.tr).forEach(([key, turkishLabel]) => {
    const englishLabel = moduleLabelsByLanguage.en[key as ModuleKey];
    const russianLabel = moduleLabelsByLanguage.ru[key as ModuleKey];
    replacements.set(turkishLabel, language === 'en' ? englishLabel : russianLabel ?? englishLabel);

    if (language === 'ru') {
      replacements.set(englishLabel, russianLabel ?? englishLabel);
    }
  });

  Object.entries(uiText.tr).forEach(([key, turkishLabel]) => {
    const englishLabel = uiText.en[key] ?? turkishLabel;
    const russianLabel = uiText.ru[key] ?? englishLabel;
    replacements.set(turkishLabel, language === 'en' ? englishLabel : russianLabel);

    if (language === 'ru') {
      replacements.set(englishLabel, russianLabel);
    }
  });

  const extraReplacements: Record<Exclude<Language, 'tr'>, Array<[string, string]>> = {
    en: [
      ['Haz', 'Jun'],
      ['Tem', 'Jul'],
      ['Agu', 'Aug'],
      ['Eyl', 'Sep'],
      ['Eki', 'Oct'],
      ['Kas', 'Nov'],
      ['Ara', 'Dec'],
      ['Artış', 'Increase'],
      ['Azalış', 'Decrease'],
      ['Toplam', 'Total'],
      ['Ortalama', 'Average'],
      ['Son', 'Last'],
      ['Güçlü', 'Strong'],
      ['Kritik', 'Critical'],
      ['Sabit', 'Stable'],
      ['Servis Dışı', 'Out of Service'],
      ['Açık Bulgu', 'Open Finding'],
      ['Kapalı Bulgu', 'Closed Finding'],
      ['Pozitif Gözlem', 'Positive Observation'],
      ['Dağılım', 'Distribution']
    ],
    ru: [
      ['Haz', 'Июн'],
      ['Tem', 'Июл'],
      ['Agu', 'Авг'],
      ['Eyl', 'Сен'],
      ['Eki', 'Окт'],
      ['Kas', 'Ноя'],
      ['Ara', 'Дек'],
      ['Artış', 'Рост'],
      ['Azalış', 'Снижение'],
      ['Toplam: ', 'Всего: '],
      ['Ortalama: ', 'Среднее: '],
      ['Son: ', 'Последнее: '],
      ['Güçlü', 'Сильный'],
      ['Kritik', 'Критично'],
      ['Sabit', 'Стабильно'],
      ['Periyot', 'Период'],
      ['Günlük', 'Ежедневно'],
      ['Haftalık', 'Еженедельно'],
      ['Aylık', 'Ежемесячно'],
      ['Yıllık', 'Ежегодно'],
      ['Özel Tarih Aralığı', 'Произвольный период'],
      ['Karşılaştırma', 'Сравнение'],
      ['Yok', 'Нет'],
      ['Projeleri Seç', 'Выберите проекты'],
      ['Departmanları Seç', 'Выберите отделы'],
      ['Filtreleri Uygula', 'Применить фильтры'],
      ['Filtreleri Sıfırla', 'Сбросить фильтры'],
      ['Yönetici PDF Çıktısı', 'PDF-отчет для руководства'],
      ['Güvenlik Performansı', 'Показатели безопасности'],
      ['Güvenlik Skoru Trendi', 'Тренд показателя безопасности'],
      ['Genel Güvenlik Endeksi', 'Общий индекс безопасности'],
      ['Kontrol Et', 'Проверить'],
      ['Güvenli Adam-Saat - Günlük', 'Безопасные человеко-часы - день'],
      ['Güvenli Adam-Saat - Haftalık', 'Безопасные человеко-часы - неделя'],
      ['Güvenli Adam-Saat - Aylık', 'Безопасные человеко-часы - месяц'],
      ['Güvenli Adam-Saat - Toplam', 'Безопасные человеко-часы - всего'],
      ['Günlük, haftalık ve aylık güvenli kapasitenin yıllık toplamı', 'Годовой итог безопасной мощности на основе дневных, недельных и месячных значений'],
      ['Sıcak Çalışma', 'Огневые работы'],
      ['Kapalı Alan', 'Замкнутое пространство'],
      ['Yüksekte Çalışma', 'Работы на высоте'],
      ['Denetim Sayısı Trendi', 'Тренд количества инспекций'],
      ['Bulgu Durumu', 'Статус замечаний'],
      ['Toplam Olay', 'Всего инцидентов'],
      ['Hedef <= 10', 'Цель <= 10'],
      ['Sıfır yaralanma hedefi', 'Цель: ноль травм'],
      ['Vardiya bazlı izle', 'Отслеживать по сменам'],
      ['Davranış aksiyonları', 'Поведенческие действия'],
      ['Mühendislik kontrolü', 'Инженерные меры контроля'],
      ['Bildirim ivmesi', 'Динамика уведомлений'],
      ['Anında eskalasyon', 'Немедленная эскалация'],
      ['Toolbox Konuşmaları', 'Тулбокс-брифинги'],
      ['Süresi Dolan Sertifikalar', 'Сертификаты с истекшим сроком'],
      ['Dağılım', 'Распределение'],
      ['Mekanik', 'Механика'],
      ['İnşaat', 'Строительство'],
      ['Elektrik', 'Электрика'],
      ['İdari', 'Административный'],
      ['Servis Dışı', 'Вне эксплуатации'],
      ['Kronik Hastalık', 'Хронические заболевания'],
      ['Bulaşıcı Hastalık', 'Инфекционные заболевания'],
      ['Aşılama', 'Вакцинация'],
      ['Kritik Risk Adedi', 'Количество критических рисков'],
      ['Atık Üretimi Trendi', 'Тренд образования отходов'],
      ['Geri Dönüşüm Trendi', 'Тренд переработки'],
      ['Mevzuata Uyum', 'Соответствие законодательству'],
      ['Uyum Skoru', 'Оценка соответствия'],
      ['Tetkik ve Uygunsuzluk Görünümü', 'Обзор аудитов и несоответствий'],
      ['Yasal Uygunluk', 'Юридическое соответствие'],
      ['Açık Uygunsuzluk', 'Открытое несоответствие'],
      ['Açık Bulgu', 'Открытое замечание'],
      ['Kapalı Bulgu', 'Закрытое замечание'],
      ['Pozitif Gözlem', 'Позитивное предписание'],
      ['Uygunluk', 'Соответствие'],
      ['Rüzgar Yönü', 'Направление ветра'],
      ['Açık', 'Открыто'],
      ['Proje performansı ve yönetici özeti.', 'Производительность проекта и сводка для руководства.'],
      ['Erdem Cetin tarafından tasarlandı ve geliştirildi', 'Разработано и спроектировано Erdem Cetin'],
      ['Tasarım ve Geliştirme: Erdem Cetin © 2026 Tüm Hakları Saklıdır.', 'Дизайн и разработка: Erdem Cetin © 2026 Все права защищены.'],
      ['Total Inspections', 'Всего инспекций'],
      ['Weekly Performance', 'Еженедельная эффективность'],
      ['Actions Overdue', 'Просроченные действия'],
      ['Near Misses', 'Почти произошедшие инциденты'],
      ['Unsafe Acts', 'Небезопасные действия'],
      ['Unsafe Conditions', 'Небезопасные условия'],
      ['Near Misses Bildirimleri', 'Сообщения о почти случившихся инцидентах'],
      ['Weekly Inductions', 'Еженедельные инструктажи'],
      ['Monthly Inductions', 'Ежемесячные инструктажи'],
      ['Total Inductions', 'Всего инструктажей'],
      ['Upcoming Trainings', 'Предстоящие обучения'],
      ['Stokta', 'На складе'],
      ['Dağıtılan', 'Выдано'],
      ['Low Stock Items', 'Позиции с низким запасом'],
      ['Teslim Alınan', 'Получено'],
      ['Tüketim Trendi', 'Тренд потребления'],
      ['Total Equipment', 'Всего оборудования'],
      ['Due Inspections', 'Проверки к выполнению'],
      ['Yüksek Riskler', 'Высокие риски'],
      ['Orta Riskler', 'Средние риски'],
      ['Düşük Riskler', 'Низкие риски'],
      ['Экологияsel Инциденты', 'Экологические инциденты'],
      ['Su Tüketimi', 'Потребление воды'],
      ['Elektrik Tüketimi', 'Потребление электроэнергии'],
      ['Yakıt Tüketimi', 'Расход топлива'],
      ['Yasal Uyum', 'Юридическое соответствие'],
      ['CAPA Open', 'Открытые CAPA'],
      ['CAPA Closed', 'Закрытые CAPA'],
      ['PTW Active', 'Активные наряды-допуски'],
      ['PTW Closed', 'Закрытые наряды-допуски'],
      ['PTW Volume Trend', 'Тренд объема нарядов-допусков'],
      ['PTW Status Distribution', 'Распределение статусов нарядов-допусков'],
      ['PTW Violation', 'Нарушение наряда-допуска'],
      ['Taslak', 'Черновик'],
      ['Onay Bekliyor', 'Ожидает согласования'],
      ['Aktif', 'Активен'],
      ['Askıya Alındı', 'Приостановлен'],
      ['Tamamlandı', 'Завершен'],
      ['İptal', 'Отменен'],
      ['Hot Work', 'Огневые работы'],
      ['Confined Space', 'Работы в замкнутом пространстве'],
      ['Work at Height', 'Работы на высоте'],
      ['Yüksekte Çalışma', 'Работы на высоте'],
      ['İşi Talep Eden', 'Инициатор работ'],
      ['İşi Veren', 'Выдающий наряд'],
      ['HSE Sorumlusu', 'Ответственный HSE'],
      ['Yetkili Onaylayan', 'Уполномоченный утверждающий'],
      ['Ad Soyad', 'ФИО'],
      ['Görevi', 'Должность'],
      ['Firma', 'Компания'],
      ['Eğitim Durumu', 'Статус обучения'],
      ['İmza', 'Подпись'],
      ['Yeni Personel Ekle', 'Добавить сотрудника'],
      ['İşin Adı', 'Наименование работы'],
      ['Çalışma Alanı', 'Зона выполнения работ'],
      ['İşin Açıklaması', 'Описание работы'],
      ['Yapılacak İş', 'Планируемая работа'],
      ['Çalışma Koşulları', 'Условия выполнения работ'],
      ['Başlangıç Tarihi', 'Дата начала'],
      ['Başlangıç Saati', 'Время начала'],
      ['Bitiş Tarihi', 'Дата окончания'],
      ['Bitiş Saati', 'Время окончания'],
      ['Düşme Riski', 'Риск падения'],
      ['Askıda Yük', 'Подвешенный груз'],
      ['Kimyasal', 'Химические вещества'],
      ['Gürültü', 'Шум'],
      ['Toz', 'Пыль'],
      ['Hareketli Makine', 'Движущееся оборудование'],
      ['Hava Koşulları', 'Погодные условия'],
      ['Diğer', 'Другое'],
      ['Yaşam Hattı', 'Страховочная линия'],
      ['Düşüş Durdurma Sistemi', 'Система остановки падения'],
      ['Konumlandırma Sistemi', 'Система позиционирования'],
      ['Tutucu Sistem', 'Удерживающая система'],
      ['Kurtarma Sistemi', 'Система спасения'],
      ['Emniyet Kemeri', 'Страховочная привязь'],
      ['Çift Lanyard', 'Двойной строп'],
      ['Ankraj Noktası', 'Анкерная точка'],
      ['Kurtarma Ekipmanı', 'Спасательное оборудование'],
      ['Malzemeler', 'Материалы'],
      ['El Aletleri', 'Ручные инструменты'],
      ['Makineler', 'Машины'],
      ['Platformlar', 'Платформы'],
      ['İskele', 'Леса'],
      ['Merdiven', 'Лестница'],
      ['Vinç', 'Кран'],
      ['Mobil Platform', 'Мобильная платформа'],
      ['Kontrol', 'Проверка'],
      ['Açıklama', 'Описание'],
      ['Kontrol Satırı Ekle', 'Добавить строку проверки'],
      ['Tedbir', 'Мера'],
      ['Sorumlu', 'Ответственный'],
      ['Termin', 'Срок'],
      ['Tedbir Ekle', 'Добавить меру'],
      ['PTW Hazırlayan', 'Подготовил PTW'],
      ['HSE Onayı', 'Согласование HSE'],
      ['Proje Müdürü', 'Руководитель проекта'],
      ['İşveren Temsilcisi', 'Представитель заказчика'],
      ['Dijital İmza', 'Цифровая подпись'],
      ['Tarih', 'Дата'],
      ['Çalışma Başladı', 'Работа начата'],
      ['Çalışma Bitti', 'Работа завершена'],
      ['Günlük Kayıt Ekle', 'Добавить ежедневную запись'],
      ['Eklenen Personel', 'Добавленный сотрудник'],
      ['Ayrılan Personel', 'Выбывший сотрудник'],
      ['Onaylayan', 'Согласовал'],
      ['Ekip Değişikliği Ekle', 'Добавить изменение состава'],
      ['İş Tamamlandı', 'Работа завершена'],
      ['Alan Güvenli', 'Зона безопасна'],
      ['Malzemeler Toplandı', 'Материалы собраны'],
      ['PTW Kapatıldı', 'Наряд-допуск закрыт'],
      ['Kapatan', 'Закрыл'],
      ['Ek Türü', 'Тип вложения'],
      ['Dosya Adı', 'Имя файла'],
      ['Henüz dosya eklenmedi.', 'Файлы пока не добавлены.'],
      ['Risk Score', 'Оценка риска'],
      ['PPE Stock Health', 'Состояние запасов СИЗ'],
      ['Safe Man-Hours - Daily', 'Безопасные человеко-часы - день'],
      ['Safe Man-Hours - Weekly', 'Безопасные человеко-часы - неделя'],
      ['Safe Man-Hours - Monthly', 'Безопасные человеко-часы - месяц'],
      ['Safe Man-Hours - Total', 'Безопасные человеко-часы - всего'],
      ['Daily x weekly x monthly annualized capacity', 'Годовой эквивалент на основе дневной, недельной и месячной мощности'],
      ['Near Miss', 'Почти произошедший инцидент'],
      ['Проект Сравнение', 'Сравнение проектов'],
      ['Departman Сравнение', 'Сравнение отделов'],
      ['Еженедельно Performans', 'Еженедельная эффективность'],
      ['İlk Yardım Vakaları', 'Случаи первой помощи'],
      ['Hedef <= 5', 'Цель <= 5'],
      ['Почти произошедшие инциденты Bildirimleri', 'Сообщения о почти случившихся инцидентах'],
      ['Kritично Инциденты', 'Критические инциденты'],
      ['Sertifikasyon Uyum Oranı', 'Уровень соответствия сертификации'],
      ['Denetimi Gecikmiş', 'Просроченные инспекции'],
      ['Mobil', 'Мобильное'],
      ['Ağır', 'Тяжелое'],
      ['Kaldırma', 'Подъемное'],
      ['Тренд открытых рисковi', 'Тренд открытых рисков']
    ]
  };

  extraReplacements[language].forEach(([from, to]) => {
    replacements.set(from, to);
  });

  return replacements;
};

const replaceByDictionary = (input: string, replacements: Array<[string, string]>): string => {
  let output = input;
  replacements.forEach(([from, to]) => {
    if (!from || from === to || !output.includes(from)) {
      return;
    }

    // Avoid breaking words like "Temizlik" while still replacing standalone tokens like "Tem".
    if (/^[\p{L}\p{N}]+$/u.test(from)) {
      const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const boundedPattern = new RegExp(`(^|[^\\p{L}\\p{N}])(${escaped})(?=[^\\p{L}\\p{N}]|$)`, 'gu');
      output = output.replace(boundedPattern, `$1${to}`);
      return;
    }

    output = output.split(from).join(to);
  });
  return output;
};

const applyTextReplacementsToNode = (root: Node, replacementEntries: Array<[string, string]>) => {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();

  while (node) {
    const textNode = node as Text;
    const current = textNode.nodeValue ?? '';
    if (current.trim()) {
      const replaced = replaceByDictionary(current, replacementEntries);
      if (replaced !== current) {
        textNode.nodeValue = replaced;
      }
    }
    node = walker.nextNode();
  }
};

function kpiTermHint(label: string, language: Language): string {
  if (language !== 'tr') {
    return '';
  }

  const normalized = label.trim().toLowerCase();
  if (normalized === 'trir') {
    return '(toplam kayda geçen olay sıklık oranı)';
  }
  if (normalized === 'ltifr') {
    return '(kayıp zamanlı yaralanma sıklık oranı)';
  }
  if (normalized === 'lti') {
    return '(kayıp zamanlı yaralanma)';
  }
  if (normalized === 'mti') {
    return '(tıbbi müdahale gerektiren yaralanma)';
  }
  return '';
}

const groupedMenu: Array<{ key: GroupKey; title: string; modules: ModuleKey[] }> = [
  {
    key: 'operations',
    title: 'OPERASYONLAR',
    modules: [
      'inspections',
      'observations',
      'risk-assessments',
      'permit-to-work',
      'incidents',
      'equipment-management',
      'environmental',
      'emergency-preparedness'
    ]
  },
  {
    key: 'workforce',
    title: 'İŞ GÜCÜ',
    modules: ['employees', 'trainings', 'ppe-stocks', 'occupational-health']
  },
  {
    key: 'compliance',
    title: 'Compliance',
    modules: ['legal-register', 'documents', 'action-tracker', 'kpis-analytics']
  },
  {
    key: 'reports',
    title: 'RAPORLAR',
    modules: ['reports']
  },
  {
    key: 'system',
    title: 'SİSTEM',
    modules: ['projects', 'departments', 'contractors', 'settings']
  }
];

const locationTypes = [
  'Ofis',
  'Depo',
  'Atölye',
  'Showroom',
  'Konaklama',
  'Proje Sahası',
  'Alt Yüklenici Çalışma Alanı',
  'Diğer'
];

const workforceDepartmentOptions = [
  'Elektrik',
  'Mekanik',
  'İnşaat',
  'İK',
  'İSG',
  'Proje Yönetimi',
  'Depo',
  'Bakım',
  'Atölye',
  'Teknik Ofis',
  'Finans',
  'PTO',
  'Dizayn'
];

const workforceDepartmentTranslations: Record<string, string> = {
  'Elektrik': 'Электротехнические',
  'Mekanik': 'Механика',
  'İnşaat': 'Строительство',
  'İK': 'Управление кадрами',
  'İSG': 'Охрана труда',
  'Proje Yönetimi': 'Управление проектами',
  'Depo': 'Склад',
  'Bakım': 'Техническое обслуживание',
  'Atölye': 'Мастерская',
  'Teknik Ofis': 'Техническое управление',
  'Finans': 'Финансы',
  'PTO': 'ПТО',
  'Dizayn': 'Дизайн'
};

const translateDepartment = (dept: string, lang: string): string => {
  if (lang === 'ru' && workforceDepartmentTranslations[dept]) {
    return workforceDepartmentTranslations[dept];
  }
  return dept;
};

const trainingTypeOptions = [
  'İSG Oryantasyonu',
  'Yüksekte Çalışma',
  'Kapalı Alan Çalışması',
  'Kaldırma Operasyonları',
  'Elektrik Güvenliği',
  'Yangın Güvenliği',
  'İlk Yardım',
  'Savunma Odaklı Sürüş',
  'İskele Kurulumu/Kullanımı',
  'Kazı Güvenliği',
  'Kaynak İşleri Güvenliği',
  'Rostechnadzor'
];

const trainingCategoryOptions = ['Zorunlu İSG', 'Teknik Güvenlik', 'Operasyonel Güvenlik', 'Sertifikasyon', 'Yasal Uyum'];

const trainingDepartmentOptions = ['İnşaat', 'Mekanik', 'Elektrik', 'İSG', 'Depo', 'Atölye', 'Proje Yönetimi'];

const trainingPositionOptions = ['Saha Mühendisi', 'Usta', 'Operatör', 'Teknisyen', 'Formen', 'İSG Uzmanı'];

const ppeCategoryOptions = ['Baş Koruma', 'Göz/Yüz Koruma', 'Solunum Koruma', 'El Koruma', 'Ayak Koruma', 'Vücut Koruma', 'Yüksekten Düşüş Koruma'];
const ppeItemOptions = ['Baret', 'Koruyucu Gözlük', 'Toz Maskesi FFP3', 'Nitril Eldiven', 'Çelik Burunlu Bot', 'Reflektif Yelek', 'Tam Vücut Emniyet Kemeri'];
const ppeUnitOptions = ['Adet', 'Çift', 'Kutu', 'Set'];
const ppeWarehouseOptions = ['Merkez Depo', 'Tünel Deposu', 'Saha Konteyner Depo', 'Lojistik Geçici Depo'];
const ppeTransactionTypeOptions: Array<{ value: PpeTransactionType; label: string }> = [
  { value: 'STOK_GIRISI', label: 'Stok Girişi (Satın Alma)' },
  { value: 'STOK_CIKISI', label: 'Stok Çıkışı (Personele Dağıtım)' },
  { value: 'STOGA_IADE', label: 'Stoğa İade' },
  { value: 'HASARLI_HURDA', label: 'Hasarlı / Hurda' },
  { value: 'STOK_DUZELTME', label: 'Stok Düzeltme' }
];

const ppeLifecycleOptions: Array<{ value: PpeTransactionLifecycle; label: string }> = [
  { value: 'TAMAMLANDI', label: 'Tamamlandı' },
  { value: 'ONAY_BEKLIYOR', label: 'Onay Bekliyor' },
  { value: 'IPTAL', label: 'İptal' }
];

const ptwDurumlar: PtwStatus[] = ['Taslak', 'Onay Bekliyor', 'Aktif', 'Askıya Alındı', 'Tamamlandı', 'İptal'];

const ptwStatusRuLabel: Record<PtwStatus, string> = {
  Taslak: 'Черновик',
  'Onay Bekliyor': 'Ожидает согласования',
  Aktif: 'Активен',
  'Askıya Alındı': 'Приостановлен',
  Tamamlandı: 'Завершен',
  'İptal': 'Отменен'
};

const ptwTypeRuLabel: Record<string, string> = {
  'Yüksekte Çalışma': 'Работы на высоте',
  'Kapalı Alan Çalışması': 'Работы в замкнутом пространстве',
  'Sıcak Çalışma': 'Огневые работы',
  'Kazı Çalışması': 'Земляные работы',
  'Elektrik Çalışması': 'Электромонтажные работы',
  'Kaldırma Operasyonu': 'Подъемные операции',
  'Bakım-Onarım': 'Техническое обслуживание и ремонт'
};

const ptwTehlikeSecenekleri = [
  'Yüksekte Çalışma',
  'Düşme Riski',
  'Elektrik',
  'Askıda Yük',
  'Kapalı Alan',
  'Sıcak Çalışma',
  'Kimyasal',
  'Gürültü',
  'Toz',
  'Hareketli Makine',
  'Hava Koşulları',
  'Diğer'
];

const ptwGuvenlikSistemleri = [
  'Yaşam Hattı',
  'Düşüş Durdurma Sistemi',
  'Konumlandırma Sistemi',
  'Tutucu Sistem',
  'Kurtarma Sistemi',
  'Emniyet Kemeri',
  'Çift Lanyard',
  'Ankraj Noktası',
  'Kurtarma Ekipmanı'
];

const ptwEkipmanSecenekleri = [
  'Malzemeler',
  'El Aletleri',
  'Makineler',
  'Platformlar',
  'İskele',
  'Merdiven',
  'Vinç',
  'Mobil Platform'
];

const ptwDosyaTipleri = ['Fotoğraf', 'PDF', 'Risk Assessment', 'Method Statement', 'Toolbox Talk', 'Sertifikalar'];

const ptwAttachmentTypeByLabel: Record<string, PtwAttachmentApiType> = {
  'Фото': 'PHOTO',
  'Фотографии': 'PHOTO',
  Fotoğraf: 'PHOTO',
  PDF: 'PDF',
  'Risk Assessment': 'RISK_ASSESSMENT',
  'Оценка рисков': 'RISK_ASSESSMENT',
  'Method Statement': 'METHOD_STATEMENT',
  'Методика работ': 'METHOD_STATEMENT',
  'Toolbox Talk': 'TOOLBOX_TALK',
  'Инструктаж Toolbox': 'TOOLBOX_TALK',
  Sertifikalar: 'CERTIFICATE',
  Сертификаты: 'CERTIFICATE'
};

const ptwAttachmentLabelByType: Record<PtwAttachmentApiType, string> = {
  PHOTO: 'Фотографии',
  PDF: 'PDF',
  RISK_ASSESSMENT: 'Оценка рисков',
  METHOD_STATEMENT: 'Методика работ',
  TOOLBOX_TALK: 'Инструктаж Toolbox',
  CERTIFICATE: 'Сертификаты',
  OTHER: 'Прочее'
};

const ptwStatusFlow: PtwStatus[] = ['Taslak', 'Onay Bekliyor', 'Aktif', 'Tamamlandı'];

const ptwStatusFlowRuLabel: Record<PtwStatus, string> = {
  Taslak: 'Черновик',
  'Onay Bekliyor': 'На согласовании',
  Aktif: 'Активен',
  'Askıya Alındı': 'Приостановлен',
  Tamamlandı: 'Закрыт',
  'İptal': 'Отменен'
};

const legalCategoryOptions = [
  'İş Sağlığı ve Güvenliği',
  'Çevre',
  'Yangın Güvenliği',
  'Elektrik Güvenliği',
  'Kaldırma Ekipmanları',
  'Makineler',
  'Atık Yönetimi',
  'Kimyasal Güvenlik'
];

const legalAuthorityOptions = [
  'Çalışma ve Sosyal Güvenlik Bakanlığı',
  'Çevre, Şehircilik ve İklim Değişikliği Bakanlığı',
  'Enerji ve Tabii Kaynaklar Bakanlığı',
  'İtfaiye Daire Başkanlığı',
  'TSE',
  'Rostechnadzor'
];

const legalRoleOptions: Array<{ value: LegalUserRole; label: string }> = [
  { value: 'HSE_MANAGER', label: 'İSG Yöneticisi' },
  { value: 'CORPORATE_HSE_MANAGER', label: 'Kurumsal İSG Yöneticisi' },
  { value: 'COMPLIANCE_MANAGER', label: 'Uyumluluk Yöneticisi' },
  { value: 'ISO_REPRESENTATIVE', label: 'ISO Yönetim Temsilcisi' },
  { value: 'VIEWER', label: 'Görüntüleyici' }
];

const controlledDocumentCategoryOptions = [
  'Çalışma İzinleri',
  'Toolbox Talk',
  'İSG Prosedürleri',
  'SOP',
  'İş Talimatları',
  'Risk Değerlendirme Şablonları',
  'JSA Şablonları',
  'Denetim Kontrol Listeleri',
  'Olay Bildirim Formları',
  'Acil Durum Müdahale Planı',
  'LOTO Prosedürü',
  'Yüksekte Çalışma Prosedürü',
  'Kazı Prosedürü',
  'İskele Prosedürü',
  'Çevre Prosedürleri',
  'Şirket Politikaları',
  'Kurumsal Standartlar'
];

const controlledDocumentTypeOptions = ['Form', 'Prosedür', 'Talimat', 'Şablon', 'Kontrol Listesi', 'Plan', 'Politika', 'Standart'];

const controlledDocumentStatusOptions: Array<{ value: ControlledDocumentStatus; label: string }> = [
  { value: 'TASLAK', label: 'Taslak' },
  { value: 'GOZDEN_GECIRMEDE', label: 'Gözden Geçirme Aşamasında' },
  { value: 'ONAYLANDI', label: 'Onaylandı' },
  { value: 'GECERSIZ', label: 'Geçersiz (Obsolete)' }
];

const corporateReportOptions: CorporateReportOption[] = [
  { key: 'executive-dashboard', label: 'Yönetici HSE Panosu', group: 'module' },
  { key: 'inspections', label: 'Denetimler', group: 'module' },
  { key: 'observations', label: 'Gözlemler', group: 'module' },
  { key: 'risk-assessments', label: 'Risk Değerlendirmeleri', group: 'module' },
  { key: 'incidents', label: 'Olaylar / Kazalar', group: 'module' },
  { key: 'trainings', label: 'Eğitimler', group: 'module' },
  { key: 'ppe-stocks', label: 'KKD Envanteri', group: 'module' },
  { key: 'equipment-management', label: 'Ekipman ve Periyodik Kontrol', group: 'module' },
  { key: 'employees', label: 'Çalışanlar', group: 'module' },
  { key: 'emergency-preparedness', label: 'Acil Durum Hazırlığı', group: 'module' },
  { key: 'capa', label: 'Düzeltici Faaliyetler (CAPA)', group: 'module' },
  { key: 'legal-register', label: 'Yasal Mevzuat ve Uyumluluk Kaydı', group: 'module' },
  { key: 'documents', label: 'Doküman Kontrolü', group: 'module' },
  { key: 'daily-hse', label: 'Günlük HSE Raporu', group: 'management' },
  { key: 'weekly-hse', label: 'Haftalık HSE Raporu', group: 'management' },
  { key: 'monthly-hse', label: 'Aylık HSE Raporu', group: 'management' },
  { key: 'quarterly-hse', label: 'Üç Aylık HSE Raporu', group: 'management' },
  { key: 'annual-hse', label: 'Yıllık HSE Raporu', group: 'management' },
  { key: 'client-hse', label: 'Müşteri HSE Raporu', group: 'management' },
  { key: 'corporate-hse', label: 'Kurumsal HSE Raporu', group: 'management' },
  { key: 'employer-warning-letters', label: 'İşverenden Gelen Uyarı Mektupları', group: 'correspondence' },
  { key: 'employer-penalty-letters', label: 'İşverenden Gelen Ceza Mektupları', group: 'correspondence' },
  { key: 'hse-environment-fire-letters', label: 'İSG / Çevre / Yangın ile İlgili Mektuplar', group: 'correspondence' },
  { key: 'state-inspection-warnings', label: 'Devlet Denetimi Uyarıları', group: 'correspondence' },
  { key: 'state-inspection-letters', label: 'Devlet Denetimi Mektupları', group: 'correspondence' },
  { key: 'state-inspection-penalties', label: 'Devlet Denetimi Cezaları', group: 'correspondence' }
];

const reportFormatOptions: Array<{ value: CorporateReportFormat; label: string }> = [
  { value: 'PDF', label: 'PDF' },
  { value: 'EXCEL', label: 'Excel' }
];

const initialGeneratedReports: GeneratedCorporateReport[] = [
  {
    id: 'generated-report-1',
    reportName: 'Kurumsal HSE Raporu - Haziran 2026',
    reportTypeKey: 'corporate-hse',
    reportTypeLabel: 'Kurumsal HSE Raporu',
    projectId: 'all',
    projectName: 'Tüm Projeler',
    department: 'Tüm Departmanlar',
    periodStart: '2026-06-01',
    periodEnd: '2026-06-30',
    format: 'PDF',
    createdAt: '2026-07-10T08:00:00.000Z',
    createdBy: 'Kurumsal İSG Yöneticisi',
    summary: 'Proje genelinde güvenlik performansı, uyumluluk ve açık aksiyonların kurumsal özeti.',
    tableRows: [
      { label: 'Uyumluluk', value: '%94', note: 'Platform genel ortalama' },
      { label: 'Açık Aksiyon', value: '12', note: 'Takipte' }
    ],
    recommendations: ['Yüksek riskli projelerde haftalık takip planı sürdürülmelidir.', 'Açık CAPA maddeleri yönetim gözden geçirmesine taşınmalıdır.'],
    html: '<p>Kurumsal HSE raporu önizlemesi.</p>',
    csvContent: 'Rapor Türü,Değer\nUyumluluk,%94\nAçık Aksiyon,12',
    downloadName: 'kurumsal-hse-raporu-2026-06.pdf'
  }
];

const initialControlledDocumentRecords: ControlledDocumentRecord[] = [
  {
    id: 'doc-1',
    projectId: 'all',
    documentId: 'DOC-2026-0001',
    title: 'Çalışma İzin Formları Standardı',
    category: 'Çalışma İzinleri',
    documentType: 'Form',
    revisionNumber: 2,
    status: 'ONAYLANDI',
    effectiveDate: '2026-06-15',
    reviewDate: '2026-08-01',
    department: 'İSG',
    preparedBy: 'Kurumsal İSG Müdürü',
    reviewedBy: 'Uyumluluk Yöneticisi',
    approvedBy: 'İSG Direktörü',
    notes: 'Tüm proje sahalarında kullanılan standart PTW şablonu.',
    fileName: 'ptw-standart-rev2.pdf',
    fileType: 'application/pdf',
    fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
    uploadedAt: '2026-06-15T09:00:00.000Z',
    uploadedBy: 'Kurumsal İSG Müdürü',
    revisions: [
      {
        id: 'doc-1-rev1',
        revisionNumber: 1,
        status: 'GOZDEN_GECIRMEDE',
        fileName: 'ptw-standart-rev1.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-05-20T09:00:00.000Z',
        uploadedBy: 'İSG Uzmanı',
        note: 'İlk yayın'
      },
      {
        id: 'doc-1-rev2',
        revisionNumber: 2,
        status: 'ONAYLANDI',
        fileName: 'ptw-standart-rev2.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-06-15T09:00:00.000Z',
        uploadedBy: 'Kurumsal İSG Müdürü',
        note: 'Revizyon 2 onaylandı'
      }
    ]
  },
  {
    id: 'doc-2',
    projectId: 'metro',
    documentId: 'DOC-2026-0002',
    title: 'LOTO Prosedürü',
    category: 'LOTO Prosedürü',
    documentType: 'Prosedür',
    revisionNumber: 3,
    status: 'ONAYLANDI',
    effectiveDate: '2026-04-01',
    reviewDate: '2026-07-30',
    department: 'Elektrik',
    preparedBy: 'Elektrik İşletme Şefi',
    reviewedBy: 'Uyumluluk Yöneticisi',
    approvedBy: 'Kurumsal İSG Müdürü',
    notes: 'Enerji izolasyonu ve doğrulama adımlarını kapsar.',
    fileName: 'loto-proseduru-rev3.pdf',
    fileType: 'application/pdf',
    fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
    uploadedAt: '2026-06-28T10:00:00.000Z',
    uploadedBy: 'Elektrik İşletme Şefi',
    revisions: [
      {
        id: 'doc-2-rev1',
        revisionNumber: 1,
        status: 'ONAYLANDI',
        fileName: 'loto-proseduru-rev1.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-04-01T10:00:00.000Z',
        uploadedBy: 'Elektrik İşletme Şefi',
        note: 'İlk onaylı sürüm'
      },
      {
        id: 'doc-2-rev2',
        revisionNumber: 2,
        status: 'GOZDEN_GECIRMEDE',
        fileName: 'loto-proseduru-rev2.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-05-12T10:00:00.000Z',
        uploadedBy: 'İSG Uzmanı',
        note: 'Ek saha adımı eklendi'
      },
      {
        id: 'doc-2-rev3',
        revisionNumber: 3,
        status: 'ONAYLANDI',
        fileName: 'loto-proseduru-rev3.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-06-28T10:00:00.000Z',
        uploadedBy: 'Elektrik İşletme Şefi',
        note: 'Onaylı yürürlük sürümü'
      }
    ]
  },
  {
    id: 'doc-3',
    projectId: 'port',
    documentId: 'DOC-2026-0003',
    title: 'Toolbox Talk - İskele Güvenliği',
    category: 'Toolbox Talk',
    documentType: 'Form',
    revisionNumber: 1,
    status: 'GOZDEN_GECIRMEDE',
    effectiveDate: '2026-07-05',
    reviewDate: '2026-07-22',
    department: 'İnşaat',
    preparedBy: 'Şantiye Şefi',
    reviewedBy: 'İSG Uzmanı',
    approvedBy: '-',
    notes: 'Günlük saha konuşmalarında kullanılan kısa form.',
    fileName: 'toolbox-talk-iskele-sagligi.pdf',
    fileType: 'application/pdf',
    fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
    uploadedAt: '2026-07-05T08:30:00.000Z',
    uploadedBy: 'Şantiye Şefi',
    revisions: [
      {
        id: 'doc-3-rev1',
        revisionNumber: 1,
        status: 'GOZDEN_GECIRMEDE',
        fileName: 'toolbox-talk-iskele-sagligi.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-07-05T08:30:00.000Z',
        uploadedBy: 'Şantiye Şefi',
        note: 'İnceleme aşamasında'
      }
    ]
  }
];

const initialLegalRecords: LegalRecord[] = [
  {
    id: 'legal-1',
    projectId: 'metro',
    regulationId: 'LG-2026-0001',
    category: 'İş Sağlığı ve Güvenliği',
    title: 'Yapı İşlerinde İş Sağlığı ve Güvenliği Yönetmeliği',
    authority: 'Çalışma ve Sosyal Güvenlik Bakanlığı',
    department: 'İSG',
    legalRequirement: 'Yüksekte çalışma, iskele, düşmeye karşı koruma ve saha denetim şartlarının periyodik doğrulaması.',
    responsiblePerson: 'Kurumsal İSG Müdürü',
    complianceStatus: 'KISMEN_UYUMLU',
    effectiveDate: '2024-01-01',
    lastReviewDate: '2026-06-10',
    nextReviewDate: '2026-08-05',
    openActions: 3,
    riskLevel: 'YUKSEK',
    notes: 'Tünel şaft alanlarında düşüş önleme planı güncelleniyor.',
    documents: [
      {
        id: 'legal-doc-1',
        kind: 'MEVZUAT_BELGESI',
        fileName: 'yapi-islerinde-isg-yonetmeligi.pdf',
        fileType: 'application/pdf',
        fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
        uploadedAt: '2026-06-10',
        uploadedBy: 'Kurumsal İSG Müdürü',
        versions: [
          {
            version: 1,
            fileName: 'yapi-islerinde-isg-yonetmeligi.pdf',
            fileType: 'application/pdf',
            fileUrl: 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf',
            uploadedAt: '2026-06-10',
            uploadedBy: 'Kurumsal İSG Müdürü'
          }
        ]
      },
      {
        id: 'legal-doc-2',
        kind: 'DENETIM_RAPORU',
        fileName: 'metro-yuksekte-calisma-denetim-raporu.docx',
        fileType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        fileUrl: '#',
        uploadedAt: '2026-07-01',
        uploadedBy: 'İSG Uzmanı',
        versions: [
          {
            version: 1,
            fileName: 'metro-yuksekte-calisma-denetim-raporu.docx',
            fileType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            fileUrl: '#',
            uploadedAt: '2026-07-01',
            uploadedBy: 'İSG Uzmanı'
          }
        ]
      }
    ],
    createdBy: 'Kurumsal İSG Müdürü',
    createdAt: '2026-06-10',
    modifiedBy: 'İSG Uzmanı',
    modifiedAt: '2026-07-01',
    reviewedBy: 'Uyumluluk Yöneticisi',
    reviewedAt: '2026-07-08',
    exportHistory: [{ format: 'PDF_SINGLE', actor: 'Uyumluluk Yöneticisi', date: '2026-07-09' }],
    downloadHistory: [{ fileName: 'yapi-islerinde-isg-yonetmeligi.pdf', actor: 'İSG Uzmanı', date: '2026-07-02' }],
    auditTrail: [
      { id: 'legal-audit-1', eventType: 'CREATED', actor: 'Kurumsal İSG Müdürü', eventDate: '2026-06-10', detail: 'Mevzuat kaydı oluşturuldu.' },
      { id: 'legal-audit-2', eventType: 'REVIEWED', actor: 'Uyumluluk Yöneticisi', eventDate: '2026-07-08', detail: 'Gözden geçirme tamamlandı.' }
    ]
  },
  {
    id: 'legal-2',
    projectId: 'port',
    regulationId: 'LG-2026-0002',
    category: 'Çevre',
    title: 'Atık Yönetimi Yönetmeliği',
    authority: 'Çevre, Şehircilik ve İklim Değişikliği Bakanlığı',
    department: 'Çevre',
    legalRequirement: 'Tehlikeli atıkların sınıflandırılması, geçici depolama ve lisanslı bertaraf süreçlerinin kayıt altına alınması.',
    responsiblePerson: 'Çevre Mühendisi',
    complianceStatus: 'UYUMSUZ',
    effectiveDate: '2023-04-12',
    lastReviewDate: '2026-05-20',
    nextReviewDate: '2026-07-25',
    openActions: 5,
    riskLevel: 'KRITIK',
    notes: 'Atık transfer formlarının dijital arşive aktarımı gecikmiş durumda.',
    documents: [],
    createdBy: 'Çevre Mühendisi',
    createdAt: '2026-05-20',
    modifiedBy: 'Çevre Mühendisi',
    modifiedAt: '2026-07-10',
    reviewedBy: 'ISO Yönetim Temsilcisi',
    reviewedAt: '2026-07-11',
    exportHistory: [],
    downloadHistory: [],
    auditTrail: [
      { id: 'legal-audit-3', eventType: 'CREATED', actor: 'Çevre Mühendisi', eventDate: '2026-05-20', detail: 'Kayıt eklendi.' }
    ]
  },
  {
    id: 'legal-3',
    projectId: 'solar',
    regulationId: 'LG-2026-0003',
    category: 'Elektrik Güvenliği',
    title: 'Elektrik Tesislerinde İşletme Güvenliği',
    authority: 'Enerji ve Tabii Kaynaklar Bakanlığı',
    department: 'Elektrik',
    legalRequirement: 'Yüksek gerilim operasyonları için yetkin personel, LOTO prosedürü ve yıllık test raporlarının takibi.',
    responsiblePerson: 'Elektrik İşletme Şefi',
    complianceStatus: 'UYUMLU',
    effectiveDate: '2022-09-01',
    lastReviewDate: '2026-07-01',
    nextReviewDate: '2026-10-01',
    openActions: 1,
    riskLevel: 'ORTA',
    notes: 'Bir adet bakım prosedürü revizyonu açık.',
    documents: [],
    createdBy: 'Elektrik İşletme Şefi',
    createdAt: '2026-07-01',
    modifiedBy: 'Elektrik İşletme Şefi',
    modifiedAt: '2026-07-01',
    reviewedBy: 'Kurumsal İSG Müdürü',
    reviewedAt: '2026-07-12',
    exportHistory: [],
    downloadHistory: [],
    auditTrail: [{ id: 'legal-audit-4', eventType: 'CREATED', actor: 'Elektrik İşletme Şefi', eventDate: '2026-07-01', detail: 'Kayıt eklendi.' }]
  }
];

const ptwTablar: Array<{ key: PtwTabKey; label: string }> = [
  { key: 'kayit', label: 'PTW Kayıt Bilgileri' },
  { key: 'sorumlular', label: 'Sorumlular' },
  { key: 'ekip', label: 'Ekip Bilgileri' },
  { key: 'is-bilgileri', label: 'İş Bilgileri' },
  { key: 'riskler', label: 'Riskler' },
  { key: 'guvenlik-sistemleri', label: 'Yüksekte Çalışma Güvenlik Sistemleri' },
  { key: 'ekipman', label: 'Kullanılacak Ekipman' },
  { key: 'on-kontroller', label: 'Çalışma Öncesi Kontroller' },
  { key: 'tedbirler', label: 'Çalışma Sırasında Alınacak Tedbirler' },
  { key: 'ozel-sartlar', label: 'Özel Şartlar' },
  { key: 'onaylar', label: 'Onaylar' },
  { key: 'gunluk-takip', label: 'Günlük PTW Takibi' },
  { key: 'ekip-degisiklik', label: 'Ekip Değişiklikleri' },
  { key: 'kapanis', label: 'Kapanış' },
  { key: 'dosya-ekleri', label: 'Dosya Ekleri' }
];

const inspectionChecklistSections: InspectionSection[] = [
  {
    id: 'housekeeping',
    title: 'Düzen ve Temizlik',
    description: 'Temizlik, depolama, atık kontrolü ve genel düzeni değerlendirir.',
    questions: [
      'Çalışma alanları temiz mi ve dağınık atıktan arındırılmış mı?',
      'Malzemeler erişimi engellemeden ve düşme riski oluşturmadan güvenli depolanıyor mu?',
      'Dökülmeler hızlıca temizleniyor veya kontrol altına alınıyor mu?',
      'Alet ve ekipmanlar kullanım sonrası uygun alana geri konuluyor mu?',
      'Atık kutuları uygun mu ve taşmadan kullanılıyor mu?'
    ]
  },
  {
    id: 'walking-surfaces',
    title: 'Yürüyüş Yolları ve Erişim',
    description: 'Güvenli hareket, yürüyüş yolları, kapılar, basamaklar ve erişim güzergahlarını değerlendirir.',
    questions: [
      'Yürüyüş yolları takılma riski ve engellerden arındırılmış mı?',
      'Zeminler kuru, düzgün ve güvenli yürüyüş için uygun mu?',
      'Kapılar, çıkışlar, panolar ve kritik ekipmanlara erişim kolay mı?',
      'Merdivenler, rampalar ve korkuluklar güvenli durumda mı?',
      'Aydınlatma güvenli çalışma ve hareket için yeterli mi?'
    ]
  },
  {
    id: 'fire-safety',
    title: 'Yangın Güvenliği',
    description: 'Kaçış yolları, söndürücüler, yangın kapıları ve tutuşturucu kaynak kontrolünü değerlendirir.',
    questions: [
      'Yangın çıkışları açık, kilitsiz ve kullanılabilir durumda mı?',
      'Yangın söndürücüler görünür, erişilebilir ve geçerli tarihte mi?',
      'Yangın kapıları kapalı tutuluyor ve açık sabitlenmiyor mu?',
      'Yanıcı malzemeler kontrol altında ve tutuşturucu kaynaklardan uzak mı?',
      'Sigara yasağı ve sıcak çalışma kurallarına uyuluyor mu?'
    ]
  },
  {
    id: 'emergency-readiness',
    title: 'Acil Durum Hazırlığı',
    description: 'Acil durum bilgileri, güzergahlar, iletişim ve müdahale hazırlıklarını değerlendirir.',
    questions: [
      'Acil durum iletişim numaraları görünür şekilde mevcut mu?',
      'Tahliye yolları ve toplanma alanları açıkça işaretli mi?',
      'Çalışanlar alarm verme veya acil durum bildirme sürecini biliyor mu?',
      'Acil durum ekipmanlarına erişim açık mı, önü kapalı değil mi?',
      'Gerekli yerlerde acil durum tatbikatı veya bilgilendirme yapılmış mı?'
    ]
  },
  {
    id: 'first-aid-welfare',
    title: 'İlk Yardım ve Refah',
    description: 'İlk yardım, içme suyu, sosyal alanlar ve temel sağlık desteğini değerlendirir.',
    questions: [
      'Dolu bir ilk yardım çantası mevcut ve erişilebilir mi?',
      'İlk yardımcı bilgisi görünür veya çalışanlarca biliniyor mu?',
      'Sosyal alanlar temiz ve çalışma alanına uygun mu?',
      'Gereken noktalarda içme suyu mevcut mu?',
      'Sıcaklık, yorgunluk ve temel sağlık riskleri kontrol altında mı?'
    ]
  },
  {
    id: 'electrical-safety',
    title: 'Elektrik Güvenliği',
    description: 'Kablo, priz, pano, geçici enerji ve izolasyon uygulamalarını değerlendirir.',
    questions: [
      'Fiş, priz ve kablolarda görünür hasar bulunmuyor mu?',
      'Elektrik panoları kapalı, etiketli ve erişilebilir mi?',
      'Uzatmalar ve geçici enerji güvenli ve aşırı yüklenmeden kullanılıyor mu?',
      'Kablolar takılma ve hasar riskini önleyecek şekilde yönlendirilmiş mi?',
      'Elektrik bakımından önce izolasyon/LOTO uygulanıyor mu?'
    ]
  },
  {
    id: 'ppe',
    title: 'KKD',
    description: 'KKD bulunurluğu, doğru kullanım, kondisyon ve ziyaretçi korumasını değerlendirir.',
    questions: [
      'İş/alan için gerekli KKD mevcut mu?',
      'Çalışanlar gerekli yerlerde doğru KKD kullanıyor mu?',
      'KKD temiz, uygun ve iyi durumda mı?',
      'Göz, el, işitme veya solunum riskleri için işe özel KKD kullanılıyor mu?',
      'Kontrollü alanlara girişten önce ziyaretçilere gerekli KKD veriliyor mu?'
    ]
  },
  {
    id: 'plant-tools-equipment',
    title: 'Makine, Alet ve Ekipman',
    description: 'Makine, el aleti, koruyucu, periyodik kontrol ve arızalı ekipman yönetimini değerlendirir.',
    questions: [
      'Gerekli yerlerde makine koruyucuları takılı ve sağlam mı?',
      'Acil durdurma ve güvenlik ekipmanları erişilebilir ve çalışır durumda mı?',
      'Alet ve ekipmanlar kullanım öncesi kontrol ediliyor mu?',
      'Operatörler kullandıkları ekipman için yetkin ve yetkilendirilmiş mi?',
      'Arızalı ekipman etiketleniyor, kullanım dışına alınıyor veya kontrol ediliyor mu?'
    ]
  },
  {
    id: 'working-at-height',
    title: 'Yüksekte Çalışma',
    description: 'Merdiven, iskele, kenar koruma, düşen cisim ve hava koşulu kontrollerini değerlendirir.',
    questions: [
      'Merdivenler uygun, kontrol edilmiş ve güvenli kullanılıyor mu?',
      'İskele/platformlar tam, stabil ve kontrol edilmiş mi?',
      'Açık kenarlar, döşeme boşlukları ve kırılgan yüzeyler korunuyor mu?',
      'Düşme riski olan yerlerde düşüş önleme sistemleri kullanılıyor mu?',
      'Alet ve malzemeler düşen cisim riskini önleyecek şekilde kontrol ediliyor mu?'
    ]
  },
  {
    id: 'manual-handling',
    title: 'Manuel Taşıma ve Ergonomi',
    description: 'Kaldırma, duruş, istasyon düzeni, tekrarlı iş ve taşıma yardımcılarını değerlendirir.',
    questions: [
      'Ağır/zor yükler uygun yöntem veya yardımcı ekipmanla taşınıyor mu?',
      'Çalışanlar güvensiz kaldırma, dönme veya aşırı uzanmadan kaçınıyor mu?',
      'Çalışma istasyonları zorlanmayı ve kötü postürü azaltacak şekilde düzenlenmiş mi?',
      'Tekrarlı işler mola, rotasyon veya görev değişimi ile yönetiliyor mu?',
      'Arabalar ve taşıma yardımcıları güvenli durumda mı?'
    ]
  },
  {
    id: 'chemicals',
    title: 'Kimyasallar ve Tehlikeli Maddeler',
    description: 'Etiketleme, depolama, GBF, döküntü kontrolü ve maruziyet önlemlerini değerlendirir.',
    questions: [
      'Kimyasal kaplar açık ve doğru etiketlenmiş mi?',
      'Kimyasal kullanılan/depolanan alanlarda GBF mevcut mu?',
      'Uyumsuz kimyasallar güvenli şekilde ayrı depolanıyor mu?',
      'Gereken alanlarda döküntü kiti veya döküntü kontrolü var mı?',
      'Kullanıcılar solunum, cilt veya göz maruziyetine karşı korunuyor mu?'
    ]
  },
  {
    id: 'vehicle-pedestrian',
    title: 'Araç ve Yaya Güvenliği',
    description: 'Araç trafiği, yaya yolları, görüş, geri manevra ve yükleme alanlarını değerlendirir.',
    questions: [
      'Mümkün olan yerlerde yaya yolları araç trafiğinden ayrılmış mı?',
      'Geri manevra ve manevra faaliyetleri güvenli şekilde kontrol ediliyor mu?',
      'Yükleme/boşaltma alanları kontrollü ve gereksiz kişilerden arındırılmış mı?',
      'Hız limitleri, levhalar ve trafik işaretleri görünür mü?',
      'Sürücü ve yayalar saha trafik kurallarına uyuyor mu?'
    ]
  },
  {
    id: 'contractor-maintenance',
    title: 'Alt Yüklenici / Bakım Çalışmaları',
    description: 'İş kontrolü, izinler, izolasyon, gözetim ve alan teslim süreçlerini değerlendirir.',
    questions: [
      'Alt yüklenici veya bakım işi başlamadan önce yetkilendirme yapılıyor mu?',
      'Yüksek riskli işlerde izin veya görev kontrolü uygulanıyor mu?',
      'Çalışma alanları başkalarını koruyacak şekilde bariyerleniyor veya kontrol ediliyor mu?',
      'Gereken yerlerde izolasyon, LOTO veya servis kesme uygulanıyor mu?',
      'Bakım sonrası alan kontrol edilip güvenli hale getiriliyor mu?'
    ]
  },
  {
    id: 'environmental-controls',
    title: 'Çevresel Kontroller ve Atık',
    description: 'Atık ayrıştırma, sızıntı, emisyon, drenaj ve çevresel etkileri değerlendirir.',
    questions: [
      'Atıklar uygun kaplarda ayrıştırılarak depolanıyor mu?',
      'Sızıntı, döküntü veya drenaja deşarj riski önleniyor/kontrol ediliyor mu?',
      'Tehlikeli atıklar etiketli ve kontrol altında mı?',
      'Gürültü, toz, duman veya koku etkileri kontrol ediliyor mu?',
      'Açık alan/depo sahaları çevresel etkiyi önleyecek şekilde yönetiliyor mu?'
    ]
  },
  {
    id: 'safety-communication',
    title: 'Güvenlik İletişimi ve Bildirim',
    description: 'Levhalar, bilgilendirme, bildirim, gözetim ve görünür güvenlik beklentilerini değerlendirir.',
    questions: [
      'Güvenlik levhaları ve talimatları açık ve görünür mü?',
      'Çalışanlar tehlike, ramak kala ve olay bildirim sürecini biliyor mu?',
      'Gerekli yerlerde güvenlik bilgilendirmesi/toolbox konuşması yapılıyor mu?',
      'Yöneticiler güvensiz durum ve davranışlara hızlı müdahale ediyor mu?',
      'Ziyaretçi veya yeni çalışanlara alana girişten önce gerekli güvenlik bilgilendirmesi yapılıyor mu?'
    ]
  }
];

const inspectionChecklistSectionsRu: InspectionSection[] = [
  {
    id: 'housekeeping',
    title: 'Порядок и чистота',
    description: 'Оценивает чистоту, хранение, контроль отходов и общий порядок.',
    questions: [
      'Рабочие зоны чистые и свободны от беспорядка и мусора?',
      'Материалы хранятся безопасно, не блокируя доступ и не создавая риск падения?',
      'Разливы оперативно убираются или берутся под контроль?',
      'Инструменты и оборудование после использования возвращаются в предназначенное место?',
      'Контейнеры для отходов подходят по назначению и не переполнены?'
    ]
  },
  {
    id: 'walking-surfaces',
    title: 'Пешеходные пути и доступ',
    description: 'Оценивает безопасное перемещение, проходы, двери, ступени и маршруты доступа.',
    questions: [
      'Пешеходные пути свободны от препятствий и риска споткнуться?',
      'Покрытие пола сухое, ровное и безопасное для передвижения?',
      'Доступ к дверям, выходам, щитам и критическому оборудованию свободен?',
      'Лестницы, пандусы и перила находятся в безопасном состоянии?',
      'Освещение достаточно для безопасной работы и перемещения?'
    ]
  },
  {
    id: 'fire-safety',
    title: 'Пожарная безопасность',
    description: 'Оценивает эвакуационные пути, огнетушители, противопожарные двери и контроль источников возгорания.',
    questions: [
      'Пожарные выходы открыты, не заперты и доступны для использования?',
      'Огнетушители видимы, доступны и имеют действующий срок?',
      'Противопожарные двери закрываются и не фиксируются в открытом положении?',
      'Горючие материалы под контролем и удалены от источников возгорания?',
      'Соблюдаются запрет курения и правила горячих работ?'
    ]
  },
  {
    id: 'emergency-readiness',
    title: 'Готовность к ЧС',
    description: 'Оценивает доступность аварийной информации, маршрутов, связи и готовность к реагированию.',
    questions: [
      'Контактные номера экстренных служб размещены на видном месте?',
      'Маршруты эвакуации и точки сбора четко обозначены?',
      'Сотрудники знают порядок подачи тревоги и сообщения о ЧС?',
      'Доступ к аварийному оборудованию свободен и не загроможден?',
      'В необходимых зонах проводятся учения или инструктаж по ЧС?'
    ]
  },
  {
    id: 'first-aid-welfare',
    title: 'Первая помощь и бытовые условия',
    description: 'Оценивает первую помощь, питьевую воду, бытовые зоны и базовую поддержку здоровья.',
    questions: [
      'Аптечка укомплектована и доступна?',
      'Информация о назначенных ответственных за первую помощь доступна сотрудникам?',
      'Бытовые помещения чистые и пригодны к использованию?',
      'В необходимых местах обеспечен доступ к питьевой воде?',
      'Риски, связанные с температурой, усталостью и базовым здоровьем, находятся под контролем?'
    ]
  },
  {
    id: 'electrical-safety',
    title: 'Электробезопасность',
    description: 'Оценивает кабели, розетки, щиты, временное электроснабжение и изоляцию.',
    questions: [
      'На вилках, розетках и кабелях отсутствуют видимые повреждения?',
      'Электрощиты закрыты, промаркированы и доступны для обслуживания?',
      'Удлинители и временное питание используются безопасно, без перегрузок?',
      'Кабели проложены так, чтобы исключить спотыкание и повреждения?',
      'Перед электротехническим обслуживанием применяется изоляция/LOTO?'
    ]
  },
  {
    id: 'ppe',
    title: 'СИЗ',
    description: 'Оценивает наличие СИЗ, правильность использования, состояние и защиту посетителей.',
    questions: [
      'Необходимые СИЗ для работ/зоны доступны?',
      'Сотрудники используют требуемые СИЗ в нужных местах?',
      'СИЗ чистые, подходят по размеру и в исправном состоянии?',
      'Для рисков зрения, рук, слуха и дыхания используются соответствующие СИЗ?',
      'Посетителям выдаются необходимые СИЗ до входа в контролируемые зоны?'
    ]
  },
  {
    id: 'plant-tools-equipment',
    title: 'Машины, инструмент и оборудование',
    description: 'Оценивает ограждения, ручной инструмент, защиту, периодические проверки и управление неисправным оборудованием.',
    questions: [
      'Защитные ограждения машин установлены и исправны там, где это необходимо?',
      'Аварийные остановы и защитные устройства доступны и работоспособны?',
      'Инструменты и оборудование проверяются перед использованием?',
      'Операторы компетентны и допущены к работе с используемым оборудованием?',
      'Неисправное оборудование маркируется, выводится из эксплуатации или контролируется?'
    ]
  },
  {
    id: 'working-at-height',
    title: 'Работы на высоте',
    description: 'Оценивает лестницы, леса, защиту краев, падение предметов и погодные условия.',
    questions: [
      'Лестницы исправны, проверены и используются безопасно?',
      'Леса/платформы полностью укомплектованы, устойчивы и проверены?',
      'Открытые края, проемы и хрупкие поверхности защищены?',
      'В зонах риска падения применяются системы предотвращения падения?',
      'Инструменты и материалы контролируются для предотвращения падения предметов?'
    ]
  },
  {
    id: 'manual-handling',
    title: 'Ручное перемещение и эргономика',
    description: 'Оценивает подъем, рабочую позу, организацию места, повторяющиеся операции и вспомогательные средства.',
    questions: [
      'Тяжелые/неудобные грузы перемещаются безопасным методом или с использованием вспомогательных средств?',
      'Сотрудники избегают небезопасного подъема, скручивания и чрезмерного вытягивания?',
      'Рабочие места организованы так, чтобы снижать перегрузку и неудобную позу?',
      'Повторяющиеся операции управляются через перерывы, ротацию или смену задач?',
      'Тележки и вспомогательные средства перемещения находятся в безопасном состоянии?'
    ]
  },
  {
    id: 'chemicals',
    title: 'Химические и опасные вещества',
    description: 'Оценивает маркировку, хранение, SDS, контроль разливов и меры по предотвращению воздействия.',
    questions: [
      'Емкости с химическими веществами четко и правильно промаркированы?',
      'В зонах использования/хранения химии доступны SDS?',
      'Несовместимые химические вещества хранятся раздельно и безопасно?',
      'В необходимых зонах есть наборы для ликвидации разливов или другие меры контроля?',
      'Пользователи защищены от воздействия на дыхание, кожу и глаза?'
    ]
  },
  {
    id: 'vehicle-pedestrian',
    title: 'Безопасность транспорта и пешеходов',
    description: 'Оценивает движение транспорта, пешеходные пути, обзор, маневры задним ходом и зоны погрузки.',
    questions: [
      'Где возможно, пешеходные пути отделены от транспортного потока?',
      'Маневры и движение задним ходом контролируются безопасным образом?',
      'Зоны погрузки/разгрузки контролируются и очищены от посторонних?',
      'Ограничения скорости, знаки и дорожная разметка видимы?',
      'Водители и пешеходы соблюдают правила движения на площадке?'
    ]
  },
  {
    id: 'contractor-maintenance',
    title: 'Работы подрядчиков / техническое обслуживание',
    description: 'Оценивает контроль работ, разрешения, изоляцию, надзор и процессы передачи зоны.',
    questions: [
      'Перед началом работ подрядчика/обслуживания проводится допуск?',
      'Для работ повышенного риска применяется разрешительная система или контроль задания?',
      'Рабочие зоны ограждаются/контролируются для защиты других лиц?',
      'Где требуется, применяется изоляция, LOTO или отключение сервисов?',
      'После обслуживания зона проверяется и приводится в безопасное состояние?'
    ]
  },
  {
    id: 'environmental-controls',
    title: 'Экологический контроль и отходы',
    description: 'Оценивает сортировку отходов, утечки, выбросы, дренаж и экологические воздействия.',
    questions: [
      'Отходы сортируются и хранятся в соответствующих контейнерах?',
      'Риск утечки, разлива или сброса в дренаж предотвращается/контролируется?',
      'Опасные отходы промаркированы и находятся под контролем?',
      'Шум, пыль, дым и запахи контролируются?',
      'Открытые площадки/склады управляются так, чтобы предотвращать экологическое воздействие?'
    ]
  },
  {
    id: 'safety-communication',
    title: 'Коммуникация по безопасности и уведомления',
    description: 'Оценивает знаки, информирование, уведомления, надзор и видимые требования безопасности.',
    questions: [
      'Знаки и инструкции по безопасности понятны и хорошо видимы?',
      'Сотрудники знают порядок сообщения об опасностях, near miss и инцидентах?',
      'В необходимых местах проводятся инструктажи по безопасности/toolbox talks?',
      'Руководители оперативно реагируют на небезопасные условия и поведение?',
      'Посетители и новые сотрудники получают необходимый инструктаж перед входом в зону?'
    ]
  }
];

const recordsByModule: Record<ModuleKey, ModuleRecord[]> = {
  dashboard: [],
  inspections: [
    { projectId: 'metro', date: '2026-07-17', title: 'İskele denetimi', valueA: 14, valueB: 5, status: 'IN_PROGRESS' }
  ],
  observations: [],
  'risk-assessments': [
    { projectId: 'solar', date: '2026-07-15', title: 'Elektrik risk değerlendirmesi', valueA: 12, valueB: 4, status: 'CLOSED' }
  ],
  'permit-to-work': [
    { projectId: 'metro', date: '2026-07-17', title: 'Sıcak çalışma izni', valueA: 6, valueB: 2, status: 'OPEN' }
  ],
  incidents: [
    { projectId: 'port', date: '2026-07-14', title: 'Hafif parmak yaralanması', valueA: 1, valueB: 0, status: 'CLOSED' }
  ],
  audits: [
    { projectId: 'solar', date: '2026-07-13', title: 'Haftalık iç tetkik', valueA: 10, valueB: 2, status: 'IN_PROGRESS' }
  ],
  'equipment-management': [
    { projectId: 'metro', date: '2026-07-12', title: 'Kaldırma ekipmanı kontrolü', valueA: 42, valueB: 5, status: 'OPEN' }
  ],
  environmental: [
    { projectId: 'port', date: '2026-07-12', title: 'Atık ayrıştırma takibi', valueA: 26, valueB: 18, status: 'IN_PROGRESS' }
  ],
  'emergency-preparedness': [
    { projectId: 'solar', date: '2026-07-10', title: 'Tahliye tatbikatı uygulaması', valueA: 7, valueB: 2, status: 'CLOSED' }
  ],
  employees: [
    { projectId: 'metro', date: '2026-07-17', title: 'Yeni çalışan oryantasyonu', valueA: 9, valueB: 0, status: 'CLOSED' }
  ],
  trainings: [
    { projectId: 'port', date: '2026-07-12', title: 'Yüksekte çalışma eğitimi', valueA: 34, valueB: 6, status: 'IN_PROGRESS' }
  ],
  'ppe-stocks': [
    { projectId: 'solar', date: '2026-07-17', title: 'Baret stok güncellemesi', valueA: 500, valueB: 380, status: 'OPEN' }
  ],
  'occupational-health': [
    { projectId: 'metro', date: '2026-07-11', title: 'Sağlık taraması', valueA: 63, valueB: 7, status: 'CLOSED' }
  ],
  'legal-register': [
    { projectId: 'all', date: '2026-07-10', title: 'Mevzuat güncellemesi', valueA: 4, valueB: 1, status: 'OPEN' }
  ],
  documents: [
    { projectId: 'all', date: '2026-07-09', title: 'SOP revizyonu', valueA: 16, valueB: 3, status: 'IN_PROGRESS' }
  ],
  'action-tracker': [
    { projectId: 'port', date: '2026-07-08', title: 'DÖF kapatma paketi', valueA: 12, valueB: 4, status: 'IN_PROGRESS' }
  ],
  'kpis-analytics': [],
  reports: [
    { projectId: 'all', date: '2026-07-17', title: 'Haftalık rapor oluşturuldu', valueA: 1, valueB: 1, status: 'CLOSED' }
  ],
  'export-center': [
    { projectId: 'all', date: '2026-07-17', title: 'Dışa aktarımlar tamamlandı', valueA: 22, valueB: 4, status: 'IN_PROGRESS' }
  ],
  projects: [
    { projectId: 'all', date: '2026-07-01', title: 'Aktif projeler', valueA: 3, valueB: 0, status: 'OPEN' }
  ],
  departments: [],
  contractors: [
    { projectId: 'all', date: '2026-07-01', title: 'Onaylı alt yükleniciler', valueA: 18, valueB: 2, status: 'OPEN' }
  ],
  settings: [
    { projectId: 'all', date: '2026-07-01', title: 'Sistem politika güncellemeleri', valueA: 5, valueB: 1, status: 'CLOSED' }
  ]
};

const initialRiskRecords: RiskRecord[] = [
  {
    id: 'risk-1',
    riskId: 'RK-2026-0001',
    projectId: 'solar',
    departmentActivity: 'Elektrik / Kablo sonlandırma',
    assessmentDate: '2026-07-15',
    hazard: 'Açık uçlu kablo ile temas',
    potentialConsequence: 'Elektrik çarpması ve ciddi yaralanma',
    existingControls: 'LOTO prosedürü, izole eldiven, gözetmen kontrolü',
    recommendedControls: 'Enerji izolasyon doğrulama formu ve ikinci kişi kontrolü',
    likelihood: 4,
    severity: 5,
    initialRiskScore: 20,
    residualLikelihood: 2,
    residualSeverity: 4,
    residualRiskScore: 8,
    responsiblePerson: 'Saha Elektrik Şefi',
    targetCompletionDate: '2026-07-22',
    status: 'IN_PROGRESS',
    attachments: ['risk-analiz-formu.pdf'],
    photos: ['pano-alani-1.jpg'],
    notes: 'Kontrol sonrası tekrar ölçüm yapılacak.'
  }
];

const initialEquipmentRecords: EquipmentRecord[] = [
  {
    id: 'eq-1',
    equipmentId: 'EQ-2026-0001',
    projectId: 'metro',
    equipmentName: 'Mobil Vinç 30T',
    equipmentType: 'Kaldırma Ekipmanı',
    brandModel: 'Liebherr LTM 1030',
    serialNumber: 'LTM1030-88921',
    location: 'Kuzey Şaft Alanı',
    responsiblePerson: 'Mekanik Bakım Şefi',
    lastInspectionDate: '2026-06-15',
    nextInspectionDate: '2026-07-25',
    inspectionStatus: 'UPCOMING',
    certificateNumber: 'CRT-CRANE-2026-118',
    certificateExpiryDate: '2026-08-10',
    equipmentStatus: 'ACTIVE',
    riskLevel: 'MEDIUM',
    attachments: ['vinç-kontrol-listesi.pdf'],
    inspectionReports: ['vinç-periyodik-muayene-20260615.pdf'],
    equipmentPhotos: ['mobil-vinc-01.jpg'],
    notes: 'Bir sonraki periyodik muayene için rezervasyon oluşturuldu.'
  },
  {
    id: 'eq-2',
    equipmentId: 'EQ-2026-0002',
    projectId: 'port',
    equipmentName: 'Forklift 3T',
    equipmentType: 'Taşıma Ekipmanı',
    brandModel: 'Toyota 8FG30',
    serialNumber: '8FG30-54100',
    location: 'Lojistik Depo 2',
    responsiblePerson: 'Depo Operasyon Sorumlusu',
    lastInspectionDate: '2026-05-20',
    nextInspectionDate: '2026-07-05',
    inspectionStatus: 'OVERDUE',
    certificateNumber: 'CRT-FLT-2026-304',
    certificateExpiryDate: '2026-07-21',
    equipmentStatus: 'UNDER_MAINTENANCE',
    riskLevel: 'HIGH',
    attachments: ['forklift-bakim-kaydi.xlsx'],
    inspectionReports: ['forklift-inspection-20260520.pdf'],
    equipmentPhotos: ['forklift-3t.jpg'],
    notes: 'Fren sistemi revizyonu tamamlanmadan sahaya çıkış kapalı.'
  },
  {
    id: 'eq-3',
    equipmentId: 'EQ-2026-0003',
    projectId: 'solar',
    equipmentName: 'İskele Platform Seti',
    equipmentType: 'Erişim Ekipmanı',
    brandModel: 'Layher Allround',
    serialNumber: 'LAY-PLT-77321',
    location: 'Panel Blok C',
    responsiblePerson: 'Saha Formeni',
    lastInspectionDate: '2026-07-01',
    nextInspectionDate: '2026-09-01',
    inspectionStatus: 'COMPLIANT',
    certificateNumber: 'CRT-SCF-2026-450',
    certificateExpiryDate: '2026-12-01',
    equipmentStatus: 'OUT_OF_SERVICE',
    riskLevel: 'LOW',
    attachments: ['iskele-envanter-karti.pdf'],
    inspectionReports: ['iskele-denetim-raporu-20260701.pdf'],
    equipmentPhotos: ['iskele-c-blok.jpg'],
    notes: 'Revizyon sonrası tekrar devreye alma kontrolü bekleniyor.'
  }
];

const initialEmergencyDrillRecords: EmergencyDrillRecord[] = [
  {
    id: 'drill-1',
    drillId: 'ADT-2026-0001',
    projectId: 'metro',
    emergencyType: 'TAHLIYE',
    drillName: 'T1 Şaft Tahliye Tatbikatı',
    drillDate: '2026-07-10',
    participantCount: 48,
    drillResult: 'BASARILI',
    openActions: 2,
    closedActions: 5,
    responsiblePerson: 'HSE Müdürü',
    nextPlannedDrillDate: '2026-08-15',
    status: 'TAMAMLANDI',
    attachments: ['tahliye-katilim-listesi.pdf'],
    drillReports: ['tahliye-tatbikati-raporu-20260710.pdf'],
    photos: ['tahliye-toplanma-alani.jpg'],
    notes: 'Toplanma alanı yönlendirme levhaları güncellenecek.'
  },
  {
    id: 'drill-2',
    drillId: 'ADT-2026-0002',
    projectId: 'port',
    emergencyType: 'KIMYASAL_SIZINTI',
    drillName: 'Kimyasal Döküntü Müdahale Tatbikatı',
    drillDate: '2026-07-18',
    participantCount: 26,
    drillResult: 'KISMEN_BASARILI',
    openActions: 4,
    closedActions: 1,
    responsiblePerson: 'Çevre ve HSE Şefi',
    nextPlannedDrillDate: '2026-08-05',
    status: 'PLANLANDI',
    attachments: ['kimyasal-setup-checklist.pdf'],
    drillReports: ['kimyasal-on-degerlendirme.docx'],
    photos: ['sizinti-kiti-kontrol.jpg'],
    notes: 'Spill kit yerleşimi yeniden düzenlenecek.'
  },
  {
    id: 'drill-3',
    drillId: 'ADT-2026-0003',
    projectId: 'solar',
    emergencyType: 'YANGIN',
    drillName: 'Yangın Alarm ve İlk Müdahale Tatbikatı',
    drillDate: '2026-07-02',
    participantCount: 31,
    drillResult: 'BASARISIZ',
    openActions: 6,
    closedActions: 0,
    responsiblePerson: 'İşletme Sorumlusu',
    nextPlannedDrillDate: '2026-07-28',
    status: 'TAMAMLANDI',
    attachments: ['yangin-alarim-test-formu.xlsx'],
    drillReports: ['yangin-tatbikati-kok-neden-analizi.pdf'],
    photos: ['yangin-ekipmani-nokta-a.jpg'],
    notes: 'Alarm siren kapsaması düşük, aksiyon açıldı.'
  }
];

const initialTrainingRecords: TrainingRecord[] = [
  {
    id: 'training-1',
    projectId: 'metro',
    trainingId: 'EGT-2026-0001',
    trainingType: 'Yüksekte Çalışma',
    trainingTitle: 'Yüksekte Çalışma ve Düşüş Önleme',
    trainingCategory: 'Zorunlu İSG',
    trainingDate: '2026-07-12',
    provider: 'Akredite İSG Akademi',
    department: 'İnşaat',
    position: 'Usta',
    projectEmployeeCount: 58,
    certifiedEmployeeCount: 44,
    certificateRequired: true,
    certificateValidityDate: '2027-07-12',
    totalTrainingCost: 168000,
    costPerEmployee: 2897,
    status: 'TAMAMLANDI',
    attachments: ['egitim-programi.pdf'],
    participantList: ['katilim-listesi-metro.xlsx'],
    certificates: ['sertifika-paketi-metro.zip'],
    notes: 'Yeni gelen ekip için ek oturum planlandı.'
  },
  {
    id: 'training-2',
    projectId: 'port',
    trainingId: 'EGT-2026-0002',
    trainingType: 'Kapalı Alan Çalışması',
    trainingTitle: 'Kapalı Alan İzin ve Kurtarma Prosedürü',
    trainingCategory: 'Teknik Güvenlik',
    trainingDate: '2026-07-18',
    provider: 'Saha HSE Ekibi',
    department: 'Mekanik',
    position: 'Teknisyen',
    projectEmployeeCount: 36,
    certifiedEmployeeCount: 21,
    certificateRequired: true,
    certificateValidityDate: '2027-01-15',
    totalTrainingCost: 94000,
    costPerEmployee: 2611,
    status: 'DEVAM_EDIYOR',
    attachments: ['kapali-alan-plan.docx'],
    participantList: ['katilim-listesi-liman.xlsx'],
    certificates: [],
    notes: 'Gaz ölçüm ekipmanı uygulaması ikinci oturumda tamamlanacak.'
  },
  {
    id: 'training-3',
    projectId: 'solar',
    trainingId: 'EGT-2026-0003',
    trainingType: 'Elektrik Güvenliği',
    trainingTitle: 'LOTO ve Elektriksel Tehlike Yönetimi',
    trainingCategory: 'Sertifikasyon',
    trainingDate: '2026-07-08',
    provider: 'Yetkili Elektrik Kuruluşu',
    department: 'Elektrik',
    position: 'Saha Mühendisi',
    projectEmployeeCount: 28,
    certifiedEmployeeCount: 19,
    certificateRequired: true,
    certificateValidityDate: '2026-07-16',
    totalTrainingCost: 76000,
    costPerEmployee: 2714,
    status: 'SURESI_DOLDU',
    attachments: ['loto-egitim-icerigi.pdf'],
    participantList: ['katilim-listesi-gunes.xlsx'],
    certificates: ['sertifika-listesi-gunes.pdf'],
    notes: 'Süresi dolan sertifikalar için yenileme eğitimi açıldı.'
  }
];

const initialPpeTransactions: PpeTransactionRecord[] = [];

const initialDepartmentRecords: DepartmentRecord[] = [];

const isDepartmentRecord = (value: unknown): value is DepartmentRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.name === 'string'
    && typeof candidate.code === 'string'
    && typeof candidate.description === 'string'
  );
};

const loadDepartmentRecords = (): DepartmentRecord[] => {
  if (typeof window === 'undefined') {
    return initialDepartmentRecords;
  }

  try {
    const raw = window.localStorage.getItem(DEPARTMENTS_STORAGE_KEY);
    if (!raw) {
      return initialDepartmentRecords;
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isDepartmentRecord)) {
      return parsed;
    }
  } catch {
    // Fall back to seeded departments if local storage is unavailable or corrupted.
  }

  return initialDepartmentRecords;
};

const apiBaseUrl = import.meta.env.VITE_API_BASE_URL ?? '';

const apiRequest = async <T,>(path: string, init?: RequestInit): Promise<T> => {
  const method = (init?.method ?? 'GET').toUpperCase();
  if (
    typeof window !== 'undefined'
    && window.__HSE_GUEST_MODE__
    && method !== 'GET'
    && method !== 'HEAD'
  ) {
    throw new Error('GUEST_MODE_WRITE_BLOCKED');
  }

  const response = await fetch(`${apiBaseUrl}${path}`, {
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers ?? {})
    },
    ...init
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || `Request failed with status ${response.status}`);
  }

  if (response.status === 204) {
    return undefined as T;
  }

  return response.json() as Promise<T>;
};

const isPpeTransactionRecord = (value: unknown): value is PpeTransactionRecord => {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.id === 'string'
    && typeof candidate.transactionId === 'string'
    && typeof candidate.transactionType === 'string'
    && typeof candidate.lifecycle === 'string'
    && typeof candidate.projectId === 'string'
    && typeof candidate.warehouse === 'string'
    && typeof candidate.date === 'string'
    && typeof candidate.category === 'string'
    && typeof candidate.itemName === 'string'
    && typeof candidate.brandModel === 'string'
    && typeof candidate.unit === 'string'
    && typeof candidate.quantity === 'number'
    && typeof candidate.minimumStockLevel === 'number'
    && typeof candidate.responsiblePerson === 'string'
    && typeof candidate.supplier === 'string'
    && typeof candidate.targetPersonDepartment === 'string'
    && typeof candidate.unitPrice === 'number'
    && typeof candidate.totalCost === 'number'
    && typeof candidate.notes === 'string'
  );
};

const loadPpeTransactions = (): PpeTransactionRecord[] => {
  if (typeof window === 'undefined') {
    return initialPpeTransactions;
  }

  try {
    const raw = window.localStorage.getItem(PPE_TRANSACTIONS_STORAGE_KEY);
    if (!raw) {
      return initialPpeTransactions;
    }

    const parsed = JSON.parse(raw) as unknown;
    if (Array.isArray(parsed) && parsed.every(isPpeTransactionRecord)) {
      return parsed;
    }
  } catch {
    // Fall back to empty PPE records if local storage is unavailable or corrupted.
  }

  return initialPpeTransactions;
};

const initialWorkforceRecords: WorkforceRecord[] = [];

function MiniBars({ values }: { values: number[] }) {
  const max = Math.max(...values, 1);
  return (
    <div className="mini-bars">
      {values.map((value, index) => (
        <div key={index} className="mini-bar-wrap">
          <div className="mini-bar" style={{ height: `${Math.max((value / max) * 90, 8)}px` }} />
          <span>{value}</span>
        </div>
      ))}
    </div>
  );
}

function formatChartValue(value: number): string {
  if (value >= 1000) {
    return `${(value / 1000).toFixed(1)}k`;
  }
  return String(Math.round(value));
}

const chartThemes: Record<ChartThemeName, ChartTheme> = {
  operations: {
    primary: '#0284c7',
    secondary: '#06b6d4',
    accent: '#0ea5e9',
    danger: '#ef4444',
    neutral: '#64748b'
  },
  safety: {
    primary: '#0f766e',
    secondary: '#22c55e',
    accent: '#14b8a6',
    danger: '#dc2626',
    neutral: '#64748b'
  },
  risk: {
    primary: '#c2410c',
    secondary: '#f97316',
    accent: '#f59e0b',
    danger: '#b91c1c',
    neutral: '#64748b'
  },
  compliance: {
    primary: '#1d4ed8',
    secondary: '#0ea5e9',
    accent: '#22c55e',
    danger: '#ef4444',
    neutral: '#64748b'
  }
};

function getSectionChartTheme(sectionKey: string): ChartThemeName {
  if (sectionKey === 'safety-performance' || sectionKey === 'training' || sectionKey === 'occupational-health') {
    return 'safety';
  }
  if (sectionKey === 'incidents' || sectionKey === 'risk-management') {
    return 'risk';
  }
  if (sectionKey === 'compliance' || sectionKey === 'environment') {
    return 'compliance';
  }
  return 'operations';
}

function chartTrend(values: number[]): { label: string; tone: 'up' | 'down' | 'flat' } {
  const first = values[0] ?? 0;
  const last = values[values.length - 1] ?? 0;
  const diff = last - first;
  if (diff > 0) {
    const percent = first > 0 ? Math.round((diff / first) * 100) : 0;
    return { label: percent > 0 ? `Artış %${percent}` : `Artış ${formatChartValue(diff)}`, tone: 'up' };
  }
  if (diff < 0) {
    const percent = first > 0 ? Math.round((Math.abs(diff) / first) * 100) : 0;
    return { label: percent > 0 ? `Azalış %${percent}` : `Azalış ${formatChartValue(Math.abs(diff))}`, tone: 'down' };
  }
  return { label: 'Sabit', tone: 'flat' };
}

function LineChart({ values, labels, theme }: { values: number[]; labels?: string[]; theme: ChartTheme }) {
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const span = Math.max(max - min, 1);
  const trend = chartTrend(values);
  const points = values
    .map((value, index) => {
      const x = 6 + (index / Math.max(values.length - 1, 1)) * 88;
      const y = 92 - ((value - min) / span) * 78;
      return `${x},${y}`;
    })
    .join(' ');
  const latest = values[values.length - 1] ?? 0;

  return (
    <div className="viz-shell">
      <svg viewBox="0 0 100 100" className="viz-svg" preserveAspectRatio="none">
        <defs>
          <linearGradient id="lineGlow" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stopColor={theme.secondary} />
            <stop offset="100%" stopColor={theme.primary} />
          </linearGradient>
        </defs>
        {[20, 40, 60, 80].map((y) => (
          <line key={y} x1="6" y1={y} x2="94" y2={y} className="viz-grid-line" />
        ))}
        <polyline points={points} fill="none" stroke="url(#lineGlow)" strokeWidth="2.8" strokeLinejoin="round" strokeLinecap="round" />
        {values.map((value, index) => {
          const x = 6 + (index / Math.max(values.length - 1, 1)) * 88;
          const y = 92 - ((value - min) / span) * 78;
          return <circle key={`${index}-${value}`} cx={x} cy={y} r="1.7" className="viz-point" />;
        })}
      </svg>
      <div className="viz-summary-row">
        <span>Min: {formatChartValue(min)}</span>
        <span>Son: {formatChartValue(latest)}</span>
        <span>Max: {formatChartValue(max)}</span>
      </div>
      <div className={`viz-trend-chip ${trend.tone}`}>{trend.label}</div>
      {labels && labels.length === values.length ? (
        <div className="viz-axis-row">
          {labels.map((label, index) => (
            <span key={`${label}-${index}`}>{label}</span>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function AreaChart({ values, labels, theme }: { values: number[]; labels?: string[]; theme: ChartTheme }) {
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  const span = Math.max(max - min, 1);
  const trend = chartTrend(values);
  const linePoints = values
    .map((value, index) => {
      const x = 6 + (index / Math.max(values.length - 1, 1)) * 88;
      const y = 92 - ((value - min) / span) * 78;
      return `${x},${y}`;
    })
    .join(' ');
  const areaPoints = `6,92 ${linePoints} 94,92`;

  return (
    <div className="viz-shell">
      <svg viewBox="0 0 100 100" className="viz-svg" preserveAspectRatio="none">
        <defs>
          <linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor={theme.secondary} stopOpacity="0.46" />
            <stop offset="100%" stopColor={theme.primary} stopOpacity="0.06" />
          </linearGradient>
        </defs>
        {[20, 40, 60, 80].map((y) => (
          <line key={y} x1="6" y1={y} x2="94" y2={y} className="viz-grid-line" />
        ))}
        <polygon points={areaPoints} fill="url(#areaFill)" />
        <polyline points={linePoints} fill="none" stroke={theme.primary} strokeWidth="2.5" strokeLinejoin="round" strokeLinecap="round" />
      </svg>
      <div className="viz-summary-row">
        <span>Min: {formatChartValue(min)}</span>
        <span>Ortalama: {formatChartValue(values.reduce((sum, value) => sum + value, 0) / Math.max(values.length, 1))}</span>
        <span>Max: {formatChartValue(max)}</span>
      </div>
      <div className={`viz-trend-chip ${trend.tone}`}>{trend.label}</div>
      {labels && labels.length === values.length ? (
        <div className="viz-axis-row">
          {labels.map((label, index) => (
            <span key={`${label}-${index}`}>{label}</span>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function BarChart({
  values,
  labels,
  theme,
  showTrend = true,
  showAverage = true
}: {
  values: number[];
  labels?: string[];
  theme: ChartTheme;
  showTrend?: boolean;
  showAverage?: boolean;
}) {
  const max = Math.max(...values, 1);
  const trend = chartTrend(values);
  const columnCount = Math.max(values.length, 1);
  return (
    <div className="viz-shell">
      <div className="bar-viz" style={{ gridTemplateColumns: `repeat(${columnCount}, minmax(18px, 1fr))` }}>
        {values.map((value, index) => (
          <div key={index} className="bar-viz-item">
            <span className="bar-viz-value">{formatChartValue(value)}</span>
            <div className="bar-viz-col">
              <div
                className="bar-viz-fill"
                style={{
                  height: `${Math.max((value / max) * 100, 6)}%`,
                  background: `linear-gradient(180deg, ${theme.secondary} 0%, ${theme.primary} 100%)`
                }}
              />
            </div>
            {labels && labels.length === values.length ? <span className="bar-viz-axis-label">{labels[index]}</span> : null}
          </div>
        ))}
      </div>
      <div className="viz-summary-row">
        <span>Toplam: {formatChartValue(values.reduce((sum, value) => sum + value, 0))}</span>
        {showAverage ? <span>Ortalama: {formatChartValue(values.reduce((sum, value) => sum + value, 0) / Math.max(values.length, 1))}</span> : null}
      </div>
      {showTrend ? <div className={`viz-trend-chip ${trend.tone}`}>{trend.label}</div> : null}
    </div>
  );
}

function DonutChart({ values, theme, size = 'normal' }: { values: number[]; theme: ChartTheme; size?: 'normal' | 'large' }) {
  const palette = [theme.primary, theme.secondary, theme.accent, theme.danger, '#6366f1'];
  const safeValues = values.map((value) => Math.max(value, 0));
  const total = safeValues.reduce((sum, value) => sum + value, 0) || 1;
  let cursor = 0;
  const segments = safeValues.map((value, index) => {
    const share = (value / total) * 100;
    const start = cursor;
    const end = cursor + share;
    cursor = end;
    return {
      color: palette[index % palette.length],
      value,
      percent: Math.round(share),
      start,
      end
    };
  });
  const gradient = `conic-gradient(${segments.map((segment) => `${segment.color} ${segment.start}% ${segment.end}%`).join(', ')})`;

  return (
    <div className="donut-wrap enhanced-donut-wrap">
      <div className={`donut ${size === 'large' ? 'donut-large' : ''}`} style={{ background: gradient }}>
        <span>{formatChartValue(total)}</span>
      </div>
      <div className="chart-legend-row">
        {segments.slice(0, 4).map((segment, index) => (
          <div key={`${index}-${segment.value}`} className="chart-legend-item">
            <span className="chart-legend-dot" style={{ background: segment.color }} />
            <span>{segment.percent}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function GaugeChart({ values, theme }: { values: number[]; theme: ChartTheme }) {
  const value = Math.max(0, Math.min(100, values[0] ?? 0));
  const status = value >= 85 ? 'Güçlü' : value >= 65 ? 'Kontrol Et' : 'Kritik';
  const gaugeColor = value >= 85 ? theme.secondary : value >= 65 ? theme.accent : theme.danger;
  const handleStatusClick = () => {
    const target = document.getElementById('dashboard-safety-performance');
    target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    target?.focus({ preventScroll: true });
  };
  return (
    <div className="gauge-wrap">
      <div className="gauge-track" style={{ background: `conic-gradient(${gaugeColor} 0 ${value * 1.8}deg, #e2e8f0 ${value * 1.8}deg 180deg)` }} />
      <strong>{value}%</strong>
      {status === 'Kontrol Et' ? (
        <button type="button" className="gauge-status gauge-status-button" onClick={handleStatusClick}>
          {status}
        </button>
      ) : (
        <small className="gauge-status">{status}</small>
      )}
    </div>
  );
}

function HeatmapChart({ values, labels, theme }: { values: number[]; labels?: string[]; theme: ChartTheme }) {
  const max = Math.max(...values, 1);
  const min = Math.min(...values, 0);
  return (
    <div className="viz-shell">
      <div className="heatmap-grid">
        {values.map((value, index) => {
          const intensity = value / max;
          return (
            <div
              key={index}
              className="heat-cell"
              style={{ opacity: 0.22 + intensity * 0.78, background: theme.primary }}
              title={`${labels?.[index] ?? `Periyot ${index + 1}`}: ${formatChartValue(value)}`}
            />
          );
        })}
      </div>
      <div className="viz-summary-row">
        <span>Düşük: {formatChartValue(min)}</span>
        <span>Yüksek: {formatChartValue(max)}</span>
      </div>
      {labels && labels.length === values.length ? (
        <div className="viz-axis-row">
          {labels.map((label, index) => (
            <span key={`${label}-${index}`}>{label}</span>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function DashboardChart({
  type,
  values,
  xLabels,
  themeName = 'operations',
  showTrend = true,
  showAverage = true,
  donutSize = 'normal'
}: {
  type: 'line' | 'bar' | 'area' | 'donut' | 'gauge' | 'heatmap';
  values: number[];
  xLabels?: string[];
  themeName?: ChartThemeName;
  showTrend?: boolean;
  showAverage?: boolean;
  donutSize?: 'normal' | 'large';
}) {
  const theme = chartThemes[themeName];
  if (type === 'line') {
    return <LineChart values={values} labels={xLabels} theme={theme} />;
  }
  if (type === 'bar') {
    return <BarChart values={values} labels={xLabels} theme={theme} showTrend={showTrend} showAverage={showAverage} />;
  }
  if (type === 'area') {
    return <AreaChart values={values} labels={xLabels} theme={theme} />;
  }
  if (type === 'donut') {
    return <DonutChart values={values} theme={theme} size={donutSize} />;
  }
  if (type === 'gauge') {
    return <GaugeChart values={values} theme={theme} />;
  }
  return <HeatmapChart values={values} labels={xLabels} theme={theme} />;
}

function daysUntil(dateText: string): number {
  const now = new Date();
  const date = new Date(dateText);
  const diff = date.getTime() - now.getTime();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

function localizeStatus(status: Status): string {
  if (status === 'OPEN') {
    return 'Açık';
  }
  if (status === 'IN_PROGRESS') {
    return 'Devam Ediyor';
  }
  return 'Kapalı';
}

const cityCoordinates: Record<string, { lat: number; lon: number }> = {
  Istanbul: { lat: 41.0082, lon: 28.9784 },
  Izmir: { lat: 38.4237, lon: 27.1428 },
  Ankara: { lat: 39.9334, lon: 32.8597 }
};

const weatherCodeMeta = (code: number, language: Language) => {
  const labels = {
    clear: language === 'en' ? 'Clear' : language === 'ru' ? 'Ясно' : 'Açık',
    mostlyClear: language === 'en' ? 'Mostly Clear' : language === 'ru' ? 'Малооблачно' : 'Az Bulutlu',
    cloudy: language === 'en' ? 'Cloudy' : language === 'ru' ? 'Облачно' : 'Bulutlu',
    fog: language === 'en' ? 'Fog' : language === 'ru' ? 'Туман' : 'Sisli',
    drizzle: language === 'en' ? 'Drizzle' : language === 'ru' ? 'Морось' : 'Çiseleme',
    rain: language === 'en' ? 'Rain' : language === 'ru' ? 'Дождь' : 'Yağmurlu',
    snow: language === 'en' ? 'Snow' : language === 'ru' ? 'Снег' : 'Karlı',
    storm: language === 'en' ? 'Thunderstorm' : language === 'ru' ? 'Гроза' : 'Fırtına'
  };

  if (code === 0) return { icon: '☀', text: labels.clear };
  if (code === 1 || code === 2) return { icon: '⛅', text: labels.mostlyClear };
  if (code === 3) return { icon: '☁', text: labels.cloudy };
  if (code === 45 || code === 48) return { icon: '🌫', text: labels.fog };
  if ([51, 53, 55, 56, 57].includes(code)) return { icon: '🌦', text: labels.drizzle };
  if ([61, 63, 65, 66, 67, 80, 81, 82].includes(code)) return { icon: '🌧', text: labels.rain };
  if ([71, 73, 75, 77, 85, 86].includes(code)) return { icon: '❄', text: labels.snow };
  return { icon: '⛈', text: labels.storm };
};

const directionFromDegrees = (degrees: number, language: Language) => {
  const trDirections = ['K', 'KD', 'D', 'GD', 'G', 'GB', 'B', 'KB'];
  const enDirections = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW'];
  const ruDirections = ['С', 'СВ', 'В', 'ЮВ', 'Ю', 'ЮЗ', 'З', 'СЗ'];
  const labels = language === 'en' ? enDirections : language === 'ru' ? ruDirections : trDirections;
  const normalized = ((degrees % 360) + 360) % 360;
  const index = Math.round(normalized / 45) % 8;
  return labels[index];
};

const AUTH_EMAIL_SESSION_KEY = 'HSE_AUTH_EMAIL';

const detectLanguageFromPath = (pathname: string): Language => {
  const normalized = pathname.toLowerCase();
  if (normalized.startsWith('/ru')) {
    return 'ru';
  }
  if (normalized.startsWith('/eng') || normalized.startsWith('/en')) {
    return 'en';
  }
  return 'tr';
};

export function App() {
  const [language, setLanguage] = useState<Language>(() => {
    if (typeof window === 'undefined') {
      return 'tr';
    }
    return detectLanguageFromPath(window.location.pathname);
  });
  const [accessLevel, setAccessLevel] = useState<'locked' | 'authorized' | 'guest'>('locked');
  const [loginEmail, setLoginEmail] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');

  const allowedCredentials = useMemo(() => {
    const credentials: Record<string, string> = {};

    const pairRaw = (import.meta.env.VITE_AUTH_CREDENTIALS ?? '').trim();
    if (pairRaw.length > 0) {
      pairRaw
        .split(';')
        .map((entry: string) => entry.trim())
        .filter((entry: string) => entry.length > 0)
        .forEach((entry: string) => {
          const dividerIndex = entry.indexOf(':');
          if (dividerIndex <= 0) {
            return;
          }
          const email = entry.slice(0, dividerIndex).trim().toLowerCase();
          const password = entry.slice(dividerIndex + 1).trim();
          if (email && password) {
            credentials[email] = password;
          }
        });
    }

    if (Object.keys(credentials).length > 0) {
      return credentials;
    }

    const fallbackEmails = (import.meta.env.VITE_ALLOWED_EMAILS ?? '')
      .split(',')
      .map((item: string) => item.trim().toLowerCase())
      .filter((item: string) => item.length > 0);
    const fallbackPassword = (import.meta.env.VITE_ACCESS_PASSWORD ?? 'ChangeMe123!').trim();

    if (fallbackEmails.length > 0) {
      fallbackEmails.forEach((email: string) => {
        credentials[email] = fallbackPassword;
      });
      return credentials;
    }

    credentials['admin@gohse.blog'] = fallbackPassword;
    return credentials;
  }, []);

  const allowedEmails = useMemo(() => Object.keys(allowedCredentials), [allowedCredentials]);

  const isAuthorizedViewer = accessLevel === 'authorized';
  const t = uiText[language];
  const moduleLabels = moduleLabelsByLanguage[language];

  const [activeModule, setActiveModule] = useState<ModuleKey>('dashboard');
  const [projectFilter, setProjectFilter] = useState<string>('all');
  const [expandedGroups, setExpandedGroups] = useState<Record<GroupKey, boolean>>({
    operations: true,
    workforce: true,
    compliance: true,
    reports: true,
    system: true
  });

  const [moduleData, setModuleData] = useState<Record<ModuleKey, ModuleRecord[]>>(() => {
    if (isAuthorizedViewer) {
      return recordsByModule;
    }

    const emptyModuleData = {} as Record<ModuleKey, ModuleRecord[]>;
    (Object.keys(recordsByModule) as ModuleKey[]).forEach((key) => {
      emptyModuleData[key] = [];
    });
    return emptyModuleData;
  });
  const [editingEnvironmentalIndex, setEditingEnvironmentalIndex] = useState<number | null>(null);
  const [editingModuleRecord, setEditingModuleRecord] = useState<{ module: Exclude<ModuleKey, 'dashboard' | 'occupational-health' | 'projects' | 'legal-register' | 'documents' | 'reports' | 'export-center' | 'departments' | 'contractors' | 'settings' | 'inspections' | 'observations' | 'risk-assessments' | 'permit-to-work' | 'incidents' | 'environmental' | 'action-tracker' | 'kpis-analytics' | 'equipment-management' | 'emergency-preparedness' | 'employees' | 'trainings' | 'ppe-stocks'>; index: number } | null>(null);
  const [projectCatalog, setProjectCatalog] = useState<Project[]>(() => (isAuthorizedViewer ? loadProjectCatalog() : []));
  const [projectForm, setProjectForm] = useState<Omit<Project, 'id'>>({
    name: '',
    country: '',
    city: '',
    address: '',
    contractScope: ''
  });
  const [editingProjectId, setEditingProjectId] = useState<string | null>(null);
  const [projectSaveFeedback, setProjectSaveFeedback] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [departmentRecords, setDepartmentRecords] = useState<DepartmentRecord[]>(() => (isAuthorizedViewer ? loadDepartmentRecords() : []));
  const [departmentForm, setDepartmentForm] = useState<DepartmentForm>({
    name: '',
    code: '',
    description: ''
  });
  const [editingDepartmentId, setEditingDepartmentId] = useState<string | null>(null);
  const departmentNames = useMemo(() => departmentRecords.map((department) => department.name), [departmentRecords]);
  const [contractorRecords, setContractorRecords] = useState<ContractorMasterRecord[]>(() => (isAuthorizedViewer ? loadContractorRecords() : []));
  const [contractorForm, setContractorForm] = useState<ContractorMasterForm>({
    companyName: '',
    projectId: projectCatalog[0]?.id ?? '',
    projectName: projectCatalog[0]?.name ?? '',
    country: projectCatalog[0]?.country ?? '',
    city: projectCatalog[0]?.city ?? '',
    projectLocation: projectCatalog[0]?.address ?? '',
    contractScope: projectCatalog[0]?.contractScope ?? '',
    hseWarningCount: 0,
    hseWarningDate: '',
    fireWarningCount: 0,
    fireWarningDate: '',
    environmentWarningCount: 0,
    environmentWarningDate: '',
    penaltyCount: 0,
    penaltyLegalClause: '',
    totalPenaltyAmount: 0
  });
  const [editingContractorId, setEditingContractorId] = useState<string | null>(null);
  const [activeSettingsTab, setActiveSettingsTab] = useState<SettingsTabKey>('company-info');
  const [settingsConfig, setSettingsConfig] = useState<SettingsConfig>(initialSettingsConfig);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    const modeFromQuery = new URLSearchParams(window.location.search).get('mode');
    const rememberedEmail = window.sessionStorage.getItem(AUTH_EMAIL_SESSION_KEY)?.toLowerCase() ?? '';

    if (rememberedEmail && allowedEmails.includes(rememberedEmail)) {
      setAccessLevel('authorized');
      setLoginEmail(rememberedEmail);
      return;
    }

    if (modeFromQuery === 'guest') {
      setAccessLevel('guest');
      return;
    }

    setAccessLevel('locked');
  }, [allowedEmails]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    window.__HSE_GUEST_MODE__ = accessLevel === 'guest';
  }, [accessLevel]);

  useEffect(() => {
    if (typeof window === 'undefined' || language === 'tr') {
      return;
    }

    const replacementMap = buildTurkishReplacementMap(language);
    const replacementEntries = Array.from(replacementMap.entries()).sort((a, b) => b[0].length - a[0].length);
    const root = document.querySelector('.page') ?? document.body;

    applyTextReplacementsToNode(root, replacementEntries);

    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((addedNode) => {
          applyTextReplacementsToNode(addedNode, replacementEntries);
        });

        if (mutation.type === 'characterData' && mutation.target.nodeType === Node.TEXT_NODE) {
          const textNode = mutation.target as Text;
          const current = textNode.nodeValue ?? '';
          const replaced = replaceByDictionary(current, replacementEntries);
          if (replaced !== current) {
            textNode.nodeValue = replaced;
          }
        }
      });
    });

    observer.observe(root, { childList: true, subtree: true, characterData: true });

    return () => {
      observer.disconnect();
    };
  }, [language, activeModule, projectFilter, accessLevel]);

  const inspectionSections = language === 'ru' ? inspectionChecklistSectionsRu : inspectionChecklistSections;

  const inspectionCopy = {
    dataEntryTitle: language === 'ru' ? 'Ввод данных инспекции' : 'Denetim Veri Girişi',
    summaryTitle: language === 'ru' ? 'Сводка инспекции' : 'Denetim Özeti',
    checklistTitle: language === 'ru' ? 'Контрольный список инспекции' : 'Denetim Kontrol Listesi',
    correctiveActionsTitle: language === 'ru' ? 'Корректирующие действия' : 'Düzeltici Faaliyetler',
    reportsTitle: language === 'ru' ? 'Операции с отчетом' : 'Rapor İşlemleri',
    titleLabel: language === 'ru' ? 'Заголовок' : 'Başlık',
    businessUnitLabel: language === 'ru' ? 'Бизнес-единица / Компания' : 'İş Birimi / Şirket',
    siteAreaLabel: language === 'ru' ? 'Площадка / Зона' : 'Saha / Alan',
    locationTypeLabel: language === 'ru' ? 'Тип локации' : 'Lokasyon Türü',
    selectOption: language === 'ru' ? 'Выберите' : 'Seçiniz',
    inspectionDateLabel: language === 'ru' ? 'Дата инспекции' : 'Denetim Tarihi',
    inspectorLabel: language === 'ru' ? 'ФИО инспектора' : 'Denetçi Ad Soyad',
    departmentLabel: language === 'ru' ? 'Отдел' : 'Departman',
    positiveObservationsLabel: language === 'ru' ? 'Позитивные предписания' : 'Pozitif Gözlemler',
    inspectionNotesLabel: language === 'ru' ? 'Примечания инспекции' : 'Denetim Notları',
    progressLabel: language === 'ru' ? 'Прогресс' : 'İlerleme',
    totalFindingsLabel: language === 'ru' ? 'Всего замечаний' : 'Toplam Bulgular',
    openActionsLabel: language === 'ru' ? 'Открытые действия' : 'Açık Aksiyonlar',
    answeredProgressLabel: language === 'ru' ? 'Отвеченный прогресс' : 'Yanıtlanan İlerleme',
    complianceScoreLabel: language === 'ru' ? 'Оценка соответствия' : 'Uygunluk skoru',
    highFindingsLabel: language === 'ru' ? 'Высокие замечания' : 'Yüksek Bulgular',
    mediumFindingsLabel: language === 'ru' ? 'Средние замечания' : 'Orta Bulgular',
    lowFindingsLabel: language === 'ru' ? 'Низкие замечания' : 'Düşük Bulgular',
    skipChecklistLabel: language === 'ru' ? 'Пропустить этот контрольный список' : 'Bu kontrol listesini atla',
    yesLabel: language === 'ru' ? 'Да' : 'Evet',
    noLabel: language === 'ru' ? 'Нет' : 'Hayır',
    naLabel: language === 'ru' ? 'Неприменимо' : 'Uygulanamaz',
    correctiveHint: language === 'ru' ? 'Отслеживайте замечания, требующие действий.' : 'Takip gerektiren bulguları izleyin.',
    statusOpen: language === 'ru' ? 'Открыто' : 'Açık',
    statusInProgress: language === 'ru' ? 'В процессе' : 'Devam Ediyor',
    statusClosed: language === 'ru' ? 'Закрыто' : 'Kapalı',
    addLabel: language === 'ru' ? 'Добавить' : 'Ekle',
    manualActionPlaceholder: language === 'ru' ? 'Добавить действие вручную' : 'Manuel aksiyon ekle',
    noActionsYet: language === 'ru' ? 'Корректирующие действия еще не добавлены.' : 'Henüz düzeltici faaliyet eklenmedi.',
    reportHint:
      language === 'ru'
        ? 'Экспортируйте краткий отчет: сводка, замечания, неприменимые пункты и корректирующие действия.'
        : 'Özet, bulgular, uygun değil/uygulanamaz maddeler ve düzeltici faaliyetleri içeren kısa raporu dışa aktarın.',
    exportDocx: language === 'ru' ? 'Экспорт DOCX' : 'DOCX Dışa Aktar',
    exportActionCsv: language === 'ru' ? 'Экспорт CSV действий' : 'Aksiyon CSV Dışa Aktar',
    exportFullCsv: language === 'ru' ? 'Экспорт полного CSV' : 'Tam CSV Dışa Aktar',
    sendEmailSummary: language === 'ru' ? 'Отправить сводку по e-mail' : 'Özeti E-posta ile Gönder',
    printPdf: language === 'ru' ? 'Печать / PDF' : 'Yazdır / PDF',
    resetInspection: language === 'ru' ? 'Сбросить инспекцию' : 'Denetimi Sıfırla',
    preparedBy: language === 'ru' ? 'Подготовлено Erdem Cetin' : 'Erdem Cetin tarafından hazırlandı'
  };

  const observationCopy = {
    moduleSubtitle: language === 'ru' ? 'Ввод данных и аналитика предписаний' : 'Gözlem veri girişi ve analizi',
    formTitle: language === 'ru' ? 'Корпоративная форма регистрации предписаний по охране труда' : 'Kurumsal İSG Gözlem Giriş Formu',
    observationNo: language === 'ru' ? '№ предписания' : 'Gözlem No',
    project: language === 'ru' ? 'Проект' : 'Proje',
    selectProject: language === 'ru' ? 'Выберите проект' : 'Proje seçin',
    projectLocation: language === 'ru' ? 'Местоположение проекта' : 'Proje Konumu',
    inspectionDate: language === 'ru' ? 'Дата проверки' : 'Denetim Tarihi',
    inspectionTime: language === 'ru' ? 'Время проверки' : 'Denetim Saati',
    inspectorName: language === 'ru' ? 'Инспектор' : 'Denetçi Adı',
    contractor: language === 'ru' ? 'Генеральный подрядчик' : 'Yüklenici',
    subcontractor: language === 'ru' ? 'Субподрядчик' : 'Alt Yüklenici',
    responsiblePerson: language === 'ru' ? 'Ответственный' : 'Sorumlu Kişi',
    category: language === 'ru' ? 'Категория предписания' : 'Gözlem Kategorisi',
    subject: language === 'ru' ? 'Тема предписания' : 'Gözlem Konusu',
    observationLocation: language === 'ru' ? 'Место выявления' : 'Gözlem Yeri',
    violatedRequirement: language === 'ru' ? 'Нарушенное требование' : 'İhlal Edilen Gereklilik',
    observationDescription: language === 'ru' ? 'Описание нарушения' : 'Gözlem Açıklaması',
    correctiveAction: language === 'ru' ? 'Корректирующие мероприятия' : 'Gereken Düzeltici Faaliyet',
    dueDate: language === 'ru' ? 'Срок устранения' : 'Bitiş Tarihi',
    priority: language === 'ru' ? 'Приоритет' : 'Öncelik',
    status: language === 'ru' ? 'Статус' : 'Durum',
    comment: language === 'ru' ? 'Комментарий' : 'Yorum',
    attachmentsTitle: language === 'ru' ? 'Фото-приложения' : 'Fotoğraf Ekleri',
    dragDropHint: language === 'ru' ? 'Перетащите файлы сюда или выберите файл (jpg, jpeg, png, pdf)' : 'Sürükle ve Bırak veya dosya seçin (jpg, jpeg, png, pdf)',
    browseFiles: language === 'ru' ? 'Выбрать файлы' : 'Dosyalara Göz At',
    removeImage: language === 'ru' ? 'Удалить файл' : 'Resmi Kaldır',
    saveObservation: language === 'ru' ? 'Сохранить предписание' : 'Kaydet',
    updateObservation: language === 'ru' ? 'Обновить предписание' : 'Güncelle',
    cancel: language === 'ru' ? 'Отмена' : 'Vazgeç',
    registryTitle: language === 'ru' ? 'Реестр предписаний' : 'Gözlem Kayıtları',
    tableCategory: language === 'ru' ? 'Категория' : 'Kategori',
    tableSubject: language === 'ru' ? 'Тема' : 'Konu',
    tableActions: language === 'ru' ? 'Действия' : 'İşlemler',
    view: language === 'ru' ? 'Просмотр' : 'Görüntüle',
    edit: language === 'ru' ? 'Редактировать' : 'Düzenle',
    delete: language === 'ru' ? 'Удалить' : 'Sil',
    reportActions: language === 'ru' ? 'Операции с отчетом' : 'Rapor İşlemleri',
    generatePdf: language === 'ru' ? 'Сформировать PDF' : 'PDF Oluştur',
    printReport: language === 'ru' ? 'Печать отчета' : 'Raporu Yazdır',
    noPhotoAttached: language === 'ru' ? 'Фото не приложены.' : 'Ekli fotoğraf bulunmamaktadır.',
    pdfAttachment: language === 'ru' ? 'PDF-приложение' : 'PDF Ek',
    confirmationDelete: language === 'ru' ? 'Удалить предписание №' : 'numaralı gözlem kaydı silinsin mi?',
    reportTitle: language === 'ru' ? 'Форма предписания' : 'Gözlem Formu',
    reportHeaderTitle: language === 'ru' ? 'КОРПОРАТИВНАЯ ФОРМА ПРЕДПИСАНИЯ ПО ОХРАНЕ ТРУДА' : 'KURUMSAL İSG GÖZLEM FORMU',
    formNo: language === 'ru' ? 'Форма №' : 'Form No',
    revision: language === 'ru' ? 'Ревизия' : 'Revizyon',
    reportSystemName: language === 'ru' ? 'Корпоративная система управления предписаниями по ОТ и ПБ' : 'Kurumsal İSG Gözlem Yönetim Sistemi',
    page: language === 'ru' ? 'Страница' : 'Sayfa',
    preparedBy: language === 'ru' ? 'Подготовил' : 'Hazırlayan',
    reviewedBy: language === 'ru' ? 'Проверил' : 'Kontrol Eden',
    approvedBy: language === 'ru' ? 'Утвердил' : 'Onaylayan',
    nameSignature: language === 'ru' ? 'ФИО / Подпись' : 'Ad Soyad / İmza'
  };

  const riskCopy = {
    formTitle: language === 'ru' ? 'Форма оценки профессиональных рисков' : 'Risk Değerlendirmeleri Veri Girişi',
    introLine1: language === 'ru'
      ? 'Оценка риска — показатель уровня риска, автоматически рассчитываемый как произведение вероятности и тяжести последствий.'
      : 'Risk Skoru: Tehlikenin ciddiyetini gösteren puandır ve Olasılık x Şiddet ile otomatik hesaplanır.',
    introLine2: language === 'ru'
      ? 'Остаточный риск — уровень риска после внедрения мероприятий по управлению рисками, автоматически рассчитываемый как произведение остаточной вероятности и остаточной тяжести последствий.'
      : 'Kalan Risk Puanı: Kontroller uygulandıktan sonra kalan risk seviyesidir ve Kalan Olasılık x Kalan Şiddet ile hesaplanır.',
    riskId: language === 'ru' ? '№ оценки риска' : 'Risk Kimliği',
    initialRiskScore: language === 'ru' ? 'Уровень риска' : 'Başlangıç Risk Puanı',
    residualRiskScore: language === 'ru' ? 'Остаточный риск' : 'Kalan Risk Puanı',
    project: language === 'ru' ? 'Проект' : 'Proje',
    departmentActivity: language === 'ru' ? 'Подразделение / Вид деятельности' : 'Departman / Faaliyet',
    departmentActivityPlaceholder: language === 'ru' ? 'Пример: Электромонтаж / оконцевание кабеля' : 'Orn: Elektrik / Kablo sonlandırma',
    assessmentDate: language === 'ru' ? 'Дата оценки' : 'Değerlendirme Tarihi',
    targetDate: language === 'ru' ? 'Плановая дата устранения' : 'Hedef Tamamlanma Tarihi',
    hazard: language === 'ru' ? 'Опасность' : 'Tehlike',
    hazardPlaceholder: language === 'ru' ? 'Кратко и четко опишите опасность' : 'Tehlikeyi kısa ve net yazın',
    potentialConsequence: language === 'ru' ? 'Возможные последствия' : 'Olası Sonuç',
    existingControls: language === 'ru' ? 'Существующие меры контроля' : 'Mevcut Kontroller',
    recommendedControls: language === 'ru' ? 'Рекомендуемые мероприятия' : 'Önerilen Kontrol Önlemleri',
    likelihood: language === 'ru' ? 'Вероятность (1-5)' : 'Olasılık (1-5)',
    severity: language === 'ru' ? 'Тяжесть последствий (1-5)' : 'Şiddet (1-5)',
    residualLikelihood: language === 'ru' ? 'Остаточная вероятность (1-5)' : 'Kalan Olasılık (1-5)',
    residualSeverity: language === 'ru' ? 'Остаточная тяжесть (1-5)' : 'Kalan Şiddet (1-5)',
    responsiblePerson: language === 'ru' ? 'Ответственный' : 'Sorumlu Kişi',
    status: language === 'ru' ? 'Статус' : 'Durum',
    attachments: language === 'ru' ? 'Вложения' : 'Ekler',
    photos: language === 'ru' ? 'Фотографии' : 'Fotoğraflar',
    notes: language === 'ru' ? 'Примечания' : 'Notlar',
    save: language === 'ru' ? 'Сохранить оценку риска' : 'Kaydet',
    tableTitle: language === 'ru' ? 'Реестр оценок рисков' : 'Risk Değerlendirmeleri Kayıtlar',
    activity: language === 'ru' ? 'Вид деятельности' : 'Faaliyet',
    detailTitle: language === 'ru' ? 'Карточка оценки риска' : 'Risk Detayı',
    initialLevel: language === 'ru' ? 'Начальный уровень' : 'Başlangıç Seviye',
    residualLevel: language === 'ru' ? 'Остаточный уровень' : 'Kalan Seviye',
    due: language === 'ru' ? 'Срок' : 'Termin',
    hazardAndConsequence: language === 'ru' ? 'Опасность и последствия' : 'Tehlike ve Sonuç',
    controlPlan: language === 'ru' ? 'План мероприятий' : 'Kontrol Planı',
    scores: language === 'ru' ? 'Показатели риска' : 'Puanlar',
    responsibilityAndDue: language === 'ru' ? 'Ответственность и срок' : 'Sorumluluk ve Termin',
    noAttachment: language === 'ru' ? 'Вложения отсутствуют.' : 'Ek bulunmuyor.',
    noPhoto: language === 'ru' ? 'Фотографии отсутствуют.' : 'Fotoğraf bulunmuyor.'
  };

  const ptwCopy = {
    formTitle: language === 'ru' ? 'Корпоративная система регистрации нарядов-допусков (PTW)' : 'Çalışma İzni (PTW) - Kurumsal Narad-Dopusk Kayıt Ekranı',
    tabMain: language === 'ru' ? 'Основная информация' : 'PTW Kayıt Bilgileri',
    tabResponsible: language === 'ru' ? 'Ответственные лица' : 'Sorumlular',
    tabTeam: language === 'ru' ? 'Состав бригады' : 'Ekip Bilgileri',
    tabWork: language === 'ru' ? 'Информация о работе' : 'İş Bilgileri',
    tabRisks: language === 'ru' ? 'Риски' : 'Riskler',
    tabHeightSafety: language === 'ru' ? 'Системы безопасности при работах на высоте' : 'Yüksekte Çalışma Güvenlik Sistemleri',
    tabEquipment: language === 'ru' ? 'Используемое оборудование' : 'Kullanılacak Ekipman',
    tabPreChecks: language === 'ru' ? 'Проверки перед началом работ' : 'Çalışma Öncesi Kontroller',
    tabPrecautions: language === 'ru' ? 'Меры безопасности во время выполнения работ' : 'Çalışma Sırasında Alınacak Tedbirler',
    tabSpecial: language === 'ru' ? 'Особые условия' : 'Özel Şartlar',
    tabApprovals: language === 'ru' ? 'Согласования' : 'Onaylar',
    tabDaily: language === 'ru' ? 'Ежедневный контроль наряда-допуска' : 'Günlük PTW Takibi',
    tabTeamChange: language === 'ru' ? 'Изменения состава бригады' : 'Ekip Değişiklikleri',
    tabClose: language === 'ru' ? 'Закрытие наряда' : 'Kapanış',
    tabAttachments: language === 'ru' ? 'Вложения' : 'Dosya Ekleri',
    sectionMain: language === 'ru' ? 'Информация о наряде-допуске' : 'PTW Kayıt Bilgileri',
    ptwNo: language === 'ru' ? '№ наряда-допуска (автоматически)' : 'PTW No (Otomatik)',
    organization: language === 'ru' ? 'Организация' : 'Organizasyon',
    department: language === 'ru' ? 'Подразделение' : 'Departman',
    location: language === 'ru' ? 'Место выполнения работ' : 'Lokasyon',
    ptwType: language === 'ru' ? 'Тип наряда-допуска' : 'PTW Türü',
    issueDate: language === 'ru' ? 'Дата оформления' : 'Düzenlenme Tarihi',
    validFrom: language === 'ru' ? 'Дата начала действия' : 'Geçerlilik Başlangıcı',
    validTo: language === 'ru' ? 'Дата окончания действия' : 'Geçerlilik Bitişi',
    status: language === 'ru' ? 'Статус' : 'Durum',
    reportActions: language === 'ru' ? 'Отчёты' : 'Rapor İşlemleri',
    summaryReport: language === 'ru' ? 'Сводный отчёт' : 'Özet Rapor',
    exportDocx: language === 'ru' ? 'Экспорт в DOCX' : 'DOCX Dışa Aktar',
    exportPdf: language === 'ru' ? 'Экспорт в PDF' : 'PDF Dışa Aktar',
    exportActionCsv: language === 'ru' ? 'Экспорт действий в CSV' : 'Aksiyon CSV Dışa Aktar',
    exportFullCsv: language === 'ru' ? 'Полный экспорт CSV' : 'Tam CSV Dışa Aktar',
    sendMail: language === 'ru' ? 'Отправить отчёт по электронной почте' : 'Özeti E-posta ile Gönder',
    printPdf: language === 'ru' ? 'Печать / PDF' : 'Yazdır / PDF',
    downloadAttachments: language === 'ru' ? 'Скачать вложения' : 'Ekleri İndir',
    reset: language === 'ru' ? 'Сбросить форму' : 'PTW Sıfırla',
    saveDraft: language === 'ru' ? 'Сохранить как черновик' : 'Taslak Kaydet',
    submitApproval: language === 'ru' ? 'Отправить на согласование' : 'Onaya Gönder',
    closePtw: language === 'ru' ? 'Закрыть наряд-допуск' : 'PTW Kapat'
  };

  const locationTypesLocalized =
    language === 'ru'
      ? [
          'Офис',
          'Склад',
          'Мастерская',
          'Шоурум',
          'Проживание',
          'Проектная площадка',
          'Зона работ субподрядчика',
          'Другое'
        ]
      : locationTypes;

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(projectCatalog));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [isAuthorizedViewer, projectCatalog]);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    const syncProjectCatalog = async () => {
      try {
        const rows = await apiRequest<Array<{ id: string; name: string; country: string; city: string; address: string; contractScope: string }>>('/api/master/projects');
        if (Array.isArray(rows) && rows.length > 0) {
          setProjectCatalog(rows.map((row) => ({
            id: row.id,
            name: row.name,
            country: row.country,
            city: row.city,
            address: row.address,
            contractScope: row.contractScope
          })));
          return;
        }

        if (rows.length === 0) {
          setProjectCatalog([]);
        }
      } catch {
        // Keep local fallback if API is unavailable.
      }
    };

    void syncProjectCatalog();
  }, [isAuthorizedViewer]);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(DEPARTMENTS_STORAGE_KEY, JSON.stringify(departmentRecords));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [departmentRecords, isAuthorizedViewer]);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    const syncDepartments = async () => {
      try {
        const rows = await apiRequest<Array<{ id: string; name: string; code: string; description: string | null }>>('/api/master/departments');
        if (Array.isArray(rows) && rows.length > 0) {
          setDepartmentRecords(rows.map((row) => ({
            id: row.id,
            name: row.name,
            code: row.code,
            description: row.description ?? ''
          })));
        }
      } catch {
        // Keep local fallback if API is unavailable.
      }
    };

    void syncDepartments();
  }, [isAuthorizedViewer]);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(CONTRACTORS_STORAGE_KEY, JSON.stringify(contractorRecords));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [contractorRecords, isAuthorizedViewer]);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    const syncContractors = async () => {
      try {
        const rows = await apiRequest<Array<{
          id: string;
          companyName: string;
          projectId: string | null;
          country: string;
          city: string;
          projectLocation: string;
          contractScope: string;
          hseWarningCount: number;
          hseWarningDate: string | null;
          fireWarningCount: number;
          fireWarningDate: string | null;
          environmentWarningCount: number;
          environmentWarningDate: string | null;
          penaltyCount: number;
          penaltyLegalClause: string | null;
          totalPenaltyAmount: number;
        }>>('/api/master/contractors');
        if (Array.isArray(rows) && rows.length > 0) {
          setContractorRecords(rows.map((row) => ({
            id: row.id,
            companyName: row.companyName,
            projectId: row.projectId ?? projectCatalog[0]?.id ?? '',
            projectName: projectCatalog.find((project) => project.id === (row.projectId ?? ''))?.name ?? '',
            country: row.country,
            city: row.city,
            projectLocation: row.projectLocation,
            contractScope: row.contractScope,
            hseWarningCount: row.hseWarningCount,
            hseWarningDate: row.hseWarningDate ?? '',
            fireWarningCount: row.fireWarningCount,
            fireWarningDate: row.fireWarningDate ?? '',
            environmentWarningCount: row.environmentWarningCount,
            environmentWarningDate: row.environmentWarningDate ?? '',
            penaltyCount: row.penaltyCount,
            penaltyLegalClause: row.penaltyLegalClause ?? '',
            totalPenaltyAmount: row.totalPenaltyAmount
          })));
        }
      } catch {
        // Keep local fallback if API is unavailable.
      }
    };

    void syncContractors();
  }, [isAuthorizedViewer, projectCatalog]);

  const [healthRecords, setHealthRecords] = useState<HealthRecord[]>(() => (isAuthorizedViewer ? loadHealthRecords() : []));

  const [healthForm, setHealthForm] = useState<HealthRecord>({
    employeeName: '',
    employeeId: '',
    department: '',
    position: '',
    bloodGroup: '',
    allergies: '',
    chronicDisease: '',
    communicableDisease: '',
    medication: '',
    disabilityStatus: '',
    fitForWork: 'Yes',
    restrictedWork: '',
    medicalExaminationDate: new Date().toISOString().slice(0, 10),
    nextMedicalExamination: new Date().toISOString().slice(0, 10),
    vaccinationStatus: '',
    remarks: ''
  });
  const [editingHealthEmployeeId, setEditingHealthEmployeeId] = useState<string | null>(null);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(HEALTH_RECORDS_STORAGE_KEY, JSON.stringify(healthRecords));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [healthRecords, isAuthorizedViewer]);

  const [legalRecords, setLegalRecords] = useState<LegalRecord[]>(() => (isAuthorizedViewer ? initialLegalRecords : []));
  const [selectedLegalRecordId, setSelectedLegalRecordId] = useState<string | null>(() => (isAuthorizedViewer ? initialLegalRecords[0]?.id ?? null : null));
  const [legalUserRole, setLegalUserRole] = useState<LegalUserRole>('CORPORATE_HSE_MANAGER');
  const [legalViewerZoom, setLegalViewerZoom] = useState<number>(120);
  const [legalViewerPage, setLegalViewerPage] = useState<number>(1);
  const [legalViewerSearch, setLegalViewerSearch] = useState<string>('');
  const [selectedLegalDocumentId, setSelectedLegalDocumentId] = useState<string | null>(null);
  const [legalFilters, setLegalFilters] = useState<{
    projectId: string;
    category: string;
    authority: string;
    complianceStatus: 'ALL' | LegalComplianceStatus;
    keyword: string;
  }>({
    projectId: 'all',
    category: 'all',
    authority: 'all',
    complianceStatus: 'ALL',
    keyword: ''
  });
  const [legalForm, setLegalForm] = useState<LegalRecordForm>({
    projectId: projectCatalog[0]?.id ?? '',
    category: legalCategoryOptions[0],
    title: '',
    authority: legalAuthorityOptions[0],
    department: departmentNames[0] ?? '',
    legalRequirement: '',
    responsiblePerson: '',
    complianceStatus: 'UYUMLU',
    effectiveDate: new Date().toISOString().slice(0, 10),
    lastReviewDate: new Date().toISOString().slice(0, 10),
    nextReviewDate: new Date().toISOString().slice(0, 10),
    openActions: 0,
    riskLevel: 'ORTA',
    notes: ''
  });

  const [controlledDocuments, setControlledDocuments] = useState<ControlledDocumentRecord[]>(() => (isAuthorizedViewer ? initialControlledDocumentRecords : []));
  const [selectedControlledDocumentId, setSelectedControlledDocumentId] = useState<string | null>(() => (isAuthorizedViewer ? initialControlledDocumentRecords[0]?.id ?? null : null));
  const [controlledDocumentFilters, setControlledDocumentFilters] = useState<{
    projectId: string;
    category: string;
    status: 'ALL' | ControlledDocumentStatus;
    revision: string;
    department: string;
    keyword: string;
  }>({
    projectId: 'all',
    category: 'all',
    status: 'ALL',
    revision: 'all',
    department: 'all',
    keyword: ''
  });
  const [controlledDocumentForm, setControlledDocumentForm] = useState<ControlledDocumentForm>({
    projectId: projectCatalog[0]?.id ?? '',
    title: '',
    category: controlledDocumentCategoryOptions[0],
    documentType: controlledDocumentTypeOptions[0],
    effectiveDate: new Date().toISOString().slice(0, 10),
    reviewDate: new Date().toISOString().slice(0, 10),
    department: departmentNames[0] ?? '',
    preparedBy: '',
    reviewedBy: '',
    approvedBy: '',
    status: 'TASLAK',
    notes: ''
  });
  const [controlledDocumentFileList, setControlledDocumentFileList] = useState<FileList | null>(null);
  const [controlledDocumentRevisionFileList, setControlledDocumentRevisionFileList] = useState<FileList | null>(null);
  const [controlledDocumentRevisionStatus, setControlledDocumentRevisionStatus] = useState<ControlledDocumentStatus>('GOZDEN_GECIRMEDE');
  const [controlledDocumentRevisionNote, setControlledDocumentRevisionNote] = useState('');

  const filesToFileList = (files: File[]) => {
    const transfer = new DataTransfer();
    files.forEach((file) => transfer.items.add(file));
    return transfer.files;
  };

  const [reportFilters, setReportFilters] = useState<{
    reportTypeKey: CorporateReportKey;
    projectId: string;
    department: string;
    periodStart: string;
    periodEnd: string;
    format: CorporateReportFormat;
  }>({
    reportTypeKey: 'corporate-hse',
    projectId: 'all',
    department: 'all',
    periodStart: new Date(new Date().setDate(new Date().getDate() - 30)).toISOString().slice(0, 10),
    periodEnd: new Date().toISOString().slice(0, 10),
    format: 'PDF'
  });
  const [generatedCorporateReports, setGeneratedCorporateReports] = useState<GeneratedCorporateReport[]>(() => (isAuthorizedViewer ? initialGeneratedReports : []));
  const [selectedCorporateReportId, setSelectedCorporateReportId] = useState<string | null>(() => (isAuthorizedViewer ? initialGeneratedReports[0]?.id ?? null : null));

  const [form, setForm] = useState<ModuleRecord>({
    projectId: projectCatalog[0]?.id ?? '',
    date: new Date().toISOString().slice(0, 10),
    title: '',
    valueA: 0,
    valueB: 0,
    status: 'OPEN'
  });

  const [incidentForm, setIncidentForm] = useState<IncidentForm>({
    projectId: projectCatalog[0]?.id ?? '',
    date: new Date().toISOString().slice(0, 10),
    title: '',
    incidentCount: 1,
    lostWorkDays: 0,
    status: 'OPEN'
  });
  const [editingIncidentIndex, setEditingIncidentIndex] = useState<number | null>(null);

  const [emergencyDrillRecords, setEmergencyDrillRecords] = useState<EmergencyDrillRecord[]>(() => (isAuthorizedViewer ? initialEmergencyDrillRecords : []));
  const [selectedEmergencyDrillId, setSelectedEmergencyDrillId] = useState<string | null>(() => (isAuthorizedViewer ? initialEmergencyDrillRecords[0]?.id ?? null : null));
  const [emergencyDrillForm, setEmergencyDrillForm] = useState<EmergencyDrillForm>({
    projectId: projectCatalog[0]?.id ?? '',
    emergencyType: 'TAHLIYE',
    drillName: '',
    drillDate: new Date().toISOString().slice(0, 10),
    participantCount: 10,
    drillResult: 'BASARILI',
    openActions: 0,
    closedActions: 0,
    responsiblePerson: '',
    nextPlannedDrillDate: new Date().toISOString().slice(0, 10),
    status: 'PLANLANDI',
    attachments: [],
    drillReports: [],
    photos: [],
    notes: ''
  });
  const [workforceRecords, setWorkforceRecords] = useState<WorkforceRecord[]>(() => (isAuthorizedViewer ? initialWorkforceRecords : []));
  const [editingWorkforceId, setEditingWorkforceId] = useState<string | null>(null);
  const [workforceForm, setWorkforceForm] = useState<WorkforceForm>({
    projectId: projectCatalog[0]?.id ?? '',
    date: new Date().toISOString().slice(0, 10),
    departmentArea: workforceDepartmentOptions[0],
    contractor: '',
    totalWorkforce: 0,
    newEmployees: 0,
    maleEmployees: 0,
    femaleEmployees: 0,
    dayShiftWorkers: 0,
    nightShiftWorkers: 0,
    overtimeWorkers: 0,
    age18_25: 0,
    age26_35: 0,
    age36_45: 0,
    age46_55: 0,
    age56Plus: 0,
    status: 'OPEN',
    notes: ''
  });
  const [trainingRecords, setTrainingRecords] = useState<TrainingRecord[]>(() => (isAuthorizedViewer ? initialTrainingRecords : []));
  const [trainingForm, setTrainingForm] = useState<TrainingForm>({
    projectId: projectCatalog[0]?.id ?? '',
    trainingType: trainingTypeOptions[0],
    trainingTitle: '',
    trainingCategory: trainingCategoryOptions[0],
    trainingDate: new Date().toISOString().slice(0, 10),
    provider: '',
    department: trainingDepartmentOptions[0],
    position: trainingPositionOptions[0],
    projectEmployeeCount: 0,
    certifiedEmployeeCount: 0,
    certificateRequired: true,
    certificateValidityDate: new Date().toISOString().slice(0, 10),
    totalTrainingCost: 0,
    status: 'PLANLANDI',
    attachments: [],
    participantList: [],
    certificates: [],
    notes: ''
  });
  const [editingTrainingId, setEditingTrainingId] = useState<string | null>(null);

  const [ppeTransactions, setPpeTransactions] = useState<PpeTransactionRecord[]>(() => (isAuthorizedViewer ? loadPpeTransactions() : []));
  const [ppeDashboardProjectFilter, setPpeDashboardProjectFilter] = useState<string>('all');
  const [ppeTransactionForm, setPpeTransactionForm] = useState<PpeTransactionForm>({
    transactionType: 'STOK_GIRISI',
    lifecycle: 'TAMAMLANDI',
    projectId: projectCatalog[0]?.id ?? '',
    warehouse: ppeWarehouseOptions[0],
    date: new Date().toISOString().slice(0, 10),
    category: ppeCategoryOptions[0],
    itemName: ppeItemOptions[0],
    brandModel: '',
    unit: ppeUnitOptions[0],
    quantity: 1,
    minimumStockLevel: 0,
    responsiblePerson: '',
    supplier: '',
    targetPersonDepartment: '',
    unitPrice: 0,
    notes: ''
  });
  const [editingPpeTransactionId, setEditingPpeTransactionId] = useState<string | null>(null);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(PPE_TRANSACTIONS_STORAGE_KEY, JSON.stringify(ppeTransactions));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [isAuthorizedViewer, ppeTransactions]);

  const [equipmentRecords, setEquipmentRecords] = useState<EquipmentRecord[]>(() => (isAuthorizedViewer ? initialEquipmentRecords : []));
  const [selectedEquipmentId, setSelectedEquipmentId] = useState<string | null>(() => (isAuthorizedViewer ? initialEquipmentRecords[0]?.id ?? null : null));
  const [equipmentForm, setEquipmentForm] = useState<EquipmentForm>({
    projectId: projectCatalog[0]?.id ?? '',
    equipmentName: '',
    equipmentType: '',
    brandModel: '',
    serialNumber: '',
    location: '',
    responsiblePerson: '',
    lastInspectionDate: new Date().toISOString().slice(0, 10),
    nextInspectionDate: new Date().toISOString().slice(0, 10),
    inspectionStatus: 'UPCOMING',
    certificateNumber: '',
    certificateExpiryDate: new Date().toISOString().slice(0, 10),
    equipmentStatus: 'ACTIVE',
    riskLevel: 'MEDIUM',
    attachments: [],
    inspectionReports: [],
    equipmentPhotos: [],
    notes: ''
  });

  const [riskRecords, setRiskRecords] = useState<RiskRecord[]>(() => (isAuthorizedViewer ? initialRiskRecords : []));
  const [selectedRiskId, setSelectedRiskId] = useState<string | null>(() => (isAuthorizedViewer ? initialRiskRecords[0]?.id ?? null : null));
  const [riskForm, setRiskForm] = useState<RiskForm>({
    projectId: projectCatalog[0]?.id ?? '',
    departmentActivity: '',
    assessmentDate: new Date().toISOString().slice(0, 10),
    hazard: '',
    potentialConsequence: '',
    existingControls: '',
    recommendedControls: '',
    likelihood: 3,
    severity: 3,
    residualLikelihood: 2,
    residualSeverity: 2,
    responsiblePerson: '',
    targetCompletionDate: new Date().toISOString().slice(0, 10),
    status: 'OPEN',
    attachments: [],
    photos: [],
    notes: ''
  });

  const [inspectionForm, setInspectionForm] = useState<InspectionForm>({
    title: 'Workplace Safety Check',
    businessUnit: '',
    siteArea: '',
    locationType: '',
    inspectionDate: new Date().toISOString().slice(0, 10),
    inspectorName: '',
    department: '',
    positiveObservations: '',
    inspectionNotes: ''
  });
  const [checklistAnswers, setChecklistAnswers] = useState<Record<string, ChecklistAnswer>>({});
  const [skippedChecklistSections, setSkippedChecklistSections] = useState<Record<string, boolean>>({});
  const [manualActionText, setManualActionText] = useState('');
  const [manualActions, setManualActions] = useState<CorrectiveAction[]>([]);
  const [observationRecords, setObservationRecords] = useState<ObservationRecord[]>(() => (isAuthorizedViewer ? loadObservationRecords() : []));
  const [observationForm, setObservationForm] = useState<ObservationForm>(() => createEmptyObservationForm(projectCatalog[0]));
  const [observationDraftAttachments, setObservationDraftAttachments] = useState<ObservationAttachmentRecord[]>([]);
  const [editingObservationId, setEditingObservationId] = useState<string | null>(null);
  const [viewingObservationId, setViewingObservationId] = useState<string | null>(null);
  const [observationDragActive, setObservationDragActive] = useState(false);

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }
    try {
      window.localStorage.setItem(OBSERVATION_RECORDS_STORAGE_KEY, JSON.stringify(observationRecords));
    } catch {
      // Ignore storage write failures and keep in-memory behavior.
    }
  }, [isAuthorizedViewer, observationRecords]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    const syncFromPath = () => {
      setLanguage(detectLanguageFromPath(window.location.pathname));
    };

    syncFromPath();
    window.addEventListener('popstate', syncFromPath);

    return () => {
      window.removeEventListener('popstate', syncFromPath);
    };
  }, []);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    if (accessLevel !== 'authorized') {
      return;
    }

    const rememberedEmail = window.sessionStorage.getItem(AUTH_EMAIL_SESSION_KEY);
    if (!rememberedEmail) {
      return;
    }

    if (allowedEmails.includes(rememberedEmail.toLowerCase())) {
      setLoginEmail(rememberedEmail);
    }
  }, [accessLevel, allowedEmails]);

  const handleAuthorizedLogin = () => {
    const normalizedEmail = loginEmail.trim().toLowerCase();
    if (!normalizedEmail || !loginPassword.trim()) {
      setLoginError(language === 'en' ? 'Email and password are required.' : language === 'ru' ? 'Требуются email и пароль.' : 'E-posta ve parola zorunludur.');
      return;
    }

    if (!allowedEmails.includes(normalizedEmail)) {
      setLoginError(language === 'en' ? 'This email is not authorized.' : language === 'ru' ? 'Этот email не авторизован.' : 'Bu e-posta yetkili değil.');
      return;
    }

    const expectedPassword = allowedCredentials[normalizedEmail] ?? '';
    if (loginPassword !== expectedPassword) {
      setLoginError(language === 'en' ? 'Incorrect password.' : language === 'ru' ? 'Неверный пароль.' : 'Parola hatalı.');
      return;
    }

    if (typeof window !== 'undefined') {
      window.sessionStorage.setItem(AUTH_EMAIL_SESSION_KEY, normalizedEmail);
      window.location.reload();
    }
  };

  const continueAsGuest = () => {
    if (typeof window !== 'undefined') {
      window.sessionStorage.removeItem(AUTH_EMAIL_SESSION_KEY);
    }
    setAccessLevel('guest');
    setLoginError('');
  };

  useEffect(() => {
    if (!observationForm.projectId && projectCatalog.length > 0 && !editingObservationId) {
      setObservationForm(createEmptyObservationForm(projectCatalog[0]));
    }
  }, [editingObservationId, observationForm.projectId, projectCatalog]);

  const [dashboardFilterDraft, setDashboardFilterDraft] = useState<{
    period: DashboardPeriod;
    customStart: string;
    customEnd: string;
    projectId: string;
    comparisonType: ComparisonMode;
    selectedProjects: string[];
    selectedDepartments: string[];
  }>({
    period: 'monthly',
    customStart: '',
    customEnd: '',
    projectId: 'all',
    comparisonType: 'none',
    selectedProjects: [],
    selectedDepartments: []
  });
  const [dashboardFilters, setDashboardFilters] = useState(dashboardFilterDraft);
  const [activePtwTab, setActivePtwTab] = useState<PtwTabKey>('kayit');

  const [ptwForm, setPtwForm] = useState<PtwForm>({
    ptwNo: `PTW-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`,
    organizasyon: '',
    departman: '',
    proje: '',
    lokasyon: '',
    ptwTuru: 'Yüksekte Çalışma',
    duzenlenmeTarihi: new Date().toISOString().slice(0, 10),
    gecerlilikBaslangici: new Date().toISOString().slice(0, 10),
    gecerlilikBitisi: new Date().toISOString().slice(0, 10),
    durum: 'Taslak',
    isiTalepEden: '',
    isiVeren: '',
    isSorumlusu: '',
    sahaSorumlusu: '',
    hseSorumlusu: '',
    yetkiliOnaylayan: '',
    isinAdi: '',
    isinAciklamasi: '',
    calismaAlani: '',
    yapilacakIs: '',
    calismaKosullari: '',
    baslangicTarihi: new Date().toISOString().slice(0, 10),
    baslangicSaati: '08:00',
    bitisTarihi: new Date().toISOString().slice(0, 10),
    bitisSaati: '17:00',
    ozelSartlar: '',
    ptwHazirlayan: '',
    hseOnayi: '',
    projeMuduru: '',
    isverenTemsilcisi: '',
    dijitalImza: '',
    onayTarihi: new Date().toISOString().slice(0, 10),
    isTamamlandi: false,
    alanGuvenli: false,
    malzemelerToplandi: false,
    ptwKapatildi: false,
    kapatan: '',
    kapanisTarihi: new Date().toISOString().slice(0, 10),
    kapanisAciklama: ''
  });

  const [ptwTehlikeler, setPtwTehlikeler] = useState<string[]>([]);
  const [ptwGuvenlikSecimleri, setPtwGuvenlikSecimleri] = useState<string[]>([]);
  const [ptwEkipmanSecimleri, setPtwEkipmanSecimleri] = useState<string[]>([]);

  const [ptwEkipListesi, setPtwEkipListesi] = useState<PtwTeamMember[]>([
    { adSoyad: '', gorevi: '', firma: '', egitimDurumu: '', imza: '' }
  ]);
  const [ptwKontroller, setPtwKontroller] = useState<PtwControlItem[]>([
    { kontrol: 'Çalışma alanı bariyerleme kontrolü', durum: 'Uygun', aciklama: '' },
    { kontrol: 'Düşüş önleme sistemi kontrolü', durum: 'Uygun', aciklama: '' },
    { kontrol: 'Ekipman sertifika kontrolü', durum: 'N/A', aciklama: '' }
  ]);
  const [ptwTedbirler, setPtwTedbirler] = useState<PtwPrecautionItem[]>([
    { tedbir: '', sorumlu: '', termin: new Date().toISOString().slice(0, 10), durum: 'Açık' }
  ]);
  const [ptwGunlukKayitlar, setPtwGunlukKayitlar] = useState<PtwDailyLog[]>([
    { tarih: new Date().toISOString().slice(0, 10), saat: '08:00', calismaBasladi: '', calismaBitti: '', aciklama: '', sorumlu: '' }
  ]);
  const [ptwEkipDegisiklikleri, setPtwEkipDegisiklikleri] = useState<PtwTeamChange[]>([
    { eklenenPersonel: '', ayrilanPersonel: '', tarih: new Date().toISOString().slice(0, 10), onaylayan: '' }
  ]);
  const [ptwEkler, setPtwEkler] = useState<PtwAttachment[]>([]);
  const [ptwRecordId, setPtwRecordId] = useState<string | null>(null);
  const [ptwSaving, setPtwSaving] = useState(false);
  const [ptwFeedback, setPtwFeedback] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [ptwLastSavedAt, setPtwLastSavedAt] = useState<string | null>(null);
  const [ptwEmailTo, setPtwEmailTo] = useState('');
  const [liveWeather, setLiveWeather] = useState<{
    icon: string;
    temperature: string;
    condition: string;
    location: string;
    updatedAt: string;
    windSpeed: string;
    windDirection: string;
  }>({
    icon: '☀',
    temperature: '--°C',
    condition: language === 'en' ? 'Loading...' : language === 'ru' ? 'Загрузка...' : 'Yükleniyor...',
    location: language === 'en' ? 'All Projects' : language === 'ru' ? 'Все проекты' : 'Tüm Projeler',
    updatedAt: '--:--',
    windSpeed: '-- km/s',
    windDirection: '--'
  });

  const scopedRows = useMemo(() => {
    const rows = moduleData[activeModule] ?? [];
    if (projectFilter === 'all') {
      return rows;
    }
    return rows.filter((row) => row.projectId === projectFilter);
  }, [activeModule, moduleData, projectFilter]);

  const scopedModuleRecords = useMemo(() => {
    const rows = moduleData[activeModule] ?? [];
    return rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => projectFilter === 'all' || row.projectId === projectFilter);
  }, [activeModule, moduleData, projectFilter]);

  const selectedObservationRecord = useMemo(
    () => observationRecords.find((record) => record.id === viewingObservationId) ?? null,
    [observationRecords, viewingObservationId]
  );

  const nextObservationNo = useMemo(() => {
    const year = new Date().getFullYear();
    const maxSequence = observationRecords.reduce((max, record) => {
      const match = record.observationNo.match(/^(?:OBS|GOZ)-\d{4}-(\d{4})$/i);
      if (!match) {
        return max;
      }
      const sequence = Number(match[1]);
      return Number.isFinite(sequence) ? Math.max(max, sequence) : max;
    }, 0);
    return `OBS-${year}-${String(maxSequence + 1).padStart(4, '0')}`;
  }, [observationRecords]);

  const legalScopedRows = useMemo(() => {
    return legalRecords.filter((record) => {
      if (legalFilters.projectId !== 'all' && record.projectId !== legalFilters.projectId) {
        return false;
      }
      if (legalFilters.category !== 'all' && record.category !== legalFilters.category) {
        return false;
      }
      if (legalFilters.authority !== 'all' && record.authority !== legalFilters.authority) {
        return false;
      }
      if (legalFilters.complianceStatus !== 'ALL' && record.complianceStatus !== legalFilters.complianceStatus) {
        return false;
      }
      if (legalFilters.keyword.trim()) {
        const keyword = legalFilters.keyword.trim().toLocaleLowerCase('tr-TR');
        const text = `${record.title} ${record.legalRequirement} ${record.responsiblePerson} ${record.regulationId}`.toLocaleLowerCase('tr-TR');
        if (!text.includes(keyword)) {
          return false;
        }
      }
      return true;
    });
  }, [legalFilters, legalRecords]);

  const nextLegalRegulationId = useMemo(() => {
    const year = new Date().getFullYear();
    return `LG-${year}-${String(legalRecords.length + 1).padStart(4, '0')}`;
  }, [legalRecords.length]);

  const selectedLegalRecord = useMemo(
    () => legalRecords.find((record) => record.id === selectedLegalRecordId) ?? null,
    [legalRecords, selectedLegalRecordId]
  );

  const selectedLegalDocument = useMemo(
    () => selectedLegalRecord?.documents.find((documentRow) => documentRow.id === selectedLegalDocumentId) ?? null,
    [selectedLegalDocumentId, selectedLegalRecord]
  );

  const legalSummary = useMemo(() => {
    const total = legalScopedRows.length;
    const compliant = legalScopedRows.filter((record) => record.complianceStatus === 'UYUMLU').length;
    const nonCompliant = legalScopedRows.filter((record) => record.complianceStatus === 'UYUMSUZ').length;
    const inReview = legalScopedRows.filter((record) => record.complianceStatus === 'KISMEN_UYUMLU').length;
    const upcomingReviews = legalScopedRows.filter((record) => {
      const days = daysUntil(record.nextReviewDate);
      return days >= 0 && days <= 30;
    }).length;
    const openActions = legalScopedRows.reduce((sum, record) => sum + record.openActions, 0);
    return { total, compliant, nonCompliant, inReview, upcomingReviews, openActions };
  }, [legalScopedRows]);

  const legalAlerts = useMemo(() => {
    const upcoming = legalScopedRows.filter((record) => {
      const days = daysUntil(record.nextReviewDate);
      return days >= 0 && days <= 30;
    });
    const expired = legalScopedRows.filter((record) => daysUntil(record.nextReviewDate) < 0);
    return { upcoming, expired };
  }, [legalScopedRows]);

  const legalComplianceChart = useMemo(
    () => [
      legalScopedRows.filter((record) => record.complianceStatus === 'UYUMLU').length,
      legalScopedRows.filter((record) => record.complianceStatus === 'KISMEN_UYUMLU').length,
      legalScopedRows.filter((record) => record.complianceStatus === 'UYUMSUZ').length
    ],
    [legalScopedRows]
  );

  const legalCategoryBars = useMemo(() => {
    const map = new Map<string, number>();
    legalCategoryOptions.forEach((category) => map.set(category, 0));
    legalScopedRows.forEach((record) => {
      map.set(record.category, (map.get(record.category) ?? 0) + 1);
    });
    return Array.from(map.entries());
  }, [legalScopedRows]);

  const controlledDocumentFilteredRows = useMemo(() => {
    return controlledDocuments.filter((record) => {
      if (controlledDocumentFilters.projectId !== 'all' && record.projectId !== controlledDocumentFilters.projectId) {
        return false;
      }
      if (controlledDocumentFilters.category !== 'all' && record.category !== controlledDocumentFilters.category) {
        return false;
      }
      if (controlledDocumentFilters.status !== 'ALL' && record.status !== controlledDocumentFilters.status) {
        return false;
      }
      if (controlledDocumentFilters.department !== 'all' && record.department !== controlledDocumentFilters.department) {
        return false;
      }
      if (controlledDocumentFilters.revision !== 'all' && String(record.revisionNumber) !== controlledDocumentFilters.revision) {
        return false;
      }
      if (controlledDocumentFilters.keyword.trim()) {
        const keyword = controlledDocumentFilters.keyword.trim().toLocaleLowerCase('tr-TR');
        const searchable = `${record.documentId} ${record.title} ${record.category} ${record.department}`.toLocaleLowerCase('tr-TR');
        if (!searchable.includes(keyword)) {
          return false;
        }
      }
      return true;
    });
  }, [controlledDocumentFilters, controlledDocuments]);

  const selectedControlledDocument = useMemo(
    () => controlledDocuments.find((record) => record.id === selectedControlledDocumentId) ?? null,
    [controlledDocuments, selectedControlledDocumentId]
  );

  const selectedControlledDocumentRevision = useMemo(
    () => (selectedControlledDocument ? controlledDocumentCurrentRevision(selectedControlledDocument) : null),
    [selectedControlledDocument]
  );

  const controlledDocumentRevisionOptions = useMemo(() => {
    const revisions = Array.from(new Set(controlledDocuments.map((record) => String(record.revisionNumber))));
    return ['all', ...revisions.sort((left, right) => Number(left) - Number(right))];
  }, [controlledDocuments]);

  const nextControlledDocumentId = useMemo(() => {
    const year = new Date().getFullYear();
    return `DOC-${year}-${String(controlledDocuments.length + 1).padStart(4, '0')}`;
  }, [controlledDocuments.length]);

  const selectedGeneratedCorporateReport = useMemo(
    () => generatedCorporateReports.find((report) => report.id === selectedCorporateReportId) ?? null,
    [generatedCorporateReports, selectedCorporateReportId]
  );

  const legalReviewTrend = useMemo(() => {
    const now = new Date();
    const monthKeys = Array.from({ length: 6 }, (_, index) => {
      const date = new Date(now.getFullYear(), now.getMonth() + index, 1);
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    });
    const values = monthKeys.map((key) =>
      legalScopedRows.filter((record) => record.nextReviewDate.slice(0, 7) === key).length
    );
    const labels = monthKeys.map((key) => {
      const [year, month] = key.split('-');
      return `${month}/${year.slice(2)}`;
    });
    return { values, labels };
  }, [legalScopedRows]);

  const legalOpenActionsByProject = useMemo(() => {
    const labels = projectCatalog.map((project) => project.name);
    const values = projectCatalog.map((project) =>
      legalScopedRows
        .filter((record) => record.projectId === project.id)
        .reduce((sum, record) => sum + record.openActions, 0)
    );
    return { labels, values };
  }, [legalScopedRows, projectCatalog]);

  const legalCanDeleteDocument = useMemo(
    () => legalUserRole === 'COMPLIANCE_MANAGER' || legalUserRole === 'CORPORATE_HSE_MANAGER',
    [legalUserRole]
  );

  const nextRiskId = useMemo(() => {
    const year = new Date().getFullYear();
    const nextIndex = riskRecords.length + 1;
    return `RK-${year}-${String(nextIndex).padStart(4, '0')}`;
  }, [riskRecords.length]);

  const scopedRiskRows = useMemo(() => {
    if (projectFilter === 'all') {
      return riskRecords;
    }
    return riskRecords.filter((row) => row.projectId === projectFilter);
  }, [projectFilter, riskRecords]);

  const equipmentScopedRows = useMemo(() => {
    if (projectFilter === 'all') {
      return equipmentRecords;
    }
    return equipmentRecords.filter((row) => row.projectId === projectFilter);
  }, [equipmentRecords, projectFilter]);

  const emergencyScopedRows = useMemo(() => {
    if (projectFilter === 'all') {
      return emergencyDrillRecords;
    }
    return emergencyDrillRecords.filter((row) => row.projectId === projectFilter);
  }, [emergencyDrillRecords, projectFilter]);

  const workforceScopedRows = useMemo(() => {
    if (projectFilter === 'all') {
      return workforceRecords;
    }
    return workforceRecords.filter((row) => row.projectId === projectFilter);
  }, [projectFilter, workforceRecords]);

  const trainingScopedRows = useMemo(() => {
    if (projectFilter === 'all') {
      return trainingRecords;
    }
    return trainingRecords.filter((row) => row.projectId === projectFilter);
  }, [projectFilter, trainingRecords]);

  const ppeScopedTransactions = useMemo(() => {
    if (projectFilter === 'all') {
      return ppeTransactions;
    }
    return ppeTransactions.filter((row) => row.projectId === projectFilter);
  }, [ppeTransactions, projectFilter]);

  const ppeDashboardTransactions = useMemo(() => {
    if (ppeDashboardProjectFilter === 'all') {
      return ppeTransactions;
    }
    return ppeTransactions.filter((row) => row.projectId === ppeDashboardProjectFilter);
  }, [ppeDashboardProjectFilter, ppeTransactions]);

  const nextEquipmentId = useMemo(() => {
    const year = new Date().getFullYear();
    return `EQ-${year}-${String(equipmentRecords.length + 1).padStart(4, '0')}`;
  }, [equipmentRecords.length]);

  const nextEmergencyDrillId = useMemo(() => {
    const year = new Date().getFullYear();
    return `ADT-${year}-${String(emergencyDrillRecords.length + 1).padStart(4, '0')}`;
  }, [emergencyDrillRecords.length]);

  const nextPpeTransactionId = useMemo(() => {
    const year = new Date().getFullYear();
    return `KKD-TRX-${year}-${String(ppeTransactions.length + 1).padStart(4, '0')}`;
  }, [ppeTransactions.length]);

  const incidentScopedRows = useMemo(() => {
    const rows = moduleData.incidents;
    if (projectFilter === 'all') {
      return rows;
    }
    return rows.filter((row) => row.projectId === projectFilter);
  }, [moduleData.incidents, projectFilter]);

  const incidentScopedRecords = useMemo(() => {
    return moduleData.incidents
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => projectFilter === 'all' || row.projectId === projectFilter);
  }, [moduleData.incidents, projectFilter]);

  const selectedRiskRecord = useMemo(
    () => riskRecords.find((row) => row.id === selectedRiskId) ?? null,
    [riskRecords, selectedRiskId]
  );

  const selectedEquipmentRecord = useMemo(
    () => equipmentRecords.find((row) => row.id === selectedEquipmentId) ?? null,
    [equipmentRecords, selectedEquipmentId]
  );

  const selectedEmergencyDrillRecord = useMemo(
    () => emergencyDrillRecords.find((row) => row.id === selectedEmergencyDrillId) ?? null,
    [emergencyDrillRecords, selectedEmergencyDrillId]
  );

  const getRiskBand = (score: number) => {
    if (score >= 15) {
      return { label: 'Yüksek', className: 'risk-band-high', range: '15-25' };
    }
    if (score >= 6) {
      return { label: 'Orta', className: 'risk-band-medium', range: '6-12' };
    }
    return { label: 'Düşük', className: 'risk-band-low', range: '1-5' };
  };

  const getStatusBadgeClass = (status: Status) => {
    if (status === 'CLOSED') {
      return 'status-badge-closed';
    }
    if (status === 'IN_PROGRESS') {
      return 'status-badge-progress';
    }
    return 'status-badge-open';
  };

  const incidentSummary = useMemo(() => {
    const rows = incidentScopedRows;
    const totalIncidents = rows.reduce((sum, row) => sum + row.valueA, 0);
    const totalLostDays = rows.reduce((sum, row) => sum + row.valueB, 0);
    const openCases = rows.filter((row) => row.status === 'OPEN').length;
    const inProgressCases = rows.filter((row) => row.status === 'IN_PROGRESS').length;
    const closedCases = rows.filter((row) => row.status === 'CLOSED').length;
    const severeCases = rows.filter((row) => row.valueB >= 3).length;

    return {
      totalIncidents,
      totalLostDays,
      openCases,
      inProgressCases,
      closedCases,
      severeCases
    };
  }, [incidentScopedRows]);

  const equipmentSummary = useMemo(() => {
    const rows = equipmentScopedRows;
    const total = rows.length;
    const active = rows.filter((row) => row.equipmentStatus === 'ACTIVE').length;
    const upcoming = rows.filter((row) => row.inspectionStatus === 'UPCOMING').length;
    const overdue = rows.filter((row) => row.inspectionStatus === 'OVERDUE').length;
    const expiringCertificates = rows.filter((row) => daysUntil(row.certificateExpiryDate) >= 0 && daysUntil(row.certificateExpiryDate) <= 30).length;
    const outOfService = rows.filter((row) => row.equipmentStatus === 'OUT_OF_SERVICE').length;
    const compliant = rows.filter((row) => row.inspectionStatus === 'COMPLIANT').length;
    const complianceRate = total === 0 ? 0 : Math.round((compliant / total) * 100);

    return {
      total,
      active,
      upcoming,
      overdue,
      expiringCertificates,
      outOfService,
      compliant,
      complianceRate
    };
  }, [equipmentScopedRows]);

  const emergencySummary = useMemo(() => {
    const totalDrills = emergencyScopedRows.length;
    const totalParticipants = emergencyScopedRows.reduce((sum, row) => sum + row.participantCount, 0);
    const openActions = emergencyScopedRows.reduce((sum, row) => sum + row.openActions, 0);
    const closedActions = emergencyScopedRows.reduce((sum, row) => sum + row.closedActions, 0);
    const upcomingDrills = emergencyScopedRows.filter((row) => daysUntil(row.nextPlannedDrillDate) >= 0 && daysUntil(row.nextPlannedDrillDate) <= 30).length;
    const completedDrills = emergencyScopedRows.filter((row) => row.status === 'TAMAMLANDI').length;

    return {
      totalDrills,
      totalParticipants,
      openActions,
      closedActions,
      upcomingDrills,
      completedDrills
    };
  }, [emergencyScopedRows]);

  const workforceSummary = useMemo(() => {
    return {
      totalWorkforce: workforceScopedRows.reduce((sum, row) => sum + row.totalWorkforce, 0),
      newEmployees: workforceScopedRows.reduce((sum, row) => sum + row.newEmployees, 0),
      maleEmployees: workforceScopedRows.reduce((sum, row) => sum + row.maleEmployees, 0),
      femaleEmployees: workforceScopedRows.reduce((sum, row) => sum + row.femaleEmployees, 0),
      dayShiftWorkers: workforceScopedRows.reduce((sum, row) => sum + row.dayShiftWorkers, 0),
      nightShiftWorkers: workforceScopedRows.reduce((sum, row) => sum + row.nightShiftWorkers, 0),
      overtimeWorkers: workforceScopedRows.reduce((sum, row) => sum + row.overtimeWorkers, 0),
      age18_25: workforceScopedRows.reduce((sum, row) => sum + row.age18_25, 0),
      age26_35: workforceScopedRows.reduce((sum, row) => sum + row.age26_35, 0),
      age36_45: workforceScopedRows.reduce((sum, row) => sum + row.age36_45, 0),
      age46_55: workforceScopedRows.reduce((sum, row) => sum + row.age46_55, 0),
      age56Plus: workforceScopedRows.reduce((sum, row) => sum + row.age56Plus, 0)
    };
  }, [workforceScopedRows]);

  const workforceDepartmentDistribution = useMemo(() => {
    const totals = workforceDepartmentOptions.reduce<Record<string, number>>((accumulator, department) => {
      accumulator[department] = 0;
      return accumulator;
    }, {});

    workforceScopedRows.forEach((row) => {
      if (!(row.departmentArea in totals)) {
        totals[row.departmentArea] = 0;
      }
      totals[row.departmentArea] += row.totalWorkforce;
    });

    return Object.entries(totals).map(([label, value]) => ({ label, value }));
  }, [workforceScopedRows]);

  const workforceDepartmentMax = useMemo(
    () => Math.max(...workforceDepartmentDistribution.map((item) => item.value), 1),
    [workforceDepartmentDistribution]
  );

  const workforceRates = useMemo(() => {
    const total = Math.max(workforceSummary.totalWorkforce, 1);
    return {
      womenRatio: Math.round((workforceSummary.femaleEmployees / total) * 100),
      newEmployeeRatio: Math.round((workforceSummary.newEmployees / total) * 100),
      overtimeRatio: Math.round((workforceSummary.overtimeWorkers / total) * 100)
    };
  }, [workforceSummary]);

  const nextTrainingId = useMemo(() => {
    const year = new Date().getFullYear();
    return `EGT-${year}-${String(trainingRecords.length + 1).padStart(4, '0')}`;
  }, [trainingRecords.length]);

  const trainingSummary = useMemo(() => {
    const totalTrainings = trainingScopedRows.length;
    const totalParticipants = trainingScopedRows.reduce((sum, row) => sum + row.projectEmployeeCount, 0);
    const certifiedEmployees = trainingScopedRows.reduce((sum, row) => sum + row.certifiedEmployeeCount, 0);
    const pendingEmployees = trainingScopedRows.reduce(
      (sum, row) => sum + Math.max(row.projectEmployeeCount - row.certifiedEmployeeCount, 0),
      0
    );
    const totalCost = trainingScopedRows.reduce((sum, row) => sum + row.totalTrainingCost, 0);
    const avgCostPerEmployee = totalParticipants === 0 ? 0 : Math.round(totalCost / totalParticipants);

    return {
      totalTrainings,
      totalParticipants,
      certifiedEmployees,
      pendingEmployees,
      totalCost,
      avgCostPerEmployee
    };
  }, [trainingScopedRows]);

  const trainingTypeDistribution = useMemo(() => {
    const totals = trainingTypeOptions.reduce<Record<string, number>>((accumulator, label) => {
      accumulator[label] = 0;
      return accumulator;
    }, {});

    trainingScopedRows.forEach((row) => {
      if (!(row.trainingType in totals)) {
        totals[row.trainingType] = 0;
      }
      totals[row.trainingType] += 1;
    });

    return Object.entries(totals).filter((entry) => entry[1] > 0);
  }, [trainingScopedRows]);

  const trainingCostByProject = useMemo(
    () =>
      projectCatalog.map((project) => ({
        label: project.name,
        value: trainingScopedRows
          .filter((row) => row.projectId === project.id)
          .reduce((sum, row) => sum + row.totalTrainingCost, 0)
      })),
    [projectCatalog, trainingScopedRows]
  );

  const trainingProjectWorkforce = useMemo(
    () =>
      projectCatalog.map((project) => ({
        label: project.name,
        workforce: trainingScopedRows
          .filter((row) => row.projectId === project.id)
          .reduce((sum, row) => sum + row.projectEmployeeCount, 0),
        certified: trainingScopedRows
          .filter((row) => row.projectId === project.id)
          .reduce((sum, row) => sum + row.certifiedEmployeeCount, 0)
      })),
    [projectCatalog, trainingScopedRows]
  );

  const trainingProjectWorkforceMax = useMemo(
    () => Math.max(...trainingProjectWorkforce.map((row) => row.workforce), 1),
    [trainingProjectWorkforce]
  );

  const trainingStatusDistribution = useMemo(() => {
    return {
      TAMAMLANDI: trainingScopedRows.filter((row) => row.status === 'TAMAMLANDI').length,
      PLANLANDI: trainingScopedRows.filter((row) => row.status === 'PLANLANDI').length,
      DEVAM_EDIYOR: trainingScopedRows.filter((row) => row.status === 'DEVAM_EDIYOR').length,
      SURESI_DOLDU: trainingScopedRows.filter((row) => row.status === 'SURESI_DOLDU').length
    };
  }, [trainingScopedRows]);

  const ppeInventoryByKey = useMemo(() => {
    const map = new Map<string, {
      projectId: string;
      category: string;
      itemName: string;
      unit: string;
      warehouse: string;
      minimumStockLevel: number;
      incoming: number;
      outgoing: number;
      returned: number;
      damaged: number;
      adjusted: number;
      stock: number;
      averageUnitPrice: number;
      value: number;
      lastDate: string;
    }>();

    ppeDashboardTransactions.forEach((record) => {
      if (record.lifecycle !== 'TAMAMLANDI') {
        return;
      }

      const key = `${record.projectId}__${record.warehouse}__${record.category}__${record.itemName}__${record.unit}`;
      const current = map.get(key) ?? {
        projectId: record.projectId,
        category: record.category,
        itemName: record.itemName,
        unit: record.unit,
        warehouse: record.warehouse,
        minimumStockLevel: record.minimumStockLevel,
        incoming: 0,
        outgoing: 0,
        returned: 0,
        damaged: 0,
        adjusted: 0,
        stock: 0,
        averageUnitPrice: 0,
        value: 0,
        lastDate: record.date
      };

      if (record.minimumStockLevel > 0) {
        current.minimumStockLevel = record.minimumStockLevel;
      }

      if (record.transactionType === 'STOK_GIRISI') {
        current.incoming += record.quantity;
        if (record.unitPrice > 0) {
          const prevIncoming = current.incoming - record.quantity;
          const prevValue = current.averageUnitPrice * Math.max(prevIncoming, 0);
          current.averageUnitPrice = (prevValue + record.quantity * record.unitPrice) / Math.max(current.incoming, 1);
        }
      }
      if (record.transactionType === 'STOK_CIKISI') {
        current.outgoing += record.quantity;
      }
      if (record.transactionType === 'STOGA_IADE') {
        current.returned += record.quantity;
      }
      if (record.transactionType === 'HASARLI_HURDA') {
        current.damaged += record.quantity;
      }
      if (record.transactionType === 'STOK_DUZELTME') {
        current.adjusted += record.quantity;
      }

      const movement =
        record.transactionType === 'STOK_GIRISI' || record.transactionType === 'STOGA_IADE' || record.transactionType === 'STOK_DUZELTME'
          ? record.quantity
          : -record.quantity;
      current.stock += movement;
      current.lastDate = record.date > current.lastDate ? record.date : current.lastDate;

      map.set(key, current);
    });

    const rows = Array.from(map.values());
    rows.forEach((row) => {
      row.stock = Math.max(0, row.stock);
      row.value = Math.round(row.stock * row.averageUnitPrice);
    });
    return rows;
  }, [ppeDashboardTransactions]);

  const ppeMonthlySummary = useMemo(() => {
    const now = new Date();
    const month = now.getMonth();
    const year = now.getFullYear();

    const rows = ppeDashboardTransactions.filter((row) => {
      if (row.lifecycle !== 'TAMAMLANDI') {
        return false;
      }
      const date = new Date(row.date);
      return date.getFullYear() === year && date.getMonth() === month;
    });

    return {
      monthlyIncoming: rows
        .filter((row) => row.transactionType === 'STOK_GIRISI' || row.transactionType === 'STOGA_IADE')
        .reduce((sum, row) => sum + row.quantity, 0),
      monthlyOutgoing: rows
        .filter((row) => row.transactionType === 'STOK_CIKISI' || row.transactionType === 'HASARLI_HURDA')
        .reduce((sum, row) => sum + row.quantity, 0)
    };
  }, [ppeDashboardTransactions]);

  const ppeSummary = useMemo(() => {
    const totalStock = ppeInventoryByKey.reduce((sum, row) => sum + row.stock, 0);
    const lowStockItems = ppeInventoryByKey.filter((row) => row.stock > 0 && row.stock <= row.minimumStockLevel).length;
    const outOfStockItems = ppeInventoryByKey.filter((row) => row.stock <= 0).length;
    const totalInventoryValue = ppeInventoryByKey.reduce((sum, row) => sum + row.value, 0);
    const purchaseOrdersNeeded = ppeInventoryByKey.filter((row) => row.stock <= row.minimumStockLevel).length;

    return {
      totalStock,
      monthlyIncoming: ppeMonthlySummary.monthlyIncoming,
      monthlyOutgoing: ppeMonthlySummary.monthlyOutgoing,
      lowStockItems,
      outOfStockItems,
      totalInventoryValue,
      purchaseOrdersNeeded
    };
  }, [ppeInventoryByKey, ppeMonthlySummary]);

  const ppeCategoryStockData = useMemo(() => {
    const map = new Map<string, number>();
    ppeInventoryByKey.forEach((row) => {
      map.set(row.category, (map.get(row.category) ?? 0) + row.stock);
    });
    return Array.from(map.entries()).map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
  }, [ppeInventoryByKey]);

  const ppeStatusDistribution = useMemo(() => {
    return ppeInventoryByKey.reduce(
      (accumulator, row) => {
        if (row.stock <= 0) {
          accumulator.STOKTA_YOK += 1;
        } else if (row.stock <= row.minimumStockLevel) {
          accumulator.DUSUK_STOK += 1;
        } else {
          accumulator.YETERLI += 1;
        }
        return accumulator;
      },
      { YETERLI: 0, DUSUK_STOK: 0, STOKTA_YOK: 0 }
    );
  }, [ppeInventoryByKey]);

  const ppeMonthlyTrend = useMemo(() => {
    const now = new Date();
    const buckets = Array.from({ length: 6 }, (_, index) => {
      const d = new Date(now.getFullYear(), now.getMonth() - (5 - index), 1);
      return {
        key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
        label: d.toLocaleDateString('tr-TR', { month: 'short' }),
        incoming: 0,
        outgoing: 0
      };
    });

    ppeDashboardTransactions.forEach((row) => {
      if (row.lifecycle !== 'TAMAMLANDI') {
        return;
      }
      const date = new Date(row.date);
      const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      const bucket = buckets.find((entry) => entry.key === key);
      if (!bucket) {
        return;
      }
      if (row.transactionType === 'STOK_GIRISI' || row.transactionType === 'STOGA_IADE') {
        bucket.incoming += row.quantity;
      }
      if (row.transactionType === 'STOK_CIKISI' || row.transactionType === 'HASARLI_HURDA') {
        bucket.outgoing += row.quantity;
      }
    });

    return buckets;
  }, [ppeDashboardTransactions]);

  const ppeMonthlyTrendMax = useMemo(
    () => Math.max(...ppeMonthlyTrend.map((row) => Math.max(row.incoming, row.outgoing)), 1),
    [ppeMonthlyTrend]
  );

  const ppeLastTransactions = useMemo(() => {
    return [...ppeScopedTransactions]
      .sort((a, b) => `${b.date}-${b.transactionId}`.localeCompare(`${a.date}-${a.transactionId}`))
      .slice(0, 8);
  }, [ppeScopedTransactions]);

  const projectFilterOptions = useMemo(
    () => [{ id: 'all', name: t.allProjects, country: '', city: '', address: '' }, ...projectCatalog],
    [projectCatalog, t.allProjects]
  );

  const departmentOptions = useMemo(() => ['Mechanical', 'Civil', 'Electrical', 'Operations', 'Admin'], []);

  const inspectionSummary = useMemo(() => {
    let answeredProgress = 0;
    let yesCount = 0;
    let noCount = 0;
    let naCount = 0;

    const highRiskSections = new Set(['fire-safety', 'electrical-safety', 'working-at-height', 'chemicals', 'vehicle-pedestrian']);
    const mediumRiskSections = new Set(['housekeeping', 'walking-surfaces', 'emergency-readiness', 'plant-tools-equipment', 'contractor-maintenance']);
    let highFindings = 0;
    let mediumFindings = 0;
    let lowFindings = 0;

    inspectionSections.forEach((section) => {
      if (skippedChecklistSections[section.id]) {
        return;
      }
      section.questions.forEach((_, questionIndex) => {
        const key = `${section.id}-${questionIndex}`;
        const answer = checklistAnswers[key] ?? '';
        if (answer !== '') {
          answeredProgress += 1;
        }
        if (answer === 'YES') {
          yesCount += 1;
        }
        if (answer === 'NO') {
          noCount += 1;
          if (highRiskSections.has(section.id)) {
            highFindings += 1;
          } else if (mediumRiskSections.has(section.id)) {
            mediumFindings += 1;
          } else {
            lowFindings += 1;
          }
        }
        if (answer === 'NA') {
          naCount += 1;
        }
      });
    });

    const denominator = yesCount + noCount;
    const complianceScore = denominator === 0 ? 0 : Math.round((yesCount / denominator) * 100);
    const urgentLabel =
      complianceScore <= 50
        ? 'Acil iyileştirme gerekli (0-50%)'
        : complianceScore <= 75
          ? 'İyileştirme gerekli (51-75%)'
          : 'Kontrol altında (76-100%)';
    const openActions = manualActions.filter((action) => action.status === 'OPEN').length;
    const inProgressActions = manualActions.filter((action) => action.status === 'IN_PROGRESS').length;
    const closedActions = manualActions.filter((action) => action.status === 'CLOSED').length;

    return {
      urgentLabel,
      complianceScore,
      answeredProgress,
      totalFindings: noCount,
      openActions,
      inProgressActions,
      closedActions,
      highFindings,
      mediumFindings,
      lowFindings,
      naCount
    };
  }, [checklistAnswers, inspectionSections, manualActions, skippedChecklistSections]);

  const inspectionUrgencyLabel =
    language === 'ru'
      ? inspectionSummary.complianceScore <= 50
        ? 'Требуется срочное улучшение (0-50%)'
        : inspectionSummary.complianceScore <= 75
          ? 'Требуется улучшение (51-75%)'
          : 'Под контролем (76-100%)'
      : inspectionSummary.urgentLabel;

  const chartValues = useMemo(() => {
    if (activeModule === 'dashboard') {
      const allRows = Object.values(moduleData).flat();
      const byProject = projectCatalog
        .map((project) => allRows.filter((row) => row.projectId === project.id).length);
      return byProject.length > 0 ? byProject : [0, 0, 0];
    }

    if (activeModule === 'occupational-health') {
      const fit = healthRecords.filter((record) => record.fitForWork === 'Yes').length;
      const restricted = healthRecords.filter((record) => record.restrictedWork.trim().toLowerCase() !== 'no' && record.restrictedWork.trim() !== '').length;
      const chronic = healthRecords.filter((record) => record.chronicDisease.trim().toLowerCase() !== 'none' && record.chronicDisease.trim() !== '').length;
      const communicable = healthRecords.filter((record) => record.communicableDisease.trim().toLowerCase() !== 'none' && record.communicableDisease.trim() !== '').length;
      const expiring = healthRecords.filter((record) => daysUntil(record.nextMedicalExamination) <= 30).length;
      return [fit, restricted, chronic, communicable, expiring, healthRecords.length];
    }

    if (activeModule === 'incidents') {
      return [
        incidentSummary.totalIncidents,
        incidentSummary.openCases,
        incidentSummary.inProgressCases,
        incidentSummary.closedCases,
        incidentSummary.totalLostDays,
        incidentSummary.severeCases
      ];
    }

    if (activeModule === 'equipment-management') {
      return [
        equipmentSummary.compliant,
        equipmentSummary.upcoming,
        equipmentSummary.overdue,
        equipmentSummary.expiringCertificates,
        equipmentSummary.outOfService,
        equipmentSummary.total
      ];
    }

    if (activeModule === 'emergency-preparedness') {
      return [
        emergencySummary.completedDrills,
        emergencySummary.upcomingDrills,
        emergencySummary.openActions,
        emergencySummary.closedActions,
        emergencySummary.totalParticipants,
        emergencySummary.totalDrills
      ];
    }

    const values = scopedRows.slice(0, 6).map((row) => row.valueA);
    if (values.length > 0) {
      return values;
    }
    return [0, 0, 0, 0, 0, 0];
  }, [activeModule, emergencySummary, equipmentSummary, healthRecords, incidentSummary, moduleData, projectCatalog, scopedRows]);

  const healthAnalytics = useMemo(() => {
    const totalEmployees = healthRecords.length;
    const fitForWork = healthRecords.filter((record) => record.fitForWork === 'Yes').length;
    const notFitForWork = Math.max(totalEmployees - fitForWork, 0);
    const restrictedDuty = healthRecords.filter((record) => record.restrictedWork.trim().toLowerCase() !== 'no' && record.restrictedWork.trim() !== '').length;
    const chronicCount = healthRecords.filter((record) => record.chronicDisease.trim().toLowerCase() !== 'none' && record.chronicDisease.trim() !== '').length;
    const communicableCount = healthRecords.filter((record) => record.communicableDisease.trim().toLowerCase() !== 'none' && record.communicableDisease.trim() !== '').length;
    const medicalDue = healthRecords.filter((record) => daysUntil(record.nextMedicalExamination) <= 30).length;
    const vaccinationUpToDate = healthRecords.filter((record) => record.vaccinationStatus.toLowerCase().includes('up')).length;
    const vaccinationRate = totalEmployees > 0 ? Math.round((vaccinationUpToDate / totalEmployees) * 100) : 0;

    const metricLabels = [
      'İşe Uygun',
      'Kısıtlı Görev',
      'Kronik Hastalık',
      'Bulaşıcı Hastalık',
      '30 Gün İçinde Muayene',
      'Toplam Kayıt'
    ];
    const metricValues = [fitForWork, restrictedDuty, chronicCount, communicableCount, medicalDue, totalEmployees];

    const departmentMap = new Map<string, number>();
    healthRecords.forEach((record) => {
      const key = record.department?.trim() || 'Belirtilmedi';
      departmentMap.set(key, (departmentMap.get(key) ?? 0) + 1);
    });
    const departmentRows = Array.from(departmentMap.entries()).sort((a, b) => b[1] - a[1]);
    const departmentMax = Math.max(...departmentRows.map((row) => row[1]), 1);

    return {
      totalEmployees,
      fitForWork,
      notFitForWork,
      restrictedDuty,
      chronicCount,
      communicableCount,
      medicalDue,
      vaccinationUpToDate,
      vaccinationRate,
      metricLabels,
      metricValues,
      departmentRows,
      departmentMax
    };
  }, [healthRecords]);

  const emergencyTypeLabel = (value: EmergencyType) => {
    const map: Record<EmergencyType, string> = {
      YANGIN: language === 'ru' ? 'Пожар' : 'Yangın',
      TAHLIYE: language === 'ru' ? 'Эвакуация' : 'Tahliye',
      ILK_YARDIM: language === 'ru' ? 'Первая помощь' : 'İlk Yardım',
      KIMYASAL_SIZINTI: language === 'ru' ? 'Химический разлив' : 'Kimyasal Sızıntı',
      KURTARMA: language === 'ru' ? 'Спасательные работы' : 'Kurtarma',
      DEPREM: language === 'ru' ? 'Землетрясение' : 'Deprem',
      DIGER: language === 'ru' ? 'Другое' : 'Diğer'
    };
    return map[value];
  };

  const emergencyResultLabel = (value: EmergencyDrillResult) => {
    const map: Record<EmergencyDrillResult, string> = {
      BASARILI: language === 'ru' ? 'Успешно' : 'Başarılı',
      KISMEN_BASARILI: language === 'ru' ? 'Частично успешно' : 'Kısmen Başarılı',
      BASARISIZ: language === 'ru' ? 'Неуспешно' : 'Başarısız'
    };
    return map[value];
  };

  const emergencyStatusLabel = (value: EmergencyDrillStatus) => {
    const map: Record<EmergencyDrillStatus, string> = {
      PLANLANDI: language === 'ru' ? 'Запланировано' : 'Planlandı',
      TAMAMLANDI: language === 'ru' ? 'Завершено' : 'Tamamlandı',
      IPTAL_EDILDI: language === 'ru' ? 'Отменено' : 'İptal Edildi'
    };
    return map[value];
  };

  const emergencyStatusClass = (value: EmergencyDrillStatus) => {
    if (value === 'TAMAMLANDI') {
      return 'status-badge-closed';
    }
    if (value === 'IPTAL_EDILDI') {
      return 'traffic-red';
    }
    return 'status-badge-open';
  };

  const trainingStatusLabel = (status: TrainingStatus) => {
    const map: Record<TrainingStatus, string> = {
      PLANLANDI: 'Planlandı',
      TAMAMLANDI: 'Tamamlandı',
      DEVAM_EDIYOR: 'Devam Ediyor',
      SURESI_DOLDU: 'Süresi Doldu'
    };
    return map[status];
  };

  const trainingStatusClass = (status: TrainingStatus) => {
    if (status === 'TAMAMLANDI') {
      return 'status-badge-closed';
    }
    if (status === 'DEVAM_EDIYOR') {
      return 'status-badge-progress';
    }
    if (status === 'SURESI_DOLDU') {
      return 'traffic-red';
    }
    return 'status-badge-open';
  };

  const trainingCertificateTrafficClass = (row: TrainingRecord) => {
    if (!row.certificateRequired) {
      return 'traffic-green';
    }
    const days = daysUntil(row.certificateValidityDate);
    if (days < 0) {
      return 'traffic-red';
    }
    if (days <= 30) {
      return 'traffic-amber';
    }
    return 'traffic-green';
  };

  const trainingCertificateTrafficLabel = (row: TrainingRecord) => {
    if (!row.certificateRequired) {
      return 'Sertifika Gerekli Değil';
    }
    const days = daysUntil(row.certificateValidityDate);
    if (days < 0) {
      return 'Süresi Doldu';
    }
    if (days <= 30) {
      return `Yaklaşan Süre (${days} gün)`;
    }
    return `Geçerli (${days} gün)`;
  };

  const ppeStatusLabel = (status: PpeStockStatus) => {
    if (status === 'YETERLI') {
      return 'Yeterli';
    }
    if (status === 'DUSUK_STOK') {
      return 'Düşük Stok';
    }
    return 'Stokta Yok';
  };

  const ppeStatusClass = (status: PpeStockStatus) => {
    if (status === 'YETERLI') {
      return 'status-badge-closed';
    }
    if (status === 'DUSUK_STOK') {
      return 'traffic-amber';
    }
    return 'traffic-red';
  };

  const ppeTransactionTypeLabel = (type: PpeTransactionType) => {
    if (type === 'STOK_GIRISI') {
      return 'Stok Girişi';
    }
    if (type === 'STOK_CIKISI') {
      return 'Stok Çıkışı';
    }
    if (type === 'STOGA_IADE') {
      return 'Stoğa İade';
    }
    if (type === 'HASARLI_HURDA') {
      return 'Hasarlı / Hurda';
    }
    return 'Stok Düzeltme';
  };

  const equipmentInspectionStatusLabel = (status: EquipmentInspectionStatus) => {
    if (status === 'COMPLIANT') {
      return language === 'ru' ? 'Соответствует' : 'Uygun';
    }
    if (status === 'UPCOMING') {
      return language === 'ru' ? 'Скоро истекает' : 'Yaklaşan';
    }
    return language === 'ru' ? 'Просрочено' : 'Gecikmiş';
  };

  const equipmentStatusLabel = (status: EquipmentStatus) => {
    if (status === 'ACTIVE') {
      return language === 'ru' ? 'Активно' : 'Aktif';
    }
    if (status === 'UNDER_MAINTENANCE') {
      return language === 'ru' ? 'На техническом обслуживании' : 'Bakımda';
    }
    return language === 'ru' ? 'Вне эксплуатации' : 'Servis Dışı';
  };

  const equipmentRiskLabel = (risk: EquipmentRiskLevel) => {
    if (risk === 'HIGH') {
      return language === 'ru' ? 'Высокий' : 'Yüksek';
    }
    if (risk === 'MEDIUM') {
      return language === 'ru' ? 'Средний' : 'Orta';
    }
    return language === 'ru' ? 'Низкий' : 'Düşük';
  };

  const equipmentStatusClassName = (status: EquipmentStatus) => {
    if (status === 'ACTIVE') {
      return 'equipment-status-active';
    }
    if (status === 'OUT_OF_SERVICE') {
      return 'equipment-status-out';
    }
    if (status === 'UNDER_MAINTENANCE') {
      return 'equipment-status-maintenance';
    }
    return '';
  };

  const equipmentStatusBadgeClassName = (record: EquipmentRecord) => {
    if (record.equipmentStatus === 'ACTIVE') {
      return 'status-badge-active-equipment';
    }
    if (record.equipmentStatus === 'UNDER_MAINTENANCE') {
      return 'status-badge-maintenance';
    }
    return getEquipmentTrafficClass(record);
  };

  const getEquipmentTrafficClass = (record: EquipmentRecord) => {
    const certDays = daysUntil(record.certificateExpiryDate);
    if (record.inspectionStatus === 'OVERDUE' || certDays < 0) {
      return 'traffic-red';
    }
    if (record.inspectionStatus === 'UPCOMING' || certDays <= 30) {
      return 'traffic-amber';
    }
    return 'traffic-green';
  };

  const getEquipmentTrafficLabel = (record: EquipmentRecord) => {
    const certDays = daysUntil(record.certificateExpiryDate);
    if (record.inspectionStatus === 'OVERDUE' || certDays < 0) {
      return language === 'ru' ? 'Красный' : 'Kırmızı';
    }
    if (record.inspectionStatus === 'UPCOMING' || certDays <= 30) {
      return language === 'ru' ? 'Жёлтый' : 'Sarı';
    }
    return language === 'ru' ? 'Зелёный' : 'Yeşil';
  };

  const kpiAnalyticsCards = useMemo(() => [
    {
      label: 'Denetim Tamamlanma',
      value: `${inspectionSummary.complianceScore}%`,
      note: `${inspectionSummary.totalFindings} bulgu`
    },
    {
      label: 'Açık Olaylar',
      value: String(incidentSummary.openCases),
      note: `${incidentSummary.totalIncidents} toplam olay`
    },
    {
      label: 'Eğitim Katılımcıları',
      value: String(trainingSummary.totalParticipants),
      note: `${trainingSummary.pendingEmployees} bekleyen çalışan`
    },
    {
      label: 'KKD Stok Sağlığı',
      value: String(ppeSummary.totalStock),
      note: `${ppeSummary.lowStockItems + ppeSummary.outOfStockItems} kritik stok kalemi`
    },
    {
      label: language === 'ru' ? 'Соответствие оборудования' : 'Ekipman Uygunluğu',
      value: String(equipmentSummary.compliant),
      note: language === 'ru' ? `${equipmentSummary.overdue} просроченных проверок` : `${equipmentSummary.overdue} gecikmiş muayene`
    },
    {
      label: language === 'ru' ? 'Годность к работе' : 'İşe Uygunluk',
      value: String(healthRecords.filter((record) => record.fitForWork === 'Yes').length),
      note: language === 'ru' ? `${healthRecords.filter((record) => daysUntil(record.nextMedicalExamination) <= 30).length} предстоящих проверок` : `${healthRecords.filter((record) => daysUntil(record.nextMedicalExamination) <= 30).length} yaklaşan muayene`
    }
  ], [equipmentSummary, healthRecords, incidentSummary, inspectionSummary, ppeSummary, trainingSummary]);

  const dashboardData = useMemo(() => {
    const periodMultiplier: Record<DashboardPeriod, number> = {
      daily: 0.25,
      weekly: 0.6,
      monthly: 1,
      yearly: 1.7,
      custom: 1
    };

    const customMultiplier =
      dashboardFilters.period === 'custom' && dashboardFilters.customStart && dashboardFilters.customEnd
        ? Math.max(
            0.2,
            Math.min(
              2,
              (Math.max(
                1,
                (new Date(dashboardFilters.customEnd).getTime() - new Date(dashboardFilters.customStart).getTime()) /
                  (1000 * 60 * 60 * 24) +
                  1
              ) as number) / 30
            )
          )
        : 1;
    const activePeriodMultiplier = dashboardFilters.period === 'custom' ? customMultiplier : periodMultiplier[dashboardFilters.period];
    const scaleSeries = (values: number[]) => values.map((value) => Math.max(1, Math.round(value * activePeriodMultiplier)));

    const projectScopedRows = (rows: ModuleRecord[]) =>
      dashboardFilters.projectId === 'all' ? rows : rows.filter((row) => row.projectId === dashboardFilters.projectId);

    const inspectionsRows = projectScopedRows(moduleData.inspections);
    const observationRows = projectScopedRows(moduleData.observations);
    const riskRows = projectScopedRows(moduleData['risk-assessments']);
    const ptwRows = projectScopedRows(moduleData['permit-to-work']);
    const incidentRows = projectScopedRows(moduleData.incidents);
    const auditRows = projectScopedRows(moduleData.audits);
    const trainingRows = projectScopedRows(moduleData.trainings);
    const ppeRows = projectScopedRows(moduleData['ppe-stocks']);
    const equipmentRows = projectScopedRows(moduleData['equipment-management']);
    const environmentalRows = projectScopedRows(moduleData.environmental);
    const actionRows = projectScopedRows(moduleData['action-tracker']);

    const allRows = [
      ...inspectionsRows,
      ...observationRows,
      ...riskRows,
      ...ptwRows,
      ...incidentRows,
      ...auditRows,
      ...trainingRows,
      ...ppeRows,
      ...equipmentRows,
      ...environmentalRows,
      ...actionRows
    ];
    const totalWorkforce = Math.max(
      allRows.filter((row) => row.projectId !== 'all' && row.title.toLowerCase().includes('employee')).reduce((sum, row) => sum + row.valueA, 0) + 245,
      250
    );
    const dailyWorkforce = Math.round(totalWorkforce * 0.91);
    const safeDaily = dailyWorkforce * 10;
    const safeWeekly = safeDaily * 7;
    const safeMonthly = safeDaily * 30;
    const safeTotal = safeDaily * 180;

    const lti = incidentRows.filter((row) => row.status === 'CLOSED').length;
    const mti = incidentRows.filter((row) => row.status !== 'CLOSED').length + 1;
    const firstAidCases = 4;
    const nearMisses = observationRows.reduce((sum, row) => sum + row.valueB, 0) + 2;
    const unsafeActs = observationRows.reduce((sum, row) => sum + row.valueA, 0);
    const unsafeConditions = observationRows.reduce((sum, row) => sum + row.valueB * 2, 0);
    const ltifr = ((lti * 1_000_000) / Math.max(safeTotal, 1)).toFixed(2);
    const trir = (((lti + mti + firstAidCases) * 200_000) / Math.max(safeTotal, 1)).toFixed(2);

    const totalInspections = inspectionsRows.reduce((sum, row) => sum + row.valueA, 0);
    const openFindings = inspectionsRows.filter((row) => row.status !== 'CLOSED').length + auditRows.filter((row) => row.status !== 'CLOSED').length;
    const closedFindings = inspectionsRows.filter((row) => row.status === 'CLOSED').length + auditRows.filter((row) => row.status === 'CLOSED').length;
    const safetyObservations = observationRows.reduce((sum, row) => sum + row.valueA, 0);
    const positiveObservations = observationRows.reduce((sum, row) => sum + row.valueB, 0);
    const overdueActions = actionRows.reduce((sum, row) => sum + row.valueB, 0);
    const weeklyInspectionPerformance = `${Math.round((closedFindings / Math.max(closedFindings + openFindings, 1)) * 100)}%`;

    const employeesTrainedWeekly = trainingRows.reduce((sum, row) => sum + row.valueA, 0);
    const dailyInductions = 11;
    const weeklyInductions = dailyInductions * 6;
    const monthlyInductions = weeklyInductions * 4;
    const totalInductions = monthlyInductions * 6;

    const ppeReceived = ppeRows.reduce((sum, row) => sum + row.valueA, 0);
    const ppeIssued = ppeRows.reduce((sum, row) => sum + row.valueB, 0);
    const ppeInStock = ppeReceived - ppeIssued;

    const equipmentTotal = equipmentRows.reduce((sum, row) => sum + row.valueA, 0);
    const healthTotal = healthRecords.length;
    const healthFit = healthRecords.filter((record) => record.fitForWork === 'Yes').length;
    const healthRestricted = healthRecords.filter((record) => record.restrictedWork.trim().toLowerCase() !== 'no' && record.restrictedWork.trim() !== '').length;
    const medicalDue = healthRecords.filter((record) => daysUntil(record.nextMedicalExamination) <= 30).length;
    const completedMedical = healthRecords.filter((record) => record.medicalExaminationDate.startsWith('2026')).length;
    const chronicEmployees = healthRecords.filter((record) => record.chronicDisease.trim().toLowerCase() !== 'none' && record.chronicDisease.trim() !== '').length;
    const communicableEmployees = healthRecords.filter((record) => record.communicableDisease.trim().toLowerCase() !== 'none' && record.communicableDisease.trim() !== '').length;
    const vaccinationRatio = `${Math.round((healthRecords.filter((record) => record.vaccinationStatus.toLowerCase().includes('up')).length / Math.max(healthTotal, 1)) * 100)}%`;

    const openRisks = riskRows.filter((row) => row.status !== 'CLOSED').length;
    const closedRisks = riskRows.filter((row) => row.status === 'CLOSED').length;
    const highRisks = riskRows.filter((row) => row.valueA >= 10).length;
    const mediumRisks = riskRows.filter((row) => row.valueA >= 6 && row.valueA < 10).length;
    const lowRisks = riskRows.filter((row) => row.valueA < 6).length;

    const wasteGenerated = environmentalRows.reduce((sum, row) => sum + row.valueA, 0);
    const wasteRecycled = environmentalRows.reduce((sum, row) => sum + row.valueB, 0);

    const internalAudits = auditRows.length;
    const externalAudits = 2;
    const openNcrs = openFindings + 1;
    const closedNcrs = closedFindings;

    const emergencyDrills = moduleData['emergency-preparedness'].reduce((sum, row) => sum + row.valueA, 0);
    const totalIncidents = moduleData.incidents.length;
    const nearMissReports = Math.max(0, nearMisses + incidentRows.filter((row) => row.status === 'OPEN').length);
    const criticalIncidents = incidentRows.filter((row) => row.valueA >= 10 || row.status !== 'CLOSED').length;

    const incidentStatus = (value: number, watchAt: number, riskAt: number, criticalAt: number): IncidentKpiStatus => {
      if (value >= criticalAt) {
        return 'critical';
      }
      if (value >= riskAt) {
        return 'risk';
      }
      if (value >= watchAt) {
        return 'monitor';
      }
      return 'normal';
    };

    const incidentsExecutiveKpis: IncidentExecutiveKpi[] = [
      {
        label: 'Total Incidents',
        value: String(totalIncidents),
        status: incidentStatus(totalIncidents, 6, 10, 14),
        note: 'Target <= 10'
      },
      {
        label: 'LTI',
        value: String(lti),
        status: incidentStatus(lti, 1, 2, 3),
        note: 'Zero harm target'
      },
      {
        label: 'Near Miss',
        value: String(nearMisses),
        status: incidentStatus(nearMisses, 8, 14, 20),
        note: 'Track by shift'
      },
      {
        label: 'First Aid Cases',
        value: String(firstAidCases),
        status: incidentStatus(firstAidCases, 4, 7, 10),
        note: 'Target <= 5'
      },
      {
        label: 'Unsafe Acts',
        value: String(unsafeActs),
        status: incidentStatus(unsafeActs, 12, 18, 24),
        note: 'Behavior actions'
      },
      {
        label: 'Unsafe Conditions',
        value: String(unsafeConditions),
        status: incidentStatus(unsafeConditions, 10, 15, 20),
        note: 'Engineering controls'
      },
      {
        label: 'Near Miss Reports',
        value: String(nearMissReports),
        status: incidentStatus(nearMissReports, 10, 16, 24),
        note: 'Reporting momentum'
      },
      {
        label: 'Critical Incidents',
        value: String(criticalIncidents),
        status: incidentStatus(criticalIncidents, 1, 2, 3),
        note: 'Immediate escalation'
      }
    ];

    const summaryKpis: SummaryKpi[] = [
      { label: 'Total Workforce', value: String(totalWorkforce), trend: '+2.1%' },
      { label: 'TRIR', value: trir, trend: '-0.3' },
      { label: 'LTIFR', value: ltifr, trend: '-0.1' },
      { label: 'Open Findings', value: String(openFindings), trend: '-4' },
      {
        label: 'PTW Active',
        value: String(ptwRows.filter((row) => row.status !== 'CLOSED').length),
        trend: '+1'
      },
      { label: 'Training Completion', value: `${Math.min(100, Math.round((employeesTrainedWeekly / Math.max(totalWorkforce, 1)) * 100))}%`, trend: '+5%' },
      { label: 'PPE Stock Health', value: `${Math.max(0, Math.min(100, Math.round((ppeInStock / Math.max(ppeReceived, 1)) * 100)))}%`, trend: '+2%' },
      { label: 'Risk Score', value: `${openRisks + highRisks * 2}`, trend: '-1' },
      { label: 'Compliance', value: '94%', trend: '+1%' },
      { label: 'Environmental Incidents', value: '1', trend: '0' }
    ];

    const sections: ExecutiveSection[] = [
      {
        key: 'safety-performance',
        title: 'Güvenlik Performansı',
        chartA: { type: 'line', values: scaleSeries([72, 74, 75, 77, 79, 82, 84]), label: 'Güvenlik Skoru Trendi' },
        chartB: { type: 'gauge', values: [84], label: 'Genel Güvenlik Endeksi' },
        highlights: [
          { label: 'TRIR', value: trir },
          { label: 'LTIFR', value: ltifr },
          { label: 'Near Miss', value: String(nearMisses) }
        ],
        details: [
          { label: 'LTI', value: String(lti) },
          { label: 'MTI', value: String(mti) },
          { label: 'First Aid Cases', value: String(firstAidCases) },
          { label: 'Safe Man-Hours - Daily', value: String(safeDaily) },
          { label: 'Safe Man-Hours - Weekly', value: String(safeWeekly) },
          { label: 'Safe Man-Hours - Monthly', value: String(safeMonthly) },
          { label: 'Safe Man-Hours - Total', value: String(safeTotal), note: 'Daily x weekly x monthly annualized capacity' }
        ]
      },
      {
        key: 'ptw',
        title: 'Permit to Work',
        chartA: {
          type: 'area',
          values: scaleSeries([6, 8, 7, 9, 10, 9, 11]),
          label: 'PTW Volume Trend'
        },
        chartB: {
          type: 'donut',
          values: [9, 2, 1],
          label: 'PTW Status Distribution'
        },
        highlights: [
          {
            label: 'PTW Active',
            value: String(ptwRows.filter((row) => row.status !== 'CLOSED').length)
          },
          {
            label: 'PTW Closed',
            value: String(ptwRows.filter((row) => row.status === 'CLOSED').length)
          },
          { label: 'Compliance', value: '91%' }
        ],
        details: [
          { label: 'Hot Work', value: '4' },
          { label: 'Confined Space', value: '2' },
          { label: 'Work at Height', value: '3' },
          { label: 'PTW Violation', value: '0' }
        ]
      },
      {
        key: 'inspections',
        title: 'Denetimler',
        chartA: { type: 'bar', values: scaleSeries([12, 14, 16, 15, 17, 14, totalInspections]), label: 'Denetim Sayısı Trendi' },
        chartB: { type: 'donut', values: [closedFindings, openFindings, overdueActions], label: 'Bulgu Durumu' },
        highlights: [
          { label: 'Toplam Denetim', value: String(totalInspections) },
          { label: 'Açık Bulgu', value: String(openFindings) },
          { label: 'Haftalık Performans', value: weeklyInspectionPerformance }
        ],
        details: [
          { label: 'Kapalı Bulgu', value: String(closedFindings) },
          { label: 'Geciken Aksiyonlar', value: String(overdueActions) },
          { label: 'Pozitif Gözlem', value: String(positiveObservations) }
        ]
      },
      {
        key: 'incidents',
        title: 'Olaylar',
        chartA: {
          type: 'bar',
          values: scaleSeries([
            Math.max(totalIncidents + 4, 1),
            Math.max(totalIncidents + 2, 1),
            Math.max(totalIncidents + 3, 1),
            Math.max(totalIncidents + 1, 1),
            Math.max(totalIncidents, 1),
            Math.max(totalIncidents - 1, 1),
            Math.max(totalIncidents, 1)
          ]),
          label: language === 'en' ? 'Incident Trend (Last 7 Periods)' : 'Olay Trendi (Son 7 Dönem)'
        },
        chartB: {
          type: 'donut',
          values: [Math.max(lti, 0), Math.max(firstAidCases, 0), Math.max(nearMissReports, 0), Math.max(criticalIncidents, 0)],
          label: language === 'en' ? 'Incident Category Distribution' : 'Olay Kategori Dağılımı'
        },
        highlights: [
          { label: 'Toplam Olay', value: String(moduleData.incidents.length) },
          { label: 'LTI', value: String(lti) },
          { label: language === 'en' ? 'Near Miss' : 'Ramak Kala Olay', value: String(nearMisses) }
        ],
        incidentExecutiveKpis: incidentsExecutiveKpis,
        details: []
      },
      {
        key: 'training',
        title: 'Eğitim',
        chartA: { type: 'area', values: scaleSeries([22, 28, 24, 30, 26, 34, employeesTrainedWeekly]), label: 'Eğitime Katılım Trendi' },
        chartB: { type: 'gauge', values: [Math.min(100, Math.round((employeesTrainedWeekly / Math.max(totalWorkforce, 1)) * 100))], label: 'Eğitim Tamamlama Oranı' },
        highlights: [
          { label: 'Haftalık Oryantasyon', value: String(weeklyInductions) },
          { label: 'Toolbox Konuşmaları', value: '28' },
          { label: 'Süresi Dolan Sertifikalar', value: '3' }
        ],
        details: [
          { label: 'Aylık Oryantasyon', value: String(monthlyInductions) },
          { label: 'Toplam Oryantasyon', value: String(totalInductions) },
          { label: 'Yaklaşan Eğitimler', value: '7' }
        ]
      },
      {
        key: 'ppe',
        title: 'KKD Yönetimi',
        chartA: { type: 'bar', values: scaleSeries([420, 380, 410, 395, 430, ppeReceived]), label: 'KKD Teslim Alım Trendi' },
        chartB: { type: 'donut', values: [ppeInStock, ppeIssued, 4], label: 'KKD Stok Dağılımı' },
        highlights: [
          { label: 'Stokta', value: String(ppeInStock) },
          { label: 'Dağıtılan', value: String(ppeIssued) },
          { label: 'Düşük Stok Kalemleri', value: '4' }
        ],
        details: [
          { label: 'Teslim Alınan', value: String(ppeReceived) },
          { label: 'Tüketim Trendi', value: '+6%' },
          { label: 'Dağılım', value: 'Mekanik 35% | İnşaat 30% | Elektrik 20% | İdari 15%' }
        ]
      },
      {
        key: 'equipment',
        title: language === 'ru' ? 'Оборудование' : 'Ekipman',
        chartA: { type: 'bar', values: scaleSeries([31, 34, 36, 38, 40, equipmentTotal]), label: language === 'ru' ? 'Инвентаризация оборудования' : 'Ekipman Envanteri' },
        chartB: { type: 'gauge', values: [88], label: language === 'ru' ? 'Уровень соответствия сертификации' : 'Sertifikasyon Uyum Oranı' },
        highlights: [
          { label: language === 'ru' ? 'Всего оборудования' : 'Toplam Ekipman', value: String(equipmentTotal) },
          { label: language === 'ru' ? 'Просроченные проверки' : 'Denetimi Gecikmiş', value: '2' },
          { label: language === 'ru' ? 'Вне эксплуатации' : 'Servis Dışı', value: '3' }
        ],
        details: [
          { label: language === 'ru' ? 'Мобильное' : 'Mobil', value: '18' },
          { label: language === 'ru' ? 'Тяжелое' : 'Ağır', value: '9' },
          { label: language === 'ru' ? 'Подъемное' : 'Kaldırma', value: '15' }
        ]
      },
      {
        key: 'occupational-health',
        title: 'İş Sağlığı',
        chartA: { type: 'line', values: scaleSeries([82, 84, 85, 86, 87, 89, 90]), label: 'İşe Uygunluk Trendi' },
        chartB: { type: 'donut', values: [healthFit, healthRestricted, Math.max(healthTotal - healthFit, 0)], label: 'Sağlık Durum Dağılımı' },
        highlights: [
          { label: 'İşe Uygun', value: String(healthFit) },
          { label: 'Muayene Zamanı Gelen', value: String(medicalDue) },
          { label: 'Aşılama', value: vaccinationRatio }
        ],
        details: [
          { label: 'Toplam Çalışan', value: String(healthTotal) },
          { label: 'Kronik Hastalık', value: String(chronicEmployees) },
          { label: 'Bulaşıcı Hastalık', value: String(communicableEmployees) }
        ]
      },
      {
        key: 'risk-management',
        title: 'Risk Yönetimi',
        chartA: { type: 'area', values: scaleSeries([9, 10, 8, 7, 6, 6, openRisks + highRisks]), label: 'Açık Risk Trendi' },
        chartB: { type: 'donut', values: [highRisks, mediumRisks, lowRisks], label: 'Risk Seviyesi Dağılımı' },
        highlights: [
          { label: 'Açık Riskler', value: String(openRisks) },
          { label: 'Yüksek Riskler', value: String(highRisks) },
          { label: 'Kapanan Riskler', value: String(closedRisks) }
        ],
        details: [
          { label: 'Orta Riskler', value: String(mediumRisks) },
          { label: 'Düşük Riskler', value: String(lowRisks) },
          { label: 'Kritik Risk Adedi', value: String(Math.min(10, moduleData['risk-assessments'].length)) }
        ]
      },
      {
        key: 'environment',
        title: 'Çevre',
        chartA: { type: 'line', values: scaleSeries([28, 26, 25, 24, 23, 24, wasteGenerated]), label: 'Atık Üretimi Trendi' },
        chartB: { type: 'bar', values: scaleSeries([14, 16, 17, 18, 17, wasteRecycled]), label: 'Geri Dönüşüm Trendi' },
        highlights: [
          { label: 'Oluşan Atık', value: `${wasteGenerated} ton` },
          { label: 'Geri Dönüştürülen Atık', value: `${wasteRecycled} ton` },
          { label: 'Çevresel Olaylar', value: '1' }
        ],
        details: [
          { label: 'Su Tüketimi', value: '1280 m3' },
          { label: 'Elektrik Tüketimi', value: '48,400 kWh' },
          { label: 'Yakıt Tüketimi', value: '6,300 L' }
        ]
      },
      {
        key: 'compliance',
        title: 'Mevzuata Uyum',
        chartA: { type: 'gauge', values: [94], label: 'Uyum Skoru' },
        chartB: { type: 'bar', values: scaleSeries([internalAudits, externalAudits, openNcrs, closedNcrs]), label: 'Tetkik ve Uygunsuzluk Görünümü' },
        highlights: [
          { label: 'Yasal Uyum', value: '94%' },
          { label: 'Açık Uygunsuzluk', value: String(openNcrs) },
          { label: 'Açık DÖF', value: String(actionRows.filter((row) => row.status !== 'CLOSED').length) }
        ],
        details: [
          { label: 'İç Tetkikler', value: String(internalAudits) },
          { label: 'Dış Tetkikler', value: String(externalAudits) },
          { label: 'Kapanan DÖF', value: String(actionRows.filter((row) => row.status === 'CLOSED').length) }
        ]
      },
    ];

    const comparisonProjects =
      dashboardFilters.selectedProjects.length > 0
        ? projectCatalog.filter((project) => dashboardFilters.selectedProjects.includes(project.id))
        : projectCatalog;
    const projectComparison = comparisonProjects.map((project) => ({
      label: project.name,
      value: allRows.filter((row) => row.projectId === project.id).length
    }));
    const departmentBase = [
      { label: 'Mechanical', value: 28 },
      { label: 'Civil', value: 24 },
      { label: 'Electrical', value: 19 },
      { label: 'Operations', value: 16 },
      { label: 'Admin', value: 11 }
    ];
    const departmentComparison =
      dashboardFilters.selectedDepartments.length > 0
        ? departmentBase.filter((department) => dashboardFilters.selectedDepartments.includes(department.label))
        : departmentBase;

    return { summaryKpis, sections, projectComparison, departmentComparison };
  }, [dashboardFilters, healthRecords, moduleData, projectCatalog]);

  const environmentalSummary = useMemo(() => {
    const records = moduleData.environmental;
    const wasteGenerated = records.reduce((sum, row) => sum + row.valueA, 0);
    const wasteRecycled = records.reduce((sum, row) => sum + row.valueB, 0);

    return {
      recordCount: records.length,
      wasteGenerated,
      wasteRecycled,
      netWaste: Math.max(wasteGenerated - wasteRecycled, 0)
    };
  }, [moduleData.environmental]);

  const dashboardXAxisLabels = useMemo(() => {
    const monthlyTr = ['Oca', 'Sub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Agu', 'Eyl', 'Eki', 'Kas', 'Ara'];
    const monthlyEn = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthlyRu = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

    const buildWeekly = (count: number) => Array.from({ length: count }, (_, index) => `W${index + 1}`);

    const buildDaily = (count: number) =>
      Array.from({ length: count }, (_, index) =>
        language === 'en' ? `D${index + 1}` : language === 'ru' ? `Д${index + 1}` : `${index + 1}.G`
      );

    const buildYearly = (count: number) => {
      const startYear = new Date().getFullYear() - Math.max(count - 1, 0);
      return Array.from({ length: count }, (_, index) => String(startYear + index));
    };

    return (count: number): string[] => {
      if (count <= 0) {
        return [];
      }

      if (dashboardFilters.period === 'weekly') {
        return buildWeekly(count);
      }
      if (dashboardFilters.period === 'daily') {
        return buildDaily(count);
      }
      if (dashboardFilters.period === 'yearly') {
        return buildYearly(count);
      }
      if (dashboardFilters.period === 'monthly') {
        const labels = language === 'en' ? monthlyEn : language === 'ru' ? monthlyRu : monthlyTr;
        return labels.slice(Math.max(labels.length - count, 0));
      }

      return buildWeekly(count);
    };
  }, [dashboardFilters.period, language]);

  const readFileAsDataUrl = useCallback((file: File) => {
    return new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }, []);

  const addObservationFiles = useCallback(async (files: File[]) => {
    const allowedMimeTypes = new Set(['image/jpeg', 'image/png', 'application/pdf']);
    const acceptedFiles = files.filter((file) => {
      const byMime = allowedMimeTypes.has(file.type);
      const byExt = /\.(jpg|jpeg|png|pdf)$/i.test(file.name);
      return byMime || byExt;
    });

    if (acceptedFiles.length === 0) {
      return;
    }

    const converted = await Promise.all(
      acceptedFiles.map(async (file) => ({
        id: `obs-attachment-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        name: file.name,
        mimeType: file.type || (file.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/jpeg'),
        dataUrl: await readFileAsDataUrl(file)
      }))
    );

    setObservationDraftAttachments((prev) => {
      const existing = new Set(prev.map((item) => `${item.name}-${item.dataUrl.length}`));
      const unique = converted.filter((item) => !existing.has(`${item.name}-${item.dataUrl.length}`));
      return [...prev, ...unique];
    });
  }, [readFileAsDataUrl]);

  const removeObservationAttachment = (attachmentId: string) => {
    setObservationDraftAttachments((prev) => prev.filter((attachment) => attachment.id !== attachmentId));
  };

  const setObservationProject = (projectId: string) => {
    const selectedProject = projectCatalog.find((project) => project.id === projectId);
    setObservationForm((prev) => ({
      ...prev,
      projectId,
      country: selectedProject?.country ?? prev.country,
      city: selectedProject?.city ?? prev.city,
      projectLocation: selectedProject?.address ?? prev.projectLocation
    }));
  };

  const resetObservationForm = () => {
    const selectedProject = projectCatalog.find((project) => project.id === observationForm.projectId) ?? projectCatalog[0];
    setObservationForm(createEmptyObservationForm(selectedProject));
    setObservationDraftAttachments([]);
    setEditingObservationId(null);
  };

  const saveObservationRecord = () => {
    if (
      !observationForm.projectId
      || !observationForm.inspectorName.trim()
      || !observationForm.responsiblePerson.trim()
      || !observationForm.subject.trim()
      || !observationForm.observationDescription.trim()
      || !observationForm.correctiveAction.trim()
    ) {
      return;
    }

    if (editingObservationId) {
      setObservationRecords((prev) => prev.map((record) => (
        record.id === editingObservationId
          ? {
            ...record,
            ...observationForm,
            attachments: observationDraftAttachments
          }
          : record
      )));
      setViewingObservationId(editingObservationId);
      resetObservationForm();
      return;
    }

    const nextRecord: ObservationRecord = {
      id: `observation-${Date.now()}`,
      observationNo: nextObservationNo,
      ...observationForm,
      attachments: observationDraftAttachments
    };

    setObservationRecords((prev) => [nextRecord, ...prev]);
    setViewingObservationId(nextRecord.id);
    resetObservationForm();
  };

  const editObservationRecord = (recordId: string) => {
    const target = observationRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }

    setEditingObservationId(recordId);
    setViewingObservationId(recordId);
    setObservationForm({
      projectId: target.projectId,
      country: target.country,
      city: target.city,
      projectLocation: target.projectLocation,
      inspectionDate: target.inspectionDate,
      inspectionTime: target.inspectionTime,
      inspectorName: target.inspectorName,
      contractor: target.contractor,
      subcontractor: target.subcontractor,
      responsiblePerson: target.responsiblePerson,
      category: target.category,
      subject: target.subject,
      observationLocation: target.observationLocation,
      violatedRequirement: target.violatedRequirement,
      observationDescription: target.observationDescription,
      correctiveAction: target.correctiveAction,
      dueDate: target.dueDate,
      comment: target.comment,
      priority: target.priority,
      status: target.status
    });
    setObservationDraftAttachments(target.attachments);
  };

  const deleteObservationRecord = (recordId: string) => {
    const target = observationRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    if (!window.confirm(language === 'ru' ? `Удалить предписание № \"${target.observationNo}\"?` : `\"${target.observationNo}\" numaralı gözlem kaydı silinsin mi?`)) {
      return;
    }
    setObservationRecords((prev) => prev.filter((record) => record.id !== recordId));
    if (viewingObservationId === recordId) {
      setViewingObservationId(null);
    }
    if (editingObservationId === recordId) {
      resetObservationForm();
    }
  };

  const buildObservationReportHtml = (record: ObservationRecord) => {
    const projectName = projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId;
    const imageAttachments = record.attachments.filter((attachment) => attachment.mimeType.startsWith('image/'));
    const pdfAttachments = record.attachments.filter((attachment) => attachment.mimeType === 'application/pdf');

    const reportPriorityLabel = language === 'ru' ? observationPriorityLabelRu[record.priority] : observationPriorityLabel[record.priority];
    const reportStatusLabel = language === 'ru' ? observationStatusLabelRu[record.status] : observationStatusLabel[record.status];

    return `
      <!doctype html>
      <html lang="${language === 'ru' ? 'ru' : 'tr'}">
      <head>
        <meta charset="utf-8" />
        <title>${observationCopy.reportTitle} - ${record.observationNo}</title>
        <style>
          @page { size: A4 portrait; margin: 10mm 10mm 12mm 10mm; }
          body { margin: 0; font-family: Arial, sans-serif; color: #0f172a; }
          .page { width: 100%; }
          table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          .sheet th, .sheet td {
            border: 1px solid #111827;
            padding: 6px 7px;
            font-size: 11px;
            line-height: 1.25;
            vertical-align: top;
            word-break: break-word;
          }
          .sheet th {
            font-weight: 700;
            background: #ffffff;
            text-align: left;
          }
          .header-table td {
            border: 1px solid #111827;
            height: 30px;
            font-size: 10px;
            padding: 4px 6px;
          }
          .header-logo {
            text-align: center;
            font-weight: 700;
            letter-spacing: 0.03em;
          }
          .header-title {
            text-align: center;
            font-size: 17px;
            font-weight: 700;
          }
          .meta-cell {
            font-size: 10px;
            white-space: nowrap;
          }
          .section-gap { height: 8px; border: 0 !important; padding: 0 !important; }
          .row-tall td { height: 58px; }
          .row-medium td { height: 40px; }
          .row-photos td { height: 230px; }
          .attachments-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 10px;
          }
          .attachment-card {
            border: 1px solid #0f172a;
            padding: 4px;
            background: #fff;
          }
          .attachment-card img {
            width: 100%;
            height: 220px;
            object-fit: cover;
            display: block;
          }
          .attachment-card p {
            margin: 4px 0 0;
            font-size: 12px;
          }
          .pdf-list {
            margin-top: 6px;
            font-size: 12px;
          }
          .signature-grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 8px;
            margin-top: 10px;
          }
          .signature-card {
            border: 1px solid #111827;
            min-height: 78px;
            padding: 6px;
            font-size: 11px;
          }
          .signature-line {
            margin-top: 36px;
            border-top: 1px solid #111827;
            padding-top: 4px;
            font-size: 10px;
          }
          .footer {
            border: 1px solid #111827;
            border-top: 0;
            margin-top: 0;
            padding: 6px 8px;
            font-size: 10px;
            display: flex;
            justify-content: space-between;
          }
        </style>
      </head>
      <body>
        <main class="page">
          <table class="header-table" aria-label="header">
            <colgroup>
              <col style="width: 20%" />
              <col style="width: 58%" />
              <col style="width: 22%" />
            </colgroup>
            <tr>
              <td class="header-logo" rowspan="3">ŞİRKET LOGOSU</td>
              <td class="header-title" rowspan="3">${observationCopy.reportHeaderTitle}</td>
              <td class="meta-cell"><strong>${observationCopy.formNo}:</strong> HSE-OBS-001</td>
            </tr>
            <tr>
              <td class="meta-cell"><strong>${observationCopy.revision}:</strong> 00</td>
            </tr>
            <tr>
              <td class="meta-cell"><strong>${observationCopy.observationNo}:</strong> ${record.observationNo}</td>
            </tr>
          </table>

          <table class="sheet" aria-label="observation report">
            <colgroup>
              <col style="width: 18%" />
              <col style="width: 32%" />
              <col style="width: 18%" />
              <col style="width: 32%" />
            </colgroup>
            <tr><td class="section-gap" colspan="4"></td></tr>
            <tr><th>${observationCopy.project}</th><td>${projectName}</td><th>${observationCopy.projectLocation}</th><td>${record.projectLocation}</td></tr>
            <tr><th>${observationCopy.inspectionDate}</th><td>${record.inspectionDate}</td><th>${observationCopy.inspectionTime}</th><td>${record.inspectionTime}</td></tr>
            <tr><th>${observationCopy.inspectorName}</th><td>${record.inspectorName}</td><th>${observationCopy.responsiblePerson}</th><td>${record.responsiblePerson}</td></tr>
            <tr><th>${observationCopy.contractor}</th><td>${record.contractor}</td><th>${observationCopy.subcontractor}</th><td>${record.subcontractor}</td></tr>
            <tr><th>${observationCopy.category}</th><td>${record.category}</td><th>${observationCopy.subject}</th><td>${record.subject}</td></tr>
            <tr><th>${observationCopy.observationLocation}</th><td colspan="3">${record.observationLocation}</td></tr>
            <tr class="row-medium"><th>${observationCopy.violatedRequirement}</th><td colspan="3">${record.violatedRequirement}</td></tr>
            <tr class="row-tall"><th>${observationCopy.observationDescription}</th><td colspan="3">${record.observationDescription}</td></tr>
            <tr class="row-tall"><th>${observationCopy.correctiveAction}</th><td colspan="3">${record.correctiveAction}</td></tr>
            <tr><th>${observationCopy.dueDate}</th><td>${record.dueDate}</td><th>${observationCopy.priority} / ${observationCopy.status}</th><td>${reportPriorityLabel} / ${reportStatusLabel}</td></tr>
            <tr class="row-medium"><th>${observationCopy.comment}</th><td colspan="3">${record.comment}</td></tr>
            <tr class="row-photos">
              <th>${observationCopy.attachmentsTitle}</th>
              <td colspan="3">
                ${imageAttachments.length > 0
                  ? `<div class="attachments-grid">${imageAttachments
                    .map((attachment) => `<div class="attachment-card"><img src="${attachment.dataUrl}" alt="${attachment.name}" /><p>${attachment.name}</p></div>`)
                    .join('')}</div>`
                  : observationCopy.noPhotoAttached}
                ${pdfAttachments.length > 0
                  ? `<div class="pdf-list">${pdfAttachments.map((attachment) => `<div>${attachment.name} (${observationCopy.pdfAttachment})</div>`).join('')}</div>`
                  : ''}
              </td>
            </tr>
          </table>

          <section class="signature-grid">
            <article class="signature-card"><strong>${observationCopy.preparedBy}</strong><div class="signature-line">${observationCopy.nameSignature}</div></article>
            <article class="signature-card"><strong>${observationCopy.reviewedBy}</strong><div class="signature-line">${observationCopy.nameSignature}</div></article>
            <article class="signature-card"><strong>${observationCopy.approvedBy}</strong><div class="signature-line">${observationCopy.nameSignature}</div></article>
          </section>

          <footer class="footer">
            <span>${observationCopy.reportSystemName}</span>
            <span>${observationCopy.page} 1/1</span>
          </footer>
        </main>
      </body>
      </html>
    `;
  };

  const openObservationReport = (record: ObservationRecord, autoPrint: boolean) => {
    const reportWindow = window.open('', '_blank', 'width=1200,height=900');
    if (!reportWindow) {
      return;
    }
    reportWindow.document.write(buildObservationReportHtml(record));
    reportWindow.document.close();
    reportWindow.focus();
    if (autoPrint) {
      reportWindow.print();
    }
  };

  const saveEntry = () => {
    if (activeModule === 'dashboard' || activeModule === 'occupational-health' || activeModule === 'projects' || activeModule === 'legal-register') {
      return;
    }
    if (editingModuleRecord && editingModuleRecord.module === activeModule) {
      setModuleData((prev) => ({
        ...prev,
        [activeModule]: prev[activeModule].map((record, index) => (
          index === editingModuleRecord.index
            ? {
                ...record,
                ...form,
                projectId: form.projectId,
                title: form.title || `${moduleLabels[activeModule]} entry`
              }
            : record
        ))
      }));
      setEditingModuleRecord(null);
      setForm((prev) => ({ ...prev, title: '', valueA: 0, valueB: 0, status: 'OPEN' }));
      return;
    }
    if (activeModule === 'environmental' && editingEnvironmentalIndex !== null) {
      setModuleData((prev) => ({
        ...prev,
        environmental: prev.environmental.map((record, index) => (
          index === editingEnvironmentalIndex
            ? {
                ...record,
                projectId: form.projectId,
                date: form.date,
                title: form.title || `${moduleLabels[activeModule]} entry`,
                valueA: form.valueA,
                valueB: form.valueB,
                status: form.status
              }
            : record
        ))
      }));
      setEditingEnvironmentalIndex(null);
      setForm((prev) => ({ ...prev, title: '', valueA: 0, valueB: 0, status: 'OPEN' }));
      return;
    }
    setModuleData((prev) => ({
      ...prev,
      [activeModule]: [
        {
          ...form,
          projectId: form.projectId,
          title: form.title || `${moduleLabels[activeModule]} entry`
        },
        ...prev[activeModule]
      ]
    }));
    setForm((prev) => ({ ...prev, title: '', valueA: 0, valueB: 0, status: 'OPEN' }));
  };

  const editModuleRecord = (recordIndex: number) => {
    const record = moduleData[activeModule]?.[recordIndex];
    if (!record || activeModule === 'environmental') {
      return;
    }

    setEditingModuleRecord({ module: activeModule as any, index: recordIndex });
    setForm({
      projectId: record.projectId,
      date: record.date,
      title: record.title,
      valueA: record.valueA,
      valueB: record.valueB,
      status: record.status
    });
  };

  const deleteModuleRecord = (recordIndex: number) => {
    const record = moduleData[activeModule]?.[recordIndex];
    if (!record || activeModule === 'environmental') {
      return;
    }

    if (!window.confirm(`"${record.title}" kaydı silinsin mi?`)) {
      return;
    }

    setModuleData((prev) => ({
      ...prev,
      [activeModule]: prev[activeModule].filter((_, index) => index !== recordIndex)
    }));

    if (editingModuleRecord?.module === activeModule && editingModuleRecord.index === recordIndex) {
      setEditingModuleRecord(null);
      setForm((prev) => ({ ...prev, title: '', valueA: 0, valueB: 0, status: 'OPEN' }));
    }
  };

  const editEnvironmentalRecord = (recordIndex: number) => {
    const record = moduleData.environmental[recordIndex];
    if (!record) {
      return;
    }
    setEditingEnvironmentalIndex(recordIndex);
    setForm({
      projectId: record.projectId,
      date: record.date,
      title: record.title,
      valueA: record.valueA,
      valueB: record.valueB,
      status: record.status
    });
  };

  const deleteEnvironmentalRecord = (recordIndex: number) => {
    const record = moduleData.environmental[recordIndex];
    if (!record) {
      return;
    }
    if (!window.confirm(language === 'ru' ? `Удалить экологическую запись "${record.title}"?` : `"${record.title}" çevre kaydı silinsin mi?`)) {
      return;
    }
    setModuleData((prev) => ({
      ...prev,
      environmental: prev.environmental.filter((_, index) => index !== recordIndex)
    }));
    if (editingEnvironmentalIndex === recordIndex) {
      setEditingEnvironmentalIndex(null);
      setForm((prev) => ({ ...prev, title: '', valueA: 0, valueB: 0, status: 'OPEN' }));
    }
  };

  const saveRiskEntry = () => {
    if (!riskForm.departmentActivity.trim() || !riskForm.hazard.trim() || !riskForm.responsiblePerson.trim()) {
      return;
    }

    const initialRiskScore = riskForm.likelihood * riskForm.severity;
    const residualRiskScore = riskForm.residualLikelihood * riskForm.residualSeverity;
    const newRecord: RiskRecord = {
      id: `risk-${Date.now()}`,
      riskId: nextRiskId,
      projectId: riskForm.projectId,
      departmentActivity: riskForm.departmentActivity,
      assessmentDate: riskForm.assessmentDate,
      hazard: riskForm.hazard,
      potentialConsequence: riskForm.potentialConsequence,
      existingControls: riskForm.existingControls,
      recommendedControls: riskForm.recommendedControls,
      likelihood: riskForm.likelihood,
      severity: riskForm.severity,
      initialRiskScore,
      residualLikelihood: riskForm.residualLikelihood,
      residualSeverity: riskForm.residualSeverity,
      residualRiskScore,
      responsiblePerson: riskForm.responsiblePerson,
      targetCompletionDate: riskForm.targetCompletionDate,
      status: riskForm.status,
      attachments: riskForm.attachments,
      photos: riskForm.photos,
      notes: riskForm.notes
    };

    setRiskRecords((prev) => [newRecord, ...prev]);
    setSelectedRiskId(newRecord.id);

    setRiskForm((prev) => ({
      ...prev,
      departmentActivity: '',
      hazard: '',
      potentialConsequence: '',
      existingControls: '',
      recommendedControls: '',
      likelihood: 3,
      severity: 3,
      residualLikelihood: 2,
      residualSeverity: 2,
      responsiblePerson: '',
      attachments: [],
      photos: [],
      notes: '',
      status: 'OPEN'
    }));
  };

  const saveIncidentEntry = () => {
    if (!incidentForm.title.trim()) {
      return;
    }

    if (editingIncidentIndex !== null) {
      setModuleData((prev) => ({
        ...prev,
        incidents: prev.incidents.map((row, index) => (
          index === editingIncidentIndex
            ? {
                ...row,
                projectId: incidentForm.projectId,
                date: incidentForm.date,
                title: incidentForm.title,
                valueA: incidentForm.incidentCount,
                valueB: incidentForm.lostWorkDays,
                status: incidentForm.status
              }
            : row
        ))
      }));
      setEditingIncidentIndex(null);
      setIncidentForm((prev) => ({ ...prev, title: '', incidentCount: 1, lostWorkDays: 0, status: 'OPEN' }));
      return;
    }

    const nextRow: ModuleRecord = {
      projectId: incidentForm.projectId,
      date: incidentForm.date,
      title: incidentForm.title,
      valueA: incidentForm.incidentCount,
      valueB: incidentForm.lostWorkDays,
      status: incidentForm.status
    };

    setModuleData((prev) => ({
      ...prev,
      incidents: [nextRow, ...prev.incidents]
    }));

    setIncidentForm((prev) => ({
      ...prev,
      title: '',
      incidentCount: 1,
      lostWorkDays: 0,
      status: 'OPEN'
    }));
  };

  const editIncidentRecord = (recordIndex: number) => {
    const record = moduleData.incidents[recordIndex];
    if (!record) {
      return;
    }

    setEditingIncidentIndex(recordIndex);
    setIncidentForm({
      projectId: record.projectId,
      date: record.date,
      title: record.title,
      incidentCount: record.valueA,
      lostWorkDays: record.valueB,
      status: record.status
    });
  };

  const deleteIncidentRecord = (recordIndex: number) => {
    const record = moduleData.incidents[recordIndex];
    if (!record) {
      return;
    }

    if (!window.confirm(`"${record.title}" olay kaydı silinsin mi?`)) {
      return;
    }

    setModuleData((prev) => ({
      ...prev,
      incidents: prev.incidents.filter((_, index) => index !== recordIndex)
    }));

    if (editingIncidentIndex === recordIndex) {
      setEditingIncidentIndex(null);
      setIncidentForm((prev) => ({ ...prev, title: '', incidentCount: 1, lostWorkDays: 0, status: 'OPEN' }));
    }
  };

  const saveEquipmentEntry = () => {
    if (!equipmentForm.equipmentName.trim() || !equipmentForm.equipmentType.trim() || !equipmentForm.responsiblePerson.trim()) {
      return;
    }

    const nextRecord: EquipmentRecord = {
      id: `eq-${Date.now()}`,
      equipmentId: nextEquipmentId,
      projectId: equipmentForm.projectId,
      equipmentName: equipmentForm.equipmentName,
      equipmentType: equipmentForm.equipmentType,
      brandModel: equipmentForm.brandModel,
      serialNumber: equipmentForm.serialNumber,
      location: equipmentForm.location,
      responsiblePerson: equipmentForm.responsiblePerson,
      lastInspectionDate: equipmentForm.lastInspectionDate,
      nextInspectionDate: equipmentForm.nextInspectionDate,
      inspectionStatus: equipmentForm.inspectionStatus,
      certificateNumber: equipmentForm.certificateNumber,
      certificateExpiryDate: equipmentForm.certificateExpiryDate,
      equipmentStatus: equipmentForm.equipmentStatus,
      riskLevel: equipmentForm.riskLevel,
      attachments: equipmentForm.attachments,
      inspectionReports: equipmentForm.inspectionReports,
      equipmentPhotos: equipmentForm.equipmentPhotos,
      notes: equipmentForm.notes
    };

    setEquipmentRecords((prev) => [nextRecord, ...prev]);
    setSelectedEquipmentId(nextRecord.id);

    setEquipmentForm((prev) => ({
      ...prev,
      equipmentName: '',
      equipmentType: '',
      brandModel: '',
      serialNumber: '',
      location: '',
      responsiblePerson: '',
      inspectionStatus: 'UPCOMING',
      certificateNumber: '',
      equipmentStatus: 'ACTIVE',
      riskLevel: 'MEDIUM',
      attachments: [],
      inspectionReports: [],
      equipmentPhotos: [],
      notes: ''
    }));
  };

  const saveEmergencyDrillEntry = () => {
    if (!emergencyDrillForm.drillName.trim() || !emergencyDrillForm.responsiblePerson.trim()) {
      return;
    }

    const nextRecord: EmergencyDrillRecord = {
      id: `drill-${Date.now()}`,
      drillId: nextEmergencyDrillId,
      projectId: emergencyDrillForm.projectId,
      emergencyType: emergencyDrillForm.emergencyType,
      drillName: emergencyDrillForm.drillName,
      drillDate: emergencyDrillForm.drillDate,
      participantCount: emergencyDrillForm.participantCount,
      drillResult: emergencyDrillForm.drillResult,
      openActions: emergencyDrillForm.openActions,
      closedActions: emergencyDrillForm.closedActions,
      responsiblePerson: emergencyDrillForm.responsiblePerson,
      nextPlannedDrillDate: emergencyDrillForm.nextPlannedDrillDate,
      status: emergencyDrillForm.status,
      attachments: emergencyDrillForm.attachments,
      drillReports: emergencyDrillForm.drillReports,
      photos: emergencyDrillForm.photos,
      notes: emergencyDrillForm.notes
    };

    setEmergencyDrillRecords((prev) => [nextRecord, ...prev]);
    setSelectedEmergencyDrillId(nextRecord.id);

    setEmergencyDrillForm((prev) => ({
      ...prev,
      drillName: '',
      participantCount: 10,
      openActions: 0,
      closedActions: 0,
      responsiblePerson: '',
      status: 'PLANLANDI',
      attachments: [],
      drillReports: [],
      photos: [],
      notes: ''
    }));
  };

  const saveWorkforceEntry = () => {
    if (!workforceForm.departmentArea.trim() || workforceForm.totalWorkforce <= 0) {
      return;
    }

    if (editingWorkforceId) {
      setWorkforceRecords((prev) => prev.map((record) => (
        record.id === editingWorkforceId
          ? {
              ...record,
              projectId: workforceForm.projectId,
              date: workforceForm.date,
              departmentArea: workforceForm.departmentArea,
              contractor: workforceForm.contractor,
              totalWorkforce: workforceForm.totalWorkforce,
              newEmployees: workforceForm.newEmployees,
              maleEmployees: workforceForm.maleEmployees,
              femaleEmployees: workforceForm.femaleEmployees,
              dayShiftWorkers: workforceForm.dayShiftWorkers,
              nightShiftWorkers: workforceForm.nightShiftWorkers,
              overtimeWorkers: workforceForm.overtimeWorkers,
              age18_25: workforceForm.age18_25,
              age26_35: workforceForm.age26_35,
              age36_45: workforceForm.age36_45,
              age46_55: workforceForm.age46_55,
              age56Plus: workforceForm.age56Plus,
              status: workforceForm.status,
              notes: workforceForm.notes
            }
          : record
      )));

      setEditingWorkforceId(null);
      setWorkforceForm((prev) => ({
        ...prev,
        totalWorkforce: 0,
        newEmployees: 0,
        maleEmployees: 0,
        femaleEmployees: 0,
        dayShiftWorkers: 0,
        nightShiftWorkers: 0,
        overtimeWorkers: 0,
        age18_25: 0,
        age26_35: 0,
        age36_45: 0,
        age46_55: 0,
        age56Plus: 0,
        status: 'OPEN',
        notes: ''
      }));
      return;
    }

    const nextRecord: WorkforceRecord = {
      id: `workforce-${Date.now()}`,
      projectId: workforceForm.projectId,
      date: workforceForm.date,
      departmentArea: workforceForm.departmentArea,
      contractor: workforceForm.contractor,
      totalWorkforce: workforceForm.totalWorkforce,
      newEmployees: workforceForm.newEmployees,
      maleEmployees: workforceForm.maleEmployees,
      femaleEmployees: workforceForm.femaleEmployees,
      dayShiftWorkers: workforceForm.dayShiftWorkers,
      nightShiftWorkers: workforceForm.nightShiftWorkers,
      overtimeWorkers: workforceForm.overtimeWorkers,
      age18_25: workforceForm.age18_25,
      age26_35: workforceForm.age26_35,
      age36_45: workforceForm.age36_45,
      age46_55: workforceForm.age46_55,
      age56Plus: workforceForm.age56Plus,
      status: workforceForm.status,
      notes: workforceForm.notes
    };

    setWorkforceRecords((prev) => [nextRecord, ...prev]);

    setWorkforceForm((prev) => ({
      ...prev,
      totalWorkforce: 0,
      newEmployees: 0,
      maleEmployees: 0,
      femaleEmployees: 0,
      dayShiftWorkers: 0,
      nightShiftWorkers: 0,
      overtimeWorkers: 0,
      age18_25: 0,
      age26_35: 0,
      age36_45: 0,
      age46_55: 0,
      age56Plus: 0,
      status: 'OPEN',
      notes: ''
    }));
  };

  const editWorkforceEntry = (recordId: string) => {
    const target = workforceRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }

    setEditingWorkforceId(recordId);
    setWorkforceForm({
      projectId: target.projectId,
      date: target.date,
      departmentArea: target.departmentArea,
      contractor: target.contractor,
      totalWorkforce: target.totalWorkforce,
      newEmployees: target.newEmployees,
      maleEmployees: target.maleEmployees,
      femaleEmployees: target.femaleEmployees,
      dayShiftWorkers: target.dayShiftWorkers,
      nightShiftWorkers: target.nightShiftWorkers,
      overtimeWorkers: target.overtimeWorkers,
      age18_25: target.age18_25,
      age26_35: target.age26_35,
      age36_45: target.age36_45,
      age46_55: target.age46_55,
      age56Plus: target.age56Plus,
      status: target.status,
      notes: target.notes
    });
  };

  const deleteWorkforceEntry = (recordId: string) => {
    const target = workforceRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }

    if (!window.confirm(`"${target.departmentArea} / ${target.contractor || 'Yüklenici yok'}" işgücü kaydı silinsin mi?`)) {
      return;
    }

    setWorkforceRecords((prev) => prev.filter((record) => record.id !== recordId));

    if (editingWorkforceId === recordId) {
      setEditingWorkforceId(null);
      setWorkforceForm((prev) => ({
        ...prev,
        totalWorkforce: 0,
        newEmployees: 0,
        maleEmployees: 0,
        femaleEmployees: 0,
        dayShiftWorkers: 0,
        nightShiftWorkers: 0,
        overtimeWorkers: 0,
        age18_25: 0,
        age26_35: 0,
        age36_45: 0,
        age46_55: 0,
        age56Plus: 0,
        status: 'OPEN',
        notes: ''
      }));
    }
  };

  const resetTrainingForm = () => {
    setTrainingForm((prev) => ({
      ...prev,
      trainingTitle: '',
      provider: '',
      projectEmployeeCount: 0,
      certifiedEmployeeCount: 0,
      totalTrainingCost: 0,
      status: 'PLANLANDI',
      attachments: [],
      participantList: [],
      certificates: [],
      notes: ''
    }));
    setEditingTrainingId(null);
  };

  const saveTrainingEntry = () => {
    if (!trainingForm.trainingTitle.trim() || !trainingForm.provider.trim()) {
      return;
    }

    const costPerEmployee =
      trainingForm.projectEmployeeCount > 0
        ? Math.round(trainingForm.totalTrainingCost / trainingForm.projectEmployeeCount)
        : 0;

    if (editingTrainingId) {
      setTrainingRecords((prev) =>
        prev.map((record) =>
          record.id === editingTrainingId
            ? {
                ...record,
                projectId: trainingForm.projectId,
                trainingType: trainingForm.trainingType,
                trainingTitle: trainingForm.trainingTitle,
                trainingCategory: trainingForm.trainingCategory,
                trainingDate: trainingForm.trainingDate,
                provider: trainingForm.provider,
                department: trainingForm.department,
                position: trainingForm.position,
                projectEmployeeCount: trainingForm.projectEmployeeCount,
                certifiedEmployeeCount: trainingForm.certifiedEmployeeCount,
                certificateRequired: trainingForm.certificateRequired,
                certificateValidityDate: trainingForm.certificateValidityDate,
                totalTrainingCost: trainingForm.totalTrainingCost,
                costPerEmployee,
                status: trainingForm.status,
                attachments: trainingForm.attachments,
                participantList: trainingForm.participantList,
                certificates: trainingForm.certificates,
                notes: trainingForm.notes
              }
            : record
        )
      );
      resetTrainingForm();
      return;
    }

    const nextRecord: TrainingRecord = {
      id: `training-${Date.now()}`,
      projectId: trainingForm.projectId,
      trainingId: nextTrainingId,
      trainingType: trainingForm.trainingType,
      trainingTitle: trainingForm.trainingTitle,
      trainingCategory: trainingForm.trainingCategory,
      trainingDate: trainingForm.trainingDate,
      provider: trainingForm.provider,
      department: trainingForm.department,
      position: trainingForm.position,
      projectEmployeeCount: trainingForm.projectEmployeeCount,
      certifiedEmployeeCount: trainingForm.certifiedEmployeeCount,
      certificateRequired: trainingForm.certificateRequired,
      certificateValidityDate: trainingForm.certificateValidityDate,
      totalTrainingCost: trainingForm.totalTrainingCost,
      costPerEmployee,
      status: trainingForm.status,
      attachments: trainingForm.attachments,
      participantList: trainingForm.participantList,
      certificates: trainingForm.certificates,
      notes: trainingForm.notes
    };

    setTrainingRecords((prev) => [nextRecord, ...prev]);
    resetTrainingForm();
  };

  const editTrainingEntry = (recordId: string) => {
    const target = trainingRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    setEditingTrainingId(recordId);
    setTrainingForm({
      projectId: target.projectId,
      trainingType: target.trainingType,
      trainingTitle: target.trainingTitle,
      trainingCategory: target.trainingCategory,
      trainingDate: target.trainingDate,
      provider: target.provider,
      department: target.department,
      position: target.position,
      projectEmployeeCount: target.projectEmployeeCount,
      certifiedEmployeeCount: target.certifiedEmployeeCount,
      certificateRequired: target.certificateRequired,
      certificateValidityDate: target.certificateValidityDate,
      totalTrainingCost: target.totalTrainingCost,
      status: target.status,
      attachments: target.attachments,
      participantList: target.participantList,
      certificates: target.certificates,
      notes: target.notes
    });
  };

  const deleteTrainingEntry = (recordId: string) => {
    const target = trainingRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.trainingTitle}" eğitim kaydını silmek istediğinize emin misiniz?`)) {
      return;
    }
    setTrainingRecords((prev) => prev.filter((record) => record.id !== recordId));
    if (editingTrainingId === recordId) {
      resetTrainingForm();
    }
  };

  const resetPpeTransactionForm = () => {
    setPpeTransactionForm((prev) => ({
      ...prev,
      quantity: 1,
      unitPrice: 0,
      notes: '',
      supplier: prev.transactionType === 'STOK_GIRISI' ? '' : prev.supplier,
      targetPersonDepartment: prev.transactionType === 'STOK_CIKISI' ? '' : prev.targetPersonDepartment
    }));
    setEditingPpeTransactionId(null);
  };

  const savePpeTransaction = () => {
    if (!ppeTransactionForm.itemName.trim() || !ppeTransactionForm.warehouse.trim() || !ppeTransactionForm.responsiblePerson.trim()) {
      return;
    }

    const totalCost = ppeTransactionForm.transactionType === 'STOK_GIRISI'
      ? Math.round(Math.max(0, ppeTransactionForm.quantity) * Math.max(0, ppeTransactionForm.unitPrice))
      : 0;

    if (editingPpeTransactionId) {
      setPpeTransactions((prev) =>
        prev.map((record) =>
          record.id === editingPpeTransactionId
            ? {
                ...record,
                transactionType: ppeTransactionForm.transactionType,
                lifecycle: ppeTransactionForm.lifecycle,
                projectId: ppeTransactionForm.projectId,
                warehouse: ppeTransactionForm.warehouse,
                date: ppeTransactionForm.date,
                category: ppeTransactionForm.category,
                itemName: ppeTransactionForm.itemName,
                brandModel: ppeTransactionForm.brandModel,
                unit: ppeTransactionForm.unit,
                quantity: Math.max(0, ppeTransactionForm.quantity),
                minimumStockLevel: Math.max(0, ppeTransactionForm.minimumStockLevel),
                responsiblePerson: ppeTransactionForm.responsiblePerson,
                supplier: ppeTransactionForm.supplier,
                targetPersonDepartment: ppeTransactionForm.targetPersonDepartment,
                unitPrice: ppeTransactionForm.transactionType === 'STOK_GIRISI' ? Math.max(0, ppeTransactionForm.unitPrice) : 0,
                totalCost,
                notes: ppeTransactionForm.notes
              }
            : record
        )
      );
      resetPpeTransactionForm();
      return;
    }

    const nextRecord: PpeTransactionRecord = {
      id: `ppe-tx-${Date.now()}`,
      transactionId: nextPpeTransactionId,
      transactionType: ppeTransactionForm.transactionType,
      lifecycle: ppeTransactionForm.lifecycle,
      projectId: ppeTransactionForm.projectId,
      warehouse: ppeTransactionForm.warehouse,
      date: ppeTransactionForm.date,
      category: ppeTransactionForm.category,
      itemName: ppeTransactionForm.itemName,
      brandModel: ppeTransactionForm.brandModel,
      unit: ppeTransactionForm.unit,
      quantity: Math.max(0, ppeTransactionForm.quantity),
      minimumStockLevel: Math.max(0, ppeTransactionForm.minimumStockLevel),
      responsiblePerson: ppeTransactionForm.responsiblePerson,
      supplier: ppeTransactionForm.supplier,
      targetPersonDepartment: ppeTransactionForm.targetPersonDepartment,
      unitPrice: ppeTransactionForm.transactionType === 'STOK_GIRISI' ? Math.max(0, ppeTransactionForm.unitPrice) : 0,
      totalCost,
      notes: ppeTransactionForm.notes
    };

    setPpeTransactions((prev) => [nextRecord, ...prev]);
    resetPpeTransactionForm();
  };

  const editPpeTransaction = (recordId: string) => {
    const target = ppeTransactions.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    setEditingPpeTransactionId(recordId);
    setPpeTransactionForm({
      transactionType: target.transactionType,
      lifecycle: target.lifecycle,
      projectId: target.projectId,
      warehouse: target.warehouse,
      date: target.date,
      category: target.category,
      itemName: target.itemName,
      brandModel: target.brandModel,
      unit: target.unit,
      quantity: target.quantity,
      minimumStockLevel: target.minimumStockLevel,
      responsiblePerson: target.responsiblePerson,
      supplier: target.supplier,
      targetPersonDepartment: target.targetPersonDepartment,
      unitPrice: target.unitPrice,
      notes: target.notes
    });
  };

  const deletePpeTransaction = (recordId: string) => {
    const target = ppeTransactions.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.transactionId}" KKD kaydını silmek istediğinize emin misiniz?`)) {
      return;
    }
    setPpeTransactions((prev) => prev.filter((record) => record.id !== recordId));
    if (editingPpeTransactionId === recordId) {
      resetPpeTransactionForm();
    }
  };

  const exportPpeInventoryCsv = () => {
    const headers = ['Proje', 'Tarih', 'KKD Kalemi', 'Gelen Stok', 'Cikisi Yapilan Stok', 'Kalan Stok', 'Minimum Stok', 'Durum'];
    const rows = ppeInventoryByKey.map((row) => {
      const stockStatus: PpeStockStatus = row.stock <= 0 ? 'STOKTA_YOK' : row.stock <= row.minimumStockLevel ? 'DUSUK_STOK' : 'YETERLI';
      return [
        projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId,
        row.lastDate,
        row.itemName,
        String(row.incoming + row.returned),
        String(row.outgoing + row.damaged),
        String(row.stock),
        String(row.minimumStockLevel),
        ppeStatusLabel(stockStatus)
      ];
    });

    const csv = [headers, ...rows]
      .map((line) => line.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(','))
      .join('\n');

    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `kkd-envanter-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const sendPpeSummaryEmail = () => {
    const subject = `KKD Envanter KPI Ozeti - ${new Date().toISOString().slice(0, 10)}`;
    const lines = [
      `Toplam KKD Stogu: ${ppeSummary.totalStock}`,
      `Bu Ay Gelen Stok: ${ppeSummary.monthlyIncoming}`,
      `Bu Ay Cikisi Yapilan Stok: ${ppeSummary.monthlyOutgoing}`,
      `Dusuk Stoklu Kalemler: ${ppeSummary.lowStockItems}`,
      `Toplam Envanter Degeri: ${ppeSummary.totalInventoryValue.toLocaleString('tr-TR')} Rub`,
      `Gereken Satin Alma Siparisleri: ${ppeSummary.purchaseOrdersNeeded}`
    ];
    const body = encodeURIComponent(lines.join('\n'));
    window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${body}`;
  };

  const exportPpeSummaryPdf = () => {
    const buildInventoryFromTransactions = (transactions: PpeTransactionRecord[]) => {
      const map = new Map<string, {
        category: string;
        stock: number;
        minimumStockLevel: number;
        incoming: number;
        outgoing: number;
        returned: number;
        damaged: number;
        value: number;
        averageUnitPrice: number;
      }>();

      transactions.forEach((record) => {
        if (record.lifecycle !== 'TAMAMLANDI') {
          return;
        }

        const key = `${record.category}__${record.itemName}__${record.unit}`;
        const current = map.get(key) ?? {
          category: record.category,
          stock: 0,
          minimumStockLevel: record.minimumStockLevel,
          incoming: 0,
          outgoing: 0,
          returned: 0,
          damaged: 0,
          value: 0,
          averageUnitPrice: 0
        };

        if (record.minimumStockLevel > 0) {
          current.minimumStockLevel = record.minimumStockLevel;
        }

        if (record.transactionType === 'STOK_GIRISI') {
          current.incoming += record.quantity;
          if (record.unitPrice > 0) {
            const prevIncoming = current.incoming - record.quantity;
            const prevValue = current.averageUnitPrice * Math.max(prevIncoming, 0);
            current.averageUnitPrice = (prevValue + record.quantity * record.unitPrice) / Math.max(current.incoming, 1);
          }
        }
        if (record.transactionType === 'STOK_CIKISI') {
          current.outgoing += record.quantity;
        }
        if (record.transactionType === 'STOGA_IADE') {
          current.returned += record.quantity;
        }
        if (record.transactionType === 'HASARLI_HURDA') {
          current.damaged += record.quantity;
        }

        const movement =
          record.transactionType === 'STOK_GIRISI' || record.transactionType === 'STOGA_IADE' || record.transactionType === 'STOK_DUZELTME'
            ? record.quantity
            : -record.quantity;
        current.stock += movement;

        map.set(key, current);
      });

      const items = Array.from(map.values()).map((item) => {
        const stock = Math.max(item.stock, 0);
        return {
          ...item,
          stock,
          value: Math.round(stock * item.averageUnitPrice)
        };
      });

      const totalStock = items.reduce((sum, row) => sum + row.stock, 0);
      const lowStockItems = items.filter((row) => row.stock > 0 && row.stock <= row.minimumStockLevel).length;
      const outOfStockItems = items.filter((row) => row.stock <= 0).length;
      const totalInventoryValue = items.reduce((sum, row) => sum + row.value, 0);
      const purchaseOrdersNeeded = items.filter((row) => row.stock <= row.minimumStockLevel).length;

      const statusDistribution = items.reduce(
        (accumulator, row) => {
          if (row.stock <= 0) {
            accumulator.stoktaYok += 1;
          } else if (row.stock <= row.minimumStockLevel) {
            accumulator.dusukStok += 1;
          } else {
            accumulator.yeterli += 1;
          }
          return accumulator;
        },
        { yeterli: 0, dusukStok: 0, stoktaYok: 0 }
      );

      const categoryMap = new Map<string, number>();
      items.forEach((row) => {
        categoryMap.set(row.category, (categoryMap.get(row.category) ?? 0) + row.stock);
      });
      const categoryRows = Array.from(categoryMap.entries()).sort((a, b) => b[1] - a[1]);

      const monthlyIncoming = transactions
        .filter((row) => row.lifecycle === 'TAMAMLANDI' && (row.transactionType === 'STOK_GIRISI' || row.transactionType === 'STOGA_IADE'))
        .reduce((sum, row) => sum + row.quantity, 0);

      const monthlyOutgoing = transactions
        .filter((row) => row.lifecycle === 'TAMAMLANDI' && (row.transactionType === 'STOK_CIKISI' || row.transactionType === 'HASARLI_HURDA'))
        .reduce((sum, row) => sum + row.quantity, 0);

      return {
        totalStock,
        monthlyIncoming,
        monthlyOutgoing,
        lowStockItems,
        totalInventoryValue,
        purchaseOrdersNeeded,
        statusDistribution,
        categoryRows,
        outOfStockItems
      };
    };

    const allProjectsReport = buildInventoryFromTransactions(ppeTransactions);
    const perProjectReports = projectCatalog.map((project) => {
      const scoped = ppeTransactions.filter((row) => row.projectId === project.id);
      return {
        projectName: project.name,
        ...buildInventoryFromTransactions(scoped)
      };
    });

    const reportWindow = window.open('', '_blank', 'width=1200,height=900');
    if (!reportWindow) {
      return;
    }

    const renderCategoryRows = (rows: Array<[string, number]>) => {
      if (rows.length === 0) {
        return '<p class="empty">Kategori verisi yok.</p>';
      }
      return `<table><thead><tr><th>Kategori</th><th>Mevcut Stok</th></tr></thead><tbody>${rows
        .map((row) => `<tr><td>${row[0]}</td><td>${row[1]}</td></tr>`)
        .join('')}</tbody></table>`;
    };

    const renderSection = (
      title: string,
      data: {
        totalStock: number;
        monthlyIncoming: number;
        monthlyOutgoing: number;
        lowStockItems: number;
        totalInventoryValue: number;
        purchaseOrdersNeeded: number;
        statusDistribution: { yeterli: number; dusukStok: number; stoktaYok: number };
        categoryRows: Array<[string, number]>;
        outOfStockItems: number;
      }
    ) => `
      <section class="report-section">
        <h2>${title}</h2>
        <div class="kpi-grid">
          <article><span>Toplam KKD Stoğu</span><strong>${data.totalStock}</strong></article>
          <article><span>Bu Ay Gelen Stok</span><strong>${data.monthlyIncoming}</strong></article>
          <article><span>Bu Ay Çıkışı Yapılan Stok</span><strong>${data.monthlyOutgoing}</strong></article>
          <article><span>Düşük Stoklu Kalemler</span><strong>${data.lowStockItems}</strong></article>
          <article><span>Toplam Envanter Değeri</span><strong>${data.totalInventoryValue.toLocaleString('tr-TR')} Rub</strong></article>
          <article><span>Gereken Satın Alma Siparişleri</span><strong>${data.purchaseOrdersNeeded}</strong></article>
        </div>
        <div class="viz-grid">
          <article>
            <h3>Kategori Bazında Mevcut KKD Stoğu</h3>
            ${renderCategoryRows(data.categoryRows)}
          </article>
          <article>
            <h3>Stok Durumu</h3>
            <p>Yeterli: <strong>${data.statusDistribution.yeterli}</strong></p>
            <p>Düşük Stok: <strong>${data.statusDistribution.dusukStok}</strong></p>
            <p>Stokta Yok: <strong>${data.statusDistribution.stoktaYok}</strong></p>
            <p>Stok Tükenme Uyarısı: <strong>${data.outOfStockItems}</strong></p>
          </article>
        </div>
      </section>
    `;

    reportWindow.document.write(`
      <!doctype html>
      <html lang="tr">
      <head>
        <meta charset="utf-8" />
        <title>KKD Envanter KPI Raporu</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
          h1 { margin: 0 0 8px; font-size: 22px; }
          .report-header { display: flex; justify-content: space-between; align-items: center; border: 1px solid #dbe3ee; border-radius: 10px; padding: 10px 12px; background: #f8fbff; }
          .logo-box { width: 84px; height: 34px; border: 1px solid #94a3b8; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; color: #0b2740; background: #ffffff; }
          .header-meta { text-align: right; font-size: 12px; color: #334155; }
          .date { color: #475569; margin-bottom: 14px; }
          .report-section { page-break-inside: avoid; margin-top: 18px; border: 1px solid #dbe3ee; border-radius: 10px; padding: 14px; }
          .kpi-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
          .kpi-grid article { border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; }
          .kpi-grid span { display: block; color: #475569; font-size: 12px; }
          .kpi-grid strong { display: block; margin-top: 4px; font-size: 18px; color: #0b2740; }
          .viz-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 12px; }
          .viz-grid article { border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; }
          table { width: 100%; border-collapse: collapse; margin-top: 6px; }
          th, td { border: 1px solid #e2e8f0; padding: 6px; text-align: left; font-size: 12px; }
          .empty { color: #64748b; font-size: 12px; }
          .signature-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-top: 18px; }
          .signature-card { border: 1px solid #dbe3ee; border-radius: 8px; padding: 10px; min-height: 84px; }
          .signature-card h4 { margin: 0 0 16px; font-size: 12px; color: #334155; }
          .signature-line { border-top: 1px solid #94a3b8; margin-top: 28px; padding-top: 4px; font-size: 11px; color: #64748b; }
          .report-footer { margin-top: 14px; border-top: 1px dashed #cbd5e1; padding-top: 8px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; }
          @media print {
            .report-section { page-break-after: always; }
            .report-section:last-of-type { page-break-after: auto; }
            .signature-grid { page-break-inside: avoid; }
          }
        </style>
      </head>
      <body>
        <div class="report-header">
          <div class="logo-box">HSE LOGO</div>
          <div>
            <h1>KKD Envanter KPI Raporu</h1>
            <div class="date">Tarih: ${new Date().toLocaleDateString('tr-TR')}</div>
          </div>
          <div class="header-meta">
            <div>Rapor Kodu: KKD-KPI-RPT</div>
            <div>Revizyon No: 01</div>
          </div>
        </div>
        ${renderSection('Tüm Projeler', allProjectsReport)}
        ${perProjectReports.map((section) => renderSection(section.projectName, section)).join('')}
        <div class="signature-grid">
          <article class="signature-card">
            <h4>Hazırlayan</h4>
            <div class="signature-line">Ad Soyad / İmza</div>
          </article>
          <article class="signature-card">
            <h4>Kontrol Eden</h4>
            <div class="signature-line">Ad Soyad / İmza</div>
          </article>
          <article class="signature-card">
            <h4>Onaylayan</h4>
            <div class="signature-line">Ad Soyad / İmza</div>
          </article>
        </div>
        <div class="report-footer">
          <span>Kurumsal İSG Envanter Raporu</span>
          <span>Dokuman: KKD-KPI-RPT | Rev: 01</span>
        </div>
      </body>
      </html>
    `);

    reportWindow.document.close();
    reportWindow.focus();
    reportWindow.print();
    reportWindow.close();
  };

  const resetProjectForm = () => {
    setProjectForm({ name: '', country: '', city: '', address: '', contractScope: '' });
    setEditingProjectId(null);
  };

  const saveProject = async () => {
    const missingFields: string[] = [];
    if (!projectForm.name.trim()) missingFields.push('Proje Adı');
    if (!projectForm.country.trim()) missingFields.push('Ülke');
    if (!projectForm.city.trim()) missingFields.push('Şehir');
    if (!projectForm.address.trim()) missingFields.push('Konum');
    if (!projectForm.contractScope.trim()) missingFields.push('Sözleşme / İş Kapsamı');

    if (missingFields.length > 0) {
      setProjectSaveFeedback({
        type: 'error',
        text: `Kayıt yapılmadı. Eksik alanlar: ${missingFields.join(', ')}`
      });
      return;
    }

    try {
      if (editingProjectId) {
        const row = await apiRequest<{ id: string; name: string; country: string; city: string; address: string; contractScope: string }>(`/api/master/projects/${editingProjectId}`, {
          method: 'PUT',
          body: JSON.stringify(projectForm)
        });
        setProjectCatalog((prev) =>
          prev.map((project) => (project.id === editingProjectId ? { id: row.id, name: row.name, country: row.country, city: row.city, address: row.address, contractScope: row.contractScope } : project))
        );
        setProjectSaveFeedback({ type: 'success', text: 'Proje başarıyla güncellendi ve backend’e kaydedildi.' });
        resetProjectForm();
        return;
      }

      const row = await apiRequest<{ id: string; name: string; country: string; city: string; address: string; contractScope: string }>(`/api/master/projects`, {
        method: 'POST',
        body: JSON.stringify(projectForm)
      });
      setProjectCatalog((prev) => [{ id: row.id, name: row.name, country: row.country, city: row.city, address: row.address, contractScope: row.contractScope }, ...prev]);
      setProjectSaveFeedback({ type: 'success', text: `"${projectForm.name.trim()}" projesi backend’e kaydedildi.` });
      resetProjectForm();
    } catch {
      if (editingProjectId) {
        setProjectCatalog((prev) =>
          prev.map((project) => (project.id === editingProjectId ? { ...project, ...projectForm } : project))
        );
        setProjectSaveFeedback({ type: 'success', text: 'Proje başarıyla güncellendi ve listede yenilendi.' });
        resetProjectForm();
        return;
      }

      const baseId = projectForm.name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'project';
      const id = projectCatalog.some((project) => project.id === baseId) ? `${baseId}-${Date.now()}` : baseId;
      const nextProject: Project = { id, ...projectForm };
      setProjectCatalog((prev) => [nextProject, ...prev]);
      setProjectSaveFeedback({ type: 'success', text: `"${projectForm.name.trim()}" projesi kaydedildi ve listenin başına eklendi.` });
      resetProjectForm();
    }
  };

  const editProject = (projectId: string) => {
    const target = projectCatalog.find((project) => project.id === projectId);
    if (!target) {
      return;
    }
    setProjectSaveFeedback(null);
    setEditingProjectId(projectId);
    setProjectForm({
      name: target.name,
      country: target.country,
      city: target.city,
      address: target.address,
      contractScope: target.contractScope
    });
  };

  const deleteProject = async (projectId: string) => {
    const target = projectCatalog.find((project) => project.id === projectId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.name}" projesini silmek istediğinize emin misiniz?`)) {
      return;
    }
    try {
      await apiRequest<void>(`/api/master/projects/${projectId}`, { method: 'DELETE' });
      setProjectCatalog((prev) => prev.filter((project) => project.id !== projectId));
      setProjectSaveFeedback({ type: 'success', text: `"${target.name}" projesi backend’den silindi.` });
      if (editingProjectId === projectId) {
        resetProjectForm();
      }
    } catch {
      setProjectCatalog((prev) => prev.filter((project) => project.id !== projectId));
      setProjectSaveFeedback({ type: 'success', text: `"${target.name}" projesi listeden silindi.` });
      if (editingProjectId === projectId) {
        resetProjectForm();
      }
    }
  };

  const resetDepartmentForm = () => {
    setDepartmentForm({ name: '', code: '', description: '' });
    setEditingDepartmentId(null);
  };

  const saveDepartment = async () => {
    if (!departmentForm.name.trim() || !departmentForm.code.trim()) {
      return;
    }
    try {
      if (editingDepartmentId) {
        const row = await apiRequest<{ id: string; name: string; code: string; description: string | null }>(`/api/master/departments/${editingDepartmentId}`, {
          method: 'PUT',
          body: JSON.stringify(departmentForm)
        });
        setDepartmentRecords((prev) =>
          prev.map((department) => (department.id === editingDepartmentId ? { id: row.id, name: row.name, code: row.code, description: row.description ?? '' } : department))
        );
        resetDepartmentForm();
        return;
      }
      const row = await apiRequest<{ id: string; name: string; code: string; description: string | null }>(`/api/master/departments`, {
        method: 'POST',
        body: JSON.stringify(departmentForm)
      });
      setDepartmentRecords((prev) => [{ id: row.id, name: row.name, code: row.code, description: row.description ?? '' }, ...prev]);
      resetDepartmentForm();
    } catch {
      if (editingDepartmentId) {
        setDepartmentRecords((prev) =>
          prev.map((department) => (department.id === editingDepartmentId ? { ...department, ...departmentForm } : department))
        );
        resetDepartmentForm();
        return;
      }
      const baseId = departmentForm.name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'department';
      const id = departmentRecords.some((department) => department.id === baseId) ? `${baseId}-${Date.now()}` : baseId;
      setDepartmentRecords((prev) => [{ id, ...departmentForm }, ...prev]);
      resetDepartmentForm();
    }
  };

  const editDepartment = (departmentId: string) => {
    const target = departmentRecords.find((department) => department.id === departmentId);
    if (!target) {
      return;
    }
    setEditingDepartmentId(departmentId);
    setDepartmentForm({
      name: target.name,
      code: target.code,
      description: target.description
    });
  };

  const deleteDepartment = async (departmentId: string) => {
    const target = departmentRecords.find((department) => department.id === departmentId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.name}" departmanını silmek istediğinize emin misiniz?`)) {
      return;
    }
    try {
      await apiRequest<void>(`/api/master/departments/${departmentId}`, { method: 'DELETE' });
      setDepartmentRecords((prev) => prev.filter((department) => department.id !== departmentId));
      if (editingDepartmentId === departmentId) {
        resetDepartmentForm();
      }
    } catch {
      setDepartmentRecords((prev) => prev.filter((department) => department.id !== departmentId));
      if (editingDepartmentId === departmentId) {
        resetDepartmentForm();
      }
    }
  };

  const fillContractorFormFromProject = (projectId: string) => {
    const selected = projectCatalog.find((project) => project.id === projectId);
    if (!selected) {
      return;
    }
    setContractorForm((prev) => ({
      ...prev,
      projectId: selected.id,
      projectName: selected.name,
      country: selected.country,
      city: selected.city,
      projectLocation: selected.address,
      contractScope: selected.contractScope
    }));
  };

  const resetContractorForm = () => {
    const fallbackProject = projectCatalog[0];
    setContractorForm({
      companyName: '',
      projectId: fallbackProject?.id ?? '',
      projectName: fallbackProject?.name ?? '',
      country: fallbackProject?.country ?? '',
      city: fallbackProject?.city ?? '',
      projectLocation: fallbackProject?.address ?? '',
      contractScope: fallbackProject?.contractScope ?? '',
      hseWarningCount: 0,
      hseWarningDate: '',
      fireWarningCount: 0,
      fireWarningDate: '',
      environmentWarningCount: 0,
      environmentWarningDate: '',
      penaltyCount: 0,
      penaltyLegalClause: '',
      totalPenaltyAmount: 0
    });
    setEditingContractorId(null);
  };

  const saveContractor = async () => {
    if (!contractorForm.companyName.trim() || !contractorForm.projectName.trim() || !contractorForm.country.trim() || !contractorForm.city.trim() || !contractorForm.projectLocation.trim() || !contractorForm.contractScope.trim()) {
      return;
    }

    const contractorPayload = {
      ...contractorForm,
      hseWarningDate: contractorForm.hseWarningDate ? new Date(contractorForm.hseWarningDate).toISOString() : undefined,
      fireWarningDate: contractorForm.fireWarningDate ? new Date(contractorForm.fireWarningDate).toISOString() : undefined,
      environmentWarningDate: contractorForm.environmentWarningDate ? new Date(contractorForm.environmentWarningDate).toISOString() : undefined
    };

    try {
      if (editingContractorId) {
        const row = await apiRequest<{ id: string; companyName: string; projectId: string | null; country: string; city: string; projectLocation: string; contractScope: string; hseWarningCount: number; hseWarningDate: string | null; fireWarningCount: number; fireWarningDate: string | null; environmentWarningCount: number; environmentWarningDate: string | null; penaltyCount: number; penaltyLegalClause: string | null; totalPenaltyAmount: number }>(`/api/master/contractors/${editingContractorId}`, {
          method: 'PUT',
          body: JSON.stringify(contractorPayload)
        });
        setContractorRecords((prev) => prev.map((record) => (record.id === editingContractorId ? {
          id: row.id,
          companyName: row.companyName,
          projectId: row.projectId ?? contractorForm.projectId,
          projectName: contractorForm.projectName,
          country: row.country,
          city: row.city,
          projectLocation: row.projectLocation,
          contractScope: row.contractScope,
          hseWarningCount: row.hseWarningCount,
          hseWarningDate: row.hseWarningDate ?? '',
          fireWarningCount: row.fireWarningCount,
          fireWarningDate: row.fireWarningDate ?? '',
          environmentWarningCount: row.environmentWarningCount,
          environmentWarningDate: row.environmentWarningDate ?? '',
          penaltyCount: row.penaltyCount,
          penaltyLegalClause: row.penaltyLegalClause ?? '',
          totalPenaltyAmount: row.totalPenaltyAmount
        } : record)));
        resetContractorForm();
        return;
      }

      const row = await apiRequest<{ id: string; companyName: string; projectId: string | null; country: string; city: string; projectLocation: string; contractScope: string; hseWarningCount: number; hseWarningDate: string | null; fireWarningCount: number; fireWarningDate: string | null; environmentWarningCount: number; environmentWarningDate: string | null; penaltyCount: number; penaltyLegalClause: string | null; totalPenaltyAmount: number }>(`/api/master/contractors`, {
        method: 'POST',
        body: JSON.stringify(contractorPayload)
      });
      setContractorRecords((prev) => [...prev, {
        id: row.id,
        companyName: row.companyName,
        projectId: row.projectId ?? contractorForm.projectId,
        projectName: contractorForm.projectName,
        country: row.country,
        city: row.city,
        projectLocation: row.projectLocation,
        contractScope: row.contractScope,
        hseWarningCount: row.hseWarningCount,
        hseWarningDate: row.hseWarningDate ?? '',
        fireWarningCount: row.fireWarningCount,
        fireWarningDate: row.fireWarningDate ?? '',
        environmentWarningCount: row.environmentWarningCount,
        environmentWarningDate: row.environmentWarningDate ?? '',
        penaltyCount: row.penaltyCount,
        penaltyLegalClause: row.penaltyLegalClause ?? '',
        totalPenaltyAmount: row.totalPenaltyAmount
      }]);
      resetContractorForm();
    } catch {
      if (editingContractorId) {
        setContractorRecords((prev) => prev.map((record) => (record.id === editingContractorId ? { ...record, ...contractorForm } : record)));
        resetContractorForm();
        return;
      }

      const id = `contractor-${Date.now()}`;
      setContractorRecords((prev) => [...prev, { id, ...contractorForm }]);
      resetContractorForm();
    }
  };

  const editContractor = (recordId: string) => {
    const target = contractorRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    setEditingContractorId(recordId);
    setContractorForm({
      companyName: target.companyName,
      projectId: target.projectId,
      projectName: target.projectName,
      country: target.country,
      city: target.city,
      projectLocation: target.projectLocation,
      contractScope: target.contractScope,
      hseWarningCount: target.hseWarningCount,
      hseWarningDate: target.hseWarningDate,
      fireWarningCount: target.fireWarningCount,
      fireWarningDate: target.fireWarningDate,
      environmentWarningCount: target.environmentWarningCount,
      environmentWarningDate: target.environmentWarningDate,
      penaltyCount: target.penaltyCount,
      penaltyLegalClause: target.penaltyLegalClause,
      totalPenaltyAmount: target.totalPenaltyAmount
    });
  };

  const deleteContractor = async (recordId: string) => {
    const target = contractorRecords.find((record) => record.id === recordId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.companyName}" alt yüklenici kaydını silmek istediğinize emin misiniz?`)) {
      return;
    }
    try {
      await apiRequest<void>(`/api/master/contractors/${recordId}`, { method: 'DELETE' });
      setContractorRecords((prev) => prev.filter((record) => record.id !== recordId));
      if (editingContractorId === recordId) {
        resetContractorForm();
      }
    } catch {
      setContractorRecords((prev) => prev.filter((record) => record.id !== recordId));
      if (editingContractorId === recordId) {
        resetContractorForm();
      }
    }
  };

  const resetSettingsConfig = () => {
    setSettingsConfig(initialSettingsConfig);
  };

  const saveSettingsConfig = () => {
    window.alert('Sistem yapılandırması kaydedildi.');
  };

  const setChecklistAnswer = (sectionId: string, questionIndex: number, answer: ChecklistAnswer) => {
    const key = `${sectionId}-${questionIndex}`;
    setChecklistAnswers((prev) => ({ ...prev, [key]: answer }));
  };

  const getChecklistProgress = (sectionId: string, questionCount: number) => {
    if (skippedChecklistSections[sectionId]) {
      return { answered: 0, percent: 0 };
    }
    let answered = 0;
    for (let i = 0; i < questionCount; i += 1) {
      const answer = checklistAnswers[`${sectionId}-${i}`] ?? '';
      if (answer !== '') {
        answered += 1;
      }
    }
    return { answered, percent: Math.round((answered / questionCount) * 100) };
  };

  const addManualAction = () => {
    if (!manualActionText.trim()) {
      return;
    }
    setManualActions((prev) => [{ text: manualActionText.trim(), status: 'OPEN' }, ...prev]);
    setManualActionText('');
  };

  const updateManualActionStatus = (index: number, status: ActionStatus) => {
    setManualActions((prev) => prev.map((action, actionIndex) => (actionIndex === index ? { ...action, status } : action)));
  };

  const resetInspection = () => {
    setInspectionForm((prev) => ({
      ...prev,
      businessUnit: '',
      siteArea: '',
      locationType: '',
      inspectorName: '',
      department: '',
      positiveObservations: '',
      inspectionNotes: ''
    }));
    setChecklistAnswers({});
    setSkippedChecklistSections({});
    setManualActionText('');
    setManualActions([]);
  };

  const resetHealthForm = () => {
    setHealthForm((prev) => ({
      ...prev,
      employeeName: '',
      employeeId: '',
      department: '',
      position: '',
      bloodGroup: '',
      allergies: '',
      chronicDisease: '',
      communicableDisease: '',
      medication: '',
      disabilityStatus: '',
      fitForWork: 'Yes',
      restrictedWork: '',
      vaccinationStatus: '',
      remarks: ''
    }));
    setEditingHealthEmployeeId(null);
  };

  const saveHealthEntry = () => {
    if (!healthForm.employeeName.trim() || !healthForm.employeeId.trim()) {
      return;
    }
    if (editingHealthEmployeeId) {
      setHealthRecords((prev) => prev.map((record) => (record.employeeId === editingHealthEmployeeId ? healthForm : record)));
      resetHealthForm();
      return;
    }
    setHealthRecords((prev) => [healthForm, ...prev]);
    resetHealthForm();
  };

  const editHealthEntry = (employeeId: string) => {
    const target = healthRecords.find((record) => record.employeeId === employeeId);
    if (!target) {
      return;
    }
    setEditingHealthEmployeeId(employeeId);
    setHealthForm(target);
  };

  const deleteHealthEntry = (employeeId: string) => {
    const target = healthRecords.find((record) => record.employeeId === employeeId);
    if (!target) {
      return;
    }
    if (!window.confirm(`"${target.employeeName}" iş sağlığı kaydını silmek istediğinize emin misiniz?`)) {
      return;
    }
    setHealthRecords((prev) => prev.filter((record) => record.employeeId !== employeeId));
    if (editingHealthEmployeeId === employeeId) {
      resetHealthForm();
    }
  };

  const legalComplianceLabel = (value: LegalComplianceStatus) => {
    const map: Record<LegalComplianceStatus, string> = {
      UYUMLU: 'Uyumlu',
      KISMEN_UYUMLU: 'Kısmen Uyumlu',
      UYUMSUZ: 'Uyumsuz',
      UYGULANAMAZ: 'Uygulanamaz'
    };
    return map[value];
  };

  const legalComplianceClass = (value: LegalComplianceStatus) => {
    if (value === 'UYUMLU') {
      return 'status-badge-closed';
    }
    if (value === 'KISMEN_UYUMLU') {
      return 'status-badge-progress';
    }
    if (value === 'UYUMSUZ') {
      return 'traffic-red';
    }
    return 'status-badge-open';
  };

  const legalRiskLabel = (value: LegalRiskLevel) => {
    const map: Record<LegalRiskLevel, string> = {
      DUSUK: 'Düşük',
      ORTA: 'Orta',
      YUKSEK: 'Yüksek',
      KRITIK: 'Kritik'
    };
    return map[value];
  };

  const legalDocumentKindLabel = (value: LegalDocumentKind) => {
    const map: Record<LegalDocumentKind, string> = {
      MEVZUAT_BELGESI: 'Mevzuat Belgesi',
      UYUMLULUK_KANITI: 'Uyumluluk Kanıtı',
      DENETIM_RAPORU: 'Denetim Raporu',
      IZIN_SERTIFIKA: 'İzin / Sertifika',
      DIGER: 'Diğer'
    };
    return map[value];
  };

  const legalRoleLabel = (value: LegalUserRole) => {
    return legalRoleOptions.find((option) => option.value === value)?.label ?? value;
  };

  const detectFileType = (fileName: string, fallbackType?: string) => {
    if (fallbackType && fallbackType !== 'application/octet-stream') {
      return fallbackType;
    }
    const lower = fileName.toLowerCase();
    if (lower.endsWith('.pdf')) return 'application/pdf';
    if (lower.endsWith('.docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    if (lower.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
    if (lower.endsWith('.png')) return 'image/png';
    return 'application/octet-stream';
  };

  const addLegalAudit = (record: LegalRecord, event: Omit<LegalAuditEvent, 'id'>): LegalRecord => ({
    ...record,
    auditTrail: [{ id: `legal-audit-${Date.now()}-${Math.random()}`, ...event }, ...record.auditTrail]
  });

  const createLegalDocumentsFromFiles = (files: FileList | null, kind: LegalDocumentKind, actor: string): LegalDocument[] => {
    if (!files || files.length === 0) {
      return [];
    }
    return Array.from(files)
      .filter((file) => {
        const allowed = ['application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'image/jpeg', 'image/png'];
        const detectedType = detectFileType(file.name, file.type);
        return allowed.includes(detectedType);
      })
      .map((file, index) => {
        const now = new Date().toISOString();
        const url = URL.createObjectURL(file);
        const detectedType = detectFileType(file.name, file.type);
        return {
          id: `legal-doc-${Date.now()}-${index}`,
          kind,
          fileName: file.name,
          fileType: detectedType,
          fileUrl: url,
          uploadedAt: now,
          uploadedBy: actor,
          versions: [
            {
              version: 1,
              fileName: file.name,
              fileType: detectedType,
              fileUrl: url,
              uploadedAt: now,
              uploadedBy: actor
            }
          ]
        };
      });
  };

  const updateLegalRecord = (recordId: string, updater: (record: LegalRecord) => LegalRecord) => {
    setLegalRecords((prev) => prev.map((record) => (record.id === recordId ? updater(record) : record)));
  };

  const saveLegalRecord = () => {
    if (!legalForm.title.trim() || !legalForm.legalRequirement.trim() || !legalForm.responsiblePerson.trim()) {
      return;
    }

    const now = new Date().toISOString().slice(0, 10);
    const actor = legalRoleLabel(legalUserRole);

    const newRecord: LegalRecord = {
      id: `legal-${Date.now()}`,
      projectId: legalForm.projectId,
      regulationId: nextLegalRegulationId,
      category: legalForm.category,
      title: legalForm.title,
      authority: legalForm.authority,
      department: legalForm.department,
      legalRequirement: legalForm.legalRequirement,
      responsiblePerson: legalForm.responsiblePerson,
      complianceStatus: legalForm.complianceStatus,
      effectiveDate: legalForm.effectiveDate,
      lastReviewDate: legalForm.lastReviewDate,
      nextReviewDate: legalForm.nextReviewDate,
      openActions: Math.max(0, legalForm.openActions),
      riskLevel: legalForm.riskLevel,
      notes: legalForm.notes,
      documents: [],
      createdBy: actor,
      createdAt: now,
      modifiedBy: actor,
      modifiedAt: now,
      reviewedBy: '-',
      reviewedAt: '-',
      exportHistory: [],
      downloadHistory: [],
      auditTrail: [
        {
          id: `legal-audit-${Date.now()}`,
          eventType: 'CREATED',
          actor,
          eventDate: now,
          detail: `Kayıt oluşturuldu (${nextLegalRegulationId}).`
        }
      ]
    };

    setLegalRecords((prev) => [newRecord, ...prev]);
    setSelectedLegalRecordId(newRecord.id);
    setLegalForm((prev) => ({
      ...prev,
      title: '',
      legalRequirement: '',
      responsiblePerson: '',
      openActions: 0,
      notes: ''
    }));
  };

  const uploadLegalDocuments = (recordId: string, kind: LegalDocumentKind, files: FileList | null) => {
    const actor = legalRoleLabel(legalUserRole);
    const newDocuments = createLegalDocumentsFromFiles(files, kind, actor);
    if (newDocuments.length === 0) {
      return;
    }

    updateLegalRecord(recordId, (record) => {
      const now = new Date().toISOString().slice(0, 10);
      const updated = addLegalAudit(
        {
          ...record,
          documents: [...newDocuments, ...record.documents],
          modifiedBy: actor,
          modifiedAt: now
        },
        {
          eventType: 'UPDATED',
          actor,
          eventDate: now,
          detail: `${newDocuments.length} belge yüklendi (${legalDocumentKindLabel(kind)}).`
        }
      );
      return updated;
    });

    if (!selectedLegalDocumentId && newDocuments[0]) {
      setSelectedLegalDocumentId(newDocuments[0].id);
    }
  };

  const replaceLegalDocument = (recordId: string, documentId: string, files: FileList | null) => {
    const file = files?.[0];
    if (!file) {
      return;
    }
    const actor = legalRoleLabel(legalUserRole);
    const nowIso = new Date().toISOString();
    const nowDate = nowIso.slice(0, 10);

    updateLegalRecord(recordId, (record) => {
      const documents = record.documents.map((documentRow) => {
        if (documentRow.id !== documentId) {
          return documentRow;
        }
        const nextVersion = documentRow.versions.length + 1;
        const fileType = detectFileType(file.name, file.type);
        const fileUrl = URL.createObjectURL(file);
        return {
          ...documentRow,
          fileName: file.name,
          fileType,
          fileUrl,
          uploadedAt: nowIso,
          uploadedBy: actor,
          versions: [
            ...documentRow.versions,
            {
              version: nextVersion,
              fileName: file.name,
              fileType,
              fileUrl,
              uploadedAt: nowIso,
              uploadedBy: actor
            }
          ]
        };
      });

      return addLegalAudit(
        {
          ...record,
          documents,
          modifiedBy: actor,
          modifiedAt: nowDate
        },
        {
          eventType: 'REPLACED_DOCUMENT',
          actor,
          eventDate: nowDate,
          detail: 'Belge yeni sürüm ile değiştirildi.'
        }
      );
    });
  };

  const deleteLegalDocument = (recordId: string, documentId: string) => {
    if (!legalCanDeleteDocument) {
      return;
    }

    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    updateLegalRecord(recordId, (record) => {
      const target = record.documents.find((documentRow) => documentRow.id === documentId);
      const documents = record.documents.filter((documentRow) => documentRow.id !== documentId);
      const next = addLegalAudit(
        {
          ...record,
          documents,
          modifiedBy: actor,
          modifiedAt: now
        },
        {
          eventType: 'DELETED_DOCUMENT',
          actor,
          eventDate: now,
          detail: `${target?.fileName ?? 'Belge'} silindi.`
        }
      );
      return next;
    });

    if (selectedLegalDocumentId === documentId) {
      setSelectedLegalDocumentId(null);
    }
  };

  const controlledDocumentStatusLabel = (value: ControlledDocumentStatus) => {
    const map: Record<ControlledDocumentStatus, string> = {
      TASLAK: language === 'ru' ? 'Черновик' : 'Taslak',
      GOZDEN_GECIRMEDE: language === 'ru' ? 'На рассмотрении' : 'Gözden Geçirme Aşamasında',
      ONAYLANDI: language === 'ru' ? 'Утверждено' : 'Onaylandı',
      GECERSIZ: language === 'ru' ? 'Недействительно' : 'Geçersiz (Obsolete)'
    };
    return map[value];
  };

  const controlledDocumentStatusClass = (value: ControlledDocumentStatus) => {
    if (value === 'ONAYLANDI') {
      return 'status-badge-closed';
    }
    if (value === 'GOZDEN_GECIRMEDE') {
      return 'status-badge-progress';
    }
    if (value === 'GECERSIZ') {
      return 'traffic-red';
    }
    return 'status-badge-open';
  };

  function controlledDocumentCurrentRevision(record: ControlledDocumentRecord) {
    const approvedRevisions = record.revisions.filter((revision) => revision.status === 'ONAYLANDI');
    return approvedRevisions[approvedRevisions.length - 1] ?? record.revisions[record.revisions.length - 1] ?? null;
  }

  const controlledDocumentCategoryLabel = (value: string) => value;

  const controlledDocumentTypeLabel = (value: string) => value;

  const updateControlledDocumentRecord = (recordId: string, updater: (record: ControlledDocumentRecord) => ControlledDocumentRecord) => {
    setControlledDocuments((prev) => prev.map((record) => (record.id === recordId ? updater(record) : record)));
  };

  const saveControlledDocument = () => {
    const file = controlledDocumentFileList?.[0];
    if (!file || !controlledDocumentForm.title.trim() || !controlledDocumentForm.preparedBy.trim() || !controlledDocumentForm.reviewedBy.trim() || !controlledDocumentForm.approvedBy.trim()) {
      return;
    }

    const fileType = detectFileType(file.name, file.type);
    const fileUrl = URL.createObjectURL(file);
    const nowIso = new Date().toISOString();
    const revisionNumber = 1;
    const actor = controlledDocumentForm.preparedBy.trim();
    const currentRevision: ControlledDocumentRevision = {
      id: `controlled-doc-rev-${Date.now()}`,
      revisionNumber,
      status: controlledDocumentForm.status,
      fileName: file.name,
      fileType,
      fileUrl,
      uploadedAt: nowIso,
      uploadedBy: actor,
      note: 'İlk kontrollü sürüm'
    };

    const newRecord: ControlledDocumentRecord = {
      id: `controlled-doc-${Date.now()}`,
      projectId: controlledDocumentForm.projectId,
      documentId: `DOC-${new Date().getFullYear()}-${String(controlledDocuments.length + 1).padStart(4, '0')}`,
      title: controlledDocumentForm.title.trim(),
      category: controlledDocumentForm.category,
      documentType: controlledDocumentForm.documentType,
      revisionNumber,
      status: controlledDocumentForm.status,
      effectiveDate: controlledDocumentForm.effectiveDate,
      reviewDate: controlledDocumentForm.reviewDate,
      department: controlledDocumentForm.department,
      preparedBy: controlledDocumentForm.preparedBy.trim(),
      reviewedBy: controlledDocumentForm.reviewedBy.trim(),
      approvedBy: controlledDocumentForm.approvedBy.trim(),
      notes: controlledDocumentForm.notes.trim(),
      fileName: file.name,
      fileType,
      fileUrl,
      uploadedAt: nowIso,
      uploadedBy: actor,
      revisions: [currentRevision]
    };

    setControlledDocuments((prev) => [newRecord, ...prev]);
    setSelectedControlledDocumentId(newRecord.id);
    setControlledDocumentFileList(null);
    setControlledDocumentForm((prev) => ({
      ...prev,
      title: '',
      preparedBy: '',
      reviewedBy: '',
      approvedBy: '',
      notes: ''
    }));
  };

  const addControlledDocumentRevision = (recordId: string) => {
    const file = controlledDocumentRevisionFileList?.[0];
    if (!file) {
      return;
    }

    const actor = controlledDocumentForm.preparedBy.trim() || 'Belge Sorumlusu';
    const nowIso = new Date().toISOString();

    updateControlledDocumentRecord(recordId, (record) => {
      const nextRevisionNumber = record.revisions.length > 0 ? Math.max(...record.revisions.map((revision) => revision.revisionNumber)) + 1 : 1;
      const fileType = detectFileType(file.name, file.type);
      const fileUrl = URL.createObjectURL(file);
      const nextRevision: ControlledDocumentRevision = {
        id: `controlled-doc-rev-${Date.now()}`,
        revisionNumber: nextRevisionNumber,
        status: controlledDocumentRevisionStatus,
        fileName: file.name,
        fileType,
        fileUrl,
        uploadedAt: nowIso,
        uploadedBy: actor,
        note: controlledDocumentRevisionNote.trim() || 'Revizyon eklendi'
      };
      const revisions = [...record.revisions, nextRevision];
      const visibleRevision = controlledDocumentCurrentRevision({ ...record, revisions });

      return {
        ...record,
        revisionNumber: visibleRevision?.revisionNumber ?? nextRevisionNumber,
        status: visibleRevision?.status ?? controlledDocumentRevisionStatus,
        fileName: visibleRevision?.fileName ?? nextRevision.fileName,
        fileType: visibleRevision?.fileType ?? nextRevision.fileType,
        fileUrl: visibleRevision?.fileUrl ?? nextRevision.fileUrl,
        uploadedAt: visibleRevision?.uploadedAt ?? nextRevision.uploadedAt,
        uploadedBy: visibleRevision?.uploadedBy ?? nextRevision.uploadedBy,
        revisions
      };
    });

    setControlledDocumentRevisionFileList(null);
    setControlledDocumentRevisionNote('');
    setControlledDocumentRevisionStatus('GOZDEN_GECIRMEDE');
  };

  const viewControlledDocument = (recordId: string) => {
    setSelectedControlledDocumentId(recordId);
  };

  const downloadControlledDocument = (recordId: string) => {
    const record = controlledDocuments.find((item) => item.id === recordId);
    if (!record) {
      return;
    }
    const currentRevision = controlledDocumentCurrentRevision(record);
    if (!currentRevision) {
      return;
    }

    const link = document.createElement('a');
    link.href = currentRevision.fileUrl;
    link.download = currentRevision.fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const deleteControlledDocument = (recordId: string) => {
    setControlledDocuments((prev) => prev.filter((record) => record.id !== recordId));
    if (selectedControlledDocumentId === recordId) {
      setSelectedControlledDocumentId(null);
    }
  };

  const getCorporateReportLabel = (value: CorporateReportKey) => {
    return corporateReportOptions.find((option) => option.key === value)?.label ?? value;
  };

  const getProjectDisplayName = (projectId: string) => {
    return projectCatalog.find((project) => project.id === projectId)?.name ?? (projectId === 'all' ? 'Tüm Projeler' : projectId);
  };

  const isInReportPeriod = (dateText: string, startDate: string, endDate: string) => {
    const current = new Date(dateText).getTime();
    return current >= new Date(startDate).getTime() && current <= new Date(endDate).getTime();
  };

  const buildCorporateReportSnapshot = (filters: {
    reportTypeKey: CorporateReportKey;
    projectId: string;
    department: string;
    periodStart: string;
    periodEnd: string;
    format: CorporateReportFormat;
  }) => {
    const reportTypeLabel = getCorporateReportLabel(filters.reportTypeKey);
    const projectName = getProjectDisplayName(filters.projectId);
    const departmentLabel = filters.department === 'all' ? 'Tüm Departmanlar' : filters.department;
    const creator = legalRoleLabel(legalUserRole);
    const createdAt = new Date().toISOString();
    const dateRangeLabel = `${filters.periodStart} - ${filters.periodEnd}`;

    const reportRowsByType: Record<string, Array<{ label: string; value: string; note?: string }>> = {
      'executive-dashboard': [
        { label: 'Toplam İş Gücü', value: String(workforceRecords.reduce((sum, record) => sum + record.totalWorkforce, 0)), note: 'Tüm projeler' },
        { label: 'Açık Risk Kaydı', value: String(riskRecords.filter((record) => isInReportPeriod(record.assessmentDate, filters.periodStart, filters.periodEnd)).length) },
        { label: 'Aktif Eğitim', value: String(trainingRecords.filter((record) => isInReportPeriod(record.trainingDate, filters.periodStart, filters.periodEnd)).length) },
        { label: 'Yasal Kayıt', value: String(legalRecords.filter((record) => isInReportPeriod(record.createdAt, filters.periodStart, filters.periodEnd)).length) },
        { label: 'Kontrollü Belge', value: String(controlledDocuments.filter((record) => isInReportPeriod(record.uploadedAt, filters.periodStart, filters.periodEnd)).length) }
      ],
      inspections: moduleData.inspections
        .filter((row) => (filters.projectId === 'all' || row.projectId === filters.projectId) && isInReportPeriod(row.date, filters.periodStart, filters.periodEnd))
        .map((row) => ({ label: row.title, value: `${row.valueA}`, note: row.date })),
      observations: moduleData.observations
        .filter((row) => (filters.projectId === 'all' || row.projectId === filters.projectId) && isInReportPeriod(row.date, filters.periodStart, filters.periodEnd))
        .map((row) => ({ label: row.title, value: `${row.valueA}`, note: row.date })),
      'risk-assessments': riskRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.assessmentDate, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.riskId, value: `${record.residualRiskScore}`, note: record.hazard })),
      incidents: moduleData.incidents
        .filter((row) => (filters.projectId === 'all' || row.projectId === filters.projectId) && isInReportPeriod(row.date, filters.periodStart, filters.periodEnd))
        .map((row) => ({ label: row.title, value: `${row.valueA}`, note: row.date })),
      trainings: trainingRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.trainingDate, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.trainingTitle, value: `${record.certifiedEmployeeCount}/${record.projectEmployeeCount}`, note: record.department })),
      'ppe-stocks': ppeTransactions
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.date, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.itemName, value: `${record.quantity} ${record.unit}`, note: record.category })),
      'equipment-management': equipmentRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.lastInspectionDate, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.equipmentName, value: record.inspectionStatus, note: record.location })),
      employees: workforceRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.date, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: `${record.departmentArea} / ${record.contractor}`, value: `${record.totalWorkforce}`, note: record.date })),
      'emergency-preparedness': emergencyDrillRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && isInReportPeriod(record.drillDate, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.drillName, value: record.drillResult, note: record.emergencyType })),
      capa: manualActions.map((action, index) => ({ label: `DÖF ${index + 1}`, value: action.status, note: action.text })),
      'legal-register': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && (filters.department === 'all' || record.department === filters.department) && isInReportPeriod(record.createdAt, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.regulationId, value: legalComplianceLabel(record.complianceStatus), note: record.title })),
      documents: controlledDocuments
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && (filters.department === 'all' || record.department === filters.department) && isInReportPeriod(record.uploadedAt, filters.periodStart, filters.periodEnd))
        .map((record) => ({ label: record.documentId, value: `v${record.revisionNumber}`, note: record.title })),
      'daily-hse': [
        { label: 'İş Gücü', value: String(workforceRecords.reduce((sum, record) => sum + record.totalWorkforce, 0)) },
        { label: 'Denetim', value: String(moduleData.inspections.length) },
        { label: 'Olay', value: String(moduleData.incidents.length) },
        { label: 'Risk', value: String(riskRecords.length) }
      ],
      'weekly-hse': [
        { label: 'İş Gücü', value: String(workforceRecords.length) },
        { label: 'Eğitim', value: String(trainingRecords.length) },
        { label: 'PTW', value: String(moduleData['permit-to-work'].length) },
        { label: 'Yasal', value: String(legalRecords.length) }
      ],
      'monthly-hse': [
        { label: 'Denetim', value: String(moduleData.inspections.length) },
        { label: 'Gözlem', value: String(moduleData.observations.length) },
        { label: 'Olay', value: String(moduleData.incidents.length) },
        { label: 'Eğitim', value: String(trainingRecords.length) }
      ],
      'quarterly-hse': [
        { label: 'Risk', value: String(riskRecords.length) },
        { label: 'Aksiyon', value: String(manualActions.length) },
        { label: 'Yasal', value: String(legalRecords.length) },
        { label: 'Belge', value: String(controlledDocuments.length) }
      ],
      'annual-hse': [
        { label: 'Toplam Modül', value: String(Object.keys(moduleData).length) },
        { label: 'Kayıt', value: String(Object.values(moduleData).reduce((sum, rows) => sum + rows.length, 0)) },
        { label: 'Yasal', value: String(legalRecords.length) },
        { label: 'Belge', value: String(controlledDocuments.length) }
      ],
      'client-hse': [
        { label: 'Güvenlik', value: dashboardData.summaryKpis.find((metric) => metric.label === 'Compliance')?.value ?? '94%' },
        { label: 'Uyumluluk', value: `${legalSummary.total}` },
        { label: 'Aksiyon', value: `${inspectionSummary.openActions}` }
      ],
      'corporate-hse': [
        { label: 'Denetim', value: String(moduleData.inspections.length) },
        { label: 'Risk', value: String(riskRecords.length) },
        { label: 'Yasal', value: String(legalRecords.length) },
        { label: 'Belge', value: String(controlledDocuments.length) },
        { label: 'Eğitim', value: String(trainingRecords.length) }
      ],
      'employer-warning-letters': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && record.openActions > 0)
        .map((record) => ({ label: `Uyarı - ${record.regulationId}`, value: record.department, note: record.title })),
      'employer-penalty-letters': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && record.complianceStatus === 'UYUMSUZ')
        .map((record) => ({ label: `Ceza - ${record.regulationId}`, value: record.department, note: record.title })),
      'hse-environment-fire-letters': legalRecords
        .filter(
          (record) =>
            (filters.projectId === 'all' || record.projectId === filters.projectId)
            && ['İş Sağlığı ve Güvenliği', 'Çevre', 'Yangın Güvenliği'].includes(record.category)
        )
        .map((record) => ({ label: `${record.category} - ${record.regulationId}`, value: record.authority, note: record.title })),
      'state-inspection-warnings': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && record.openActions > 0 && Boolean(record.authority))
        .map((record) => ({ label: `Denetim Uyarısı - ${record.regulationId}`, value: record.authority, note: `${record.openActions} açık aksiyon` })),
      'state-inspection-letters': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && Boolean(record.authority))
        .map((record) => ({ label: `Denetim Mektubu - ${record.regulationId}`, value: record.authority, note: record.title })),
      'state-inspection-penalties': legalRecords
        .filter((record) => (filters.projectId === 'all' || record.projectId === filters.projectId) && record.complianceStatus === 'UYUMSUZ' && Boolean(record.authority))
        .map((record) => ({ label: `Denetim Cezası - ${record.regulationId}`, value: record.authority, note: record.title }))
    };

    const tableRows = reportRowsByType[filters.reportTypeKey] ?? [];
    const summary = `${reportTypeLabel} için ${projectName} / ${departmentLabel} kapsamındaki ${tableRows.length} veri satırı derlendi.`;
    const recommendations =
      filters.reportTypeKey === 'legal-register'
        ? ['Yaklaşan gözden geçirmeler için otomatik uyarılar takip edilmelidir.', 'Uyumsuz kayıtlar CAPA akışına bağlanmalıdır.']
        : filters.reportTypeKey === 'documents'
          ? ['Onaylı revizyonlar tek kaynak olarak korunmalıdır.', 'Geçersiz belgeler arşive taşınmalıdır.']
          : ['Yüksek risk veya açık aksiyon görülen alanlar için haftalık takip önerilir.', 'Yönetici gözden geçirmesi için kısa özet korunmalıdır.'];

    const reportName = `${reportTypeLabel} - ${filters.periodStart} / ${filters.periodEnd}`;
    const csvContent = [
      ['Alan', 'Değer', 'Not'],
      ['Rapor Türü', reportTypeLabel, ''],
      ['Proje', projectName, ''],
      ['Departman', departmentLabel, ''],
      ['Dönem', dateRangeLabel, ''],
      ...tableRows.map((row) => [row.label, row.value, row.note ?? ''])
    ]
      .map((line) => line.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\n');

    const htmlRows = tableRows
      .map(
        (row) => `
          <tr>
            <td>${row.label}</td>
            <td>${row.value}</td>
            <td>${row.note ?? '-'}</td>
          </tr>
        `
      )
      .join('');

    const html = `
      <!doctype html>
      <html lang="tr">
      <head>
        <meta charset="utf-8" />
        <title>${reportName}</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
          .report-header { display: flex; justify-content: space-between; gap: 12px; align-items: center; border: 1px solid #dbe3ee; border-radius: 10px; padding: 12px; background: linear-gradient(180deg, #f8fbff 0%, #eef5fb 100%); }
          .report-brand { display: flex; align-items: center; gap: 12px; }
          .report-logo { width: 88px; height: 36px; border-radius: 8px; border: 1px solid #94a3b8; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; background: #fff; }
          .report-meta { text-align: right; font-size: 12px; color: #334155; }
          h1 { margin: 0; font-size: 22px; color: #0b2740; }
          .report-section { margin-top: 14px; border: 1px solid #dbe3ee; border-radius: 10px; padding: 14px; page-break-inside: avoid; }
          .report-section h2 { margin: 0 0 10px; font-size: 16px; color: #0b2740; }
          .report-summary { display: grid; gap: 8px; }
          .report-summary p { margin: 0; }
          .report-table { width: 100%; border-collapse: collapse; }
          .report-table th, .report-table td { border: 1px solid #dbe3ee; padding: 8px; text-align: left; font-size: 12px; }
          .report-table th { background: #f8fbff; }
          .report-footer { margin-top: 14px; border-top: 1px dashed #cbd5e1; padding-top: 8px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; }
          @media print {
            .report-section { page-break-after: always; }
            .report-section:last-of-type { page-break-after: auto; }
          }
        </style>
      </head>
      <body>
        <header class="report-header">
          <div class="report-brand">
            <div class="report-logo">HSE LOGO</div>
            <div>
              <h1>${reportName}</h1>
              <div>${projectName} | ${departmentLabel}</div>
            </div>
          </div>
          <div class="report-meta">
            <div>Rapor Dönemi: ${dateRangeLabel}</div>
            <div>Oluşturulma: ${createdAt.slice(0, 10)}</div>
            <div>Oluşturan: ${creator}</div>
          </div>
        </header>
        <section class="report-section">
          <h2>Yönetici Özeti</h2>
          <div class="report-summary">
            <p>${summary}</p>
            <p><strong>Öneriler:</strong> ${recommendations.join(' | ')}</p>
          </div>
        </section>
        <section class="report-section">
          <h2>Tablolar</h2>
          <table class="report-table">
            <thead><tr><th>Alan</th><th>Değer</th><th>Not</th></tr></thead>
            <tbody>${htmlRows || '<tr><td colspan="3">Seçili filtrelerde veri bulunmuyor.</td></tr>'}</tbody>
          </table>
        </section>
        <section class="report-section">
          <h2>Öneriler</h2>
          <ul>${recommendations.map((item) => `<li>${item}</li>`).join('')}</ul>
        </section>
        <footer class="report-footer">
          <span>${reportTypeLabel}</span>
          <span>Sayfa 1</span>
        </footer>
      </body>
      </html>
    `;

    return {
      reportName,
      reportTypeLabel,
      projectName,
      departmentLabel,
      createdAt,
      creator,
      summary,
      tableRows,
      recommendations,
      html,
      csvContent,
      downloadName: filters.format === 'PDF' ? `${reportName}.pdf` : `${reportName}.xls`
    };
  };

  const openCorporateReportPreview = (html: string, autoPrint: boolean) => {
    const previewWindow = window.open('', '_blank', 'width=1280,height=920');
    if (!previewWindow) {
      return;
    }
    previewWindow.document.write(html);
    previewWindow.document.close();
    previewWindow.focus();
    if (autoPrint) {
      previewWindow.print();
    }
  };

  const generateCorporateReport = () => {
    const snapshot = buildCorporateReportSnapshot(reportFilters);
    const nextReport: GeneratedCorporateReport = {
      id: `generated-corporate-report-${Date.now()}`,
      reportName: snapshot.reportName,
      reportTypeKey: reportFilters.reportTypeKey,
      reportTypeLabel: snapshot.reportTypeLabel,
      projectId: reportFilters.projectId,
      projectName: snapshot.projectName,
      department: snapshot.departmentLabel,
      periodStart: reportFilters.periodStart,
      periodEnd: reportFilters.periodEnd,
      format: reportFilters.format,
      createdAt: snapshot.createdAt,
      createdBy: snapshot.creator,
      summary: snapshot.summary,
      tableRows: snapshot.tableRows,
      recommendations: snapshot.recommendations,
      html: snapshot.html,
      csvContent: snapshot.csvContent,
      downloadName: snapshot.downloadName
    };

    setGeneratedCorporateReports((prev) => [nextReport, ...prev]);
    setSelectedCorporateReportId(nextReport.id);

    if (reportFilters.format === 'PDF') {
      openCorporateReportPreview(snapshot.html, true);
    } else {
      const blob = new Blob([`\uFEFF${snapshot.csvContent}`], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = snapshot.downloadName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }
  };

  const viewGeneratedCorporateReport = (report: GeneratedCorporateReport) => {
    openCorporateReportPreview(report.html, false);
  };

  const downloadGeneratedCorporateReport = (report: GeneratedCorporateReport) => {
    if (report.format === 'PDF') {
      openCorporateReportPreview(report.html, true);
      return;
    }

    const blob = new Blob([`\uFEFF${report.csvContent}`], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = report.downloadName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const deleteGeneratedCorporateReport = (reportId: string) => {
    setGeneratedCorporateReports((prev) => prev.filter((report) => report.id !== reportId));
    if (selectedCorporateReportId === reportId) {
      setSelectedCorporateReportId(null);
    }
  };

  const markLegalReviewed = (recordId: string) => {
    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    updateLegalRecord(recordId, (record) =>
      addLegalAudit(
        {
          ...record,
          reviewedBy: actor,
          reviewedAt: now,
          modifiedBy: actor,
          modifiedAt: now,
          lastReviewDate: now
        },
        {
          eventType: 'REVIEWED',
          actor,
          eventDate: now,
          detail: 'Kayıt gözden geçirildi.'
        }
      )
    );
  };

  const downloadLegalDocument = (recordId: string, documentRow: LegalDocument) => {
    if (documentRow.fileUrl === '#') {
      return;
    }

    const link = document.createElement('a');
    link.href = documentRow.fileUrl;
    link.download = documentRow.fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    updateLegalRecord(recordId, (record) =>
      addLegalAudit(
        {
          ...record,
          downloadHistory: [{ fileName: documentRow.fileName, actor, date: now }, ...record.downloadHistory],
          modifiedBy: actor,
          modifiedAt: now
        },
        {
          eventType: 'DOWNLOADED_DOCUMENT',
          actor,
          eventDate: now,
          detail: `${documentRow.fileName} indirildi.`
        }
      )
    );
  };

  const openLegalDocumentFullscreen = (documentRow: LegalDocument) => {
    if (documentRow.fileUrl === '#') {
      return;
    }
    window.open(documentRow.fileUrl, '_blank');
  };

  const openLegalReportWindow = (
    mode: 'single' | 'full' | 'summary',
    options?: { recordId?: string; title?: string; autoPrint?: boolean }
  ) => {
    const actor = legalRoleLabel(legalUserRole);
    const now = new Date();
    const nowText = now.toLocaleDateString('tr-TR');
    const reportRecords =
      mode === 'single' && options?.recordId
        ? legalRecords.filter((record) => record.id === options.recordId)
        : legalScopedRows;

    const reportWindow = window.open('', '_blank', 'width=1280,height=920');
    if (!reportWindow) {
      return;
    }

    const complianceLabel = (status: LegalComplianceStatus) => legalComplianceLabel(status);

    const recordSections = reportRecords
      .map((record) => {
        const projectName = projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId;
        return `
          <section class="legal-report-section">
            <h2>${record.regulationId} - ${record.title}</h2>
            <div class="legal-grid">
              <article><span>Proje</span><strong>${projectName}</strong></article>
              <article><span>Kategori</span><strong>${record.category}</strong></article>
              <article><span>Yetkili Merci</span><strong>${record.authority}</strong></article>
              <article><span>Sorumlu Kişi</span><strong>${record.responsiblePerson}</strong></article>
              <article><span>Uyumluluk Durumu</span><strong>${complianceLabel(record.complianceStatus)}</strong></article>
              <article><span>Risk Seviyesi</span><strong>${legalRiskLabel(record.riskLevel)}</strong></article>
              <article><span>Yürürlük Tarihi</span><strong>${record.effectiveDate}</strong></article>
              <article><span>Sonraki Gözden Geçirme</span><strong>${record.nextReviewDate}</strong></article>
              <article><span>Açık Eylemler</span><strong>${record.openActions}</strong></article>
              <article><span>Oluşturan</span><strong>${record.createdBy}</strong></article>
              <article><span>Oluşturulma Tarihi</span><strong>${record.createdAt}</strong></article>
              <article><span>Gözden Geçiren</span><strong>${record.reviewedBy}</strong></article>
            </div>
            <div class="legal-row"><strong>Yasal Gereklilik:</strong> ${record.legalRequirement}</div>
            <div class="legal-row"><strong>Notlar:</strong> ${record.notes || '-'}</div>
            <div class="legal-row"><strong>Ekli Belge Referansları:</strong> ${record.documents.map((doc) => doc.fileName).join(', ') || '-'}</div>
          </section>
        `;
      })
      .join('');

    const summaryRows = `
      <section class="legal-report-section">
        <h2>Uyumluluk Özet Raporu</h2>
        <div class="legal-grid">
          <article><span>Toplam Mevzuat</span><strong>${legalSummary.total}</strong></article>
          <article><span>Uyumlu</span><strong>${legalSummary.compliant}</strong></article>
          <article><span>Uyumsuz</span><strong>${legalSummary.nonCompliant}</strong></article>
          <article><span>Kısmen Uyumlu / İnceleme</span><strong>${legalSummary.inReview}</strong></article>
          <article><span>Yaklaşan Gözden Geçirme</span><strong>${legalSummary.upcomingReviews}</strong></article>
          <article><span>Açık Yasal Aksiyon</span><strong>${legalSummary.openActions}</strong></article>
        </div>
      </section>
    `;

    reportWindow.document.write(`
      <!doctype html>
      <html lang="tr">
      <head>
        <meta charset="utf-8" />
        <title>${options?.title ?? 'Yasal Envanter Raporu'}</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
          .report-header { border: 1px solid #dbe3ee; border-radius: 10px; padding: 12px; display: flex; justify-content: space-between; align-items: center; background: linear-gradient(170deg, #f8fbff 0%, #edf5ff 100%); }
          .logo-box { width: 90px; height: 36px; border: 1px solid #94a3b8; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; background: #fff; }
          h1 { margin: 0; font-size: 22px; color: #0b2740; }
          .meta { text-align: right; font-size: 12px; color: #334155; }
          .legal-report-section { margin-top: 14px; border: 1px solid #dbe3ee; border-radius: 10px; padding: 12px; page-break-inside: avoid; }
          .legal-report-section h2 { margin: 0 0 8px; font-size: 16px; color: #0b2740; }
          .legal-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; }
          .legal-grid article { border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; }
          .legal-grid span { display: block; color: #64748b; font-size: 12px; }
          .legal-grid strong { display: block; margin-top: 4px; font-size: 13px; color: #0b2740; }
          .legal-row { margin-top: 8px; font-size: 12px; color: #1e293b; }
          .footer { margin-top: 14px; border-top: 1px dashed #cbd5e1; padding-top: 8px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; }
          @media print {
            .legal-report-section { page-break-after: always; }
            .legal-report-section:last-of-type { page-break-after: auto; }
          }
        </style>
      </head>
      <body>
        <header class="report-header">
          <div class="logo-box">HSE LOGO</div>
          <div>
            <h1>${options?.title ?? 'Yasal Envanter Yönetici Raporu'}</h1>
          </div>
          <div class="meta">
            <div>Rapor Tarihi: ${nowText}</div>
            <div>Hazırlayan: ${actor}</div>
            <div>Standart: ISO 45001 / ISO 14001</div>
          </div>
        </header>
        ${mode === 'summary' ? summaryRows : ''}
        ${mode === 'single' || mode === 'full' ? recordSections : ''}
        ${mode === 'full' ? summaryRows : ''}
        <footer class="footer">
          <span>Kurumsal Yasal Uyum Raporu</span>
          <span>Dokuman: LEG-REG-RPT | Rev: 01</span>
        </footer>
      </body>
      </html>
    `);
    reportWindow.document.close();
    reportWindow.focus();
    if (options?.autoPrint) {
      reportWindow.print();
    }
  };

  const exportSingleLegalPdf = () => {
    if (!selectedLegalRecord) {
      return;
    }
    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    openLegalReportWindow('single', {
      recordId: selectedLegalRecord.id,
      title: `Tekil Yasal Kayıt - ${selectedLegalRecord.regulationId}`,
      autoPrint: true
    });
    updateLegalRecord(selectedLegalRecord.id, (record) =>
      addLegalAudit(
        {
          ...record,
          exportHistory: [{ format: 'PDF_SINGLE', actor, date: now }, ...record.exportHistory]
        },
        {
          eventType: 'EXPORTED_PDF_SINGLE',
          actor,
          eventDate: now,
          detail: 'Tekil yasal kayıt PDF dışa aktarıldı.'
        }
      )
    );
  };

  const exportFullLegalPdf = () => {
    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    openLegalReportWindow('full', {
      title: 'Tam Yasal Kayıt Raporu',
      autoPrint: true
    });
    setLegalRecords((prev) =>
      prev.map((record) =>
        addLegalAudit(
          {
            ...record,
            exportHistory: [{ format: 'PDF_FULL', actor, date: now }, ...record.exportHistory]
          },
          {
            eventType: 'EXPORTED_PDF_FULL',
            actor,
            eventDate: now,
            detail: 'Tam yasal kayıt raporu PDF dışa aktarıldı.'
          }
        )
      )
    );
  };

  const exportLegalExcel = () => {
    const headers = [
      'Proje',
      'Mevzuat ID',
      'Mevzuat Başlığı',
      'Kategori',
      'Yetkili Merci',
      'Departman',
      'Uyumluluk Durumu',
      'Sorumlu Kişi',
      'Yürürlük Tarihi',
      'Son Gözden Geçirme',
      'Sonraki Gözden Geçirme',
      'Açık Aksiyonlar',
      'Risk Seviyesi',
      'Ekli Belge Sayısı',
      'Oluşturan',
      'Oluşturulma Tarihi'
    ];
    const rows = legalScopedRows.map((record) => [
      projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId,
      record.regulationId,
      record.title,
      record.category,
      record.authority,
      record.department,
      legalComplianceLabel(record.complianceStatus),
      record.responsiblePerson,
      record.effectiveDate,
      record.lastReviewDate,
      record.nextReviewDate,
      String(record.openActions),
      legalRiskLabel(record.riskLevel),
      String(record.documents.length),
      record.createdBy,
      record.createdAt
    ]);

    const csv = [headers, ...rows]
      .map((line) => line.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `yasal-envanter-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    const actor = legalRoleLabel(legalUserRole);
    const now = new Date().toISOString().slice(0, 10);
    setLegalRecords((prev) =>
      prev.map((record) =>
        addLegalAudit(
          {
            ...record,
            exportHistory: [{ format: 'EXCEL', actor, date: now }, ...record.exportHistory]
          },
          {
            eventType: 'EXPORTED_EXCEL',
            actor,
            eventDate: now,
            detail: 'Yasal kayıt raporu Excel dışa aktarıldı.'
          }
        )
      )
    );
  };

  const printSingleLegalRecord = () => {
    if (!selectedLegalRecord) {
      return;
    }
    openLegalReportWindow('single', {
      recordId: selectedLegalRecord.id,
      title: `Yazdırma - ${selectedLegalRecord.regulationId}`,
      autoPrint: true
    });
  };

  const printFullLegalRegister = () => {
    openLegalReportWindow('full', {
      title: 'Yazdırma - Tam Yasal Kayıt',
      autoPrint: true
    });
  };

  const printLegalComplianceSummary = () => {
    openLegalReportWindow('summary', {
      title: 'Yazdırma - Uyumluluk Özet Raporu',
      autoPrint: true
    });
  };

  const exportHealthRecordsCsv = () => {
    const headers = [
      'Calisan Adi',
      'Calisan ID',
      'Departman',
      'Pozisyon',
      'Kan Grubu',
      'Alerjiler',
      'Kronik Hastalik',
      'Bulasici Hastalik',
      'Ilac Kullanimi',
      'Engellilik Durumu',
      'Ise Uygunluk',
      'Kisitli Calisma',
      'Muayene Tarihi',
      'Sonraki Muayene',
      'Asi Durumu',
      'Notlar'
    ];

    const rows = healthRecords.map((record) => [
      record.employeeName,
      record.employeeId,
      record.department,
      record.position,
      record.bloodGroup,
      record.allergies,
      record.chronicDisease,
      record.communicableDisease,
      record.medication,
      record.disabilityStatus,
      record.fitForWork === 'Yes' ? 'Evet' : 'Hayir',
      record.restrictedWork,
      record.medicalExaminationDate,
      record.nextMedicalExamination,
      record.vaccinationStatus,
      record.remarks
    ]);

    const csv = [headers, ...rows]
      .map((line) => line.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(','))
      .join('\n');

    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `is-sagligi-kayitlari-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportHealthExecutivePdf = () => {
    const {
      totalEmployees,
      fitForWork,
      notFitForWork,
      restrictedDuty,
      chronicCount,
      communicableCount,
      medicalDue,
      vaccinationUpToDate,
      vaccinationRate,
      metricLabels,
      metricValues,
      departmentRows,
      departmentMax
    } = healthAnalytics;

    const metricMax = Math.max(...metricValues, 1);
    const statusDonutTotal = Math.max(fitForWork + notFitForWork + restrictedDuty, 1);
    const fitPct = Math.round((fitForWork / statusDonutTotal) * 100);
    const notFitPct = Math.round((notFitForWork / statusDonutTotal) * 100);
    const restrictedPct = Math.max(0, 100 - fitPct - notFitPct);
    const statusDonutGradient = `conic-gradient(#0f766e 0% ${fitPct}%, #be123c ${fitPct}% ${fitPct + notFitPct}%, #b45309 ${fitPct + notFitPct}% 100%)`;

    const reportWindow = window.open('', '_blank', 'width=1280,height=900');
    if (!reportWindow) {
      return;
    }

    reportWindow.document.write(`
      <!doctype html>
      <html lang="tr">
      <head>
        <meta charset="utf-8" />
        <title>Is Sagligi Yonetici Raporu</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; color: #0f172a; background: #ffffff; }
          .header { border: 1px solid #dbe3ee; border-radius: 12px; padding: 14px; background: linear-gradient(160deg, #f8fbff 0%, #eef6ff 100%); }
          .header-top { display: flex; justify-content: space-between; align-items: center; gap: 10px; }
          .logo { border: 1px solid #94a3b8; border-radius: 10px; width: 90px; height: 36px; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; background: #fff; }
          h1 { margin: 0; font-size: 22px; color: #0b2740; }
          .sub { margin-top: 4px; color: #475569; font-size: 13px; }
          .meta { text-align: right; font-size: 12px; color: #334155; }
          .kpi-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-top: 14px; }
          .kpi-card { border: 1px solid #dbe3ee; border-radius: 10px; padding: 10px; background: #fff; }
          .kpi-card span { display: block; color: #526074; font-size: 12px; }
          .kpi-card strong { display: block; margin-top: 4px; font-size: 20px; color: #0b2740; }
          .section { margin-top: 14px; border: 1px solid #dbe3ee; border-radius: 10px; padding: 12px; page-break-inside: avoid; }
          .section h2 { margin: 0 0 8px; font-size: 16px; color: #0b2740; }
          .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
          .box { border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; }
          .bar-row { display: grid; grid-template-columns: 130px 1fr 36px; gap: 8px; align-items: center; margin: 6px 0; }
          .bar-track { height: 10px; border-radius: 999px; background: #e2e8f0; overflow: hidden; }
          .bar-fill { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #0284c7 0%, #0f172a 100%); }
          .status-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
          .status { border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px; }
          .status strong { font-size: 18px; color: #0b2740; }
          .kpi-bars { border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; margin-top: 10px; }
          .kpi-bars h3 { margin: 0 0 8px; font-size: 13px; color: #0b2740; }
          .kpi-row { display: grid; grid-template-columns: 170px 1fr 38px; gap: 8px; align-items: center; margin: 6px 0; }
          .kpi-track { height: 10px; border-radius: 999px; background: #e2e8f0; overflow: hidden; }
          .kpi-fill { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #1d4ed8 0%, #0f172a 100%); }
          .kpi-viz-grid { display: grid; grid-template-columns: 1.3fr 1fr; gap: 12px; margin-top: 10px; }
          .donut-panel { border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; }
          .donut-panel h3 { margin: 0 0 8px; font-size: 13px; color: #0b2740; }
          .status-donut { width: 132px; height: 132px; border-radius: 50%; margin: 0 auto; position: relative; }
          .status-donut::after { content: ''; position: absolute; inset: 22px; border-radius: 50%; background: #ffffff; border: 1px solid #e2e8f0; }
          .status-donut-total { position: relative; margin-top: -78px; text-align: center; font-size: 20px; font-weight: 700; color: #0b2740; z-index: 1; }
          .status-donut-caption { text-align: center; font-size: 11px; color: #64748b; margin-top: 2px; }
          .status-legend { margin-top: 10px; display: grid; gap: 6px; }
          .status-legend-item { display: grid; grid-template-columns: 12px 1fr auto; gap: 8px; align-items: center; font-size: 12px; }
          .status-dot { width: 10px; height: 10px; border-radius: 50%; }
          table { width: 100%; border-collapse: collapse; }
          th, td { border: 1px solid #e2e8f0; padding: 6px; font-size: 12px; text-align: left; }
          th { background: #f8fafc; }
          .footer { margin-top: 12px; border-top: 1px dashed #cbd5e1; padding-top: 8px; display: flex; justify-content: space-between; font-size: 11px; color: #64748b; }
          @media print {
            .section { page-break-after: avoid; }
          }
        </style>
      </head>
      <body>
        <header class="header">
          <div class="header-top">
            <div class="logo">HSE LOGO</div>
            <div>
              <h1>Is Sagligi Yonetici Ozeti</h1>
              <div class="sub">Ust yonetim degerlendirmesi icin saglik, uygunluk ve risk gorunumu</div>
            </div>
            <div class="meta">
              <div>Rapor Kodu: OH-EXEC-RPT</div>
              <div>Revizyon: 01</div>
              <div>Tarih: ${new Date().toLocaleDateString('tr-TR')}</div>
            </div>
          </div>
          <div class="kpi-grid">
            <article class="kpi-card"><span>Toplam Calisan</span><strong>${totalEmployees}</strong></article>
            <article class="kpi-card"><span>Ise Uygunluk Orani</span><strong>${totalEmployees > 0 ? Math.round((fitForWork / totalEmployees) * 100) : 0}%</strong></article>
            <article class="kpi-card"><span>Periyodik Muayene (30 Gun)</span><strong>${medicalDue}</strong></article>
            <article class="kpi-card"><span>Asi Uyum Orani</span><strong>${vaccinationRate}%</strong></article>
          </div>
        </header>

        <section class="section">
          <h2>Saglik Durum Analizi</h2>
          <div class="grid-2">
            <article class="box">
              <div class="status-grid">
                <div class="status"><span>Ise Uygun</span><strong>${fitForWork}</strong></div>
                <div class="status"><span>Ise Uygun Degil</span><strong>${notFitForWork}</strong></div>
                <div class="status"><span>Kisitli Gorev</span><strong>${restrictedDuty}</strong></div>
              </div>
            </article>
            <article class="box">
              <div class="status-grid">
                <div class="status"><span>Kronik Hastalik</span><strong>${chronicCount}</strong></div>
                <div class="status"><span>Bulasici Hastalik</span><strong>${communicableCount}</strong></div>
                <div class="status"><span>Asi Uyumlu</span><strong>${vaccinationUpToDate}</strong></div>
              </div>
            </article>
          </div>
          <div class="kpi-viz-grid">
            <article class="kpi-bars">
              <h3>Is Sagligi KPI Metrikleri (Grafik Ozeti)</h3>
              ${metricLabels.map((label, index) => `
                <div class="kpi-row">
                  <span>${label}</span>
                  <div class="kpi-track"><div class="kpi-fill" style="width:${Math.max((metricValues[index] / metricMax) * 100, 5)}%"></div></div>
                  <strong>${metricValues[index]}</strong>
                </div>
              `).join('')}
            </article>
            <article class="donut-panel">
              <h3>Saglik Durum Dagilimi</h3>
              <div class="status-donut" style="background:${statusDonutGradient}"></div>
              <div class="status-donut-total">${statusDonutTotal}</div>
              <div class="status-donut-caption">Toplam Durum Kaydi</div>
              <div class="status-legend">
                <div class="status-legend-item"><span class="status-dot" style="background:#0f766e"></span><span>Ise Uygun</span><strong>${fitPct}%</strong></div>
                <div class="status-legend-item"><span class="status-dot" style="background:#be123c"></span><span>Ise Uygun Degil</span><strong>${notFitPct}%</strong></div>
                <div class="status-legend-item"><span class="status-dot" style="background:#b45309"></span><span>Kisitli Gorev</span><strong>${restrictedPct}%</strong></div>
              </div>
            </article>
          </div>
        </section>

        <section class="section">
          <h2>Departman Bazli Saglik Dagilimi</h2>
          <article class="box">
            ${departmentRows.length > 0 ? departmentRows.map((row) => `
              <div class="bar-row">
                <span>${row[0]}</span>
                <div class="bar-track"><div class="bar-fill" style="width:${Math.max((row[1] / departmentMax) * 100, 5)}%"></div></div>
                <strong>${row[1]}</strong>
              </div>
            `).join('') : '<p>Departman verisi bulunmuyor.</p>'}
          </article>
        </section>

        <section class="section">
          <h2>Calisan Saglik Kayit Ozeti</h2>
          <table>
            <thead>
              <tr>
                <th>Calisan</th>
                <th>ID</th>
                <th>Departman</th>
                <th>Pozisyon</th>
                <th>Ise Uygunluk</th>
                <th>Sonraki Muayene</th>
                <th>Asi Durumu</th>
              </tr>
            </thead>
            <tbody>
              ${healthRecords.map((record) => `
                <tr>
                  <td>${record.employeeName}</td>
                  <td>${record.employeeId}</td>
                  <td>${record.department || '-'}</td>
                  <td>${record.position || '-'}</td>
                  <td>${record.fitForWork === 'Yes' ? 'Evet' : 'Hayir'}</td>
                  <td>${record.nextMedicalExamination}</td>
                  <td>${record.vaccinationStatus || '-'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </section>

        <footer class="footer">
          <span>Kurumsal Is Sagligi Raporu</span>
          <span>Dokuman: OH-EXEC-RPT | Rev: 01</span>
        </footer>
      </body>
      </html>
    `);

    reportWindow.document.close();
    reportWindow.focus();
    reportWindow.print();
    reportWindow.close();
  };

  const togglePtwArraySelection = (
    value: string,
    selectedValues: string[],
    setSelectedValues: (value: string[] | ((prev: string[]) => string[])) => void
  ) => {
    if (selectedValues.includes(value)) {
      setSelectedValues(selectedValues.filter((item) => item !== value));
      return;
    }
    setSelectedValues([...selectedValues, value]);
  };

  const updatePtwTeamMember = (index: number, key: keyof PtwTeamMember, value: string) => {
    setPtwEkipListesi((prev) => prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  };
  const addPtwTeamMember = () => {
    setPtwEkipListesi((prev) => [...prev, { adSoyad: '', gorevi: '', firma: '', egitimDurumu: '', imza: '' }]);
  };

  const updatePtwControl = (index: number, key: keyof PtwControlItem, value: string) => {
    setPtwKontroller((prev) => prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  };
  const addPtwControl = () => {
    setPtwKontroller((prev) => [...prev, { kontrol: '', durum: 'Uygun', aciklama: '' }]);
  };

  const updatePtwPrecaution = (index: number, key: keyof PtwPrecautionItem, value: string) => {
    setPtwTedbirler((prev) => prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  };
  const addPtwPrecaution = () => {
    setPtwTedbirler((prev) => [...prev, { tedbir: '', sorumlu: '', termin: new Date().toISOString().slice(0, 10), durum: 'Açık' }]);
  };

  const updatePtwDailyLog = (index: number, key: keyof PtwDailyLog, value: string) => {
    setPtwGunlukKayitlar((prev) => prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  };
  const addPtwDailyLog = () => {
    setPtwGunlukKayitlar((prev) => [
      ...prev,
      { tarih: new Date().toISOString().slice(0, 10), saat: '08:00', calismaBasladi: '', calismaBitti: '', aciklama: '', sorumlu: '' }
    ]);
  };

  const updatePtwTeamChange = (index: number, key: keyof PtwTeamChange, value: string) => {
    setPtwEkipDegisiklikleri((prev) => prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  };
  const addPtwTeamChange = () => {
    setPtwEkipDegisiklikleri((prev) => [
      ...prev,
      { eklenenPersonel: '', ayrilanPersonel: '', tarih: new Date().toISOString().slice(0, 10), onaylayan: '' }
    ]);
  };

  const addPtwAttachments = (tip: string, files: FileList | null) => {
    if (!files || files.length === 0) {
      return;
    }
    const mapped = Array.from(files).map((file) => ({
      type: ptwAttachmentTypeByLabel[tip] ?? 'OTHER',
      tip,
      dosyaAdi: file.name,
      fileType: file.type || 'application/octet-stream',
      fileUrl: '',
      previewUrl: URL.createObjectURL(file)
    }));
    setPtwEkler((prev) => [...prev, ...mapped]);
  };

  const toIso = (dateValue: string, fallback = '00:00') => {
    const trimmed = (dateValue || '').trim();
    if (!trimmed) {
      return new Date().toISOString();
    }
    const hasTime = trimmed.includes('T');
    return hasTime ? trimmed : `${trimmed}T${fallback}:00.000Z`;
  };

  const toDateOnly = (isoLike: string | undefined | null) => {
    if (!isoLike) {
      return '';
    }
    const value = String(isoLike);
    if (value.length >= 10) {
      return value.slice(0, 10);
    }
    return value;
  };

  const mapPtwApiToForm = (row: PtwApiRecord) => {
    setPtwRecordId(row.id);
    setPtwForm((prev) => ({
      ...prev,
      ptwNo: row.permitNo,
      organizasyon: row.organization,
      departman: row.department,
      proje: row.project,
      lokasyon: row.location,
      ptwTuru: row.permitType,
      duzenlenmeTarihi: toDateOnly(row.issueDate),
      gecerlilikBaslangici: toDateOnly(row.validityStart),
      gecerlilikBitisi: toDateOnly(row.validityEnd),
      durum: row.status,
      isiTalepEden: row.requesterName,
      isiVeren: row.issuerName,
      isSorumlusu: row.jobResponsibleName,
      sahaSorumlusu: row.siteResponsibleName,
      hseSorumlusu: row.hseResponsibleName,
      yetkiliOnaylayan: row.authorizedApprover,
      isinAdi: row.jobTitle,
      isinAciklamasi: row.jobDescription,
      calismaAlani: row.workArea,
      yapilacakIs: row.plannedWork,
      calismaKosullari: row.workingConditions,
      baslangicTarihi: toDateOnly(row.workStartDate),
      baslangicSaati: row.workStartTime,
      bitisTarihi: toDateOnly(row.workEndDate),
      bitisSaati: row.workEndTime,
      ozelSartlar: row.specialConditions ?? '',
      ptwHazirlayan: row.preparedBy ?? '',
      hseOnayi: row.hseApprovalBy ?? '',
      projeMuduru: row.projectManagerBy ?? '',
      isverenTemsilcisi: row.employerRepBy ?? '',
      dijitalImza: row.digitalSignature ?? '',
      onayTarihi: toDateOnly(row.approvalDate),
      isTamamlandi: row.workCompleted,
      alanGuvenli: row.areaSafe,
      malzemelerToplandi: row.materialsCollected,
      ptwKapatildi: row.ptwClosed,
      kapatan: row.closedBy ?? '',
      kapanisTarihi: toDateOnly(row.closedAt) || prev.kapanisTarihi,
      kapanisAciklama: row.closureRemarks ?? ''
    }));
    setPtwTehlikeler(row.hazards ?? []);
    setPtwGuvenlikSecimleri(row.safetySystems ?? []);
    setPtwEkipmanSecimleri(row.equipmentUsed ?? []);
    setPtwEkipListesi(row.teamMembers?.length ? row.teamMembers : [{ adSoyad: '', gorevi: '', firma: '', egitimDurumu: '', imza: '' }]);
    setPtwKontroller(row.preWorkChecks?.length ? row.preWorkChecks : [{ kontrol: '', durum: 'Uygun', aciklama: '' }]);
    setPtwTedbirler(row.precautions?.length ? row.precautions : [{ tedbir: '', sorumlu: '', termin: new Date().toISOString().slice(0, 10), durum: 'Açık' }]);
    setPtwGunlukKayitlar(row.dailyLogs?.length ? row.dailyLogs.map((log) => ({ ...log, tarih: toDateOnly(log.tarih) })) : [{ tarih: new Date().toISOString().slice(0, 10), saat: '08:00', calismaBasladi: '', calismaBitti: '', aciklama: '', sorumlu: '' }]);
    setPtwEkipDegisiklikleri(row.teamChanges?.length ? row.teamChanges.map((change) => ({ ...change, tarih: toDateOnly(change.tarih) })) : [{ eklenenPersonel: '', ayrilanPersonel: '', tarih: new Date().toISOString().slice(0, 10), onaylayan: '' }]);
    setPtwEkler(
      (row.attachments ?? []).map((file) => ({
        id: file.id,
        type: file.type,
        tip: ptwAttachmentLabelByType[file.type] ?? 'Прочее',
        dosyaAdi: file.fileName,
        fileType: file.fileType,
        fileUrl: `${apiBaseUrl}${file.fileUrl}`
      }))
    );
    setPtwLastSavedAt(new Date().toISOString());
  };

  const buildPtwPayload = () => ({
    permitNo: ptwForm.ptwNo,
    organization: ptwForm.organizasyon,
    department: ptwForm.departman,
    project: ptwForm.proje,
    location: ptwForm.lokasyon,
    permitType: ptwForm.ptwTuru,
    issueDate: toIso(ptwForm.duzenlenmeTarihi),
    validityStart: toIso(ptwForm.gecerlilikBaslangici),
    validityEnd: toIso(ptwForm.gecerlilikBitisi),
    status: ptwForm.durum,
    requesterName: ptwForm.isiTalepEden,
    issuerName: ptwForm.isiVeren,
    jobResponsibleName: ptwForm.isSorumlusu || ptwForm.isiTalepEden,
    siteResponsibleName: ptwForm.sahaSorumlusu || ptwForm.isiVeren,
    hseResponsibleName: ptwForm.hseSorumlusu,
    authorizedApprover: ptwForm.yetkiliOnaylayan,
    jobTitle: ptwForm.isinAdi,
    jobDescription: ptwForm.isinAciklamasi,
    workArea: ptwForm.calismaAlani,
    plannedWork: ptwForm.yapilacakIs,
    workingConditions: ptwForm.calismaKosullari,
    workStartDate: toIso(ptwForm.baslangicTarihi),
    workStartTime: ptwForm.baslangicSaati,
    workEndDate: toIso(ptwForm.bitisTarihi),
    workEndTime: ptwForm.bitisSaati,
    hazards: ptwTehlikeler,
    safetySystems: ptwGuvenlikSecimleri,
    equipmentUsed: ptwEkipmanSecimleri,
    teamMembers: ptwEkipListesi.filter((row) => row.adSoyad.trim() || row.gorevi.trim() || row.firma.trim() || row.egitimDurumu.trim() || row.imza.trim()),
    preWorkChecks: ptwKontroller.filter((row) => row.kontrol.trim() || row.aciklama.trim()),
    precautions: ptwTedbirler.filter((row) => row.tedbir.trim() || row.sorumlu.trim()),
    specialConditions: ptwForm.ozelSartlar,
    preparedBy: ptwForm.ptwHazirlayan,
    hseApprovalBy: ptwForm.hseOnayi,
    projectManagerBy: ptwForm.projeMuduru,
    employerRepBy: ptwForm.isverenTemsilcisi,
    digitalSignature: ptwForm.dijitalImza,
    approvalDate: ptwForm.onayTarihi ? toIso(ptwForm.onayTarihi) : undefined,
    dailyLogs: ptwGunlukKayitlar.filter((row) => row.sorumlu.trim() || row.aciklama.trim()).map((row) => ({ ...row, tarih: toIso(row.tarih) })),
    teamChanges: ptwEkipDegisiklikleri.filter((row) => row.onaylayan.trim() || row.eklenenPersonel.trim() || row.ayrilanPersonel.trim()).map((row) => ({ ...row, tarih: toIso(row.tarih) })),
    workCompleted: ptwForm.isTamamlandi,
    areaSafe: ptwForm.alanGuvenli,
    materialsCollected: ptwForm.malzemelerToplandi,
    ptwClosed: ptwForm.ptwKapatildi,
    closedBy: ptwForm.kapatan || undefined,
    closedAt: ptwForm.kapanisTarihi ? toIso(ptwForm.kapanisTarihi) : undefined,
    closureRemarks: ptwForm.kapanisAciklama || undefined,
    attachments: ptwEkler.filter((file) => file.fileUrl).map((file) => ({
      type: file.type,
      fileName: file.dosyaAdi,
      fileType: file.fileType,
      fileUrl: file.fileUrl.replace(apiBaseUrl, '')
    }))
  });

  const loadLocalPtwRecords = (): PtwLocalRecord[] => {
    if (typeof window === 'undefined') {
      return [];
    }
    try {
      const raw = window.localStorage.getItem(PTW_LOCAL_RECORDS_STORAGE_KEY);
      if (!raw) {
        return [];
      }
      const parsed = JSON.parse(raw) as unknown;
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed.filter(isPtwLocalRecord);
    } catch {
      return [];
    }
  };

  const saveLocalPtwRecords = (rows: PtwLocalRecord[]) => {
    if (typeof window === 'undefined') {
      return;
    }
    try {
      window.localStorage.setItem(PTW_LOCAL_RECORDS_STORAGE_KEY, JSON.stringify(rows));
    } catch {
      // Ignore local storage write failures.
    }
  };

  const savePtwToLocalFallback = (payload: Record<string, unknown>): string => {
    const id = ptwRecordId ?? `ptw-local-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
    const rows = loadLocalPtwRecords();
    const savedAt = new Date().toISOString();
    const nextRows = [{ id, payload, savedAt }, ...rows.filter((row) => row.id !== id)];
    saveLocalPtwRecords(nextRows);
    setPtwRecordId(id);
    setPtwLastSavedAt(savedAt);
    return id;
  };

  const extractErrorMessage = (error: unknown): string => (error instanceof Error ? error.message : String(error));

  const isPtwNetworkError = (message: string) => {
    const normalized = message.toLowerCase();
    return (
      normalized.includes('failed to fetch')
      || normalized.includes('networkerror')
      || normalized.includes('load failed')
      || normalized.includes('err_timed_out')
      || normalized.includes('err_internet_disconnected')
    );
  };

  const ptwUserFacingError = (rawMessage: string) => {
    if (rawMessage.includes('GUEST_MODE_WRITE_BLOCKED')) {
      return language === 'ru'
        ? 'Гостевой режим не позволяет запись на сервер; используйте локальное сохранение.'
        : 'Misafir modunda sunucuya yazma kapalıdır; yerel kayıt kullanılacaktır.';
    }
    if (isPtwNetworkError(rawMessage)) {
      return language === 'ru'
        ? 'Сетевое соединение недоступно. Запись сохранена локально или действие повторите при восстановлении сети.'
        : 'Ağ bağlantısı kullanılamıyor. Kayıt yerel olarak saklandı veya ağ düzelince işlemi tekrar deneyin.';
    }
    return rawMessage;
  };

  const validatePtwByTab = (tab: PtwTabKey): string[] => {
    const errors: string[] = [];
    if (tab === 'kayit') {
      if (!ptwForm.organizasyon.trim()) errors.push(language === 'ru' ? 'Организация обязательна.' : 'Organizasyon zorunludur.');
      if (!ptwForm.departman.trim()) errors.push(language === 'ru' ? 'Подразделение обязательно.' : 'Departman zorunludur.');
      if (!ptwForm.proje.trim()) errors.push(language === 'ru' ? 'Проект обязателен.' : 'Proje zorunludur.');
      if (!ptwForm.lokasyon.trim()) errors.push(language === 'ru' ? 'Место выполнения работ обязательно.' : 'Lokasyon zorunludur.');
      if (!ptwForm.ptwTuru.trim()) errors.push(language === 'ru' ? 'Тип наряда-допуска обязателен.' : 'PTW türü zorunludur.');
    }
    if (tab === 'sorumlular') {
      if (!ptwForm.isiTalepEden.trim()) errors.push(language === 'ru' ? 'Инициатор работ обязателен.' : 'İşi talep eden zorunludur.');
      if (!ptwForm.isiVeren.trim()) errors.push(language === 'ru' ? 'Выдающий наряд обязателен.' : 'İşi veren zorunludur.');
      if (!ptwForm.hseSorumlusu.trim()) errors.push(language === 'ru' ? 'Ответственный HSE обязателен.' : 'HSE sorumlusu zorunludur.');
      if (!ptwForm.yetkiliOnaylayan.trim()) errors.push(language === 'ru' ? 'Уполномоченный утверждающий обязателен.' : 'Yetkili onaylayan zorunludur.');
    }
    if (tab === 'is-bilgileri') {
      if (!ptwForm.isinAdi.trim()) errors.push(language === 'ru' ? 'Наименование работы обязательно.' : 'İşin adı zorunludur.');
      if (!ptwForm.isinAciklamasi.trim()) errors.push(language === 'ru' ? 'Описание работы обязательно.' : 'İşin açıklaması zorunludur.');
      if (!ptwForm.calismaAlani.trim()) errors.push(language === 'ru' ? 'Зона выполнения работ обязательна.' : 'Çalışma alanı zorunludur.');
      if (!ptwForm.yapilacakIs.trim()) errors.push(language === 'ru' ? 'Планируемая работа обязательна.' : 'Yapılacak iş zorunludur.');
    }
    return errors;
  };

  const syncPtwFromApi = async (id: string) => {
    const row = await apiRequest<PtwApiRecord>(`/api/ptw/${id}`);
    mapPtwApiToForm(row);
  };

  const savePtwCurrentTab = async (tab: PtwTabKey, forceStatus?: PtwStatus): Promise<string | null> => {
    setPtwFeedback(null);
    const validationErrors = validatePtwByTab(tab);
    if (validationErrors.length > 0) {
      setPtwFeedback({ type: 'error', text: validationErrors[0] });
      return null;
    }
    try {
      setPtwSaving(true);
      const payload = buildPtwPayload();
      if (forceStatus) {
        payload.status = forceStatus;
      }
      const result = ptwRecordId
        ? await apiRequest<PtwApiRecord>(`/api/ptw/${ptwRecordId}`, {
            method: 'PUT',
            body: JSON.stringify(payload)
          })
        : await apiRequest<PtwApiRecord>('/api/ptw', {
            method: 'POST',
            body: JSON.stringify(payload)
          });
      mapPtwApiToForm(result);
      await syncPtwFromApi(result.id);
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Сохранено успешно.' : 'Başarıyla kaydedildi.' });
      setPtwLastSavedAt(new Date().toISOString());
      return result.id;
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW save failed:', errorText, error);
      // If backend write is blocked/unavailable, keep PTW workflow functional with local persistence.
      if (
        errorText.includes('GUEST_MODE_WRITE_BLOCKED')
        || isPtwNetworkError(errorText)
      ) {
        const payload = buildPtwPayload() as Record<string, unknown>;
        if (forceStatus) {
          payload.status = forceStatus;
        }
        const localId = savePtwToLocalFallback(payload);
        setPtwFeedback({
          type: 'success',
          text: language === 'ru' ? 'Сохранено локально. Экспортные действия доступны.' : 'Kayıt yerel olarak kaydedildi. Rapor işlemleri kullanılabilir.'
        });
        return localId;
      }
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
      return null;
    } finally {
      setPtwSaving(false);
    }
  };

  const onSavePtwDraft = async () => {
    await savePtwCurrentTab(activePtwTab, 'Taslak');
  };

  const onSubmitPtwApproval = async () => {
    const savedId = await savePtwCurrentTab(activePtwTab);
    const id = savedId ?? ptwRecordId;
    if (!id) {
      return;
    }
    try {
      setPtwSaving(true);
      const row = await apiRequest<PtwApiRecord>(`/api/ptw/${id}/status`, {
        method: 'PATCH',
        body: JSON.stringify({ status: 'Onay Bekliyor' })
      });
      mapPtwApiToForm(row);
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Отправлено на согласование.' : 'Onaya gönderildi.' });
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW approval submit failed:', errorText, error);
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
    } finally {
      setPtwSaving(false);
    }
  };

  const onClosePtw = async () => {
    const validationErrors = validatePtwByTab('kapanis');
    if (validationErrors.length > 0) {
      setPtwFeedback({ type: 'error', text: validationErrors[0] });
      return;
    }
    const savedId = ptwRecordId ?? (await savePtwCurrentTab('kapanis'));
    if (!savedId) {
      return;
    }
    try {
      setPtwSaving(true);
      const row = await apiRequest<PtwApiRecord>(`/api/ptw/${savedId}/close`, {
        method: 'POST',
        body: JSON.stringify({
          closedBy: ptwForm.kapatan || ptwForm.yetkiliOnaylayan || 'system',
          closureRemarks: ptwForm.kapanisAciklama,
          closedAt: toIso(ptwForm.kapanisTarihi),
          workCompleted: ptwForm.isTamamlandi,
          areaSafe: ptwForm.alanGuvenli,
          materialsCollected: ptwForm.malzemelerToplandi,
          ptwClosed: ptwForm.ptwKapatildi
        })
      });
      mapPtwApiToForm(row);
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Наряд-допуск закрыт.' : 'PTW kapatıldı.' });
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW close failed:', errorText, error);
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
    } finally {
      setPtwSaving(false);
    }
  };

  useEffect(() => {
    if (!isAuthorizedViewer) {
      return;
    }

    const loadLatestPtw = async () => {
      try {
        const rows = await apiRequest<PtwApiRecord[]>('/api/ptw');
        if (rows.length > 0) {
          mapPtwApiToForm(rows[0]);
        }
      } catch (error) {
        console.error('PTW initial load failed:', extractErrorMessage(error), error);
        // Keep local form if PTW API is unavailable.
      }
    };

    void loadLatestPtw();
  }, [isAuthorizedViewer]);

  const downloadFileFromResponse = async (url: string, fileName: string) => {
    const response = await fetch(`${apiBaseUrl}${url}`);
    if (!response.ok) {
      throw new Error(await response.text());
    }
    const blob = await response.blob();
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(blobUrl);
  };

  const ensurePtwIdOrFail = () => {
    if (!ptwRecordId) {
      setPtwFeedback({ type: 'error', text: language === 'ru' ? 'Сначала сохраните форму.' : 'Önce formu kaydedin.' });
      return null;
    }
    return ptwRecordId;
  };

  const isLocalPtwId = (id: string) => id.startsWith('ptw-local-');

  const generatePtwPdf = () => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const lm = 15;
    const rm = 195;
    const pageW = rm - lm;
    let y = 20;

    const addPage = () => { doc.addPage(); y = 20; };
    const checkY = (needed = 10) => { if (y + needed > 280) { addPage(); } };

    // Header
    doc.setFillColor(30, 64, 130);
    doc.rect(0, 0, 210, 16, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.text('HSE COMPLIANCE PLATFORM — PERMIT TO WORK', 15, 10);
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.text(`Generated: ${new Date().toLocaleString()}`, rm, 10, { align: 'right' });

    y = 24;
    doc.setTextColor(0, 0, 0);

    // Title
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text(`PERMIT TO WORK — ${ptwForm.ptwNo}`, lm, y);
    y += 7;
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(80, 80, 80);
    doc.text(`Status: ${ptwForm.durum}  |  Type: ${ptwForm.ptwTuru}  |  Issued: ${ptwForm.duzenlenmeTarihi}`, lm, y);
    doc.setTextColor(0, 0, 0);
    y += 5;
    doc.setDrawColor(30, 64, 130);
    doc.setLineWidth(0.5);
    doc.line(lm, y, rm, y);
    y += 6;

    const sectionHeader = (title: string) => {
      checkY(12);
      doc.setFillColor(240, 244, 252);
      doc.rect(lm, y - 4, pageW, 7, 'F');
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(30, 64, 130);
      doc.text(title, lm + 2, y + 0.5);
      doc.setTextColor(0, 0, 0);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      y += 7;
    };

    const fieldRow = (label: string, value: string) => {
      checkY(7);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.text(label + ':', lm + 2, y);
      doc.setFont('helvetica', 'normal');
      const wrapped = doc.splitTextToSize(value || '—', pageW - 55);
      doc.text(wrapped, lm + 52, y);
      y += 5 * wrapped.length + 1;
    };

    const twoCol = (pairs: Array<[string, string]>) => {
      for (let i = 0; i < pairs.length; i += 2) {
        checkY(6);
        const [l1, v1] = pairs[i];
        const [l2, v2] = pairs[i + 1] ?? ['', ''];
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8.5);
        doc.text(l1 + ':', lm + 2, y);
        doc.setFont('helvetica', 'normal');
        doc.text(v1 || '—', lm + 42, y);
        if (l2) {
          doc.setFont('helvetica', 'bold');
          doc.text(l2 + ':', lm + pageW / 2 + 2, y);
          doc.setFont('helvetica', 'normal');
          doc.text(v2 || '—', lm + pageW / 2 + 42, y);
        }
        y += 6;
      }
    };

    // Section 1 — General Info
    sectionHeader('1. GENERAL INFORMATION');
    twoCol([
      ['Organization', ptwForm.organizasyon],
      ['Department', ptwForm.departman],
      ['Project', ptwForm.proje],
      ['Location', ptwForm.lokasyon],
      ['PTW Type', ptwForm.ptwTuru],
      ['Status', ptwForm.durum],
      ['Issue Date', ptwForm.duzenlenmeTarihi],
      ['Valid From', ptwForm.gecerlilikBaslangici],
      ['Valid To', ptwForm.gecerlilikBitisi],
      ['', ''],
    ]);
    y += 2;

    // Section 2 — Responsible Persons
    sectionHeader('2. RESPONSIBLE PERSONS');
    twoCol([
      ['Requester', ptwForm.isiTalepEden],
      ['Issuer', ptwForm.isiVeren],
      ['Job Responsible', ptwForm.isSorumlusu],
      ['Site Responsible', ptwForm.sahaSorumlusu],
      ['HSE Responsible', ptwForm.hseSorumlusu],
      ['Authorized Approver', ptwForm.yetkiliOnaylayan],
    ]);
    y += 2;

    // Section 3 — Work Information
    sectionHeader('3. WORK INFORMATION');
    twoCol([
      ['Work Title', ptwForm.isinAdi],
      ['Work Area', ptwForm.calismaAlani],
      ['Start Date', ptwForm.baslangicTarihi],
      ['Start Time', ptwForm.baslangicSaati],
      ['End Date', ptwForm.bitisTarihi],
      ['End Time', ptwForm.bitisSaati],
    ]);
    fieldRow('Work Description', ptwForm.isinAciklamasi);
    fieldRow('Planned Work', ptwForm.yapilacakIs);
    fieldRow('Working Conditions', ptwForm.calismaKosullari);
    y += 2;

    // Section 4 — Hazards
    if (ptwTehlikeler.length > 0) {
      sectionHeader('4. IDENTIFIED HAZARDS');
      ptwTehlikeler.forEach((h) => {
        checkY(5);
        doc.text(`• ${h}`, lm + 4, y);
        y += 5;
      });
      y += 2;
    }

    // Section 5 — Safety Systems
    if (ptwGuvenlikSecimleri.length > 0) {
      sectionHeader('5. SAFETY SYSTEMS');
      ptwGuvenlikSecimleri.forEach((s) => {
        checkY(5);
        doc.text(`• ${s}`, lm + 4, y);
        y += 5;
      });
      y += 2;
    }

    // Section 6 — Pre-work Checks
    if (ptwKontroller.some((c) => c.kontrol.trim())) {
      sectionHeader('6. PRE-WORK CHECKS');
      checkY(8);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.text('Check Item', lm + 2, y);
      doc.text('Status', lm + 110, y);
      doc.text('Notes', lm + 135, y);
      doc.setFont('helvetica', 'normal');
      y += 4;
      doc.setLineWidth(0.2);
      doc.line(lm, y, rm, y);
      y += 3;
      ptwKontroller.filter((c) => c.kontrol.trim()).forEach((c) => {
        checkY(6);
        const wrapped = doc.splitTextToSize(c.kontrol, 100);
        doc.text(wrapped, lm + 2, y);
        doc.text(c.durum, lm + 110, y);
        const noteWrapped = doc.splitTextToSize(c.aciklama || '—', 50);
        doc.text(noteWrapped, lm + 135, y);
        y += 5 * Math.max(wrapped.length, 1) + 1;
      });
      y += 2;
    }

    // Section 7 — Approvals
    sectionHeader('7. APPROVALS & SIGNATURES');
    twoCol([
      ['Prepared By', ptwForm.ptwHazirlayan],
      ['HSE Approval', ptwForm.hseOnayi],
      ['Project Manager', ptwForm.projeMuduru],
      ['Employer Rep.', ptwForm.isverenTemsilcisi],
      ['Digital Signature', ptwForm.dijitalImza],
      ['Approval Date', ptwForm.onayTarihi],
    ]);
    y += 2;

    // Special Conditions
    if (ptwForm.ozelSartlar.trim()) {
      sectionHeader('8. SPECIAL CONDITIONS');
      fieldRow('Conditions', ptwForm.ozelSartlar);
      y += 2;
    }

    // Footer on each page
    const totalPages = doc.getNumberOfPages();
    for (let p = 1; p <= totalPages; p++) {
      doc.setPage(p);
      doc.setFontSize(7.5);
      doc.setTextColor(120, 120, 120);
      doc.setFont('helvetica', 'normal');
      doc.text(`${ptwForm.ptwNo}  |  HSE Compliance Platform  |  Page ${p} of ${totalPages}`, 105, 292, { align: 'center' });
      doc.setDrawColor(200, 200, 200);
      doc.setLineWidth(0.3);
      doc.line(15, 288, 195, 288);
    }

    doc.save(`${ptwForm.ptwNo}.pdf`);
  };

  const generatePtwDocx = async () => {
    const makeRow = (label: string, value: string): TableRow =>
      new TableRow({
        children: [
          new TableCell({
            width: { size: 35, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20 })] })],
          }),
          new TableCell({
            width: { size: 65, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: value || '—', size: 20 })] })],
          }),
        ],
      });

    const section = (title: string, rows: Array<[string, string]>) => [
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: title, color: '1e4082', bold: true, size: 24 })],
        spacing: { before: 240, after: 120 },
      }),
      new DocxTable({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: rows.map(([l, v]) => makeRow(l, v)),
      }),
    ];

    const bulletList = (title: string, items: string[]) =>
      items.length === 0
        ? []
        : [
            new Paragraph({
              heading: HeadingLevel.HEADING_2,
              children: [new TextRun({ text: title, color: '1e4082', bold: true, size: 24 })],
              spacing: { before: 240, after: 120 },
            }),
            ...items.map(
              (item) =>
                new Paragraph({
                  bullet: { level: 0 },
                  children: [new TextRun({ text: item, size: 20 })],
                })
            ),
          ];

    const doc = new DocxDocument({
      sections: [
        {
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: 'PERMIT TO WORK', bold: true, size: 32, color: '1e4082' }),
                new TextRun({ text: `\n${ptwForm.ptwNo}`, bold: true, size: 28, break: 1 }),
              ],
              spacing: { after: 240 },
              border: {
                bottom: { style: BorderStyle.SINGLE, size: 8, color: '1e4082' },
              },
            }),

            ...section('1. GENERAL INFORMATION', [
              ['Organization', ptwForm.organizasyon],
              ['Department', ptwForm.departman],
              ['Project', ptwForm.proje],
              ['Location', ptwForm.lokasyon],
              ['PTW Type', ptwForm.ptwTuru],
              ['Status', ptwForm.durum],
              ['Issue Date', ptwForm.duzenlenmeTarihi],
              ['Valid From', ptwForm.gecerlilikBaslangici],
              ['Valid To', ptwForm.gecerlilikBitisi],
            ]),

            ...section('2. RESPONSIBLE PERSONS', [
              ['Requester', ptwForm.isiTalepEden],
              ['Issuer', ptwForm.isiVeren],
              ['Job Responsible', ptwForm.isSorumlusu],
              ['Site Responsible', ptwForm.sahaSorumlusu],
              ['HSE Responsible', ptwForm.hseSorumlusu],
              ['Authorized Approver', ptwForm.yetkiliOnaylayan],
            ]),

            ...section('3. WORK INFORMATION', [
              ['Work Title', ptwForm.isinAdi],
              ['Work Area', ptwForm.calismaAlani],
              ['Start Date/Time', `${ptwForm.baslangicTarihi} ${ptwForm.baslangicSaati}`],
              ['End Date/Time', `${ptwForm.bitisTarihi} ${ptwForm.bitisSaati}`],
              ['Work Description', ptwForm.isinAciklamasi],
              ['Planned Work', ptwForm.yapilacakIs],
              ['Working Conditions', ptwForm.calismaKosullari],
            ]),

            ...bulletList('4. IDENTIFIED HAZARDS', ptwTehlikeler),
            ...bulletList('5. SAFETY SYSTEMS', ptwGuvenlikSecimleri),

            ...(ptwKontroller.some((c) => c.kontrol.trim())
              ? [
                  new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '6. PRE-WORK CHECKS', color: '1e4082', bold: true, size: 24 })],
                    spacing: { before: 240, after: 120 },
                  }),
                  new DocxTable({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: [
                      new TableRow({
                        tableHeader: true,
                        children: [
                          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Check Item', bold: true, size: 20 })] })] }),
                          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Status', bold: true, size: 20 })] })] }),
                          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Notes', bold: true, size: 20 })] })] }),
                        ],
                      }),
                      ...ptwKontroller.filter((c) => c.kontrol.trim()).map(
                        (c) =>
                          new TableRow({
                            children: [
                              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: c.kontrol, size: 20 })] })] }),
                              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: c.durum, size: 20 })] })] }),
                              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: c.aciklama || '—', size: 20 })] })] }),
                            ],
                          })
                      ),
                    ],
                  }),
                ]
              : []),

            ...section('7. APPROVALS & SIGNATURES', [
              ['Prepared By', ptwForm.ptwHazirlayan],
              ['HSE Approval', ptwForm.hseOnayi],
              ['Project Manager', ptwForm.projeMuduru],
              ['Employer Rep.', ptwForm.isverenTemsilcisi],
              ['Digital Signature', ptwForm.dijitalImza],
              ['Approval Date', ptwForm.onayTarihi],
            ]),

            ...(ptwForm.ozelSartlar.trim()
              ? [
                  new Paragraph({
                    heading: HeadingLevel.HEADING_2,
                    children: [new TextRun({ text: '8. SPECIAL CONDITIONS', color: '1e4082', bold: true, size: 24 })],
                    spacing: { before: 240, after: 120 },
                  }),
                  new Paragraph({ children: [new TextRun({ text: ptwForm.ozelSartlar, size: 20 })] }),
                ]
              : []),

            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 480 },
              children: [
                new TextRun({ text: `${ptwForm.ptwNo}  |  HSE Compliance Platform  |  Generated: ${new Date().toLocaleString()}`, size: 16, color: '888888' }),
              ],
            }),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${ptwForm.ptwNo}.docx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const onExportPtwPdf = async () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    if (isLocalPtwId(id)) {
      generatePtwPdf();
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'PDF сформирован.' : 'PDF oluşturuldu.' });
      return;
    }
    try {
      await downloadFileFromResponse(`/api/ptw/${id}/export/pdf`, `${ptwForm.ptwNo}.pdf`);
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW PDF export failed — falling back to client PDF:', errorText);
      generatePtwPdf();
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'PDF сформирован локально.' : 'PDF yerel olarak oluşturuldu.' });
    }
  };

  const onExportPtwDocx = async () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    if (isLocalPtwId(id)) {
      await generatePtwDocx();
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Документ Word сформирован.' : 'Word belgesi oluşturuldu.' });
      return;
    }
    try {
      await downloadFileFromResponse(`/api/ptw/${id}/export/docx`, `${ptwForm.ptwNo}.docx`);
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW DOCX export failed — falling back to client DOCX:', errorText);
      await generatePtwDocx();
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Документ Word сформирован локально.' : 'Word belgesi yerel olarak oluşturuldu.' });
    }
  };

  const exportLocalCsv = (fullExport: boolean) => {
    const rows: string[][] = [
      ['PTW No', ptwForm.ptwNo],
      ['Status', ptwForm.durum],
      ['Organization', ptwForm.organizasyon],
      ['Department', ptwForm.departman],
      ['Project', ptwForm.proje],
      ['Location', ptwForm.lokasyon],
      ['Requester', ptwForm.isiTalepEden],
      ['Issuer', ptwForm.isiVeren],
      ['HSE Responsible', ptwForm.hseSorumlusu],
    ];
    if (fullExport) {
      rows.push(
        ['Job Title', ptwForm.isinAdi],
        ['Work Area', ptwForm.calismaAlani],
        ['Work Description', ptwForm.isinAciklamasi],
        ['Start Date', ptwForm.baslangicTarihi],
        ['End Date', ptwForm.bitisTarihi],
        ['Hazards', ptwTehlikeler.join('; ')],
        ['Safety Systems', ptwGuvenlikSecimleri.join('; ')],
        ['Special Conditions', ptwForm.ozelSartlar],
      );
    }
    ptwTedbirler.filter((r) => r.tedbir.trim()).forEach((r, i) => {
      rows.push([`Action ${i + 1}`, `${r.tedbir} | Responsible: ${r.sorumlu} | Due: ${r.termin} | Status: ${r.durum}`]);
    });
    const csvContent = rows.map((r) => r.map((c) => `"${c.replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fullExport ? `${ptwForm.ptwNo}-full.csv` : `${ptwForm.ptwNo}-actions.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const onExportPtwActionCsv = async () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    if (isLocalPtwId(id)) {
      exportLocalCsv(false);
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'CSV действий выгружен.' : 'Aksiyon CSV dışa aktarıldı.' });
      return;
    }
    try {
      await downloadFileFromResponse(`/api/ptw/${id}/export/action-csv`, `${ptwForm.ptwNo}-actions.csv`);
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW action CSV export failed:', errorText, error);
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
    }
  };

  const onExportPtwFullCsv = async () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    if (isLocalPtwId(id)) {
      exportLocalCsv(true);
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Полный CSV выгружен.' : 'Tam CSV dışa aktarıldı.' });
      return;
    }
    try {
      await downloadFileFromResponse(`/api/ptw/${id}/export/full-csv`, `${ptwForm.ptwNo}-full.csv`);
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW full CSV export failed:', errorText, error);
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
    }
  };

  const onPrintPtw = () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    if (isLocalPtwId(id)) {
      generatePtwPdf();
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'PDF для печати сформирован.' : 'Yazdırmak için PDF oluşturuldu.' });
      return;
    }
    window.open(`${apiBaseUrl}/api/ptw/${id}/print`, '_blank');
  };

  const onEmailPtwSummary = async () => {
    const id = ensurePtwIdOrFail();
    if (!id) return;
    const to = ptwEmailTo.trim();
    if (!to) {
      setPtwFeedback({ type: 'error', text: language === 'ru' ? 'Введите email получателя.' : 'Alıcı e-posta adresi girin.' });
      return;
    }
    if (isLocalPtwId(id)) {
      setPtwFeedback({
        type: 'success',
        text: language === 'ru'
          ? 'Локальная запись сохранена; e-mail отправка требует серверного PTW ID.'
          : 'Yerel kayıt hazır; e-posta gönderimi için sunucuya kayıtlı PTW ID gerekir.'
      });
      return;
    }
    try {
      await apiRequest<{ accepted: boolean }>(`/api/ptw/email-summary`, {
        method: 'POST',
        body: JSON.stringify({
          ptwId: id,
          to,
          subject: `${ptwForm.ptwNo} - ${language === 'ru' ? 'Сводный отчет PTW' : 'PTW Özet Raporu'}`
        })
      });
      setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Отчёт отправлен по электронной почте.' : 'Rapor e-posta ile gönderildi.' });
    } catch (error) {
      const errorText = extractErrorMessage(error);
      console.error('PTW email summary failed:', errorText, error);
      setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
    }
  };

  const onDownloadAllAttachments = () => {
    const files = ptwEkler.filter((file) => !!file.fileUrl);
    if (files.length === 0) {
      setPtwFeedback({ type: 'error', text: language === 'ru' ? 'Нет вложений для скачивания.' : 'İndirilecek ek bulunamadı.' });
      return;
    }
    files.forEach((file) => {
      const link = document.createElement('a');
      link.href = file.fileUrl;
      link.download = file.dosyaAdi;
      link.target = '_blank';
      document.body.appendChild(link);
      link.click();
      link.remove();
    });
    setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Вложения загружаются.' : 'Ekler indiriliyor.' });
  };

  const deletePtwAttachment = (index: number) => {
    setPtwEkler((prev) => prev.filter((_, fileIndex) => fileIndex !== index));
  };

  const uploadPtwFiles = async (tip: string, files: FileList | null) => {
    if (!files || files.length === 0) {
      return;
    }
    const uploaded: PtwAttachment[] = [];
    for (const file of Array.from(files)) {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('type', ptwAttachmentTypeByLabel[tip] ?? 'OTHER');
      const response = await fetch(`${apiBaseUrl}/api/uploads/ptw-attachment`, {
        method: 'POST',
        body: formData
      });
      if (!response.ok) {
        throw new Error(await response.text());
      }
      const payload = (await response.json()) as {
        type: PtwAttachmentApiType;
        fileName: string;
        fileType: string;
        fileUrl: string;
      };
      uploaded.push({
        type: payload.type,
        tip: ptwAttachmentLabelByType[payload.type] ?? tip,
        dosyaAdi: payload.fileName,
        fileType: payload.fileType,
        fileUrl: `${apiBaseUrl}${payload.fileUrl}`
      });
    }
    setPtwEkler((prev) => [...prev, ...uploaded]);
  };

  const resetPtw = () => {
    setPtwForm((prev) => ({
      ...prev,
      ptwNo: `PTW-${new Date().getFullYear()}-${String(Math.floor(Math.random() * 9000) + 1000)}`,
      organizasyon: '',
      departman: '',
      lokasyon: '',
      ptwTuru: 'Yüksekte Çalışma',
      durum: 'Taslak',
      isiTalepEden: '',
      isiVeren: '',
      isSorumlusu: '',
      sahaSorumlusu: '',
      hseSorumlusu: '',
      yetkiliOnaylayan: '',
      isinAdi: '',
      isinAciklamasi: '',
      calismaAlani: '',
      yapilacakIs: '',
      calismaKosullari: '',
      ozelSartlar: '',
      ptwHazirlayan: '',
      hseOnayi: '',
      projeMuduru: '',
      isverenTemsilcisi: '',
      dijitalImza: '',
      isTamamlandi: false,
      alanGuvenli: false,
      malzemelerToplandi: false,
      ptwKapatildi: false,
      kapatan: '',
      kapanisAciklama: ''
    }));
    setPtwTehlikeler([]);
    setPtwGuvenlikSecimleri([]);
    setPtwEkipmanSecimleri([]);
    setPtwEkipListesi([{ adSoyad: '', gorevi: '', firma: '', egitimDurumu: '', imza: '' }]);
    setPtwKontroller([
      { kontrol: 'Çalışma alanı bariyerleme kontrolü', durum: 'Uygun', aciklama: '' },
      { kontrol: 'Düşüş önleme sistemi kontrolü', durum: 'Uygun', aciklama: '' },
      { kontrol: 'Ekipman sertifika kontrolü', durum: 'N/A', aciklama: '' }
    ]);
    setPtwTedbirler([{ tedbir: '', sorumlu: '', termin: new Date().toISOString().slice(0, 10), durum: 'Açık' }]);
    setPtwGunlukKayitlar([
      { tarih: new Date().toISOString().slice(0, 10), saat: '08:00', calismaBasladi: '', calismaBitti: '', aciklama: '', sorumlu: '' }
    ]);
    setPtwEkipDegisiklikleri([
      { eklenenPersonel: '', ayrilanPersonel: '', tarih: new Date().toISOString().slice(0, 10), onaylayan: '' }
    ]);
    setPtwEkler([]);
    setPtwRecordId(null);
    setPtwLastSavedAt(null);
    setActivePtwTab('kayit');
  };

  const openDashboardExecutivePdf = () => {
    if (activeModule !== 'dashboard') {
      return;
    }

    const reportWindow = window.open('', '_blank', 'width=1280,height=920');
    if (!reportWindow) {
      return;
    }

    const escapeHtml = (value: string) =>
      value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const periodLabelMap: Record<DashboardPeriod, string> = {
      daily: 'Günlük',
      weekly: 'Haftalık',
      monthly: 'Aylık',
      yearly: 'Yıllık',
      custom: 'Özel Tarih Aralığı'
    };

    const pdfCopy = {
      title:
        language === 'ru'
          ? 'Отчет руководства по показателям HSE'
          : language === 'en'
            ? 'Executive HSE Performance Summary'
            : 'Yönetici HSE Performans Özeti',
      comparisonTitle:
        language === 'ru'
          ? 'Сравнение'
          : language === 'en'
            ? 'Comparison'
            : 'Karşılaştırma',
      strategicKpiSummary:
        language === 'ru'
          ? 'Сводка стратегических KPI'
          : language === 'en'
            ? 'Strategic KPI Summary'
            : 'Stratejik KPI Özeti',
      fullPanelChartSummary:
        language === 'ru'
          ? 'Графический обзор всей панели'
          : language === 'en'
            ? 'Full Panel Chart Summary'
            : 'Tüm Panel Grafik Özeti',
      pageLabel: language === 'ru' ? 'Страница' : language === 'en' ? 'Page' : 'Sayfa',
      footerLabel:
        language === 'ru'
          ? 'Сводка руководства CEO'
          : language === 'en'
            ? 'CEO Executive Summary'
            : 'CEO Yönetici Özeti',
      projectComparison:
        language === 'ru'
          ? 'Сравнение проектов'
          : language === 'en'
            ? 'Project Comparison'
            : 'Proje Karşılaştırma',
      departmentComparison:
        language === 'ru'
          ? 'Сравнение отделов'
          : language === 'en'
            ? 'Department Comparison'
            : 'Departman Karşılaştırma',
      projectActivity:
        language === 'ru'
          ? 'HSE-активность по проектам'
          : language === 'en'
            ? 'HSE Activity by Project'
            : 'Projelere Göre HSE Aktivitesi',
      departmentActivity:
        language === 'ru'
          ? 'HSE-активность по отделам'
          : language === 'en'
            ? 'HSE Activity by Department'
            : 'Departmanlara Göre HSE Aktivitesi',
      projectDistribution:
        language === 'ru'
          ? 'Распределение проектов'
          : language === 'en'
            ? 'Project Distribution'
            : 'Proje Dağılımı',
      departmentDistribution:
        language === 'ru'
          ? 'Распределение отделов'
          : language === 'en'
            ? 'Department Distribution'
            : 'Departman Dağılımı',
      allProjects: localizeText('All Projects', language),
      period:
        language === 'ru'
          ? {
              daily: 'Ежедневно',
              weekly: 'Еженедельно',
              monthly: 'Ежемесячно',
              yearly: 'Ежегодно',
              custom: 'Произвольный период'
            }
          : language === 'en'
            ? {
                daily: 'Daily',
                weekly: 'Weekly',
                monthly: 'Monthly',
                yearly: 'Yearly',
                custom: 'Custom Date Range'
              }
            : {
                daily: 'Günlük',
                weekly: 'Haftalık',
                monthly: 'Aylık',
                yearly: 'Yıllık',
                custom: 'Özel Tarih Aralığı'
              }
    };

    const pdfRuCopy: Record<string, string> = {
      'Güvenlik Performansı': 'Показатели безопасности',
      'Güvenlik Skoru Trendi': 'Тренд показателя безопасности',
      'Genel Güvenlik Endeksi': 'Общий индекс безопасности',
      'Çalışma İzni': 'Разрешение на работы',
      'Çalışma İzni Hacim Trendi': 'Тренд объема разрешений на работы',
      'Çalışma İzni Durum Dağılımı': 'Распределение статусов разрешений на работы',
      'Denetimler': 'Инспекции',
      'Denetim Sayısı Trendi': 'Тренд количества инспекций',
      'Bulgu Durumu': 'Статус замечаний',
      'Olaylar': 'Инциденты',
      'Olay Trendi (Son 7 Dönem)': 'Тренд инцидентов (последние 7 периодов)',
      'Olay Kategori Dağılımı': 'Распределение категорий инцидентов',
      'Eğitim': 'Обучение',
      'Eğitime Katılım Trendi': 'Тренд участия в обучении',
      'Eğitim Tamamlama Oranı': 'Уровень завершения обучения',
      'KKD Yönetimi': 'Управление СИЗ',
      'KKD Teslim Alım Trendi': 'Тренд поступления СИЗ',
      'KKD Stok Dağılımı': 'Распределение запасов СИЗ',
      'Ekipman': 'Оборудование',
      'Ekipman Envanteri': 'Инвентаризация оборудования',
      'Sertifikasyon Uyum Oranı': 'Уровень соответствия сертификации',
      'İş Sağlığı': 'Охрана труда',
      'İşe Uygunluk Trendi': 'Тренд годности к работе',
      'Sağlık Durum Dağılımı': 'Распределение статусов здоровья',
      'Risk Yönetimi': 'Управление рисками',
      'Açık Risk Trendi': 'Тренд открытых рисков',
      'Risk Seviyesi Dağılımı': 'Распределение уровней риска',
      'Çevre': 'Экология',
      'Atık Üretimi Trendi': 'Тренд образования отходов',
      'Geri Dönüşüm Trendi': 'Тренд переработки',
      'Mevzuata Uyum': 'Соответствие законодательству',
      'Uyum Skoru': 'Оценка соответствия',
      'Tetkik ve Uygunsuzluk Görünümü': 'Аудиты и несоответствия',
      'Aktif Çalışma İzni': 'Активные разрешения на работы',
      'Kapatılan Çalışma İzni': 'Закрытые разрешения на работы',
      'Uygunluk': 'Соответствие',
      'Sıcak Çalışma': 'Горячие работы',
      'Kapalı Alan': 'Замкнутое пространство',
      'Yüksekte Çalışma': 'Работа на высоте',
      'Çalışma İzni İhlali': 'Нарушение разрешения на работы',
      'Toplam Denetim': 'Всего инспекций',
      'Açık Bulgu': 'Открытые замечания',
      'Haftalık Performans': 'Недельная эффективность',
      'Kapalı Bulgu': 'Закрытые замечания',
      'Geciken Aksiyonlar': 'Просроченные действия',
      'Pozitif Gözlem': 'Позитивные предписания',
      'Toplam Olay': 'Всего инцидентов',
      'Ramak Kala Olay': 'Почти произошедший инцидент',
      'Haftalık Oryantasyon': 'Еженедельный вводный инструктаж',
      'Toolbox Konuşmaları': 'Инструктажи Toolbox',
      'Süresi Dolan Sertifikalar': 'Просроченные сертификаты',
      'Aylık Oryantasyon': 'Ежемесячный вводный инструктаж',
      'Toplam Oryantasyon': 'Всего вводных инструктажей',
      'Yaklaşan Eğitimler': 'Предстоящие обучения',
      'Stokta': 'На складе',
      'Dağıtılan': 'Выдано',
      'Düşük Stok Kalemleri': 'Позиции с низким запасом',
      'Teslim Alınan': 'Получено',
      'Tüketim Trendi': 'Тренд потребления',
      'Dağılım': 'Распределение',
      'Toplam Ekipman': 'Всего оборудования',
      'Denetimi Gecikmiş': 'Просроченная проверка',
      'Servis Dışı': 'Вне эксплуатации',
      'Mobil': 'Мобильное',
      'Ağır': 'Тяжелое',
      'Kaldırma': 'Подъемное',
      'İşe Uygun': 'Годен к работе',
      'Muayene Zamanı Gelen': 'Требуется осмотр',
      'Aşılama': 'Вакцинация',
      'Toplam Çalışan': 'Всего сотрудников',
      'Kronik Hastalık': 'Хроническое заболевание',
      'Bulaşıcı Hastalık': 'Инфекционное заболевание',
      'Açık Riskler': 'Открытые риски',
      'Yüksek Riskler': 'Высокие риски',
      'Kapanan Riskler': 'Закрытые риски',
      'Orta Riskler': 'Средние риски',
      'Düşük Riskler': 'Низкие риски',
      'Kritik Risk Adedi': 'Количество критических рисков',
      'Oluşan Atık': 'Образовано отходов',
      'Geri Dönüştürülen Atık': 'Переработано отходов',
      'Çevresel Olaylar': 'Экологические инциденты',
      'Su Tüketimi': 'Потребление воды',
      'Elektrik Tüketimi': 'Потребление электроэнергии',
      'Yakıt Tüketimi': 'Потребление топлива',
      'Yasal Uyum': 'Юридическое соответствие',
      'Açık Uygunsuzluk': 'Открытые несоответствия',
      'Açık DÖF': 'Открытые CAPA',
      'İç Tetkikler': 'Внутренние аудиты',
      'Dış Tetkikler': 'Внешние аудиты',
      'Kapanan DÖF': 'Закрытые CAPA',
      'First Aid Cases': 'Случаи первой помощи',
      'Near Miss': 'Почти произошедший инцидент',
      'Total Incidents': 'Всего инцидентов',
      'Unsafe Acts': 'Небезопасные действия',
      'Unsafe Conditions': 'Небезопасные условия',
      'Risk Score': 'Риск-скор',
      'Compliance': 'Соответствие',
      'Safety Performance': 'Показатели безопасности',
      'Permit to Work': 'Разрешение на работы',
      'Inspections': 'Инспекции',
      'Incidents': 'Инциденты',
      'Training': 'Обучение',
      'PPE Management': 'Управление СИЗ',
      'Occupational Health': 'Охрана труда',
      'Risk Management': 'Управление рисками',
      'Environmental': 'Экология',
      'Compliance Score': 'Оценка соответствия',
      'Training Completion': 'Завершение обучения',
      'PPE Stock Health': 'Состояние запасов СИЗ',
      'Environmental Incidents': 'Экологические инциденты',
      'Project Comparison': 'Сравнение проектов',
      'Department Comparison': 'Сравнение отделов',
      'HSE Activity by Project': 'HSE-активность по проектам',
      'HSE Activity by Department': 'HSE-активность по отделам',
      'Project Distribution': 'Распределение проектов',
      'Department Distribution': 'Распределение отделов',
      'Project': 'Проект',
      'Period': 'Период',
      'Date': 'Дата',
      'Page': 'Страница',
      'CEO Executive Summary': 'Сводка руководства CEO'
    };
    const pdfText = (value: string) => (language === 'ru' ? pdfRuCopy[value] ?? localizeText(value, language) : localizeText(value, language));

    const selectedProjectLabel =
      projectFilterOptions.find((project) => project.id === dashboardFilters.projectId)?.name ??
      pdfCopy.allProjects;
    const periodLabel = pdfCopy.period[dashboardFilters.period] ?? pdfCopy.period.monthly;

    const comparisonRows =
      dashboardFilters.comparisonType === 'project'
        ? dashboardData.projectComparison
        : dashboardFilters.comparisonType === 'department'
          ? dashboardData.departmentComparison
          : [];

    const totalPages = 1 + (dashboardFilters.comparisonType === 'none' ? 0 : 1) + dashboardData.sections.length;
    const pageNumberText = (pageNumber: number) => `${pdfCopy.pageLabel} ${pageNumber} / ${totalPages}`;

    const sectionRows = dashboardData.sections.map((section) => {
      const metricSet = section.incidentExecutiveKpis
        ? section.incidentExecutiveKpis.slice(0, 4).map((metric) => `${metric.label}: ${metric.value}`)
        : section.highlights.slice(0, 3).map((metric) => `${metric.label}: ${metric.value}`);

      const detailSet = section.details.slice(0, 3).map((metric) => `${metric.label}: ${metric.value}`);
      return {
        title: section.title,
        trend: section.chartA.values[section.chartA.values.length - 1] ?? 0,
        metrics: metricSet,
        details: detailSet
      };
    });

    const kpiRows = dashboardData.summaryKpis;

    const renderMiniBars = (values: number[]) => {
      const max = Math.max(...values, 1);
      return `
        <div class="mini-bars">
          ${values
            .map((value) => `<span class="mini-bar" style="height:${Math.max(10, (value / max) * 56)}px"></span>`)
            .join('')}
        </div>
      `;
    };

    const renderSectionChartCards = (rows: typeof dashboardData.sections) =>
      rows
        .map(
          (section) => `
            <article class="chart-mini-card">
              <strong>${escapeHtml(pdfText(section.title))}</strong>
              <div class="chart-mini-label">${escapeHtml(pdfText(section.chartA.label))}</div>
              ${renderMiniBars(section.chartA.values)}
            </article>
          `
        )
        .join('');

    const renderExecutiveSection = (section: (typeof dashboardData.sections)[number]) => `
      <section class="block">
        <div class="block-title">${escapeHtml(pdfText(section.title))}</div>
        <div class="chart-mini-grid chart-mini-grid-full">
          <article class="chart-mini-card chart-mini-card-large">
            <strong>${escapeHtml(pdfText(section.chartA.label))}</strong>
            ${renderMiniBars(section.chartA.values)}
          </article>
          ${section.chartB ? `
            <article class="chart-mini-card chart-mini-card-large">
              <strong>${escapeHtml(pdfText(section.chartB.label))}</strong>
              ${renderMiniBars(section.chartB.values)}
            </article>
          ` : ''}
        </div>
        ${section.incidentExecutiveKpis ? `
          <div class="pdf-kpi-bar">
            ${section.incidentExecutiveKpis
              .map(
                (kpi) => `
                  <article class="pdf-kpi-card pdf-kpi-card-compact">
                    <div class="kpi-name">${escapeHtml(pdfText(kpi.label))}</div>
                    <strong class="kpi-val">${escapeHtml(kpi.value)}</strong>
                    ${kpi.note ? `<div class="kpi-trend">${escapeHtml(pdfText(kpi.note))}</div>` : ''}
                  </article>
                `
              )
              .join('')}
          </div>
        ` : `
          <div class="pdf-highlight-grid">
            ${section.highlights
              .map(
                (highlight) => `
                  <article class="pdf-highlight-item">
                    <span>${escapeHtml(pdfText(highlight.label))}</span>
                    <strong>${escapeHtml(highlight.value)}</strong>
                  </article>
                `
              )
              .join('')}
          </div>
        `}
        ${section.details.length > 0 ? `
          <div class="pdf-detail-grid">
            ${section.details
              .map(
                (detail) => `
                  <article class="pdf-detail-item">
                    <span>${escapeHtml(pdfText(detail.label))}</span>
                    <strong>${escapeHtml(detail.value)}</strong>
                    ${detail.note ? `<small>${escapeHtml(pdfText(detail.note))}</small>` : ''}
                  </article>
                `
              )
              .join('')}
          </div>
        ` : ''}
      </section>
    `;

    const renderComparisonPage = () => {
      if (dashboardFilters.comparisonType === 'none' || comparisonRows.length === 0) {
        return '';
      }

      const comparisonTitle =
        dashboardFilters.comparisonType === 'project'
          ? pdfCopy.projectComparison
          : pdfCopy.departmentComparison;

      const comparisonChartTitle =
        dashboardFilters.comparisonType === 'project'
          ? pdfCopy.projectActivity
          : pdfCopy.departmentActivity;

      const distributionTitle =
        dashboardFilters.comparisonType === 'project'
          ? pdfCopy.projectDistribution
          : pdfCopy.departmentDistribution;

      return `
        <section class="page pdf-page">
          <section class="block">
            <div class="block-title">${escapeHtml(comparisonTitle)}</div>
            <div class="executive-charts-grid pdf-comparison-grid">
              <article class="chart-mini-card chart-mini-card-large">
                <strong>${escapeHtml(comparisonChartTitle)}</strong>
                ${renderMiniBars(comparisonRows.map((row) => row.value))}
              </article>
              <article class="chart-mini-card chart-mini-card-large">
                <strong>${escapeHtml(distributionTitle)}</strong>
                ${renderMiniBars(comparisonRows.map((row) => row.value))}
              </article>
            </div>
            <div class="pdf-highlight-grid pdf-highlight-grid-inline">
              ${comparisonRows
                .map(
                  (entry) => `
                    <article class="pdf-highlight-item">
                      <span>${escapeHtml(pdfText(entry.label))}</span>
                      <strong>${escapeHtml(String(entry.value))}</strong>
                    </article>
                  `
                )
                .join('')}
            </div>
          </section>
          <footer class="footer"><span>${escapeHtml(pdfCopy.footerLabel)}</span><span>${escapeHtml(pageNumberText(2))}</span></footer>
        </section>
      `;
    };

    const pdfPages = [
      `
        <section class="page pdf-page">
          <header class="header">
            <h1 class="title">${escapeHtml(pdfCopy.title)}</h1>
            <div class="meta">
              <div><strong>${escapeHtml(language === 'ru' ? 'Период' : language === 'en' ? 'Period' : 'Periyot')}:</strong> ${escapeHtml(periodLabel)}</div>
              <div><strong>${escapeHtml(language === 'ru' ? 'Проект' : language === 'en' ? 'Project' : 'Proje')}:</strong> ${escapeHtml(selectedProjectLabel)}</div>
              <div><strong>${escapeHtml(language === 'ru' ? 'Дата' : language === 'en' ? 'Date' : 'Tarih')}:</strong> ${escapeHtml(new Date().toLocaleDateString(language === 'ru' ? 'ru-RU' : language === 'en' ? 'en-US' : 'tr-TR'))}</div>
            </div>
          </header>

          <section class="block">
            <div class="block-title">${escapeHtml(pdfCopy.strategicKpiSummary)}</div>
            <div class="kpi-grid">
              ${kpiRows
                .map(
                  (kpi) => `
                    <article class="kpi-item">
                      <div class="kpi-name">${escapeHtml(localizeText(kpi.label, language))}</div>
                      <strong class="kpi-val">${escapeHtml(kpi.value)}</strong>
                      <div class="kpi-trend">${escapeHtml(kpi.trend ?? '-')}</div>
                    </article>
                  `
                )
                .join('')}
            </div>
          </section>

          <section class="block">
            <div class="block-title">${escapeHtml(pdfCopy.fullPanelChartSummary)}</div>
            <div class="chart-mini-grid chart-mini-grid-full">
              ${renderSectionChartCards(dashboardData.sections)}
            </div>
          </section>

          <footer class="footer"><span>${escapeHtml(pdfCopy.footerLabel)}</span><span>${escapeHtml(pageNumberText(1))}</span></footer>
        </section>
      `,
      renderComparisonPage(),
      ...dashboardData.sections.map((section, index) => {
        const pageNumber = index + 3;
        return `
          <section class="page pdf-page">
            ${renderExecutiveSection(section)}
            <footer class="footer"><span>${escapeHtml(pdfCopy.footerLabel)}</span><span>${escapeHtml(pageNumberText(pageNumber))}</span></footer>
          </section>
        `;
      })
    ].filter((page) => page.trim().length > 0);

    const renderSectionTableRows = (rows: typeof sectionRows) =>
      rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(localizeText(row.title, language))}</td>
              <td>${escapeHtml(row.metrics.join(' | '))}</td>
              <td>${escapeHtml(row.details.join(' | ') || '-')}</td>
              <td>${escapeHtml(String(row.trend))}</td>
            </tr>
          `
        )
        .join('');

    const html = `
      <!doctype html>
      <html lang="tr">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(pdfCopy.title)}</title>
        <style>
          @page { size: A4 portrait; margin: 8mm 9mm; }
          body { margin: 0; font-family: Arial, sans-serif; color: #0f172a; }
          .pdf-page { break-after: page; page-break-after: always; break-inside: avoid; page-break-inside: avoid; }
          .pdf-page:last-child { break-after: auto; page-break-after: auto; }
          .header { border: 1px solid #cfd9e6; border-radius: 8px; padding: 8px 10px; margin-bottom: 8px; }
          .title { font-size: 16px; font-weight: 700; margin: 0 0 4px; }
          .meta { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 6px; font-size: 10px; color: #334155; }
          .meta div { border: 1px solid #d9e2ee; border-radius: 6px; padding: 5px 6px; }
          .block { border: 1px solid #d4dde9; border-radius: 8px; margin-bottom: 8px; overflow: hidden; }
          .block-title { background: #f3f7fc; border-bottom: 1px solid #d4dde9; padding: 6px 8px; font-size: 11px; font-weight: 700; }
          .chart-mini-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; padding: 8px; }
          .chart-mini-grid-full { grid-template-columns: repeat(2, minmax(0, 1fr)); }
          .chart-mini-card { border: 1px solid #dbe3ee; border-radius: 8px; padding: 8px; }
          .chart-mini-card-large { min-height: 96px; }
          .chart-mini-card strong { display: block; font-size: 10px; margin-bottom: 4px; }
          .chart-mini-label { font-size: 9px; color: #475569; margin-bottom: 6px; }
          .mini-bars { height: 60px; display: flex; align-items: end; gap: 4px; }
          .mini-bar { flex: 1; border-radius: 4px 4px 0 0; background: linear-gradient(180deg, #38bdf8 0%, #0b2740 100%); min-height: 10px; }
          .kpi-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 0; }
          .kpi-item { border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; padding: 6px; min-height: 56px; }
          .kpi-item:nth-child(5n) { border-right: 0; }
          .kpi-name { font-size: 9px; color: #475569; line-height: 1.2; }
          .kpi-val { display: block; margin-top: 3px; font-size: 13px; font-weight: 700; }
          .kpi-trend { font-size: 9px; color: #0f766e; }
          .pdf-kpi-bar, .pdf-highlight-grid, .pdf-detail-grid { display: grid; gap: 6px; padding: 8px; }
          .pdf-kpi-bar { grid-template-columns: repeat(4, minmax(0, 1fr)); }
          .pdf-highlight-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
          .pdf-highlight-grid-inline { grid-template-columns: repeat(4, minmax(0, 1fr)); }
          .pdf-detail-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
          .pdf-kpi-card, .pdf-highlight-item, .pdf-detail-item { border: 1px solid #dbe3ee; border-radius: 8px; padding: 6px 7px; }
          .pdf-kpi-card-compact .kpi-name, .pdf-highlight-item span, .pdf-detail-item span { font-size: 9px; color: #475569; display: block; }
          .pdf-kpi-card-compact .kpi-val { font-size: 12px; }
          .pdf-highlight-item strong, .pdf-detail-item strong { display: block; margin-top: 2px; font-size: 11px; }
          .pdf-detail-item small { display: block; margin-top: 2px; color: #64748b; font-size: 8px; }
          table { width: 100%; border-collapse: collapse; table-layout: fixed; }
          th, td { border: 1px solid #dbe3ee; padding: 5px 6px; font-size: 9px; vertical-align: top; }
          th { background: #f8fbff; text-align: left; font-weight: 700; }
          .footer { margin-top: 10px; border-top: 1px solid #d4dde9; padding-top: 5px; font-size: 9px; color: #475569; display: flex; justify-content: space-between; }
        </style>
      </head>
      <body>
        ${pdfPages.join('')}
      </body>
      </html>
    `;

    reportWindow.document.write(html);
    reportWindow.document.close();
    reportWindow.focus();
    reportWindow.print();
  };

  const headerTitle = activeModule === 'dashboard'
    ? localizeText('Dashboard', language)
    : activeModule === 'incidents' && language === 'ru'
      ? 'Ввод данных и аналитика инцидентов'
      : moduleLabels[activeModule];
  let headerSubtitle = `${moduleLabels[activeModule]} ${t.moduleAnalytics}`;
  if (activeModule === 'dashboard') {
    headerSubtitle = language === 'ru'
      ? 'Производительность проекта и сводка для руководства.'
      : language === 'en'
        ? 'Project performance and executive summary.'
        : 'Proje performansı ve yönetici özeti.';
  } else if (activeModule === 'risk-assessments') {
    headerSubtitle = language === 'ru'
      ? 'Ввод данных и аналитика оценки рисков'
      : language === 'en'
        ? 'Risk assessment data entry and analytics.'
        : 'Risk değerlendirme veri girişi ve analitiği.';
  } else if (activeModule === 'permit-to-work') {
    headerSubtitle = language === 'ru'
      ? 'Ввод данных и аналитика нарядов-допусков'
      : language === 'en'
        ? 'Permit to Work data entry and analytics.'
        : 'Çalışma izni veri girişi ve analitiği.';
  } else if (activeModule === 'incidents') {
    headerSubtitle = language === 'ru'
      ? 'Ввод данных и аналитика инцидентов'
      : language === 'en'
        ? 'Incident data entry and analytics.'
        : 'Olay veri girişi ve analitiği.';
  } else if (activeModule === 'reports') {
    headerSubtitle = language === 'en' ? 'Corporate report center and automated output generation.' : 'Kurumsal raporlama merkezi ve otomatik çıktı üretimi.';
  } else if (activeModule === 'projects') {
    headerSubtitle = language === 'en' ? 'Project master data management.' : 'Proje ana veri yönetimi.';
  } else if (activeModule === 'contractors') {
    headerSubtitle = language === 'en' ? 'Subcontractor master data management.' : 'Alt yüklenici ana veri yönetimi.';
  } else if (activeModule === 'settings') {
    headerSubtitle = language === 'en' ? 'System configuration and governance settings.' : 'Sistem yapılandırması ve yönetişim ayarları.';
  } else if (activeModule === 'export-center') {
    headerSubtitle = '';
  }

  const selectedProject = projectFilter === 'all' ? null : projectCatalog.find((project) => project.id === projectFilter);
  const weatherCity = selectedProject?.city ?? projectCatalog[0]?.city ?? 'Istanbul';

  useEffect(() => {
    let disposed = false;

    const refreshWeather = async () => {
      const defaultLocationLabel = selectedProject
        ? (selectedProject.city || selectedProject.address || selectedProject.name)
        : (language === 'en' ? 'All Projects' : language === 'ru' ? 'Все проекты' : 'Tüm Projeler');

      let coords = cityCoordinates[weatherCity] ?? cityCoordinates.Istanbul;
      let locationLabel = defaultLocationLabel;

      if (selectedProject) {
        const geoQuery = [selectedProject.city, selectedProject.address, selectedProject.country]
          .map((value) => value.trim())
          .filter((value) => value.length > 0)
          .join(' ');

        if (geoQuery) {
          try {
            const geocodeResponse = await fetch(
              `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(geoQuery)}&count=1&language=${language === 'ru' ? 'ru' : language === 'en' ? 'en' : 'tr'}&format=json`
            );
            if (geocodeResponse.ok) {
              const geocodePayload = (await geocodeResponse.json()) as {
                results?: Array<{
                  latitude: number;
                  longitude: number;
                  name: string;
                }>;
              };
              const bestMatch = geocodePayload.results?.[0];
              if (bestMatch) {
                coords = { lat: bestMatch.latitude, lon: bestMatch.longitude };
                locationLabel = bestMatch.name;
              }
            }
          } catch {
            // Ignore geocoding failures and keep static/fallback coordinates.
          }
        }
      }

      try {
        const response = await fetch(
          `https://api.open-meteo.com/v1/forecast?latitude=${coords.lat}&longitude=${coords.lon}&current=temperature_2m,weather_code,wind_speed_10m,wind_direction_10m&timezone=auto`
        );
        if (!response.ok) {
          throw new Error('Weather request failed');
        }

        const payload = (await response.json()) as {
          current?: {
            temperature_2m?: number;
            weather_code?: number;
            wind_speed_10m?: number;
            wind_direction_10m?: number;
          };
        };

        const current = payload.current;
        if (!current || disposed) {
          return;
        }

        const weatherMeta = weatherCodeMeta(current.weather_code ?? 0, language);
        const updatedAt = new Intl.DateTimeFormat(language === 'en' ? 'en-US' : language === 'ru' ? 'ru-RU' : 'tr-TR', {
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        }).format(new Date());

        setLiveWeather({
          icon: weatherMeta.icon,
          temperature: `${Math.round(current.temperature_2m ?? 0)}°C`,
          condition: weatherMeta.text,
          location: locationLabel,
          updatedAt,
          windSpeed: `${Math.round(current.wind_speed_10m ?? 0)} km/s`,
          windDirection: directionFromDegrees(current.wind_direction_10m ?? 0, language)
        });
      } catch {
        if (disposed) {
          return;
        }

        const updatedAt = new Intl.DateTimeFormat(language === 'en' ? 'en-US' : language === 'ru' ? 'ru-RU' : 'tr-TR', {
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        }).format(new Date());

        setLiveWeather((prev) => ({
          ...prev,
          location: locationLabel,
          updatedAt
        }));
      }
    };

    void refreshWeather();
    const timer = setInterval(() => {
      void refreshWeather();
    }, 5 * 60 * 1000);

    return () => {
      disposed = true;
      clearInterval(timer);
    };
  }, [language, selectedProject?.name, selectedProject?.city, selectedProject?.address, selectedProject?.country, weatherCity]);
  const footerYear = new Date().getFullYear();
  const localizedFooterText =
    language === 'ru'
      ? `Дизайн и разработка: Erdem Cetin © ${footerYear} Все права защищены.`
      : language === 'en'
        ? `Design and Development: Erdem Cetin © ${footerYear} All Rights Reserved.`
        : `Tasarım ve Geliştirme: Erdem Cetin © ${footerYear} Tüm Hakları Saklıdır.`;

  if (accessLevel === 'locked') {
    const title =
      language === 'en'
        ? 'HSE Compliance Platform'
        : language === 'ru'
          ? 'Платформа соответствия HSE'
          : 'HSE Uyum Platformu';
    const subtitle =
      language === 'en'
        ? 'Sign in with an authorized email to view company records.'
        : language === 'ru'
          ? 'Войдите с авторизованным email, чтобы видеть корпоративные записи.'
          : 'Kurumsal kayıtları görmek için yetkili e-posta ile giriş yapın.';
    const guestHint =
      language === 'en'
        ? 'Guest mode allows test data entry but does not save records.'
        : language === 'ru'
          ? 'Гостевой режим позволяет тестовый ввод, но не сохраняет записи.'
          : 'Misafir modu test veri girişi sağlar ancak kayıtları saklamaz.';

    return (
      <div className="page">
        <section className="panel" style={{ maxWidth: 620, margin: '40px auto' }}>
          <h2 style={{ marginTop: 0 }}>{title}</h2>
          <p>{subtitle}</p>

          <div className="form-grid" style={{ marginTop: 14 }}>
            <label>
              {language === 'en' ? 'Authorized Email' : language === 'ru' ? 'Авторизованный Email' : 'Yetkili E-posta'}
              <input
                type="email"
                value={loginEmail}
                onChange={(event) => setLoginEmail(event.target.value)}
                placeholder={language === 'ru' ? 'имя@компания.com' : 'name@company.com'}
              />
            </label>

            <label>
              {language === 'en' ? 'Password' : language === 'ru' ? 'Пароль' : 'Parola'}
              <input
                type="password"
                value={loginPassword}
                onChange={(event) => setLoginPassword(event.target.value)}
                placeholder={language === 'ru' ? '********' : '********'}
              />
            </label>
          </div>

          {loginError ? <p className="message error-message">{loginError}</p> : null}

          <div className="actions" style={{ marginTop: 12 }}>
            <button type="button" onClick={handleAuthorizedLogin}>
              {language === 'en' ? 'Sign In' : language === 'ru' ? 'Войти' : 'Giriş Yap'}
            </button>
            <button type="button" className="table-action-button" onClick={continueAsGuest}>
              {language === 'en' ? 'Continue as Guest' : language === 'ru' ? 'Продолжить как гость' : 'Misafir Olarak Devam Et'}
            </button>
          </div>

          <p style={{ marginTop: 10, color: '#526074', fontSize: '0.84rem' }}>{guestHint}</p>
        </section>
      </div>
    );
  }

  return (
    <div className="page shell-layout">
      <aside className="sidebar">
        <p className="brand-subtitle">{localizeText('HSE Compliance Platform', language)}</p>

        <button
          type="button"
          className={`side-link ${activeModule === 'dashboard' ? 'active' : ''}`}
          onClick={() => setActiveModule('dashboard')}
        >
          {moduleLabels.dashboard}
        </button>

        {groupedMenu.map((group) => (
          <div key={group.key} className="group-block">
            <button
              type="button"
              className="group-title"
              onClick={() =>
                setExpandedGroups((prev) => ({
                  ...prev,
                  [group.key]: !prev[group.key]
                }))
              }
            >
              {localizeText(group.title, language)}
            </button>

            {expandedGroups[group.key] ? (
              <div className="group-items">
                {group.modules.map((moduleKey) => (
                  <button
                    key={moduleKey}
                    type="button"
                    className={`side-link ${activeModule === moduleKey ? 'active' : ''}`}
                    onClick={() => setActiveModule(moduleKey)}
                  >
                    {moduleLabels[moduleKey]}
                  </button>
                ))}
              </div>
            ) : null}
          </div>
        ))}

        <small className="credit">{localizeText('Designed and developed by Erdem Cetin', language)}</small>
      </aside>

      <main className="content">
        <header className="hero">
          <div className="executive-header-grid">
            <div className="executive-title-block">
              <h1>{headerTitle}</h1>
              {headerSubtitle ? <p>{headerSubtitle}</p> : null}
            </div>

            <div className="executive-control-stack">
              <label className="executive-field">
                <span className="visually-hidden">{t.project}</span>
                <select value={projectFilter} onChange={(event) => setProjectFilter(event.target.value)}>
                  {projectFilterOptions.map((project) => (
                    <option key={project.id} value={project.id}>
                      {project.name}
                    </option>
                  ))}
                </select>
              </label>

              <div className="executive-widget-stack weather-widget-row">
                <article className="mini-widget weather-widget" aria-label={localizeText('Weather Summary', language)}>
                  <div className="mini-widget-icon">{liveWeather.icon}</div>
                  <div className="mini-widget-body">
                    <strong>{liveWeather.temperature}</strong>
                    <span>{liveWeather.condition}</span>
                    <span>{liveWeather.location}</span>
                  </div>
                </article>

                <article className="mini-widget wind-widget" aria-label={localizeText('Wind Summary', language)}>
                  <div className="mini-widget-icon">💨</div>
                  <div className="mini-widget-body">
                    <strong>{liveWeather.windSpeed}</strong>
                    <span>{liveWeather.windDirection}</span>
                    <small>{localizeText('Wind Direction', language)}</small>
                  </div>
                </article>
              </div>
            </div>
          </div>
        </header>

        {activeModule === 'dashboard' ? (
          <section className="panel">
            <h2>{t.executiveKpi}</h2>
            <div className="dashboard-toolbar enterprise-filter-toolbar" role="toolbar" aria-label="kpi filters">
              <div className="enterprise-filter-primary-row">
                <label>
                  {localizeText('Period', language)}
                  <select
                    value={dashboardFilterDraft.period}
                    onChange={(event) =>
                      setDashboardFilterDraft((prev) => ({ ...prev, period: event.target.value as DashboardPeriod }))
                    }
                  >
                    <option value="daily">{localizeText('Daily', language)}</option>
                    <option value="weekly">{localizeText('Weekly', language)}</option>
                    <option value="monthly">{localizeText('Monthly', language)}</option>
                    <option value="yearly">{localizeText('Yearly', language)}</option>
                    <option value="custom">{localizeText('Custom Date Range', language)}</option>
                  </select>
                </label>

                <label>
                  {localizeText('Project', language)}
                  <select
                    value={dashboardFilterDraft.projectId}
                    onChange={(event) =>
                      setDashboardFilterDraft((prev) => ({ ...prev, projectId: event.target.value }))
                    }
                  >
                    {projectFilterOptions.map((project) => (
                      <option key={project.id} value={project.id}>
                        {project.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  {localizeText('Comparison', language)}
                  <select
                    value={dashboardFilterDraft.comparisonType}
                    onChange={(event) =>
                      setDashboardFilterDraft((prev) => ({ ...prev, comparisonType: event.target.value as ComparisonMode }))
                    }
                  >
                    <option value="none">{localizeText('None', language)}</option>
                    <option value="project">{localizeText('Project Comparison', language)}</option>
                    <option value="department">{localizeText('Department Comparison', language)}</option>
                  </select>
                </label>

                <div className="filter-actions">
                  <button type="button" onClick={() => setDashboardFilters(dashboardFilterDraft)}>
                    {localizeText('Apply Filters', language)}
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const resetState = {
                        period: 'monthly' as DashboardPeriod,
                        customStart: '',
                        customEnd: '',
                        projectId: 'all',
                        comparisonType: 'none' as ComparisonMode,
                        selectedProjects: [],
                        selectedDepartments: []
                      };
                      setDashboardFilterDraft(resetState);
                      setDashboardFilters(resetState);
                    }}
                  >
                    {localizeText('Reset Filters', language)}
                  </button>
                  <button type="button" onClick={openDashboardExecutivePdf}>
                    {localizeText('Executive PDF Output', language)}
                  </button>
                </div>
              </div>

              {dashboardFilterDraft.period === 'custom' ? (
                <>
                  <label>
                    {localizeText('From', language)}
                    <input
                      type="date"
                      value={dashboardFilterDraft.customStart}
                      onChange={(event) =>
                        setDashboardFilterDraft((prev) => ({ ...prev, customStart: event.target.value }))
                      }
                    />
                  </label>
                  <label>
                    {localizeText('To', language)}
                    <input
                      type="date"
                      value={dashboardFilterDraft.customEnd}
                      onChange={(event) =>
                        setDashboardFilterDraft((prev) => ({ ...prev, customEnd: event.target.value }))
                      }
                    />
                  </label>
                </>
              ) : null}

              {dashboardFilterDraft.comparisonType === 'project' ? (
                <label>
                  {localizeText('Select Projects', language)}
                  <select
                    multiple
                    value={dashboardFilterDraft.selectedProjects}
                    onChange={(event) =>
                      setDashboardFilterDraft((prev) => ({
                        ...prev,
                        selectedProjects: Array.from(event.target.selectedOptions).map((option) => option.value)
                      }))
                    }
                  >
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>
                        {project.name}
                      </option>
                    ))}
                  </select>
                </label>
              ) : null}

              {dashboardFilterDraft.comparisonType === 'department' ? (
                <label>
                  {localizeText('Select Departments', language)}
                  <select
                    multiple
                    value={dashboardFilterDraft.selectedDepartments}
                    onChange={(event) =>
                      setDashboardFilterDraft((prev) => ({
                        ...prev,
                        selectedDepartments: Array.from(event.target.selectedOptions).map((option) => option.value)
                      }))
                    }
                  >
                    {departmentOptions.map((department) => (
                      <option key={department} value={department}>
                        {department}
                      </option>
                    ))}
                  </select>
                </label>
              ) : null}

            </div>

            {(() => {
              const safetyPerformanceSection = dashboardData.sections.find((section) => section.key === 'safety-performance');
              const remainingDashboardSections = dashboardData.sections.filter((section) => section.key !== 'safety-performance');

              return (
                <>
                  {safetyPerformanceSection ? (
                    <article className="executive-section-card comparison-card" id="dashboard-safety-performance" tabIndex={-1}>
                      <div className="executive-head">
                        <h3>{localizeText(safetyPerformanceSection.title, language)}</h3>
                      </div>

                      <div className="executive-charts-grid">
                        <div className="chart-card">
                          <strong>{localizeText(safetyPerformanceSection.chartA.label, language)}</strong>
                          <DashboardChart
                            type={safetyPerformanceSection.chartA.type}
                            values={safetyPerformanceSection.chartA.values}
                            themeName={getSectionChartTheme(safetyPerformanceSection.key)}
                            xLabels={dashboardXAxisLabels(safetyPerformanceSection.chartA.values.length)}
                          />
                        </div>
                        {safetyPerformanceSection.chartB ? (
                          <div className="chart-card">
                            <strong>{localizeText(safetyPerformanceSection.chartB.label, language)}</strong>
                            <DashboardChart
                              type={safetyPerformanceSection.chartB.type}
                              values={safetyPerformanceSection.chartB.values}
                              themeName={getSectionChartTheme(safetyPerformanceSection.key)}
                              xLabels={dashboardXAxisLabels(safetyPerformanceSection.chartB.values.length)}
                            />
                          </div>
                        ) : null}
                      </div>

                      <div className="incident-executive-kpi-bar" aria-label={localizeText('Incident Executive KPI Summary', language)}>
                        {safetyPerformanceSection.incidentExecutiveKpis?.map((kpi) => (
                          <article key={kpi.label} className="incident-executive-kpi-card">
                            <div className="incident-kpi-header">
                              <span className={`incident-kpi-status incident-kpi-status-${kpi.status}`} aria-hidden="true" />
                              <span className="incident-kpi-label">{localizeText(kpi.label, language)}</span>
                            </div>
                            {kpiTermHint(localizeText(kpi.label, language), language) ? (
                              <span className="kpi-term-hint">{kpiTermHint(localizeText(kpi.label, language), language)}</span>
                            ) : null}
                            <strong className="incident-kpi-value">{kpi.value}</strong>
                            {kpi.note ? <small className="incident-kpi-note">{localizeText(kpi.note, language)}</small> : null}
                          </article>
                        ))}
                      </div>

                      {safetyPerformanceSection.details.length > 0 ? (
                        <div className="executive-detail-grid">
                          {safetyPerformanceSection.details.map((detail) => (
                            <div key={detail.label} className="executive-detail-item">
                              <span>{localizeText(detail.label, language)}</span>
                              {kpiTermHint(localizeText(detail.label, language), language) ? (
                                <span className="kpi-term-hint">{kpiTermHint(localizeText(detail.label, language), language)}</span>
                              ) : null}
                              <strong>{detail.value}</strong>
                              {detail.note ? <small className="executive-detail-note">{localizeText(detail.note, language)}</small> : null}
                            </div>
                          ))}
                        </div>
                      ) : null}
                    </article>
                  ) : null}

                  <div className="summary-kpi-grid">
                    {dashboardData.summaryKpis.map((metric) => (
                      <article className="summary-kpi-card" key={metric.label}>
                        <span className="kpi-label">{localizeText(metric.label, language)}</span>
                        {kpiTermHint(localizeText(metric.label, language), language) ? (
                          <span className="kpi-term-hint">{kpiTermHint(localizeText(metric.label, language), language)}</span>
                        ) : null}
                        <strong className="kpi-value">{metric.value}</strong>
                        <small>{metric.trend}</small>
                      </article>
                    ))}
                  </div>

            {dashboardFilters.comparisonType !== 'none' ? (
              <article className="executive-section-card comparison-card">
                <div className="executive-head">
                  <h3>
                    {dashboardFilters.comparisonType === 'project'
                      ? localizeText('Project Comparison', language)
                      : localizeText('Department Comparison', language)}
                  </h3>
                </div>
                <div className="executive-charts-grid">
                  <div className="chart-card">
                    <strong>
                      {dashboardFilters.comparisonType === 'project'
                        ? localizeText('HSE Activity by Project', language)
                        : localizeText('HSE Activity by Department', language)}
                    </strong>
                    <DashboardChart
                      type="bar"
                      values={
                        (dashboardFilters.comparisonType === 'project' ? dashboardData.projectComparison : dashboardData.departmentComparison).map(
                          (entry) => entry.value
                        )
                      }
                      themeName="operations"
                      xLabels={dashboardXAxisLabels(
                        (dashboardFilters.comparisonType === 'project' ? dashboardData.projectComparison : dashboardData.departmentComparison).length
                      )}
                    />
                  </div>
                  <div className="chart-card">
                    <strong>
                      {dashboardFilters.comparisonType === 'project'
                        ? localizeText('Project Distribution', language)
                        : localizeText('Department Distribution', language)}
                    </strong>
                    <DashboardChart
                      type="donut"
                      values={
                        (dashboardFilters.comparisonType === 'project' ? dashboardData.projectComparison : dashboardData.departmentComparison).map(
                          (entry) => entry.value
                        )
                      }
                      themeName="operations"
                    />
                  </div>
                </div>
                <div className="executive-highlight-row">
                  {(dashboardFilters.comparisonType === 'project' ? dashboardData.projectComparison : dashboardData.departmentComparison).map((entry) => (
                    <div key={entry.label} className="executive-highlight-item">
                      <span>{localizeText(entry.label, language)}</span>
                      <strong>{entry.value}</strong>
                    </div>
                  ))}
                </div>
              </article>
            ) : null}

                  <div className="executive-sections-grid">
                    {remainingDashboardSections.map((section) => (
                <article className="executive-section-card" key={section.key}>
                  <div className="executive-head">
                    <h3>{localizeText(section.title, language)}</h3>
                  </div>

                  <div className="executive-charts-grid">
                    <div className="chart-card">
                      <strong>{localizeText(section.chartA.label, language)}</strong>
                      <DashboardChart
                        type={section.chartA.type}
                        values={section.chartA.values}
                        themeName={getSectionChartTheme(section.key)}
                        xLabels={dashboardXAxisLabels(section.chartA.values.length)}
                      />
                    </div>
                    {section.chartB ? (
                      <div className="chart-card">
                        <strong>{localizeText(section.chartB.label, language)}</strong>
                        <DashboardChart
                          type={section.chartB.type}
                          values={section.chartB.values}
                          themeName={getSectionChartTheme(section.key)}
                          xLabels={dashboardXAxisLabels(section.chartB.values.length)}
                        />
                      </div>
                    ) : null}
                  </div>

                  {section.incidentExecutiveKpis ? (
                    <div className="incident-executive-kpi-bar" aria-label={localizeText('Incident Executive KPI Summary', language)}>
                      {section.incidentExecutiveKpis.map((kpi) => (
                        <article key={kpi.label} className="incident-executive-kpi-card">
                          <div className="incident-kpi-header">
                            <span className={`incident-kpi-status incident-kpi-status-${kpi.status}`} aria-hidden="true" />
                            <span className="incident-kpi-label">{localizeText(kpi.label, language)}</span>
                          </div>
                          {kpiTermHint(localizeText(kpi.label, language), language) ? (
                            <span className="kpi-term-hint">{kpiTermHint(localizeText(kpi.label, language), language)}</span>
                          ) : null}
                          <strong className="incident-kpi-value">{kpi.value}</strong>
                          {kpi.note ? <small className="incident-kpi-note">{localizeText(kpi.note, language)}</small> : null}
                        </article>
                      ))}
                    </div>
                  ) : (
                    <div className="executive-highlight-row">
                      {section.highlights.map((highlight) => (
                        <div key={highlight.label} className="executive-highlight-item">
                          <span>{localizeText(highlight.label, language)}</span>
                          {kpiTermHint(localizeText(highlight.label, language), language) ? (
                            <span className="kpi-term-hint">{kpiTermHint(localizeText(highlight.label, language), language)}</span>
                          ) : null}
                          <strong>{highlight.value}</strong>
                        </div>
                      ))}
                    </div>
                  )}

                  {section.details.length > 0 ? (
                    <div className="executive-detail-grid">
                      {section.details.map((detail) => (
                        <div key={detail.label} className="executive-detail-item">
                          <span>{localizeText(detail.label, language)}</span>
                          {kpiTermHint(localizeText(detail.label, language), language) ? (
                            <span className="kpi-term-hint">{kpiTermHint(localizeText(detail.label, language), language)}</span>
                          ) : null}
                          <strong>{detail.value}</strong>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </article>
                    ))}
                  </div>
                </>
              );
            })()}
          </section>
        ) : null}

        {activeModule === 'legal-register' ? (
          <div className="legal-register-layout">
            <section className="panel legal-panel">
              <h2>Yasal Mevzuat Kaydı Yönetici Özeti</h2>
              <div className="dashboard-toolbar legal-toolbar">
                <label>
                  Rol
                  <select value={legalUserRole} onChange={(event) => setLegalUserRole(event.target.value as LegalUserRole)}>
                    {legalRoleOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Proje
                  <select value={legalFilters.projectId} onChange={(event) => setLegalFilters((prev) => ({ ...prev, projectId: event.target.value }))}>
                    <option value="all">Tüm Projeler</option>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Kategori
                  <select value={legalFilters.category} onChange={(event) => setLegalFilters((prev) => ({ ...prev, category: event.target.value }))}>
                    <option value="all">Tümü</option>
                    {legalCategoryOptions.map((category) => (
                      <option key={category} value={category}>{category}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Yetkili Merci
                  <select value={legalFilters.authority} onChange={(event) => setLegalFilters((prev) => ({ ...prev, authority: event.target.value }))}>
                    <option value="all">Tümü</option>
                    {legalAuthorityOptions.map((authority) => (
                      <option key={authority} value={authority}>{authority}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Uyumluluk
                  <select
                    value={legalFilters.complianceStatus}
                    onChange={(event) => setLegalFilters((prev) => ({ ...prev, complianceStatus: event.target.value as 'ALL' | LegalComplianceStatus }))}
                  >
                    <option value="ALL">Tümü</option>
                    <option value="UYUMLU">Uyumlu</option>
                    <option value="KISMEN_UYUMLU">Kısmen Uyumlu</option>
                    <option value="UYUMSUZ">Uyumsuz</option>
                    <option value="UYGULANAMAZ">Uygulanamaz</option>
                  </select>
                </label>
                <label>
                  Ara
                  <input
                    placeholder="Başlık, mevzuat no, sorumlu"
                    value={legalFilters.keyword}
                    onChange={(event) => setLegalFilters((prev) => ({ ...prev, keyword: event.target.value }))}
                  />
                </label>
              </div>

              <div className="legal-kpi-grid">
                <article className="equipment-kpi-card"><span>Toplam Mevzuat Sayısı</span><strong>{legalSummary.total}</strong></article>
                <article className="equipment-kpi-card"><span>Uyumlu Mevzuatlar</span><strong>{legalSummary.compliant}</strong></article>
                <article className="equipment-kpi-card"><span>Uyumsuz Mevzuatlar</span><strong>{legalSummary.nonCompliant}</strong></article>
                <article className="equipment-kpi-card"><span>Gözden Geçirme Aşamasındaki</span><strong>{legalSummary.inReview}</strong></article>
                <article className="equipment-kpi-card"><span>Yaklaşan Gözden Geçirmeler</span><strong>{legalSummary.upcomingReviews}</strong></article>
                <article className="equipment-kpi-card"><span>Açık Yasal Aksiyonlar</span><strong>{legalSummary.openActions}</strong></article>
              </div>

              <div className="executive-charts-grid legal-chart-grid">
                <article className="chart-card">
                  <strong>Uyumluluk Durumu</strong>
                  <DashboardChart type="donut" values={legalComplianceChart} themeName="compliance" />
                  <div className="health-legend-grid legal-compliance-legend">
                    <span>Uyumlu: {legalComplianceChart[0]}</span>
                    <span>Kısmen Uyumlu: {legalComplianceChart[1]}</span>
                    <span>Uyumsuz: {legalComplianceChart[2]}</span>
                  </div>
                </article>
                <article className="chart-card">
                  <strong>Kategoriye Göre Mevzuatlar</strong>
                  <div className="workforce-horizontal-chart">
                    {legalCategoryBars.map(([label, value]) => {
                      const max = Math.max(...legalCategoryBars.map((row) => row[1]), 1);
                      return (
                        <div className="workforce-horizontal-row" key={label}>
                          <span>{label}</span>
                          <div className="workforce-horizontal-track">
                            <div className="workforce-horizontal-fill" style={{ width: `${Math.max((value / max) * 100, 4)}%` }} />
                          </div>
                          <strong>{value}</strong>
                        </div>
                      );
                    })}
                  </div>
                </article>
                <article className="chart-card">
                  <strong>Yaklaşan Yasal Gözden Geçirmeler</strong>
                  <DashboardChart type="line" values={legalReviewTrend.values} xLabels={legalReviewTrend.labels} themeName="compliance" />
                </article>
                <article className="chart-card">
                  <strong>Projeye Göre Açık Yasal Aksiyonlar</strong>
                  <DashboardChart type="bar" values={legalOpenActionsByProject.values} xLabels={legalOpenActionsByProject.labels} themeName="risk" showTrend={false} />
                </article>
              </div>

              <div className="risk-meta-row">
                <article className="risk-meta-card">
                  <span>Yaklaşan Gözden Geçirme Uyarısı</span>
                  <strong>{legalAlerts.upcoming.length}</strong>
                </article>
                <article className="risk-meta-card">
                  <span>Süresi Dolan Gereklilik Uyarısı</span>
                  <strong>{legalAlerts.expired.length}</strong>
                </article>
                <article className="risk-meta-card">
                  <span>Belge Silme Yetkisi</span>
                  <strong>{legalCanDeleteDocument ? 'Aktif' : 'Sadece Uyumluluk Yöneticisi'}</strong>
                </article>
              </div>
            </section>

            <section className="panel legal-panel">
              <h2>Yasal Mevzuat Kaydı Veri Girişi</h2>
              <div className="form-grid legal-form-grid">
                <label>
                  Proje
                  <select value={legalForm.projectId} onChange={(event) => setLegalForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Mevzuat Kimliği
                  <input value={nextLegalRegulationId} disabled />
                </label>
                <label>
                  Mevzuat Kategorisi
                  <select value={legalForm.category} onChange={(event) => setLegalForm((prev) => ({ ...prev, category: event.target.value }))}>
                    {legalCategoryOptions.map((category) => (
                      <option key={category} value={category}>{category}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Mevzuat Başlığı
                  <input value={legalForm.title} onChange={(event) => setLegalForm((prev) => ({ ...prev, title: event.target.value }))} />
                </label>
                <label>
                  İlgili Otorite / Kurum
                  <select value={legalForm.authority} onChange={(event) => setLegalForm((prev) => ({ ...prev, authority: event.target.value }))}>
                    {legalAuthorityOptions.map((authority) => (
                      <option key={authority} value={authority}>{authority}</option>
                    ))}
                  </select>
                </label>
                <label>
                  İlgili Departman
                  <select value={legalForm.department} onChange={(event) => setLegalForm((prev) => ({ ...prev, department: event.target.value }))}>
                    {departmentNames.map((department) => (
                      <option key={department} value={department}>{department}</option>
                    ))}
                  </select>
                </label>
                <label className="full-row">
                  Yasal Gereklilik
                  <textarea rows={3} value={legalForm.legalRequirement} onChange={(event) => setLegalForm((prev) => ({ ...prev, legalRequirement: event.target.value }))} />
                </label>
                <label>
                  Sorumlu Kişi
                  <input value={legalForm.responsiblePerson} onChange={(event) => setLegalForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))} />
                </label>
                <label>
                  Uyumluluk Durumu
                  <select value={legalForm.complianceStatus} onChange={(event) => setLegalForm((prev) => ({ ...prev, complianceStatus: event.target.value as LegalComplianceStatus }))}>
                    <option value="UYUMLU">Uyumlu</option>
                    <option value="KISMEN_UYUMLU">Kısmen Uyumlu</option>
                    <option value="UYUMSUZ">Uyumsuz</option>
                    <option value="UYGULANAMAZ">Uygulanamaz</option>
                  </select>
                </label>
                <label>
                  Yürürlük Tarihi
                  <input type="date" value={legalForm.effectiveDate} onChange={(event) => setLegalForm((prev) => ({ ...prev, effectiveDate: event.target.value }))} />
                </label>
                <label>
                  Son Gözden Geçirme Tarihi
                  <input type="date" value={legalForm.lastReviewDate} onChange={(event) => setLegalForm((prev) => ({ ...prev, lastReviewDate: event.target.value }))} />
                </label>
                <label>
                  Bir Sonraki Gözden Geçirme Tarihi
                  <input type="date" value={legalForm.nextReviewDate} onChange={(event) => setLegalForm((prev) => ({ ...prev, nextReviewDate: event.target.value }))} />
                </label>
                <label>
                  Açık Aksiyonlar
                  <input type="number" min={0} value={legalForm.openActions} onChange={(event) => setLegalForm((prev) => ({ ...prev, openActions: Number(event.target.value) }))} />
                </label>
                <label>
                  Risk Seviyesi
                  <select value={legalForm.riskLevel} onChange={(event) => setLegalForm((prev) => ({ ...prev, riskLevel: event.target.value as LegalRiskLevel }))}>
                    <option value="DUSUK">Düşük</option>
                    <option value="ORTA">Orta</option>
                    <option value="YUKSEK">Yüksek</option>
                    <option value="KRITIK">Kritik</option>
                  </select>
                </label>
                <label className="full-row">
                  Notlar
                  <textarea rows={3} value={legalForm.notes} onChange={(event) => setLegalForm((prev) => ({ ...prev, notes: event.target.value }))} />
                </label>
                <div className="full-row actions">
                  <button type="button" onClick={saveLegalRecord}>Yasal Kaydı Kaydet</button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap legal-panel">
              <h2>Yasal Mevzuat Kayıtları</h2>
              <table className="legal-record-table">
                <thead>
                  <tr>
                    <th>Proje</th>
                    <th>Mevzuat Başlığı</th>
                    <th>Kategori</th>
                    <th>Uyumluluk Durumu</th>
                    <th>Sorumlu Kişi</th>
                    <th>Bir Sonraki Gözden Geçirme</th>
                    <th>Açık Aksiyonlar</th>
                  </tr>
                </thead>
                <tbody>
                  {legalScopedRows.length > 0 ? (
                    legalScopedRows.map((record) => (
                      <tr
                        key={record.id}
                        className={selectedLegalRecordId === record.id ? 'risk-row-selected' : ''}
                        onClick={() => {
                          setSelectedLegalRecordId(record.id);
                          setSelectedLegalDocumentId(record.documents[0]?.id ?? null);
                        }}
                      >
                        <td>{projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId}</td>
                        <td>{record.title}</td>
                        <td>{record.category}</td>
                        <td><span className={`status-badge ${legalComplianceClass(record.complianceStatus)}`}>{legalComplianceLabel(record.complianceStatus)}</span></td>
                        <td>{record.responsiblePerson}</td>
                        <td>{record.nextReviewDate}</td>
                        <td>{record.openActions}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={7}>Filtreye uygun yasal kayıt bulunmuyor.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            {selectedLegalRecord ? (
              <section className="panel legal-panel">
                <div className="risk-detail-head legal-detail-head">
                  <div>
                    <h2>Belge Yönetimi ve Denetim - {selectedLegalRecord.regulationId}</h2>
                    <p className="risk-detail-subtitle">{selectedLegalRecord.title} | {selectedLegalRecord.authority}</p>
                  </div>
                  <div className="legal-action-toolbar">
                    <span className="legal-action-toolbar-title">{language === 'ru' ? 'Быстрые операции с отчетом' : 'Hızlı Rapor İşlemleri'}</span>
                    <div className="risk-detail-badges legal-action-bar">
                      <button type="button" onClick={() => markLegalReviewed(selectedLegalRecord.id)}>{language === 'ru' ? 'Отмечено как проверено' : 'Gözden Geçirildi'}</button>
                      <button type="button" onClick={exportSingleLegalPdf}>{language === 'ru' ? 'PDF по записи' : 'Tekil PDF'}</button>
                      <button type="button" onClick={exportFullLegalPdf}>{language === 'ru' ? 'Полный PDF' : 'Tam PDF'}</button>
                      <button type="button" onClick={exportLegalExcel}>{language === 'ru' ? 'Excel' : 'Excel'}</button>
                      <button type="button" onClick={printSingleLegalRecord}>{language === 'ru' ? 'Печать записи' : 'Tekil Yazdır'}</button>
                      <button type="button" onClick={printFullLegalRegister}>{language === 'ru' ? 'Печать полного реестра' : 'Tam Yazdır'}</button>
                      <button type="button" onClick={printLegalComplianceSummary}>{language === 'ru' ? 'Печать сводки' : 'Özet Yazdır'}</button>
                    </div>
                  </div>
                </div>

                <div className="legal-document-upload-grid">
                  <label>
                    {language === 'ru' ? 'Загрузить нормативный документ (PDF, DOCX, XLSX)' : 'Mevzuat Belgesi Yükle (PDF, DOCX, XLSX)'}
                    {language === 'ru' ? (
                      <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx" multiple onFilesChange={(files) => uploadLegalDocuments(selectedLegalRecord.id, 'MEVZUAT_BELGESI', filesToFileList(files))} />
                    ) : (
                      <input type="file" accept=".pdf,.docx,.xlsx" multiple onChange={(event) => uploadLegalDocuments(selectedLegalRecord.id, 'MEVZUAT_BELGESI', event.target.files)} />
                    )}
                  </label>
                  <label>
                    {language === 'ru' ? 'Загрузить подтверждение соответствия' : 'Uyumluluk Kanıtı Yükle'}
                    {language === 'ru' ? (
                      <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onFilesChange={(files) => uploadLegalDocuments(selectedLegalRecord.id, 'UYUMLULUK_KANITI', filesToFileList(files))} />
                    ) : (
                      <input type="file" accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onChange={(event) => uploadLegalDocuments(selectedLegalRecord.id, 'UYUMLULUK_KANITI', event.target.files)} />
                    )}
                  </label>
                  <label>
                    {language === 'ru' ? 'Загрузить отчет о проверке' : 'Denetim Raporu Yükle'}
                    {language === 'ru' ? (
                      <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onFilesChange={(files) => uploadLegalDocuments(selectedLegalRecord.id, 'DENETIM_RAPORU', filesToFileList(files))} />
                    ) : (
                      <input type="file" accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onChange={(event) => uploadLegalDocuments(selectedLegalRecord.id, 'DENETIM_RAPORU', event.target.files)} />
                    )}
                  </label>
                  <label>
                    {language === 'ru' ? 'Загрузить разрешение / сертификат' : 'İzin / Sertifika Yükle'}
                    {language === 'ru' ? (
                      <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onFilesChange={(files) => uploadLegalDocuments(selectedLegalRecord.id, 'IZIN_SERTIFIKA', filesToFileList(files))} />
                    ) : (
                      <input type="file" accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" multiple onChange={(event) => uploadLegalDocuments(selectedLegalRecord.id, 'IZIN_SERTIFIKA', event.target.files)} />
                    )}
                  </label>
                </div>

                <div className="legal-doc-layout">
                  <div className="legal-doc-column">
                    <article className="table-wrap">
                      <table className="legal-doc-table">
                        <thead>
                          <tr>
                            <th>{language === 'ru' ? 'Документ' : 'Belge'}</th>
                            <th>{language === 'ru' ? 'Тип' : 'Tür'}</th>
                            <th>{language === 'ru' ? 'Загрузил' : 'Yükleyen'}</th>
                            <th>{language === 'ru' ? 'Версия' : 'Sürüm'}</th>
                            <th>{language === 'ru' ? 'Действия' : 'İşlemler'}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {selectedLegalRecord.documents.length > 0 ? (
                            selectedLegalRecord.documents.map((documentRow) => (
                              <tr key={documentRow.id} className={selectedLegalDocumentId === documentRow.id ? 'risk-row-selected' : ''}>
                                <td>{documentRow.fileName}</td>
                                <td>{legalDocumentKindLabel(documentRow.kind)}</td>
                                <td>{documentRow.uploadedBy}</td>
                                <td>v{documentRow.versions.length}</td>
                                <td>
                                  <div className="legal-doc-actions">
                                    <button type="button" className="legal-inline-btn" onClick={() => setSelectedLegalDocumentId(documentRow.id)}>{language === 'ru' ? 'Просмотреть' : 'Önizle'}</button>
                                    <button type="button" className="legal-inline-btn" onClick={() => downloadLegalDocument(selectedLegalRecord.id, documentRow)}>{language === 'ru' ? 'Скачать' : 'İndir'}</button>
                                    <button type="button" className="legal-inline-btn" onClick={() => openLegalDocumentFullscreen(documentRow)}>{language === 'ru' ? 'На весь экран' : 'Tam Ekran'}</button>
                                    <label className="legal-replace-file legal-inline-btn">
                                      {language === 'ru' ? 'Заменить' : 'Değiştir'}
                                      {language === 'ru' ? (
                                        <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png" onFilesChange={(files) => replaceLegalDocument(selectedLegalRecord.id, documentRow.id, filesToFileList(files))} />
                                      ) : (
                                        <input
                                          type="file"
                                          accept=".pdf,.docx,.xlsx,.jpg,.jpeg,.png"
                                          onChange={(event) => replaceLegalDocument(selectedLegalRecord.id, documentRow.id, event.target.files)}
                                        />
                                      )}
                                    </label>
                                    <button type="button" className="legal-inline-btn" disabled={!legalCanDeleteDocument} onClick={() => deleteLegalDocument(selectedLegalRecord.id, documentRow.id)}>
                                      {language === 'ru' ? 'Удалить' : 'Sil'}
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={5}>{language === 'ru' ? 'Документы пока не загружены.' : 'Henüz belge yüklenmedi.'}</td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </article>

                    <div className="legal-audit-grid">
                      <article>
                        <h3>{language === 'ru' ? 'История проверки' : 'Denetim Geçmişi'}</h3>
                        <ul>
                          <li>{language === 'ru' ? 'Создал' : 'Oluşturan Kişi'}: {selectedLegalRecord.createdBy}</li>
                          <li>{language === 'ru' ? 'Дата создания' : 'Oluşturulma Tarihi'}: {selectedLegalRecord.createdAt}</li>
                          <li>{language === 'ru' ? 'Изменил' : 'Değiştiren Kişi'}: {selectedLegalRecord.modifiedBy}</li>
                          <li>{language === 'ru' ? 'Дата изменения' : 'Değiştirilme Tarihi'}: {selectedLegalRecord.modifiedAt}</li>
                          <li>{language === 'ru' ? 'Проверил' : 'Gözden Geçiren Kişi'}: {selectedLegalRecord.reviewedBy}</li>
                          <li>{language === 'ru' ? 'Дата проверки' : 'Gözden Geçirme Tarihi'}: {selectedLegalRecord.reviewedAt}</li>
                        </ul>
                      </article>
                      <article>
                        <h3>{language === 'ru' ? 'События аудита' : 'Audit Trail Olayları'}</h3>
                        <ul>
                          {selectedLegalRecord.auditTrail.map((eventRow) => (
                            <li key={eventRow.id}>{eventRow.eventDate} - {eventRow.actor} - {eventRow.detail}</li>
                          ))}
                        </ul>
                      </article>
                    </div>
                  </div>

                  <article className="legal-viewer-shell">
                    <h3>{language === 'ru' ? 'Просмотр документа' : 'Belge Görüntüleyici'}</h3>
                    {selectedLegalDocument ? (
                      <>
                        {selectedLegalDocument.fileType === 'application/pdf' ? (
                          <>
                            <div className="legal-viewer-toolbar">
                              <button type="button" onClick={() => setLegalViewerZoom((prev) => Math.max(60, prev - 10))}>-</button>
                              <span>{language === 'ru' ? 'Масштаб' : 'Yakınlaştırma'}: {legalViewerZoom}%</span>
                              <button type="button" onClick={() => setLegalViewerZoom((prev) => Math.min(220, prev + 10))}>+</button>
                              <label>
                                {language === 'ru' ? 'Страница' : 'Sayfa'}
                                <input type="number" min={1} value={legalViewerPage} onChange={(event) => setLegalViewerPage(Math.max(1, Number(event.target.value) || 1))} />
                              </label>
                              <label>
                                {language === 'ru' ? 'Поиск' : 'Ara'}
                                <input value={legalViewerSearch} onChange={(event) => setLegalViewerSearch(event.target.value)} />
                              </label>
                            </div>
                            <iframe
                              className="legal-pdf-iframe"
                              title="legal-pdf-viewer"
                              src={`${selectedLegalDocument.fileUrl}#toolbar=1&navpanes=0&page=${legalViewerPage}&zoom=${legalViewerZoom}&search=${encodeURIComponent(legalViewerSearch)}`}
                            />
                          </>
                        ) : selectedLegalDocument.fileType.startsWith('image/') ? (
                          <img src={selectedLegalDocument.fileUrl} alt={selectedLegalDocument.fileName} className="legal-preview-image" />
                        ) : (
                          <div className="inline-hint">
                            {language === 'ru' ? 'Предпросмотр Office-файлов в этой среде может быть ограничен. Используйте скачивание или полноэкранный просмотр.' : 'Office önizleme bu ortamda sınırlı olabilir. Güvenli indirme veya tam ekranda açma ile görüntüleyebilirsiniz.'}
                          </div>
                        )}

                        <div className="legal-version-panel">
                          <h4>{language === 'ru' ? 'История версий' : 'Sürüm Geçmişi'}</h4>
                          <ul>
                            {selectedLegalDocument.versions.map((version) => (
                              <li key={`${selectedLegalDocument.id}-v${version.version}`}>
                                v{version.version} - {version.fileName} - {version.uploadedAt.slice(0, 10)} - {version.uploadedBy}
                              </li>
                            ))}
                          </ul>
                        </div>
                      </>
                    ) : (
                      <p>{language === 'ru' ? 'Выберите документ для просмотра.' : 'Önizlemek için bir belge seçin.'}</p>
                    )}
                  </article>
                </div>

              </section>
            ) : null}
          </div>
        ) : null}

        {activeModule === 'documents' ? (
          <div className="document-register-layout">
            <section className="panel document-hero-panel">
              <div className="document-hero">
                <div>
                  <h2>{language === 'ru' ? 'Корпоративный контроль и управление документами' : 'Kurumsal Belge Kontrol ve Yönetim'}</h2>
                  <p>
                    {language === 'ru'
                      ? 'Стандартные корпоративные контролируемые документы, используемые в проектах HSE и EPC, управляются здесь по принципу единой записи. Для каждого документа хранится история ревизий, при этом отображается только последняя утверждённая версия.'
                      : 'İSG ve EPC projelerinde kullanılan şirket standardı kontrollü belgeler bu alanda tek kayıt mantığıyla yönetilir. Her belge için revizyon geçmişi tutulur, sadece en güncel onaylı sürüm görünür.'}
                  </p>
                </div>
                <div className="document-hero-note">
                  <strong>{language === 'ru' ? 'Принцип контроля' : 'Kontrol İlkesi'}</strong>
                  <p>{language === 'ru' ? 'Единая запись, одна актуальная версия, полный ревизионный след и лаконичный корпоративный вид.' : 'Tek kayıt, tek güncel sürüm, tam revizyon izi ve sade kurumsal görünüm.'}</p>
                </div>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Ввод данных документа' : 'Belge Veri Girişi'}</h2>
              <div className="form-grid document-form-grid">
                <label>
                  {language === 'ru' ? 'Проект' : 'Proje'}
                  <select value={controlledDocumentForm.projectId} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {language === 'ru' ? 'Идентификатор документа' : 'Belge Kimliği'}
                  <input value={nextControlledDocumentId} disabled />
                </label>
                <label>
                  {language === 'ru' ? 'Заголовок документа' : 'Belge Başlığı'}
                  <input value={controlledDocumentForm.title} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, title: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Категория документа' : 'Belge Kategorisi'}
                  <select value={controlledDocumentForm.category} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, category: event.target.value }))}>
                    {controlledDocumentCategoryOptions.map((category) => (
                      <option key={category} value={category}>{category}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {language === 'ru' ? 'Тип документа' : 'Belge Türü'}
                  <select value={controlledDocumentForm.documentType} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, documentType: event.target.value }))}>
                    {controlledDocumentTypeOptions.map((type) => (
                      <option key={type} value={type}>{type}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {language === 'ru' ? 'Номер ревизии' : 'Revizyon Numarası'}
                  <input value="1" disabled />
                </label>
                <label>
                  {language === 'ru' ? 'Дата вступления в силу' : 'Yürürlük Tarihi'}
                  <input type="date" value={controlledDocumentForm.effectiveDate} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, effectiveDate: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Дата пересмотра' : 'Gözden Geçirme Tarihi'}
                  <input type="date" value={controlledDocumentForm.reviewDate} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, reviewDate: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Ответственный отдел' : 'Sorumlu Departman'}
                  <select value={controlledDocumentForm.department} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, department: event.target.value }))}>
                    {departmentNames.map((department) => (
                      <option key={department} value={department}>{department}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {language === 'ru' ? 'Подготовил' : 'Hazırlayan'}
                  <input value={controlledDocumentForm.preparedBy} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, preparedBy: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Проверил' : 'Gözden Geçiren'}
                  <input value={controlledDocumentForm.reviewedBy} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, reviewedBy: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Утвердил' : 'Onaylayan'}
                  <input value={controlledDocumentForm.approvedBy} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, approvedBy: event.target.value }))} />
                </label>
                <label>
                  {language === 'ru' ? 'Статус' : 'Durum'}
                  <select value={controlledDocumentForm.status} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, status: event.target.value as ControlledDocumentStatus }))}>
                    {controlledDocumentStatusOptions.map((status) => (
                      <option key={status.value} value={status.value}>{status.label}</option>
                    ))}
                  </select>
                </label>
                <label className="full-row">
                  {language === 'ru' ? 'Загрузить корпоративный документ' : 'Şirket belgesini yükleyin'}
                  {language === 'ru' ? (
                    <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx" onFilesChange={(files) => setControlledDocumentFileList(filesToFileList(files))} />
                  ) : (
                    <input type="file" accept=".pdf,.docx,.xlsx" onChange={(event) => setControlledDocumentFileList(event.target.files)} />
                  )}
                </label>
                <label className="full-row">
                  {language === 'ru' ? 'Примечания' : 'Notlar'}
                  <textarea rows={3} value={controlledDocumentForm.notes} onChange={(event) => setControlledDocumentForm((prev) => ({ ...prev, notes: event.target.value }))} />
                </label>
                <div className="full-row actions">
                  <button type="button" onClick={saveControlledDocument}>{language === 'ru' ? 'Сохранить документ' : 'Belgeyi Kaydet'}</button>
                </div>
              </div>
            </section>

            <div className="document-workspace">
              <section className="panel table-wrap document-table-panel">
                <div className="document-toolbar">
                  <label>
                    {language === 'ru' ? 'Проект' : 'Proje'}
                    <select value={controlledDocumentFilters.projectId} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, projectId: event.target.value }))}>
                      <option value="all">{language === 'ru' ? 'Все проекты' : 'Tüm Projeler'}</option>
                      {projectCatalog.map((project) => (
                        <option key={project.id} value={project.id}>{project.name}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {language === 'ru' ? 'Категория' : 'Kategori'}
                    <select value={controlledDocumentFilters.category} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, category: event.target.value }))}>
                      <option value="all">{language === 'ru' ? 'Все' : 'Tümü'}</option>
                      {controlledDocumentCategoryOptions.map((category) => (
                        <option key={category} value={category}>{category}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {language === 'ru' ? 'Статус' : 'Durum'}
                    <select value={controlledDocumentFilters.status} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, status: event.target.value as 'ALL' | ControlledDocumentStatus }))}>
                      <option value="ALL">{language === 'ru' ? 'Все' : 'Tümü'}</option>
                      {controlledDocumentStatusOptions.map((status) => (
                        <option key={status.value} value={status.value}>{status.label}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {language === 'ru' ? 'Ревизия' : 'Revizyon'}
                    <select value={controlledDocumentFilters.revision} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, revision: event.target.value }))}>
                      {controlledDocumentRevisionOptions.map((revision) => (
                        <option key={revision} value={revision}>{revision === 'all' ? (language === 'ru' ? 'Все' : 'Tümü') : (language === 'ru' ? `Ревизия ${revision}` : `Revizyon ${revision}`)}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {language === 'ru' ? 'Отдел' : 'Departman'}
                    <select value={controlledDocumentFilters.department} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, department: event.target.value }))}>
                      <option value="all">{language === 'ru' ? 'Все' : 'Tümü'}</option>
                      {departmentNames.map((department) => (
                        <option key={department} value={department}>{department}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    {language === 'ru' ? 'Поиск' : 'Ara'}
                    <input placeholder={language === 'ru' ? 'Идентификатор документа или заголовок' : 'Belge Kimliği veya Başlık'} value={controlledDocumentFilters.keyword} onChange={(event) => setControlledDocumentFilters((prev) => ({ ...prev, keyword: event.target.value }))} />
                  </label>
                </div>

                <table className="document-table">
                  <thead>
                    <tr>
                      <th>{language === 'ru' ? 'Проект' : 'Proje'}</th>
                      <th>{language === 'ru' ? 'Идентификатор документа' : 'Belge Kimliği'}</th>
                      <th>{language === 'ru' ? 'Заголовок документа' : 'Belge Başlığı'}</th>
                      <th>{language === 'ru' ? 'Категория' : 'Kategori'}</th>
                      <th>{language === 'ru' ? 'Ревизия' : 'Revizyon'}</th>
                      <th>{language === 'ru' ? 'Статус' : 'Durum'}</th>
                      <th>{language === 'ru' ? 'Дата вступления в силу' : 'Yürürlük Tarihi'}</th>
                      <th>{language === 'ru' ? 'Дата пересмотра' : 'Gözden Geçirme Tarihi'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {controlledDocumentFilteredRows.length > 0 ? (
                      controlledDocumentFilteredRows.map((record) => (
                        <tr
                          key={record.id}
                          className={selectedControlledDocumentId === record.id ? 'risk-row-selected' : ''}
                          onClick={() => setSelectedControlledDocumentId(record.id)}
                        >
                          <td>{projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId}</td>
                          <td>{record.documentId}</td>
                          <td>{record.title}</td>
                          <td>{controlledDocumentCategoryLabel(record.category)}</td>
                          <td>v{record.revisionNumber}</td>
                          <td><span className={`status-badge ${controlledDocumentStatusClass(record.status)} document-status-pill`}>{controlledDocumentStatusLabel(record.status)}</span></td>
                          <td>{record.effectiveDate}</td>
                          <td>{record.reviewDate}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={8}>{language === 'ru' ? 'Контролируемых документов, соответствующих фильтрам, нет.' : 'Filtreye uygun kontrollü belge bulunmuyor.'}</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </section>

              <section className="panel document-detail-panel">
                {selectedControlledDocument ? (
                  <>
                    <div className="document-detail-header">
                      <div>
                        <h3>{selectedControlledDocument.documentId}</h3>
                        <p>{selectedControlledDocument.title} | {selectedControlledDocument.category}</p>
                      </div>
                      <div className="document-action-bar">
                        <button type="button" onClick={() => viewControlledDocument(selectedControlledDocument.id)}>{language === 'ru' ? 'Просмотреть' : 'Görüntüle'}</button>
                        <button type="button" onClick={() => downloadControlledDocument(selectedControlledDocument.id)}>{language === 'ru' ? 'Скачать' : 'İndir'}</button>
                        <button type="button" onClick={() => deleteControlledDocument(selectedControlledDocument.id)}>{language === 'ru' ? 'Удалить' : 'Sil'}</button>
                      </div>
                    </div>

                    <div className="document-detail-meta">
                      <article><span>{language === 'ru' ? 'Проект' : 'Proje'}</span><strong>{projectCatalog.find((project) => project.id === selectedControlledDocument.projectId)?.name ?? selectedControlledDocument.projectId}</strong></article>
                      <article><span>{language === 'ru' ? 'Тип документа' : 'Belge Türü'}</span><strong>{controlledDocumentTypeLabel(selectedControlledDocument.documentType)}</strong></article>
                      <article><span>{language === 'ru' ? 'Ревизия' : 'Revizyon'}</span><strong>v{selectedControlledDocument.revisionNumber}</strong></article>
                      <article><span>{language === 'ru' ? 'Статус' : 'Durum'}</span><strong>{controlledDocumentStatusLabel(selectedControlledDocument.status)}</strong></article>
                      <article><span>{language === 'ru' ? 'Ответственный отдел' : 'Sorumlu Departman'}</span><strong>{selectedControlledDocument.department}</strong></article>
                      <article><span>{language === 'ru' ? 'Утвердил' : 'Onaylayan'}</span><strong>{selectedControlledDocument.approvedBy}</strong></article>
                      <article><span>{language === 'ru' ? 'Подготовил' : 'Hazırlayan'}</span><strong>{selectedControlledDocument.preparedBy}</strong></article>
                      <article><span>{language === 'ru' ? 'Проверил' : 'Gözden Geçiren'}</span><strong>{selectedControlledDocument.reviewedBy}</strong></article>
                    </div>

                    <div className="document-preview-shell">
                      <h4>{language === 'ru' ? 'Текущий документ' : 'Güncel Belge'}</h4>
                      {selectedControlledDocumentRevision?.fileType === 'application/pdf' ? (
                        <iframe
                          className="document-preview-iframe"
                          title="controlled-document-preview"
                          src={`${selectedControlledDocumentRevision?.fileUrl}#toolbar=1&navpanes=0`}
                        />
                      ) : selectedControlledDocumentRevision?.fileType.startsWith('image/') ? (
                        <img src={selectedControlledDocumentRevision?.fileUrl} alt={selectedControlledDocumentRevision?.fileName ?? selectedControlledDocument.title} className="legal-preview-image" />
                      ) : (
                        <div className="inline-hint">{language === 'ru' ? 'Встроенный предпросмотр для Office-файлов может быть ограничен. Используйте скачивание или просмотр.' : 'Office dosyaları için yerleşik önizleme sınırlı olabilir. İndir veya görüntüle işlemini kullanabilirsiniz.'}</div>
                      )}
                      <div className="document-muted-note">
                        {language === 'ru' ? 'Отображается только последняя утверждённая ревизия. Предыдущие версии сохраняются в архиве.' : 'Yalnızca en son onaylı revizyon görüntülenir. Önceki sürümler arşivde tutulur.'}
                      </div>
                    </div>

                    <div className="document-revision-upload">
                      <h4>{language === 'ru' ? 'Загрузить новую ревизию' : 'Yeni Revizyon Yükle'}</h4>
                      <div className="form-grid document-revision-form">
                        <label className="full-row">
                          {language === 'ru' ? 'Файл ревизии' : 'Revizyon Dosyası'}
                          {language === 'ru' ? (
                            <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " accept=".pdf,.docx,.xlsx" onFilesChange={(files) => setControlledDocumentRevisionFileList(filesToFileList(files))} />
                          ) : (
                            <input type="file" accept=".pdf,.docx,.xlsx" onChange={(event) => setControlledDocumentRevisionFileList(event.target.files)} />
                          )}
                        </label>
                        <label>
                          {language === 'ru' ? 'Статус ревизии' : 'Revizyon Durumu'}
                          <select value={controlledDocumentRevisionStatus} onChange={(event) => setControlledDocumentRevisionStatus(event.target.value as ControlledDocumentStatus)}>
                            {controlledDocumentStatusOptions.map((status) => (
                              <option key={status.value} value={status.value}>{status.label}</option>
                            ))}
                          </select>
                        </label>
                        <label className="full-row">
                          {language === 'ru' ? 'Примечание к ревизии' : 'Revizyon Notu'}
                          <textarea rows={2} value={controlledDocumentRevisionNote} onChange={(event) => setControlledDocumentRevisionNote(event.target.value)} />
                        </label>
                        <div className="full-row actions">
                          <button type="button" onClick={() => addControlledDocumentRevision(selectedControlledDocument.id)}>{language === 'ru' ? 'Добавить ревизию' : 'Revizyon Ekle'}</button>
                        </div>
                      </div>
                    </div>

                    <div className="document-revision-list">
                      <h4>Revizyon Geçmişi</h4>
                      <ul>
                        {selectedControlledDocument.revisions.slice().reverse().map((revision) => (
                          <li key={revision.id}>
                            v{revision.revisionNumber} - {revision.fileName} - {revision.uploadedAt.slice(0, 10)} - {revision.uploadedBy} - {controlledDocumentStatusLabel(revision.status)}
                          </li>
                        ))}
                      </ul>
                    </div>
                  </>
                ) : (
                  <div className="inline-hint">Detayları görmek için bir kontrollü belge seçin.</div>
                )}
              </section>
            </div>
          </div>
        ) : null}

        {activeModule === 'reports' ? (
          <div className="report-center-layout">
            <section className="panel report-center-hero">
              <div>
                <h2>Kurumsal Rapor Merkezi</h2>
                <p>
                  HSE Uyumluluk Platformu içindeki tüm modüllerden gelen veriler, bu merkezde profesyonel raporlara
                  dönüştürülür. Manuel rapor kaydı yoktur; raporlar seçilen filtrelere göre otomatik oluşturulur.
                </p>
              </div>
            </section>

            <section className="panel report-center-panel">
              <h2>Rapor Üretimi</h2>
              <div className="report-center-filters">
                <label>
                  Rapor Türü
                  <select value={reportFilters.reportTypeKey} onChange={(event) => setReportFilters((prev) => ({ ...prev, reportTypeKey: event.target.value as CorporateReportKey }))}>
                    <optgroup label="Modül Raporları">
                      {corporateReportOptions.filter((option) => option.group === 'module').map((option) => (
                        <option key={option.key} value={option.key}>{option.label}</option>
                      ))}
                    </optgroup>
                    <optgroup label="Hazır Yönetim Raporları">
                      {corporateReportOptions.filter((option) => option.group === 'management').map((option) => (
                        <option key={option.key} value={option.key}>{option.label}</option>
                      ))}
                    </optgroup>
                    <optgroup label="Resmi Yazışma / Uyarı Raporları">
                      {corporateReportOptions.filter((option) => option.group === 'correspondence').map((option) => (
                        <option key={option.key} value={option.key}>{option.label}</option>
                      ))}
                    </optgroup>
                  </select>
                </label>
                <label>
                  Proje
                  <select value={reportFilters.projectId} onChange={(event) => setReportFilters((prev) => ({ ...prev, projectId: event.target.value }))}>
                    <option value="all">Tüm Projeler</option>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Departman
                  <select value={reportFilters.department} onChange={(event) => setReportFilters((prev) => ({ ...prev, department: event.target.value }))}>
                    <option value="all">Tüm Departmanlar</option>
                    {departmentNames.map((department) => (
                      <option key={department} value={department}>{department}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Tarih Aralığı - Başlangıç
                  <input type="date" value={reportFilters.periodStart} onChange={(event) => setReportFilters((prev) => ({ ...prev, periodStart: event.target.value }))} />
                </label>
                <label>
                  Tarih Aralığı - Bitiş
                  <input type="date" value={reportFilters.periodEnd} onChange={(event) => setReportFilters((prev) => ({ ...prev, periodEnd: event.target.value }))} />
                </label>
                <label>
                  Çıktı Formatı
                  <select value={reportFilters.format} onChange={(event) => setReportFilters((prev) => ({ ...prev, format: event.target.value as CorporateReportFormat }))}>
                    {reportFormatOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                </label>
              </div>

              <div className="report-template-grid">
                {corporateReportOptions.map((option) => (
                  <button
                    key={option.key}
                    type="button"
                    className={reportFilters.reportTypeKey === option.key ? 'report-template-chip active' : 'report-template-chip'}
                    onClick={() => setReportFilters((prev) => ({ ...prev, reportTypeKey: option.key }))}
                  >
                    <strong>{option.label}</strong>
                    <span>{option.group === 'module' ? 'Modül Raporu' : option.group === 'management' ? 'Hazır Yönetim Raporu' : 'Resmi Yazışma Raporu'}</span>
                  </button>
                ))}
              </div>

              <div className="report-center-actions">
                <button type="button" onClick={generateCorporateReport}>Rapor Oluştur</button>
                <p className="report-center-note">Raporda veri girişi yapılmaz. Seçilen filtrelere göre platform verilerinden snapshot alınır.</p>
              </div>
            </section>

            <section className="panel table-wrap report-center-table-panel">
              <h2>Oluşturulan Raporlar</h2>
              <table className="report-center-table">
                <thead>
                  <tr>
                    <th>Rapor Adı</th>
                    <th>Proje</th>
                    <th>Rapor Türü</th>
                    <th>Oluşturan Kişi</th>
                    <th>Oluşturulma Tarihi</th>
                    <th>Format</th>
                    <th>İşlemler</th>
                  </tr>
                </thead>
                <tbody>
                  {generatedCorporateReports.length > 0 ? (
                    generatedCorporateReports.map((report) => (
                      <tr
                        key={report.id}
                        className={selectedCorporateReportId === report.id ? 'risk-row-selected' : ''}
                        onClick={() => setSelectedCorporateReportId(report.id)}
                      >
                        <td>{report.reportName}</td>
                        <td>{report.projectName}</td>
                        <td>{report.reportTypeLabel}</td>
                        <td>{report.createdBy}</td>
                        <td>{report.createdAt.slice(0, 10)}</td>
                        <td>{report.format}</td>
                        <td>
                          <div className="report-row-actions">
                            <button type="button" onClick={(event) => { event.stopPropagation(); viewGeneratedCorporateReport(report); }}>Görüntüle</button>
                            <button type="button" onClick={(event) => { event.stopPropagation(); downloadGeneratedCorporateReport(report); }}>İndir</button>
                            <button type="button" onClick={(event) => { event.stopPropagation(); deleteGeneratedCorporateReport(report.id); }}>Sil</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={7}>Henüz oluşturulmuş rapor yok.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            <section className="panel report-center-detail-panel">
              {selectedGeneratedCorporateReport ? (
                <>
                  <div className="report-center-detail-head">
                    <div>
                      <h2>{selectedGeneratedCorporateReport.reportName}</h2>
                      <p>{selectedGeneratedCorporateReport.projectName} | {selectedGeneratedCorporateReport.department}</p>
                    </div>
                    <div className="report-row-actions">
                      <button type="button" onClick={() => viewGeneratedCorporateReport(selectedGeneratedCorporateReport)}>Görüntüle</button>
                      <button type="button" onClick={() => downloadGeneratedCorporateReport(selectedGeneratedCorporateReport)}>İndir</button>
                      <button type="button" onClick={() => deleteGeneratedCorporateReport(selectedGeneratedCorporateReport.id)}>Sil</button>
                    </div>
                  </div>

                  <div className="report-detail-meta">
                    <article><span>Rapor Türü</span><strong>{selectedGeneratedCorporateReport.reportTypeLabel}</strong></article>
                    <article><span>Proje</span><strong>{selectedGeneratedCorporateReport.projectName}</strong></article>
                    <article><span>Departman</span><strong>{selectedGeneratedCorporateReport.department}</strong></article>
                    <article><span>Rapor Dönemi</span><strong>{selectedGeneratedCorporateReport.periodStart} - {selectedGeneratedCorporateReport.periodEnd}</strong></article>
                    <article><span>Oluşturan Kişi</span><strong>{selectedGeneratedCorporateReport.createdBy}</strong></article>
                    <article><span>Oluşturulma Tarihi</span><strong>{selectedGeneratedCorporateReport.createdAt.slice(0, 10)}</strong></article>
                    <article><span>Format</span><strong>{selectedGeneratedCorporateReport.format}</strong></article>
                    <article><span>Kaynak</span><strong>Platform verileri</strong></article>
                  </div>

                  <div className="report-preview-shell">
                    <h3>Yönetici Özeti</h3>
                    <p>{selectedGeneratedCorporateReport.summary}</p>
                    <h3>Tablolar</h3>
                    <table className="report-preview-table">
                      <thead>
                        <tr>
                          <th>Alan</th>
                          <th>Değer</th>
                          <th>Not</th>
                        </tr>
                      </thead>
                      <tbody>
                        {selectedGeneratedCorporateReport.tableRows.map((row, index) => (
                          <tr key={`${row.label}-${index}`}>
                            <td>{row.label}</td>
                            <td>{row.value}</td>
                            <td>{row.note ?? '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <h3>Öneriler</h3>
                    <ul>
                      {selectedGeneratedCorporateReport.recommendations.map((recommendation) => (
                        <li key={recommendation}>{recommendation}</li>
                      ))}
                    </ul>
                  </div>
                </>
              ) : (
                <div className="inline-hint">Detayları görmek için oluşturulan bir raporu seçin.</div>
              )}
            </section>
          </div>
        ) : null}

        {activeModule !== 'dashboard' && activeModule !== 'occupational-health' && activeModule !== 'legal-register' && activeModule !== 'documents' && activeModule !== 'reports' && activeModule !== 'export-center' && activeModule !== 'projects' && activeModule !== 'departments' && activeModule !== 'contractors' && activeModule !== 'settings' && activeModule !== 'inspections' && activeModule !== 'observations' && activeModule !== 'risk-assessments' && activeModule !== 'permit-to-work' && activeModule !== 'incidents' && activeModule !== 'environmental' && activeModule !== 'action-tracker' && activeModule !== 'kpis-analytics' && activeModule !== 'equipment-management' && activeModule !== 'emergency-preparedness' && activeModule !== 'employees' && activeModule !== 'trainings' && activeModule !== 'ppe-stocks' ? (
          <section className="panel">
            <h2>{moduleLabels[activeModule]} {t.chart}</h2>
            <MiniBars values={chartValues} />
          </section>
        ) : null}

        {activeModule === 'occupational-health' ? (
          <section className="panel">
            <h2 className="health-chart-title">İş Sağlığı Analiz Grafiği</h2>
            <p className="message health-chart-note">Her sütun bir KPI metrik adını temsil eder. Bu grafik PDF raporunda da aynı başlıklarla yer alır.</p>
            <div className="executive-charts-grid">
              <div className="chart-card">
                <strong>İş Sağlığı KPI Metrikleri</strong>
                <DashboardChart
                  type="bar"
                  values={healthAnalytics.metricValues}
                  xLabels={healthAnalytics.metricLabels}
                  themeName="safety"
                  showTrend={false}
                  showAverage={false}
                />
              </div>
              <div className="chart-card">
                <strong>Sağlık Durum Dağılımı</strong>
                <DashboardChart
                  type="donut"
                  values={[healthAnalytics.fitForWork, healthAnalytics.notFitForWork, healthAnalytics.restrictedDuty]}
                  themeName="safety"
                />
                <div className="health-legend-grid">
                  <span>İşe Uygun: {healthAnalytics.fitForWork}</span>
                  <span>İşe Uygun Değil: {healthAnalytics.notFitForWork}</span>
                  <span>Kısıtlı Görev: {healthAnalytics.restrictedDuty}</span>
                </div>
              </div>
            </div>
          </section>
        ) : null}

        {activeModule === 'environmental' ? (
          <>
            <section className="panel environmental-panel">
              <div className="environmental-header">
                <div>
                  <h2>{language === 'ru' || activeModule === 'environmental' ? 'Регистрация экологических данных' : 'Регистрация экологических данных'}</h2>
                  <p>
                    {language === 'ru'
                      ? 'Данные об образовании отходов и переработке автоматически используются для формирования экологических показателей и графиков ниже.'
                      : 'Данные об образовании отходов и переработке автоматически используются для формирования экологических показателей и графиков ниже.'}
                  </p>
                </div>
                <div className="environmental-mini-grid">
                  <article className="environmental-mini-card">
                    <span>{language === 'ru' ? 'Количество записей' : 'Количество записей'}</span>
                    <strong>{environmentalSummary.recordCount}</strong>
                  </article>
                  <article className="environmental-mini-card">
                    <span>{language === 'ru' || activeModule === 'environmental' ? 'Образовано отходов' : 'Образовано отходов'}</span>
                    <strong>{environmentalSummary.wasteGenerated} {language === 'ru' || activeModule === 'environmental' ? 'т' : 'т'}</strong>
                  </article>
                  <article className="environmental-mini-card">
                    <span>{language === 'ru' ? 'Переработано отходов' : 'Переработано отходов'}</span>
                    <strong>{environmentalSummary.wasteRecycled} {language === 'ru' || activeModule === 'environmental' ? 'т' : 'т'}</strong>
                  </article>
                  <article className="environmental-mini-card">
                    <span>{language === 'ru' ? 'Чистый объём отходов' : 'Чистый объём отходов'}</span>
                    <strong>{environmentalSummary.netWaste} {language === 'ru' || activeModule === 'environmental' ? 'т' : 'т'}</strong>
                  </article>
                </div>
              </div>

              <div className="environmental-entry-grid">
                <div className="environmental-entry-card">
                  <div className="form-grid environmental-form-grid">
                    <label>
                      {language === 'ru' || activeModule === 'environmental' ? 'Проект' : 'Проект'}
                      <select
                        value={form.projectId}
                        onChange={(event) => setForm((prev) => ({ ...prev, projectId: event.target.value }))}
                      >
                        {projectCatalog.map((project) => (
                          <option key={project.id} value={project.id}>{project.name}</option>
                        ))}
                      </select>
                    </label>
                    <label>
                      {language === 'ru' || activeModule === 'environmental' ? 'Дата' : 'Дата'}
                      <input
                        type="date"
                        value={form.date}
                        onChange={(event) => setForm((prev) => ({ ...prev, date: event.target.value }))}
                      />
                    </label>
                    <label className="full-row">
                      {language === 'ru' || activeModule === 'environmental' ? 'Заголовок / Описание' : 'Заголовок / Описание'}
                      <input
                        value={form.title}
                        onChange={(event) => setForm((prev) => ({ ...prev, title: event.target.value }))}
                      />
                    </label>
                    <label>
                      {language === 'ru' || activeModule === 'environmental' ? 'Образовано отходов' : 'Образовано отходов'}
                      <input
                        type="number"
                        value={form.valueA}
                        onChange={(event) => setForm((prev) => ({ ...prev, valueA: Number(event.target.value) }))}
                      />
                    </label>
                    <label>
                      {language === 'ru' || activeModule === 'environmental' ? 'Переработано отходов' : 'Переработано отходов'}
                      <input
                        type="number"
                        value={form.valueB}
                        onChange={(event) => setForm((prev) => ({ ...prev, valueB: Number(event.target.value) }))}
                      />
                    </label>
                    <label>
                      {language === 'ru' || activeModule === 'environmental' ? 'Статус' : 'Статус'}
                      <select
                        value={form.status}
                        onChange={(event) => setForm((prev) => ({ ...prev, status: event.target.value as Status }))}
                      >
                        <option value="OPEN">{language === 'ru' || activeModule === 'environmental' ? 'Открыто' : 'Открыто'}</option>
                        <option value="IN_PROGRESS">{language === 'ru' || activeModule === 'environmental' ? 'В работе' : 'В работе'}</option>
                        <option value="CLOSED">{language === 'ru' || activeModule === 'environmental' ? 'Закрыто' : 'Закрыто'}</option>
                      </select>
                    </label>
                  </div>
                  <div className="actions environmental-actions">
                    <button type="button" onClick={saveEntry}>
                      {language === 'ru' || activeModule === 'environmental'
                        ? 'Сохранить запись'
                        : editingEnvironmentalIndex !== null
                          ? 'Сохранить запись'
                          : 'Сохранить запись'}
                    </button>
                  </div>
                </div>

                <div className="environmental-note-card">
                  <h3>{language === 'ru' ? 'Связь данных с графиками' : 'Связь данных с графиками'}</h3>
                  <p>
                    {language === 'ru'
                      ? 'Данные, введённые в этой форме, автоматически используются для построения графиков образования и переработки отходов.'
                      : 'Данные, введённые в этой форме, автоматически используются для построения графиков образования и переработки отходов.'}
                  </p>
                  <ul>
                    <li>{language === 'ru' ? 'Образовано отходов — основное значение графика образования отходов.' : 'Образовано отходов — основное значение графика образования отходов.'}</li>
                    <li>{language === 'ru' ? 'Переработано отходов — основное значение графика переработки отходов.' : 'Переработано отходов — основное значение графика переработки отходов.'}</li>
                    <li>{language === 'ru' ? 'Статус — текущее состояние записи.' : 'Статус — текущее состояние записи.'}</li>
                  </ul>
                </div>
              </div>
            </section>

            <section className="panel table-wrap environmental-table-panel">
              <h2>{language === 'ru' || activeModule === 'environmental' ? 'Реестр экологических записей' : 'Реестр экологических записей'}</h2>
              <table className="environmental-table">
                <thead>
                  <tr>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Проект' : 'Проект'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Дата' : 'Дата'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Заголовок' : 'Заголовок'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Образовано отходов' : 'Образовано отходов'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Переработано отходов' : 'Переработано отходов'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Статус' : 'Статус'}</th>
                    <th>{language === 'ru' || activeModule === 'environmental' ? 'Действия' : 'Действия'}</th>
                  </tr>
                </thead>
                <tbody>
                  {moduleData.environmental.length > 0 ? (
                    moduleData.environmental.map((row, index) => (
                      <tr key={`${row.title}-${index}`}>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.date}</td>
                        <td>{localizeText(row.title, language)}</td>
                        <td>{row.valueA}</td>
                        <td>{row.valueB}</td>
                        <td>{localizeStatus(row.status)}</td>
                        <td>
                          <div className="table-actions compact">
                            <button type="button" className="secondary" onClick={() => editEnvironmentalRecord(index)}>{language === 'ru' || activeModule === 'environmental' ? 'Редактировать' : 'Редактировать'}</button>
                            <button type="button" className="danger" onClick={() => deleteEnvironmentalRecord(index)}>{language === 'ru' || activeModule === 'environmental' ? 'Удалить' : 'Удалить'}</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={7}>{language === 'ru' || activeModule === 'environmental' ? 'Нет данных для выбранной области.' : t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'permit-to-work' ? (
          <>
            <section className="panel">
              <h2>{ptwCopy.formTitle}</h2>
              <div className="ptw-tabs">
                {ptwTablar.map((tab) => (
                  <button
                    key={tab.key}
                    type="button"
                    className={activePtwTab === tab.key ? 'ptw-tab active-ptw-tab' : 'ptw-tab'}
                    onClick={() => setActivePtwTab(tab.key)}
                  >
                    {(() => {
                      if (tab.key === 'kayit') return ptwCopy.tabMain;
                      if (tab.key === 'sorumlular') return ptwCopy.tabResponsible;
                      if (tab.key === 'ekip') return ptwCopy.tabTeam;
                      if (tab.key === 'is-bilgileri') return ptwCopy.tabWork;
                      if (tab.key === 'riskler') return ptwCopy.tabRisks;
                      if (tab.key === 'guvenlik-sistemleri') return ptwCopy.tabHeightSafety;
                      if (tab.key === 'ekipman') return ptwCopy.tabEquipment;
                      if (tab.key === 'on-kontroller') return ptwCopy.tabPreChecks;
                      if (tab.key === 'tedbirler') return ptwCopy.tabPrecautions;
                      if (tab.key === 'ozel-sartlar') return ptwCopy.tabSpecial;
                      if (tab.key === 'onaylar') return ptwCopy.tabApprovals;
                      if (tab.key === 'gunluk-takip') return ptwCopy.tabDaily;
                      if (tab.key === 'ekip-degisiklik') return ptwCopy.tabTeamChange;
                      if (tab.key === 'kapanis') return ptwCopy.tabClose;
                      if (tab.key === 'dosya-ekleri') return ptwCopy.tabAttachments;
                      return tab.label;
                    })()}
                  </button>
                ))}
              </div>
            </section>

            {activePtwTab === 'kayit' ? (
              <section className="panel">
                <h2>{ptwCopy.sectionMain}</h2>
                <div className="form-grid">
                  <label>{ptwCopy.ptwNo}<input value={ptwForm.ptwNo} disabled /></label>
                  <label>{ptwCopy.organization}<input value={ptwForm.organizasyon} onChange={(event) => setPtwForm((prev) => ({ ...prev, organizasyon: event.target.value }))} /></label>
                  <label>{ptwCopy.department}<input value={ptwForm.departman} onChange={(event) => setPtwForm((prev) => ({ ...prev, departman: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Проект' : 'Proje'}<input value={ptwForm.proje} onChange={(event) => setPtwForm((prev) => ({ ...prev, proje: event.target.value }))} /></label>
                  <label>{ptwCopy.location}<input value={ptwForm.lokasyon} onChange={(event) => setPtwForm((prev) => ({ ...prev, lokasyon: event.target.value }))} /></label>
                  <label>{ptwCopy.ptwType}<input value={language === 'ru' ? (ptwTypeRuLabel[ptwForm.ptwTuru] ?? ptwForm.ptwTuru) : ptwForm.ptwTuru} onChange={(event) => setPtwForm((prev) => ({ ...prev, ptwTuru: event.target.value }))} /></label>
                  <label>{ptwCopy.issueDate}<input type="date" value={ptwForm.duzenlenmeTarihi} onChange={(event) => setPtwForm((prev) => ({ ...prev, duzenlenmeTarihi: event.target.value }))} /></label>
                  <label>{ptwCopy.validFrom}<input type="date" value={ptwForm.gecerlilikBaslangici} onChange={(event) => setPtwForm((prev) => ({ ...prev, gecerlilikBaslangici: event.target.value }))} /></label>
                  <label>{ptwCopy.validTo}<input type="date" value={ptwForm.gecerlilikBitisi} onChange={(event) => setPtwForm((prev) => ({ ...prev, gecerlilikBitisi: event.target.value }))} /></label>
                  <label>
                    {ptwCopy.status}
                    <select value={ptwForm.durum} onChange={(event) => setPtwForm((prev) => ({ ...prev, durum: event.target.value as PtwStatus }))}>
                      {ptwDurumlar.map((durum) => (
                        <option key={durum} value={durum}>{language === 'ru' ? ptwStatusRuLabel[durum] : durum}</option>
                      ))}
                    </select>
                  </label>
                </div>
                <div className="actions">
                  <button type="button" onClick={() => void savePtwCurrentTab('kayit')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button>
                </div>
              </section>
            ) : null}

            {activePtwTab === 'sorumlular' ? (
              <section className="panel">
                <h2>{ptwCopy.tabResponsible}</h2>
                <div className="form-grid">
                  <label>{localizeText('İşi Talep Eden', language)}<input value={ptwForm.isiTalepEden} onChange={(event) => setPtwForm((prev) => ({ ...prev, isiTalepEden: event.target.value }))} /></label>
                  <label>{localizeText('İşi Veren', language)}<input value={ptwForm.isiVeren} onChange={(event) => setPtwForm((prev) => ({ ...prev, isiVeren: event.target.value }))} /></label>

                  <label>{localizeText('HSE Sorumlusu', language)}<input value={ptwForm.hseSorumlusu} onChange={(event) => setPtwForm((prev) => ({ ...prev, hseSorumlusu: event.target.value }))} /></label>
                  <label>{localizeText('Yetkili Onaylayan', language)}<input value={ptwForm.yetkiliOnaylayan} onChange={(event) => setPtwForm((prev) => ({ ...prev, yetkiliOnaylayan: event.target.value }))} /></label>
                </div>
                <div className="actions">
                  <button type="button" onClick={() => void savePtwCurrentTab('sorumlular')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button>
                </div>
              </section>
            ) : null}

            {activePtwTab === 'ekip' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabTeam}</h2>
                <table>
                  <thead>
                    <tr>
                      <th>{localizeText('Ad Soyad', language)}</th>
                      <th>{localizeText('Görevi', language)}</th>
                      <th>{localizeText('Firma', language)}</th>
                      <th>{localizeText('Eğitim Durumu', language)}</th>
                      <th>{localizeText('İmza', language)}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwEkipListesi.map((row, index) => (
                      <tr key={`ptw-team-${index}`}>
                        <td><input value={row.adSoyad} onChange={(event) => updatePtwTeamMember(index, 'adSoyad', event.target.value)} /></td>
                        <td><input value={row.gorevi} onChange={(event) => updatePtwTeamMember(index, 'gorevi', event.target.value)} /></td>
                        <td><input value={row.firma} onChange={(event) => updatePtwTeamMember(index, 'firma', event.target.value)} /></td>
                        <td><input value={row.egitimDurumu} onChange={(event) => updatePtwTeamMember(index, 'egitimDurumu', event.target.value)} /></td>
                        <td><input value={row.imza} onChange={(event) => updatePtwTeamMember(index, 'imza', event.target.value)} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={addPtwTeamMember}>{localizeText('Yeni Personel Ekle', language)}</button></div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('ekip')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'is-bilgileri' ? (
              <section className="panel">
                <h2>{ptwCopy.tabWork}</h2>
                <div className="form-grid">
                  <label>{localizeText('İşin Adı', language)}<input value={ptwForm.isinAdi} onChange={(event) => setPtwForm((prev) => ({ ...prev, isinAdi: event.target.value }))} /></label>
                  <label>{localizeText('Çalışma Alanı', language)}<input value={ptwForm.calismaAlani} onChange={(event) => setPtwForm((prev) => ({ ...prev, calismaAlani: event.target.value }))} /></label>
                  <label className="full-row">{localizeText('İşin Açıklaması', language)}<textarea rows={3} value={ptwForm.isinAciklamasi} onChange={(event) => setPtwForm((prev) => ({ ...prev, isinAciklamasi: event.target.value }))} /></label>
                  <label className="full-row">{localizeText('Yapılacak İş', language)}<textarea rows={3} value={ptwForm.yapilacakIs} onChange={(event) => setPtwForm((prev) => ({ ...prev, yapilacakIs: event.target.value }))} /></label>
                  <label className="full-row">{localizeText('Çalışma Koşulları', language)}<textarea rows={3} value={ptwForm.calismaKosullari} onChange={(event) => setPtwForm((prev) => ({ ...prev, calismaKosullari: event.target.value }))} /></label>
                  <label>{localizeText('Başlangıç Tarihi', language)}<input type="date" value={ptwForm.baslangicTarihi} onChange={(event) => setPtwForm((prev) => ({ ...prev, baslangicTarihi: event.target.value }))} /></label>
                  <label>{localizeText('Başlangıç Saati', language)}<input type="time" value={ptwForm.baslangicSaati} onChange={(event) => setPtwForm((prev) => ({ ...prev, baslangicSaati: event.target.value }))} /></label>
                  <label>{localizeText('Bitiş Tarihi', language)}<input type="date" value={ptwForm.bitisTarihi} onChange={(event) => setPtwForm((prev) => ({ ...prev, bitisTarihi: event.target.value }))} /></label>
                  <label>{localizeText('Bitiş Saati', language)}<input type="time" value={ptwForm.bitisSaati} onChange={(event) => setPtwForm((prev) => ({ ...prev, bitisSaati: event.target.value }))} /></label>
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('is-bilgileri')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'riskler' ? (
              <section className="panel">
                <h2>{ptwCopy.tabRisks}</h2>
                <div className="checkbox-grid">
                  {ptwTehlikeSecenekleri.map((tehlike) => (
                    <label key={tehlike} className="checkbox-item">
                      <input
                        type="checkbox"
                        checked={ptwTehlikeler.includes(tehlike)}
                        onChange={() => togglePtwArraySelection(tehlike, ptwTehlikeler, setPtwTehlikeler)}
                      />
                      {localizeText(tehlike, language)}
                    </label>
                  ))}
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('riskler')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'guvenlik-sistemleri' ? (
              <section className="panel">
                <h2>{ptwCopy.tabHeightSafety}</h2>
                <div className="checkbox-grid">
                  {ptwGuvenlikSistemleri.map((sistem) => (
                    <label key={sistem} className="checkbox-item">
                      <input
                        type="checkbox"
                        checked={ptwGuvenlikSecimleri.includes(sistem)}
                        onChange={() => togglePtwArraySelection(sistem, ptwGuvenlikSecimleri, setPtwGuvenlikSecimleri)}
                      />
                      {localizeText(sistem, language)}
                    </label>
                  ))}
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('guvenlik-sistemleri')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'ekipman' ? (
              <section className="panel">
                <h2>{ptwCopy.tabEquipment}</h2>
                <div className="checkbox-grid">
                  {ptwEkipmanSecenekleri.map((ekipman) => (
                    <label key={ekipman} className="checkbox-item">
                      <input
                        type="checkbox"
                        checked={ptwEkipmanSecimleri.includes(ekipman)}
                        onChange={() => togglePtwArraySelection(ekipman, ptwEkipmanSecimleri, setPtwEkipmanSecimleri)}
                      />
                      {ekipman}
                    </label>
                  ))}
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('ekipman')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'on-kontroller' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabPreChecks}</h2>
                <table>
                  <thead>
                    <tr>
                      <th>Kontrol</th>
                      <th>Durum</th>
                      <th>Açıklama</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwKontroller.map((row, index) => (
                      <tr key={`ptw-control-${index}`}>
                        <td><input value={row.kontrol} onChange={(event) => updatePtwControl(index, 'kontrol', event.target.value)} /></td>
                        <td>
                          <select value={row.durum} onChange={(event) => updatePtwControl(index, 'durum', event.target.value)}>
                            <option value="Uygun">Uygun</option>
                            <option value="Uygun Değil">Uygun Değil</option>
                            <option value="N/A">N/A</option>
                          </select>
                        </td>
                        <td><input value={row.aciklama} onChange={(event) => updatePtwControl(index, 'aciklama', event.target.value)} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={addPtwControl}>{language === 'ru' ? 'Добавить строку проверки' : 'Kontrol Satırı Ekle'}</button></div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('on-kontroller')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'tedbirler' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabPrecautions}</h2>
                <table>
                  <thead>
                    <tr>
                      <th>Tedbir</th>
                      <th>Sorumlu</th>
                      <th>Termin</th>
                      <th>Durum</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwTedbirler.map((row, index) => (
                      <tr key={`ptw-precaution-${index}`}>
                        <td><input value={row.tedbir} onChange={(event) => updatePtwPrecaution(index, 'tedbir', event.target.value)} /></td>
                        <td><input value={row.sorumlu} onChange={(event) => updatePtwPrecaution(index, 'sorumlu', event.target.value)} /></td>
                        <td><input type="date" value={row.termin} onChange={(event) => updatePtwPrecaution(index, 'termin', event.target.value)} /></td>
                        <td>
                          <select value={row.durum} onChange={(event) => updatePtwPrecaution(index, 'durum', event.target.value)}>
                            <option value="Açık">Açık</option>
                            <option value="Devam Ediyor">Devam Ediyor</option>
                            <option value="Tamamlandı">Tamamlandı</option>
                          </select>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={addPtwPrecaution}>{language === 'ru' ? 'Добавить меру' : 'Tedbir Ekle'}</button></div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('tedbirler')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'ozel-sartlar' ? (
              <section className="panel">
                <h2>{ptwCopy.tabSpecial}</h2>
                <textarea rows={8} value={ptwForm.ozelSartlar} onChange={(event) => setPtwForm((prev) => ({ ...prev, ozelSartlar: event.target.value }))} />
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('ozel-sartlar')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'onaylar' ? (
              <section className="panel">
                <h2>{ptwCopy.tabApprovals}</h2>
                <div className="form-grid">
                  <label>{language === 'ru' ? 'Подготовил PTW' : 'PTW Hazırlayan'}<input value={ptwForm.ptwHazirlayan} onChange={(event) => setPtwForm((prev) => ({ ...prev, ptwHazirlayan: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Согласование HSE' : 'HSE Onayı'}<input value={ptwForm.hseOnayi} onChange={(event) => setPtwForm((prev) => ({ ...prev, hseOnayi: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Руководитель проекта' : 'Proje Müdürü'}<input value={ptwForm.projeMuduru} onChange={(event) => setPtwForm((prev) => ({ ...prev, projeMuduru: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Представитель заказчика' : 'İşveren Temsilcisi'}<input value={ptwForm.isverenTemsilcisi} onChange={(event) => setPtwForm((prev) => ({ ...prev, isverenTemsilcisi: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Цифровая подпись' : 'Dijital İmza'}<input value={ptwForm.dijitalImza} onChange={(event) => setPtwForm((prev) => ({ ...prev, dijitalImza: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Дата' : 'Tarih'}<input type="date" value={ptwForm.onayTarihi} onChange={(event) => setPtwForm((prev) => ({ ...prev, onayTarihi: event.target.value }))} /></label>
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('onaylar')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'gunluk-takip' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabDaily}</h2>
                <table>
                  <thead>
                    <tr>
                      <th>{language === 'ru' ? 'Дата' : 'Tarih'}</th>
                      <th>{language === 'ru' ? 'Время' : 'Saat'}</th>
                      <th>{language === 'ru' ? 'Работа начата' : 'Çalışma Başladı'}</th>
                      <th>{language === 'ru' ? 'Работа завершена' : 'Çalışma Bitti'}</th>
                      <th>{language === 'ru' ? 'Описание' : 'Açıklama'}</th>
                      <th>{language === 'ru' ? 'Ответственный' : 'Sorumlu'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwGunlukKayitlar.map((row, index) => (
                      <tr key={`ptw-daily-${index}`}>
                        <td><input type="date" value={row.tarih} onChange={(event) => updatePtwDailyLog(index, 'tarih', event.target.value)} /></td>
                        <td><input type="time" value={row.saat} onChange={(event) => updatePtwDailyLog(index, 'saat', event.target.value)} /></td>
                        <td><input value={row.calismaBasladi} onChange={(event) => updatePtwDailyLog(index, 'calismaBasladi', event.target.value)} /></td>
                        <td><input value={row.calismaBitti} onChange={(event) => updatePtwDailyLog(index, 'calismaBitti', event.target.value)} /></td>
                        <td><input value={row.aciklama} onChange={(event) => updatePtwDailyLog(index, 'aciklama', event.target.value)} /></td>
                        <td><input value={row.sorumlu} onChange={(event) => updatePtwDailyLog(index, 'sorumlu', event.target.value)} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={addPtwDailyLog}>{language === 'ru' ? 'Добавить ежедневную запись' : 'Günlük Kayıt Ekle'}</button></div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('gunluk-takip')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'ekip-degisiklik' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabTeamChange}</h2>
                <table>
                  <thead>
                    <tr>
                      <th>{language === 'ru' ? 'Добавленный сотрудник' : 'Eklenen Personel'}</th>
                      <th>{language === 'ru' ? 'Выбывший сотрудник' : 'Ayrılan Personel'}</th>
                      <th>{language === 'ru' ? 'Дата' : 'Tarih'}</th>
                      <th>{language === 'ru' ? 'Согласовал' : 'Onaylayan'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwEkipDegisiklikleri.map((row, index) => (
                      <tr key={`ptw-team-change-${index}`}>
                        <td><input value={row.eklenenPersonel} onChange={(event) => updatePtwTeamChange(index, 'eklenenPersonel', event.target.value)} /></td>
                        <td><input value={row.ayrilanPersonel} onChange={(event) => updatePtwTeamChange(index, 'ayrilanPersonel', event.target.value)} /></td>
                        <td><input type="date" value={row.tarih} onChange={(event) => updatePtwTeamChange(index, 'tarih', event.target.value)} /></td>
                        <td><input value={row.onaylayan} onChange={(event) => updatePtwTeamChange(index, 'onaylayan', event.target.value)} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={addPtwTeamChange}>{language === 'ru' ? 'Добавить изменение состава' : 'Ekip Değişikliği Ekle'}</button></div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('ekip-degisiklik')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'kapanis' ? (
              <section className="panel">
                <h2>{ptwCopy.tabClose}</h2>
                <div className="form-grid">
                  <label className="checkbox-item"><input type="checkbox" checked={ptwForm.isTamamlandi} onChange={(event) => setPtwForm((prev) => ({ ...prev, isTamamlandi: event.target.checked }))} />{language === 'ru' ? 'Работа завершена' : 'İş Tamamlandı'}</label>
                  <label className="checkbox-item"><input type="checkbox" checked={ptwForm.alanGuvenli} onChange={(event) => setPtwForm((prev) => ({ ...prev, alanGuvenli: event.target.checked }))} />{language === 'ru' ? 'Зона безопасна' : 'Alan Güvenli'}</label>
                  <label className="checkbox-item"><input type="checkbox" checked={ptwForm.malzemelerToplandi} onChange={(event) => setPtwForm((prev) => ({ ...prev, malzemelerToplandi: event.target.checked }))} />{language === 'ru' ? 'Материалы собраны' : 'Malzemeler Toplandı'}</label>
                  <label className="checkbox-item"><input type="checkbox" checked={ptwForm.ptwKapatildi} onChange={(event) => setPtwForm((prev) => ({ ...prev, ptwKapatildi: event.target.checked }))} />{language === 'ru' ? 'Наряд-допуск закрыт' : 'PTW Kapatıldı'}</label>
                  <label>{language === 'ru' ? 'Закрыл' : 'Kapatan'}<input value={ptwForm.kapatan} onChange={(event) => setPtwForm((prev) => ({ ...prev, kapatan: event.target.value }))} /></label>
                  <label>{language === 'ru' ? 'Дата' : 'Tarih'}<input type="date" value={ptwForm.kapanisTarihi} onChange={(event) => setPtwForm((prev) => ({ ...prev, kapanisTarihi: event.target.value }))} /></label>
                  <label className="full-row">{language === 'ru' ? 'Описание' : 'Açıklama'}<textarea rows={4} value={ptwForm.kapanisAciklama} onChange={(event) => setPtwForm((prev) => ({ ...prev, kapanisAciklama: event.target.value }))} /></label>
                </div>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('kapanis')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            {activePtwTab === 'dosya-ekleri' ? (
              <section className="panel table-wrap">
                <h2>{ptwCopy.tabAttachments}</h2>
                <div className="ptw-attachment-grid">
                  {ptwDosyaTipleri.map((tip) => {
                    const displayTip = language === 'ru' ? (tip === 'Fotoğraf' ? 'Фотографии' : tip === 'Sertifikalar' ? 'Сертификаты' : tip === 'Risk Assessment' ? 'Оценка рисков' : tip === 'Method Statement' ? 'Методика работ' : tip === 'Toolbox Talk' ? 'Инструктаж Toolbox' : tip) : tip;
                    return (
                    <label key={tip}>
                      {displayTip}
                      <input
                        type="file"
                        multiple
                        onChange={async (event) => {
                          try {
                            await uploadPtwFiles(tip, event.target.files);
                            setPtwFeedback({ type: 'success', text: language === 'ru' ? 'Вложение загружено.' : 'Ek başarıyla yüklendi.' });
                          } catch (error) {
                            const errorText = extractErrorMessage(error);
                            console.error('PTW attachment upload failed:', errorText, error);
                            setPtwFeedback({ type: 'error', text: ptwUserFacingError(errorText) });
                          }
                        }}
                      />
                    </label>
                  );})}
                </div>
                <table>
                  <thead>
                    <tr>
                      <th>{language === 'ru' ? 'Тип вложения' : 'Ek Türü'}</th>
                      <th>{language === 'ru' ? 'Имя файла' : 'Dosya Adı'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ptwEkler.length > 0 ? (
                      ptwEkler.map((row, index) => (
                        <tr key={`ptw-attachment-${index}`}>
                          <td>{row.tip}</td>
                          <td>
                            <div className="table-actions compact">
                              <span>{row.dosyaAdi}</span>
                              <button type="button" className="secondary" onClick={() => window.open(row.fileUrl, '_blank')}>{language === 'ru' ? 'Просмотр' : 'Görüntüle'}</button>
                              <button type="button" className="secondary" onClick={() => {
                                const link = document.createElement('a');
                                link.href = row.fileUrl;
                                link.download = row.dosyaAdi;
                                document.body.appendChild(link);
                                link.click();
                                link.remove();
                              }}>{language === 'ru' ? 'Скачать' : 'İndir'}</button>
                              <button type="button" className="danger" onClick={() => deletePtwAttachment(index)}>{language === 'ru' ? 'Удалить' : 'Sil'}</button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={2}>{language === 'ru' ? 'Файлы пока не добавлены.' : 'Henüz dosya eklenmedi.'}</td>
                      </tr>
                    )}
                  </tbody>
                </table>
                <div className="actions"><button type="button" onClick={() => void savePtwCurrentTab('dosya-ekleri')} disabled={ptwSaving}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button></div>
              </section>
            ) : null}

            <section className="panel">
              <h2>{ptwCopy.reportActions}</h2>
              {ptwFeedback ? <p className="message" style={{ color: ptwFeedback.type === 'error' ? '#b91c1c' : '#166534' }}>{ptwFeedback.text}</p> : null}
              {ptwLastSavedAt ? <p className="message">{language === 'ru' ? `Последнее сохранение: ${new Date(ptwLastSavedAt).toLocaleString('ru-RU')}` : `Son kayıt: ${new Date(ptwLastSavedAt).toLocaleString('tr-TR')}`}</p> : null}
              <div className="form-grid" style={{ marginBottom: 12 }}>
                <label>
                  {language === 'ru' ? 'Email для отчета' : 'Rapor E-posta Adresi'}
                  <input value={ptwEmailTo} onChange={(event) => setPtwEmailTo(event.target.value)} placeholder="name@company.com" />
                </label>
              </div>
              {language === 'ru' ? (
                <p className="message">{ptwStatusFlow.map((status) => ptwStatusFlowRuLabel[status]).join(' -> ')} | Текущий статус: {ptwStatusRuLabel[ptwForm.durum]}</p>
              ) : null}
              <div className="report-action-buttons">
                <button type="button" onClick={() => void onExportPtwPdf()}>{ptwCopy.summaryReport}</button>
                <button type="button" onClick={() => void onExportPtwDocx()}>{ptwCopy.exportDocx}</button>
                <button type="button" onClick={() => void onExportPtwPdf()}>{ptwCopy.exportPdf}</button>
                <button type="button" onClick={() => void onExportPtwActionCsv()}>{ptwCopy.exportActionCsv}</button>
                <button type="button" onClick={() => void onExportPtwFullCsv()}>{ptwCopy.exportFullCsv}</button>
                <button type="button" onClick={() => void onEmailPtwSummary()}>{ptwCopy.sendMail}</button>
                <button type="button" onClick={onPrintPtw}>{ptwCopy.printPdf}</button>
                <button type="button" onClick={onDownloadAllAttachments}>{ptwCopy.downloadAttachments}</button>
                <button type="button" onClick={resetPtw}>{ptwCopy.reset}</button>
                <button type="button" onClick={() => void onSavePtwDraft()} disabled={ptwSaving}>{ptwCopy.saveDraft}</button>
                <button type="button" onClick={() => void onSubmitPtwApproval()} disabled={ptwSaving}>{ptwCopy.submitApproval}</button>
                <button type="button" onClick={() => void onClosePtw()} disabled={ptwSaving}>{ptwCopy.closePtw}</button>
              </div>
            </section>
          </>
        ) : null}

        {activeModule !== 'dashboard' && activeModule !== 'occupational-health' && activeModule !== 'legal-register' && activeModule !== 'documents' && activeModule !== 'reports' && activeModule !== 'export-center' && activeModule !== 'projects' && activeModule !== 'departments' && activeModule !== 'contractors' && activeModule !== 'settings' && activeModule !== 'inspections' && activeModule !== 'observations' && activeModule !== 'permit-to-work' && activeModule !== 'risk-assessments' && activeModule !== 'incidents' && activeModule !== 'equipment-management' && activeModule !== 'emergency-preparedness' && activeModule !== 'employees' && activeModule !== 'trainings' && activeModule !== 'ppe-stocks' && activeModule !== 'kpis-analytics' ? (
          <section className="panel">
            <h2>{activeModule === 'environmental' ? 'Экология Ввод данных' : `${moduleLabels[activeModule]} ${t.dataEntry}`}</h2>
            <div className="form-grid">
              <label>
                {activeModule === 'environmental' ? 'Проект' : t.project}
                <select
                  value={form.projectId}
                  onChange={(event) => setForm((prev) => ({ ...prev, projectId: event.target.value }))}
                >
                  {projectCatalog.map((project) => (
                    <option key={project.id} value={project.id}>{project.name}</option>
                  ))}
                </select>
              </label>
              <label>
                {language === 'ru' || activeModule === 'environmental' ? 'Дата' : 'Дата'}
                <input
                  type="date"
                  value={form.date}
                  onChange={(event) => setForm((prev) => ({ ...prev, date: event.target.value }))}
                />
              </label>
              <label>
                {language === 'ru' || activeModule === 'environmental' ? 'Заголовок / Описание' : 'Заголовок / Описание'}
                <input
                  value={form.title}
                  onChange={(event) => setForm((prev) => ({ ...prev, title: event.target.value }))}
                />
              </label>
              <label>
                {localizeText(metricFieldLabels[activeModule]?.primary ?? t.primaryMetric, language)}
                <input
                  type="number"
                  value={form.valueA}
                  onChange={(event) => setForm((prev) => ({ ...prev, valueA: Number(event.target.value) }))}
                />
              </label>
              <label>
                {localizeText(metricFieldLabels[activeModule]?.secondary ?? t.secondaryMetric, language)}
                <input
                  type="number"
                  value={form.valueB}
                  onChange={(event) => setForm((prev) => ({ ...prev, valueB: Number(event.target.value) }))}
                />
              </label>
              <label>
                {language === 'ru' || activeModule === 'environmental' ? 'Статус' : 'Статус'}
                <select
                  value={form.status}
                  onChange={(event) => setForm((prev) => ({ ...prev, status: event.target.value as Status }))}
                >
                  <option value="OPEN">{language === 'ru' || activeModule === 'environmental' ? 'Открыто' : 'Открыто'}</option>
                  <option value="IN_PROGRESS">{language === 'ru' || activeModule === 'environmental' ? 'В работе' : 'В работе'}</option>
                  <option value="CLOSED">{language === 'ru' || activeModule === 'environmental' ? 'Закрыто' : 'Закрыто'}</option>
                </select>
              </label>
              <div className="full-row actions">
                <button type="button" onClick={saveEntry}>{activeModule === 'environmental' ? 'Сохранить запись' : t.save}</button>
              </div>
            </div>
          </section>
        ) : null}

        {activeModule === 'kpis-analytics' ? (
          <>
            <section className="panel">
              <h2>KPI ve Analitik Özeti</h2>
              <p className="inline-hint">
                Bu ekran manuel veri girişi için değil; denetim, olay, eğitim, KKD, ekipman ve iş sağlığı modüllerindeki gerçek kayıtlardan otomatik KPI üretmek için kullanılır.
              </p>
              <div className="equipment-kpi-grid">
                {kpiAnalyticsCards.map((card) => (
                  <article key={card.label} className="equipment-kpi-card">
                    <span>{card.label}</span>
                    <strong>{card.value}</strong>
                    <small>{card.note}</small>
                  </article>
                ))}
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>KPI Kaynak Modülleri</h2>
              <table>
                <thead>
                  <tr>
                    <th>Modül</th>
                    <th>KPI Kaynağı</th>
                    <th>Nasıl Güncellenir?</th>
                  </tr>
                </thead>
                <tbody>
                  <tr><td>Denetimler</td><td>Checklist cevapları ve bulgular</td><td>Denetimler ekranında saha denetimi kaydı girilir.</td></tr>
                  <tr><td>Olaylar</td><td>Açık/kapalı olay kayıtları</td><td>Olaylar ekranında olay bildirimi kaydedilir.</td></tr>
                  <tr><td>Eğitimler</td><td>Katılımcı, sertifika ve maliyet kayıtları</td><td>Eğitimler ekranında eğitim kaydı girilir.</td></tr>
                  <tr><td>KKD Stokları</td><td>Stok giriş/çıkış hareketleri</td><td>KKD Stokları ekranında envanter işlemi kaydedilir.</td></tr>
                  <tr><td>Ekipman Yönetimi</td><td>Muayene, sertifika ve ekipman durumları</td><td>Ekipman Yönetimi ekranında ekipman kaydı girilir.</td></tr>
                  <tr><td>İş Sağlığı</td><td>İşe uygunluk ve muayene takip kayıtları</td><td>İş Sağlığı ekranında sağlık kaydı girilir.</td></tr>
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'emergency-preparedness' ? (
          <>
            <section className="panel">
              <h2>{language === 'ru' ? 'Показатели готовности к ЧС' : 'Acil Durum Hazırlık KPI'}</h2>
              <div className="equipment-kpi-grid">
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Всего учений по ЧС' : 'Toplam Acil Durum Tatbikatı'}</span><strong>{emergencySummary.totalDrills}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Всего участников' : 'Toplam Katılımcı'}</span><strong>{emergencySummary.totalParticipants}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Открытые действия' : 'Açık İşlemler'}</span><strong>{emergencySummary.openActions}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Закрытые действия' : 'Kapalı İşlemler'}</span><strong>{emergencySummary.closedActions}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Ближайшие учения' : 'Yaklaşan Tatbikatlar'}</span><strong>{emergencySummary.upcomingDrills}</strong></article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Визуализация эффективности учений по ЧС' : 'Acil Durum Tatbikatı Performans Görselleştirmesi'}</h2>
              <div className="equipment-chart-grid">
                <article className="chart-card">
                  <strong>{language === 'ru' ? 'Эффективность учений и мероприятий' : 'Tatbikat ve Eylem Performansı'}</strong>
                  <DashboardChart
                    type="bar"
                    values={[
                      emergencySummary.completedDrills,
                      emergencySummary.upcomingDrills,
                      emergencySummary.openActions,
                      emergencySummary.closedActions
                    ]}
                    themeName="safety"
                    xLabels={language === 'ru' ? ['Завершенные учения', 'Ближайшие учения', 'Открытые корректирующие действия', 'Закрытые корректирующие действия'] : ['Tamamlanan Tatbikat', 'Yaklaşan Tatbikat', 'Açık DÖF', 'Kapatılan DÖF']}
                  />
                </article>
                <article className="chart-card">
                  <strong>{language === 'ru' ? 'Распределение результатов учений' : 'Tatbikat Sonuç Dağılımı'}</strong>
                  <DashboardChart
                    type="donut"
                    values={[
                      emergencyScopedRows.filter((row) => row.drillResult === 'BASARILI').length,
                      emergencyScopedRows.filter((row) => row.drillResult === 'KISMEN_BASARILI').length,
                      emergencyScopedRows.filter((row) => row.drillResult === 'BASARISIZ').length
                    ]}
                    themeName="risk"
                  />
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Ввод данных по готовности к ЧС' : 'Acil Durum Hazırlık Veri Girişi'}</h2>
              <p className="inline-hint">
                {language === 'ru'
                  ? 'Этот экран предназначен для ведения единой записи по планированию учений, уровню готовности к реагированию и контролю корректирующих мероприятий.'
                  : 'Bu ekran acil durum tatbikat planı, müdahale hazırlığı ve düzeltici eylem takibini tek kayıtta yönetmek için tasarlanmıştır.'}
              </p>

              <div className="risk-meta-row">
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Идентификатор учения' : 'Tatbikat Kimliği'}</span>
                  <strong>{nextEmergencyDrillId}</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Открыто действий' : 'Açık İşlem'}</span>
                  <strong>{emergencySummary.openActions}</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Ближайшие учения' : 'Yaklaşan Tatbikat'}</span>
                  <strong>{emergencySummary.upcomingDrills}</strong>
                </div>
              </div>

              <div className="equipment-form-grid">
                <label>
                  {language === 'ru' ? 'Проект' : 'Proje'}
                  <select value={emergencyDrillForm.projectId} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Идентификатор учения по ЧС (авто)' : 'Acil Durum Tatbikatı Kimliği (Otomatik)'}
                  <input value={nextEmergencyDrillId} disabled />
                </label>

                <label>
                  {language === 'ru' ? 'Тип ЧС' : 'Acil Durum Türü'}
                  <select value={emergencyDrillForm.emergencyType} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, emergencyType: event.target.value as EmergencyType }))}>
                    <option value="YANGIN">{language === 'ru' ? 'Пожар' : 'Yangın'}</option>
                    <option value="TAHLIYE">{language === 'ru' ? 'Эвакуация' : 'Tahliye'}</option>
                    <option value="ILK_YARDIM">{language === 'ru' ? 'Первая помощь' : 'İlk Yardım'}</option>
                    <option value="KIMYASAL_SIZINTI">{language === 'ru' ? 'Химический разлив' : 'Kimyasal Sızıntı'}</option>
                    <option value="KURTARMA">{language === 'ru' ? 'Спасательные работы' : 'Kurtarma'}</option>
                    <option value="DEPREM">{language === 'ru' ? 'Землетрясение' : 'Deprem'}</option>
                    <option value="DIGER">{language === 'ru' ? 'Другое' : 'Diğer'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Наименование учения' : 'Tatbikat Adı'}
                  <input value={emergencyDrillForm.drillName} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, drillName: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Дата учения' : 'Tatbikat Tarihi'}
                  <input type="date" value={emergencyDrillForm.drillDate} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, drillDate: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Количество участников' : 'Katılımcı Sayısı'}
                  <input
                    type="number"
                    min={0}
                    value={emergencyDrillForm.participantCount}
                    onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, participantCount: Math.max(0, Number(event.target.value) || 0) }))}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Результат учения' : 'Tatbikat Sonucu'}
                  <select value={emergencyDrillForm.drillResult} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, drillResult: event.target.value as EmergencyDrillResult }))}>
                    <option value="BASARILI">{language === 'ru' ? 'Успешно' : 'Başarılı'}</option>
                    <option value="KISMEN_BASARILI">{language === 'ru' ? 'Частично успешно' : 'Kısmen Başarılı'}</option>
                    <option value="BASARISIZ">{language === 'ru' ? 'Неуспешно' : 'Başarısız'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Открытые действия' : 'Açık İşlemler'}
                  <input
                    type="number"
                    min={0}
                    value={emergencyDrillForm.openActions}
                    onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, openActions: Math.max(0, Number(event.target.value) || 0) }))}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Закрытые действия' : 'Kapalı İşlemler'}
                  <input
                    type="number"
                    min={0}
                    value={emergencyDrillForm.closedActions}
                    onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, closedActions: Math.max(0, Number(event.target.value) || 0) }))}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Ответственный' : 'Sorumlu Kişi'}
                  <input value={emergencyDrillForm.responsiblePerson} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Дата следующего планового учения' : 'Bir Sonraki Planlanan Tatbikat Tarihi'}
                  <input type="date" value={emergencyDrillForm.nextPlannedDrillDate} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, nextPlannedDrillDate: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Статус' : 'Durum'}
                  <select value={emergencyDrillForm.status} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, status: event.target.value as EmergencyDrillStatus }))}>
                    <option value="PLANLANDI">{language === 'ru' ? 'Запланировано' : 'Planlandı'}</option>
                    <option value="TAMAMLANDI">{language === 'ru' ? 'Завершено' : 'Tamamlandı'}</option>
                    <option value="IPTAL_EDILDI">{language === 'ru' ? 'Отменено' : 'İptal Edildi'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Вложения' : 'Ekler'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      multiple
                      onFilesChange={(files) => setEmergencyDrillForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEmergencyDrillForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label>
                  {language === 'ru' ? 'Отчеты по учению' : 'Tatbikat Raporu'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      multiple
                      onFilesChange={(files) => setEmergencyDrillForm((prev) => ({ ...prev, drillReports: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEmergencyDrillForm((prev) => ({ ...prev, drillReports: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label>
                  {language === 'ru' ? 'Фотографии' : 'Fotoğraflar'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      accept="image/*"
                      multiple
                      onFilesChange={(files) => setEmergencyDrillForm((prev) => ({ ...prev, photos: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      accept="image/*"
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEmergencyDrillForm((prev) => ({ ...prev, photos: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label className="full-row">
                  {language === 'ru' ? 'Примечания' : 'Notlar'}
                  <textarea rows={3} value={emergencyDrillForm.notes} onChange={(event) => setEmergencyDrillForm((prev) => ({ ...prev, notes: event.target.value }))} />
                </label>

                <div className="full-row actions">
                  <button type="button" onClick={saveEmergencyDrillEntry}>{language === 'ru' ? 'Сохранить' : 'Kaydet'}</button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{language === 'ru' ? 'Реестр учений по ЧС' : 'Acil Durum Tatbikat Kayıtları'}</h2>
              <table>
                <thead>
                  <tr>
                    <th>{language === 'ru' ? 'Дата учения' : 'Tatbikat Tarihi'}</th>
                    <th>{language === 'ru' ? 'Проект' : 'Proje'}</th>
                    <th>{language === 'ru' ? 'Тип ЧС' : 'Acil Durum Türü'}</th>
                    <th>{language === 'ru' ? 'Участники' : 'Katılımcılar'}</th>
                    <th>{language === 'ru' ? 'Открытые действия' : 'Açık İşlemler'}</th>
                    <th>{language === 'ru' ? 'Статус' : 'Durum'}</th>
                  </tr>
                </thead>
                <tbody>
                  {emergencyScopedRows.length > 0 ? (
                    emergencyScopedRows.map((row) => (
                      <tr key={row.id} className={selectedEmergencyDrillId === row.id ? 'risk-row-selected' : ''} onClick={() => setSelectedEmergencyDrillId(row.id)}>
                        <td>{row.drillDate}</td>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{emergencyTypeLabel(row.emergencyType)}</td>
                        <td>{row.participantCount}</td>
                        <td>{row.openActions}</td>
                        <td><span className={`status-badge ${emergencyStatusClass(row.status)}`}>{emergencyStatusLabel(row.status)}</span></td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={6}>{language === 'ru' ? 'Нет данных для выбранной области.' : t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            {selectedEmergencyDrillRecord ? (
              <section className="panel risk-detail-panel">
                <div className="risk-detail-head">
                  <div>
                    <h2>{language === 'ru' ? 'Карточка учения по ЧС' : 'Acil Durum Tatbikat Detayı'} - {selectedEmergencyDrillRecord.drillId}</h2>
                    <p className="risk-detail-subtitle">
                      {selectedEmergencyDrillRecord.drillName} - {emergencyTypeLabel(selectedEmergencyDrillRecord.emergencyType)}
                    </p>
                  </div>
                  <div className="risk-detail-badges">
                    <span className={`status-badge ${emergencyStatusClass(selectedEmergencyDrillRecord.status)}`}>{emergencyStatusLabel(selectedEmergencyDrillRecord.status)}</span>
                    <span className="risk-band-badge risk-band-medium">{language === 'ru' ? 'Результат' : 'Sonuç'}: {emergencyResultLabel(selectedEmergencyDrillRecord.drillResult)}</span>
                  </div>
                </div>

                <div className="risk-detail-grid">
                  <article>
                    <h3>{language === 'ru' ? 'Информация об учении' : 'Tatbikat Bilgisi'}</h3>
                    <p><strong>{language === 'ru' ? 'Дата учения' : 'Tatbikat Tarihi'}:</strong> {selectedEmergencyDrillRecord.drillDate}</p>
                    <p><strong>{language === 'ru' ? 'Участники' : 'Katılımcı'}:</strong> {selectedEmergencyDrillRecord.participantCount}</p>
                    <p><strong>{language === 'ru' ? 'Ответственный' : 'Sorumlu'}:</strong> {selectedEmergencyDrillRecord.responsiblePerson}</p>
                  </article>
                  <article>
                    <h3>{language === 'ru' ? 'Контроль мероприятий' : 'Eylem Takibi'}</h3>
                    <p><strong>{language === 'ru' ? 'Открыто действий' : 'Açık İşlem'}:</strong> {selectedEmergencyDrillRecord.openActions}</p>
                    <p><strong>{language === 'ru' ? 'Закрыто действий' : 'Kapalı İşlem'}:</strong> {selectedEmergencyDrillRecord.closedActions}</p>
                    <p><strong>{language === 'ru' ? 'Следующее учение' : 'Sonraki Tatbikat'}:</strong> {selectedEmergencyDrillRecord.nextPlannedDrillDate}</p>
                  </article>
                  <article>
                    <h3>{language === 'ru' ? 'Оценка готовности' : 'Hazırlık Değerlendirmesi'}</h3>
                    <p><strong>{language === 'ru' ? 'Результат учения' : 'Tatbikat Sonucu'}:</strong> {emergencyResultLabel(selectedEmergencyDrillRecord.drillResult)}</p>
                    <p><strong>{language === 'ru' ? 'Статус' : 'Durum'}:</strong> {emergencyStatusLabel(selectedEmergencyDrillRecord.status)}</p>
                    <p><strong>{language === 'ru' ? 'Осталось дней' : 'Kalan Gün'}:</strong> {daysUntil(selectedEmergencyDrillRecord.nextPlannedDrillDate)}</p>
                  </article>
                </div>

                <div className="risk-detail-files">
                  <div>
                    <h3>{language === 'ru' ? 'Вложения' : 'Ekler'}</h3>
                    {selectedEmergencyDrillRecord.attachments.length > 0 ? (
                      <ul>{selectedEmergencyDrillRecord.attachments.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Вложения отсутствуют.' : 'Ek bulunmuyor.'}</p>
                    )}
                  </div>
                  <div>
                    <h3>{language === 'ru' ? 'Отчеты по учению' : 'Tatbikat Raporları'}</h3>
                    {selectedEmergencyDrillRecord.drillReports.length > 0 ? (
                      <ul>{selectedEmergencyDrillRecord.drillReports.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Отчеты отсутствуют.' : 'Rapor bulunmuyor.'}</p>
                    )}
                  </div>
                  <div>
                    <h3>{language === 'ru' ? 'Фотографии' : 'Fotoğraflar'}</h3>
                    {selectedEmergencyDrillRecord.photos.length > 0 ? (
                      <ul>{selectedEmergencyDrillRecord.photos.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Фотографии отсутствуют.' : 'Fotoğraf bulunmuyor.'}</p>
                    )}
                  </div>
                </div>

                <div>
                  <h3>{language === 'ru' ? 'Примечания' : 'Notlar'}</h3>
                  <p>{selectedEmergencyDrillRecord.notes || '-'}</p>
                </div>
              </section>
            ) : null}
          </>
        ) : null}

        {activeModule === 'employees' ? (
          <>
            <section className="panel">
              <h2>{language === 'ru' ? 'Ключевые показатели рабочей силы' : 'İşgücü KPI'}</h2>
              <div className="workforce-kpi-grid">
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Всего рабочей силы' : 'Toplam İşgücü'}</span><strong>{workforceSummary.totalWorkforce}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Новые сотрудники' : 'Yeni Çalışanlar'}</span><strong>{workforceSummary.newEmployees}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Мужчины сотрудники' : 'Erkek Çalışanlar'}</span><strong>{workforceSummary.maleEmployees}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Женщины сотрудники' : 'Kadın Çalışanlar'}</span><strong>{workforceSummary.femaleEmployees}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Рабочая сила дневной смены' : 'Gündüz Vardiyası İş Gücü'}</span><strong>{workforceSummary.dayShiftWorkers}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Рабочая сила ночной смены' : 'Gece Vardiyası İş Gücü'}</span><strong>{workforceSummary.nightShiftWorkers}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Сотрудники сверхурочной работы' : 'Fazla Mesai Çalışanları'}</span><strong>{workforceSummary.overtimeWorkers}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Процент женских сотрудников' : 'Kadın Çalışan Oranı'}</span><strong>%{workforceRates.womenRatio}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Процент новых сотрудников' : 'Yeni Çalışan Oranı'}</span><strong>%{workforceRates.newEmployeeRatio}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Процент сверхурочной работы' : 'Fazla Mesai Oranı'}</span><strong>%{workforceRates.overtimeRatio}</strong></article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Визуализация производительности рабочей силы' : 'İşgücü Performans Görselleştirmesi'}</h2>
              <div className="workforce-chart-grid">
                <article className="chart-card workforce-chart-card workforce-age-card">
                  <strong>{language === 'ru' ? 'Распределение рабочей силы по возрасту' : 'İşgücü Yaş Dağılımı'}</strong>
                  <DashboardChart
                    type="bar"
                    values={[
                      workforceSummary.age18_25,
                      workforceSummary.age26_35,
                      workforceSummary.age36_45,
                      workforceSummary.age46_55,
                      workforceSummary.age56Plus
                    ]}
                    themeName="operations"
                    xLabels={['18-25', '26-35', '36-45', '46-55', '56+']}
                    showTrend={false}
                  />
                </article>

                <article className="chart-card workforce-chart-card workforce-department-card">
                  <strong>{language === 'ru' ? 'Рабочая сила по отделам' : 'Departmana Göre İşgücü'}</strong>
                  <div className="workforce-horizontal-chart">
                    {workforceDepartmentDistribution.map((item) => (
                      <div className="workforce-horizontal-row" key={item.label}>
                        <span>{translateDepartment(item.label, language)}</span>
                        <div className="workforce-horizontal-track">
                          <div
                            className="workforce-horizontal-fill"
                            style={{
                              width: `${item.value > 0 ? Math.max((item.value / workforceDepartmentMax) * 100, 4) : 0}%`
                            }}
                          />
                        </div>
                        <strong>{item.value}</strong>
                      </div>
                    ))}
                  </div>
                </article>

                <article className="chart-card workforce-chart-card workforce-shift-card">
                  <strong>{language === 'ru' ? 'Распределение смены' : 'Vardiya Dağılımı'}</strong>
                  <div className="workforce-shift-content">
                    <div className="workforce-shift-chart">
                      <DashboardChart
                        type="donut"
                        values={[workforceSummary.dayShiftWorkers, workforceSummary.nightShiftWorkers]}
                        themeName="safety"
                        donutSize="large"
                      />
                    </div>

                    <div className="workforce-shift-legend" aria-label="Vardiya dağılımı detayları">
                      <div className="workforce-shift-legend-item">
                        <span>{language === 'ru' ? 'Дневная смена' : 'Gündüz Vardiyası'}</span>
                        <strong>
                          {workforceSummary.dayShiftWorkers} {language === 'ru' ? 'человека' : 'kişi'} (
                          {Math.round(
                            (workforceSummary.dayShiftWorkers /
                              Math.max(workforceSummary.dayShiftWorkers + workforceSummary.nightShiftWorkers, 1)) *
                              100
                          )}
                          %)
                        </strong>
                      </div>
                      <div className="workforce-shift-legend-item">
                        <span>{language === 'ru' ? 'Ночная смена' : 'Gece Vardiyası'}</span>
                        <strong>
                          {workforceSummary.nightShiftWorkers} {language === 'ru' ? 'человека' : 'kişi'} (
                          {Math.round(
                            (workforceSummary.nightShiftWorkers /
                              Math.max(workforceSummary.dayShiftWorkers + workforceSummary.nightShiftWorkers, 1)) *
                              100
                          )}
                          %)
                        </strong>
                      </div>
                    </div>
                  </div>
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Ввод данных о рабочей силе' : 'İşgücü Veri Girişi'}</h2>
              <p className="inline-hint">
                {language === 'ru'
                  ? 'Этот модуль управляет статистикой рабочей силы, распределением смен и структурой рабочей силы на объекте для отчетности HSE, а не персональными картотеками отделов кадров.'
                  : 'Bu modül İK personel kartlarını değil; SEÇ raporlaması için işgücü istatistikleri, vardiya dağılımı ve sahaya ilişkin işgücü yapısını yönetir.'}
              </p>
              <div className="workforce-form-layout">
                <article className="workforce-form-section">
                  <h3>{language === 'ru' ? 'Основная информация' : 'Genel Bilgiler'}</h3>
                  <div className="workforce-form-grid">
                    <label>
                      {language === 'ru' ? 'Проект' : 'Proje'}
                      <select value={workforceForm.projectId} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                        {projectCatalog.map((project) => (
                          <option key={project.id} value={project.id}>{project.name}</option>
                        ))}
                      </select>
                    </label>

                    <label>
                      {language === 'ru' ? 'Дата' : 'Tarih'}
                      <input type="date" value={workforceForm.date} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, date: event.target.value }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Отдел / Рабочая область' : 'Departman / Çalışma Alanı'}
                      <select value={workforceForm.departmentArea} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, departmentArea: event.target.value }))}>
                        {workforceDepartmentOptions.map((department) => (
                          <option key={department} value={department}>{translateDepartment(department, language)}</option>
                        ))}
                      </select>
                    </label>

                    <label>
                      {language === 'ru' ? 'Подрядчик / Субподрядчик (Необязательно)' : 'Yüklenici / Alt Yüklenici (Opsiyonel)'}
                      <input value={workforceForm.contractor} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, contractor: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>{language === 'ru' ? 'Информация о рабочей силе' : 'İşgücü Bilgileri'}</h3>
                  <div className="workforce-form-grid">
                    <label>
                      {language === 'ru' ? 'Всего рабочей силы' : 'Toplam İşgücü'}
                      <input type="number" min={0} value={workforceForm.totalWorkforce} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, totalWorkforce: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Новые сотрудники' : 'Yeni Çalışanlar'}
                      <input type="number" min={0} value={workforceForm.newEmployees} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, newEmployees: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Мужчины сотрудники' : 'Erkek Çalışanlar'}
                      <input type="number" min={0} value={workforceForm.maleEmployees} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, maleEmployees: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Женщины сотрудники' : 'Kadın Çalışanlar'}
                      <input type="number" min={0} value={workforceForm.femaleEmployees} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, femaleEmployees: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Сотрудники дневной смены' : 'Gündüz Vardiyası Çalışanları'}
                      <input type="number" min={0} value={workforceForm.dayShiftWorkers} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, dayShiftWorkers: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Сотрудники ночной смены' : 'Gece Vardiyası Çalışanları'}
                      <input type="number" min={0} value={workforceForm.nightShiftWorkers} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, nightShiftWorkers: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      {language === 'ru' ? 'Сотрудники сверхурочной работы' : 'Fazla Mesai Çalışanları'}
                      <input type="number" min={0} value={workforceForm.overtimeWorkers} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, overtimeWorkers: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>{language === 'ru' ? 'Распределение по возрасту' : 'Yaş Dağılımı'}</h3>
                  <div className="workforce-form-grid">
                    <label>
                      18-25
                      <input type="number" min={0} value={workforceForm.age18_25} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, age18_25: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      26-35
                      <input type="number" min={0} value={workforceForm.age26_35} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, age26_35: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      36-45
                      <input type="number" min={0} value={workforceForm.age36_45} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, age36_45: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      46-55
                      <input type="number" min={0} value={workforceForm.age46_55} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, age46_55: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>

                    <label>
                      56+
                      <input type="number" min={0} value={workforceForm.age56Plus} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, age56Plus: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>{language === 'ru' ? 'Основной' : 'Genel'}</h3>
                  <div className="workforce-form-grid">
                    <label>
                      {language === 'ru' ? 'Статус' : 'Durum'}
                      <select value={workforceForm.status} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, status: event.target.value as Status }))}>
                        <option value="OPEN">{language === 'ru' ? 'Открыто' : 'Açık'}</option>
                        <option value="IN_PROGRESS">{language === 'ru' ? 'В процессе' : 'Devam Ediyor'}</option>
                        <option value="CLOSED">{language === 'ru' ? 'Закрыто' : 'Kapalı'}</option>
                      </select>
                    </label>

                    <label className="full-row">
                      {language === 'ru' ? 'Примечания' : 'Notlar'}
                      <textarea rows={3} value={workforceForm.notes} onChange={(event) => setWorkforceForm((prev) => ({ ...prev, notes: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <div className="actions">
                  <button type="button" onClick={saveWorkforceEntry}>{editingWorkforceId ? (language === 'ru' ? 'Обновить' : 'Güncelle') : (language === 'ru' ? 'Сохранить' : 'Kaydet')}</button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{language === 'ru' ? 'Записи о рабочей силе' : 'İşgücü Kayıtları'}</h2>
              <table>
                <thead>
                  <tr>
                    <th>{language === 'ru' ? 'Проект' : 'Proje'}</th>
                    <th>{language === 'ru' ? 'Дата' : 'Tarih'}</th>
                    <th>{language === 'ru' ? 'Отдел' : 'Departman'}</th>
                    <th>{language === 'ru' ? 'Всего рабочей силы' : 'Toplam İşgücü'}</th>
                    <th>{language === 'ru' ? 'Новые сотрудники' : 'Yeni Çalışanlar'}</th>
                    <th>{language === 'ru' ? 'Дневная смена' : 'Gündüz Vardiyası'}</th>
                    <th>{language === 'ru' ? 'Ночная смена' : 'Gece Vardiyası'}</th>
                    <th>{language === 'ru' ? 'Статус' : 'Durum'}</th>
                    <th>{language === 'ru' ? 'Операции' : 'İşlemler'}</th>
                  </tr>
                </thead>
                <tbody>
                  {workforceScopedRows.length > 0 ? (
                    workforceScopedRows.map((row) => (
                      <tr key={row.id}>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.date}</td>
                        <td>{translateDepartment(row.departmentArea, language)}</td>
                        <td>{row.totalWorkforce}</td>
                        <td>{row.newEmployees}</td>
                        <td>{row.dayShiftWorkers}</td>
                        <td>{row.nightShiftWorkers}</td>
                        <td className="workforce-status-cell"><span className={`status-badge ${getStatusBadgeClass(row.status)}`}>{localizeStatus(row.status)}</span></td>
                        <td className="actions-cell table-actions-cell">
                          <div className="table-action-group" aria-label={`${row.departmentArea} ${language === 'ru' ? 'операции рабочей силы' : 'işgücü işlemleri'}`}>
                            <button type="button" className="table-action-button" onClick={() => editWorkforceEntry(row.id)}>{language === 'ru' ? 'Редактировать' : 'Düzenle'}</button>
                            <button type="button" className="table-action-button danger" onClick={() => deleteWorkforceEntry(row.id)}>{language === 'ru' ? 'Удалить' : 'Sil'}</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={9}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'ppe-stocks' ? (
          <>
            <section className="panel">
              <h2>KKD Envanter KPI</h2>
              <div className="ppe-project-filter-row">
                <button
                  type="button"
                  className={ppeDashboardProjectFilter === 'all' ? 'ppe-project-filter-btn active' : 'ppe-project-filter-btn'}
                  onClick={() => setPpeDashboardProjectFilter('all')}
                >
                  Tüm Projeler
                </button>
                {projectCatalog.map((project) => (
                  <button
                    type="button"
                    key={project.id}
                    className={ppeDashboardProjectFilter === project.id ? 'ppe-project-filter-btn active' : 'ppe-project-filter-btn'}
                    onClick={() => setPpeDashboardProjectFilter(project.id)}
                  >
                    {project.name}
                  </button>
                ))}
              </div>
              <div className="equipment-kpi-grid">
                <article className="equipment-kpi-card"><span>Toplam KKD Stoğu</span><strong>{ppeSummary.totalStock}</strong></article>
                <article className="equipment-kpi-card"><span>Bu Ay Gelen Stok</span><strong>{ppeSummary.monthlyIncoming}</strong></article>
                <article className="equipment-kpi-card"><span>Bu Ay Çıkışı Yapılan Stok</span><strong>{ppeSummary.monthlyOutgoing}</strong></article>
                <article className="equipment-kpi-card"><span>Düşük Stoklu Kalemler</span><strong>{ppeSummary.lowStockItems}</strong></article>
                <article className="equipment-kpi-card"><span>Toplam Envanter Değeri</span><strong>{ppeSummary.totalInventoryValue.toLocaleString('tr-TR')} Rub</strong></article>
                <article className="equipment-kpi-card"><span>Gereken Satın Alma Siparişleri</span><strong>{ppeSummary.purchaseOrdersNeeded}</strong></article>
              </div>
              <div className="report-action-buttons">
                <button type="button" onClick={exportPpeSummaryPdf}>PDF Dışa Aktar</button>
                <button type="button" onClick={sendPpeSummaryEmail}>Özeti E-posta ile Gönder</button>
                <button type="button" onClick={exportPpeInventoryCsv}>Excel Dışa Aktar</button>
              </div>
            </section>

            <section className="panel">
              <h2>KKD Stok Görselleştirmesi</h2>
              <div className="ppe-project-filter-row">
                <button
                  type="button"
                  className={ppeDashboardProjectFilter === 'all' ? 'ppe-project-filter-btn active' : 'ppe-project-filter-btn'}
                  onClick={() => setPpeDashboardProjectFilter('all')}
                >
                  Tüm Projeler
                </button>
                {projectCatalog.map((project) => (
                  <button
                    type="button"
                    key={project.id}
                    className={ppeDashboardProjectFilter === project.id ? 'ppe-project-filter-btn active' : 'ppe-project-filter-btn'}
                    onClick={() => setPpeDashboardProjectFilter(project.id)}
                  >
                    {project.name}
                  </button>
                ))}
              </div>
              <div className="training-chart-grid">
                <article className="chart-card training-horizontal-card">
                  <strong>Kategori Bazında Mevcut KKD Stoğu</strong>
                  <div className="workforce-horizontal-chart">
                    {ppeCategoryStockData.length > 0 ? (
                      ppeCategoryStockData.map((row) => {
                        const maxValue = Math.max(...ppeCategoryStockData.map((item) => item.value), 1);
                        const width = row.value > 0 ? Math.max((row.value / maxValue) * 100, 4) : 0;
                        return (
                          <div key={row.label} className="workforce-horizontal-row">
                            <span>{row.label}</span>
                            <div className="workforce-horizontal-track"><div className="workforce-horizontal-fill" style={{ width: `${width}%` }} /></div>
                            <strong>{row.value}</strong>
                          </div>
                        );
                      })
                    ) : (
                      <p>{t.noData}</p>
                    )}
                  </div>
                </article>

                <article className="chart-card">
                  <strong>Aylık Stok Giriş ve Çıkış Karşılaştırması</strong>
                  <div className="ppe-line-shell">
                    <svg viewBox="0 0 100 100" className="viz-svg" preserveAspectRatio="none">
                      {[20, 40, 60, 80].map((y) => (
                        <line key={y} x1="6" y1={y} x2="94" y2={y} className="viz-grid-line" />
                      ))}
                      <polyline
                        points={ppeMonthlyTrend.map((row, index) => {
                          const x = 6 + (index / Math.max(ppeMonthlyTrend.length - 1, 1)) * 88;
                          const y = 92 - ((row.incoming / ppeMonthlyTrendMax) * 78);
                          return `${x},${y}`;
                        }).join(' ')}
                        fill="none"
                        stroke="#0284c7"
                        strokeWidth="2.4"
                        strokeLinejoin="round"
                        strokeLinecap="round"
                      />
                      <polyline
                        points={ppeMonthlyTrend.map((row, index) => {
                          const x = 6 + (index / Math.max(ppeMonthlyTrend.length - 1, 1)) * 88;
                          const y = 92 - ((row.outgoing / ppeMonthlyTrendMax) * 78);
                          return `${x},${y}`;
                        }).join(' ')}
                        fill="none"
                        stroke="#f97316"
                        strokeWidth="2.4"
                        strokeLinejoin="round"
                        strokeLinecap="round"
                      />
                    </svg>
                    <div className="ppe-line-legend">
                      <span><i className="ppe-line-dot ppe-line-dot-in" />Giriş</span>
                      <span><i className="ppe-line-dot ppe-line-dot-out" />Çıkış</span>
                    </div>
                    <div className="viz-axis-row">
                      {ppeMonthlyTrend.map((row) => <span key={row.key}>{row.label}</span>)}
                    </div>
                  </div>
                </article>

                <article className="chart-card">
                  <strong>Stok Durumu</strong>
                  <DashboardChart
                    type="donut"
                    values={[ppeStatusDistribution.YETERLI, ppeStatusDistribution.DUSUK_STOK, ppeStatusDistribution.STOKTA_YOK]}
                    themeName="compliance"
                    donutSize="large"
                  />
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>KKD Envanter İşlemi Girişi</h2>
              <p className="inline-hint">
                Mevcut stok manuel düzenlenmez; her hareket bir envanter işlemi olarak kaydedilir ve stok otomatik hesaplanır.
              </p>

              <div className="risk-meta-row">
                <div className="risk-meta-card"><span>İşlem Kimliği</span><strong>{nextPpeTransactionId}</strong></div>
                <div className="risk-meta-card"><span>Kullanılabilir Stok</span><strong>{ppeSummary.totalStock}</strong></div>
                <div className="risk-meta-card"><span>Stok Tükenme Uyarısı</span><strong>{ppeSummary.outOfStockItems}</strong></div>
              </div>

              <div className="training-form-layout">
                <article className="workforce-form-section">
                  <h3>Genel Bilgiler</h3>
                  <div className="equipment-form-grid">
                    <label>
                      İşlem Türü
                      <select value={ppeTransactionForm.transactionType} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, transactionType: event.target.value as PpeTransactionType }))}>
                        {ppeTransactionTypeOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                      </select>
                    </label>
                    <label>
                      İşlem Kimliği (Otomatik Oluşturulan)
                      <input value={nextPpeTransactionId} disabled />
                    </label>
                    <label>
                      Proje
                      <select value={ppeTransactionForm.projectId} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                        {projectCatalog.map((project) => <option key={project.id} value={project.id}>{project.name}</option>)}
                      </select>
                    </label>
                    <label>
                      Depo
                      <select value={ppeTransactionForm.warehouse} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, warehouse: event.target.value }))}>
                        {ppeWarehouseOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Tarih
                      <input type="date" value={ppeTransactionForm.date} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, date: event.target.value }))} />
                    </label>
                    <label>
                      KKD Kategorisi
                      <select value={ppeTransactionForm.category} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, category: event.target.value }))}>
                        {ppeCategoryOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      KKD Kalemi
                      <select value={ppeTransactionForm.itemName} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, itemName: event.target.value }))}>
                        {ppeItemOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Marka / Model
                      <input value={ppeTransactionForm.brandModel} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, brandModel: event.target.value }))} />
                    </label>
                    <label>
                      Birim
                      <select value={ppeTransactionForm.unit} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, unit: event.target.value }))}>
                        {ppeUnitOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>Envanter Bilgileri</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Mevcut Stok (Otomatik)
                      <input
                        disabled
                        value={(() => {
                          const found = ppeInventoryByKey.find((row) => row.projectId === ppeTransactionForm.projectId && row.warehouse === ppeTransactionForm.warehouse && row.category === ppeTransactionForm.category && row.itemName === ppeTransactionForm.itemName && row.unit === ppeTransactionForm.unit);
                          return found ? found.stock : 0;
                        })()}
                      />
                    </label>
                    <label>
                      Gelen Stok
                      <input disabled value={ppeTransactionForm.transactionType === 'STOK_GIRISI' || ppeTransactionForm.transactionType === 'STOGA_IADE' ? ppeTransactionForm.quantity : 0} />
                    </label>
                    <label>
                      Çıkışı Yapılan Stok
                      <input disabled value={ppeTransactionForm.transactionType === 'STOK_CIKISI' || ppeTransactionForm.transactionType === 'HASARLI_HURDA' ? ppeTransactionForm.quantity : 0} />
                    </label>
                    <label>
                      Kalan Stok (İşlem Sonrası)
                      <input
                        disabled
                        value={(() => {
                          const found = ppeInventoryByKey.find((row) => row.projectId === ppeTransactionForm.projectId && row.warehouse === ppeTransactionForm.warehouse && row.category === ppeTransactionForm.category && row.itemName === ppeTransactionForm.itemName && row.unit === ppeTransactionForm.unit);
                          const current = found ? found.stock : 0;
                          if (ppeTransactionForm.transactionType === 'STOK_GIRISI' || ppeTransactionForm.transactionType === 'STOGA_IADE' || ppeTransactionForm.transactionType === 'STOK_DUZELTME') {
                            return Math.max(current + ppeTransactionForm.quantity, 0);
                          }
                          return Math.max(current - ppeTransactionForm.quantity, 0);
                        })()}
                      />
                    </label>
                    <label>
                      Minimum Stok Seviyesi
                      <input type="number" min={0} value={ppeTransactionForm.minimumStockLevel} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, minimumStockLevel: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                    <label>
                      Miktar
                      <input type="number" min={0} value={ppeTransactionForm.quantity} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, quantity: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>Genel</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Durum
                      <select value={ppeTransactionForm.lifecycle} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, lifecycle: event.target.value as PpeTransactionLifecycle }))}>
                        {ppeLifecycleOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                      </select>
                    </label>
                    <label>
                      Sorumlu Kişi
                      <input value={ppeTransactionForm.responsiblePerson} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))} />
                    </label>
                    <label className="full-row">
                      Notlar
                      <textarea rows={3} value={ppeTransactionForm.notes} onChange={(event) => setPpeTransactionForm((prev) => ({ ...prev, notes: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <div className="actions">
                  <button type="button" onClick={savePpeTransaction}>{editingPpeTransactionId ? 'İşlemi Güncelle' : 'İşlemi Kaydet'}</button>
                  {editingPpeTransactionId ? (
                    <button type="button" className="secondary" onClick={resetPpeTransactionForm}>Vazgeç</button>
                  ) : null}
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>KKD Stok Özeti</h2>
              <table>
                <thead>
                  <tr>
                    <th>Proje</th>
                    <th>Tarih</th>
                    <th>KKD Kalemi</th>
                    <th>Gelen Stok</th>
                    <th>Çıkışı Yapılan Stok</th>
                    <th>Kalan Stok</th>
                    <th>Minimum Stok</th>
                    <th>Durum</th>
                  </tr>
                </thead>
                <tbody>
                  {ppeInventoryByKey.length > 0 ? (
                    ppeInventoryByKey.map((row) => {
                      const stockStatus: PpeStockStatus = row.stock <= 0 ? 'STOKTA_YOK' : row.stock <= row.minimumStockLevel ? 'DUSUK_STOK' : 'YETERLI';
                      return (
                        <tr key={`${row.projectId}-${row.warehouse}-${row.category}-${row.itemName}-${row.unit}`}>
                          <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                          <td>{row.lastDate}</td>
                          <td>{row.itemName}</td>
                          <td>{row.incoming + row.returned}</td>
                          <td>{row.outgoing + row.damaged}</td>
                          <td>{row.stock}</td>
                          <td>{row.minimumStockLevel}</td>
                          <td><span className={`status-badge ${ppeStatusClass(stockStatus)}`}>{ppeStatusLabel(stockStatus)}</span></td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={8}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            <section className="panel table-wrap">
              <h2>KKD Kayıtları</h2>
              <table>
                <thead>
                  <tr>
                    <th>İşlem Kimliği</th>
                    <th>İşlem Türü</th>
                    <th>Proje</th>
                    <th>Depo</th>
                    <th>KKD Öğesi</th>
                    <th>Miktar</th>
                    <th>Tarih</th>
                    <th>Sorumlu Kişi</th>
                    <th>İşlem</th>
                  </tr>
                </thead>
                <tbody>
                  {ppeLastTransactions.length > 0 ? (
                    ppeLastTransactions.map((row) => (
                      <tr key={row.id}>
                        <td>{row.transactionId}</td>
                        <td>{ppeTransactionTypeLabel(row.transactionType)}</td>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.warehouse}</td>
                        <td>{row.itemName}</td>
                        <td>{row.quantity} {row.unit}</td>
                        <td>{row.date}</td>
                        <td>{row.responsiblePerson}</td>
                        <td className="actions-cell table-actions-cell">
                          <div className="table-action-group" aria-label={`${row.transactionId} işlemleri`}>
                            <button type="button" className="table-action-button" onClick={() => editPpeTransaction(row.id)}>Düzenle</button>
                            <button type="button" className="table-action-button danger" onClick={() => deletePpeTransaction(row.id)}>Sil</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={9}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'trainings' ? (
          <>
            <section className="panel">
              <h2>Eğitim Yönetimi KPI</h2>
              <div className="equipment-kpi-grid">
                <article className="equipment-kpi-card"><span>Toplam Eğitim Sayısı</span><strong>{trainingSummary.totalTrainings}</strong></article>
                <article className="equipment-kpi-card"><span>Toplam Katılımcı Sayısı</span><strong>{trainingSummary.totalParticipants}</strong></article>
                <article className="equipment-kpi-card"><span>Sertifikalı Çalışan Sayısı</span><strong>{trainingSummary.certifiedEmployees}</strong></article>
                <article className="equipment-kpi-card"><span>Eğitim Bekleyen Çalışan Sayısı</span><strong>{trainingSummary.pendingEmployees}</strong></article>
                <article className="equipment-kpi-card"><span>Toplam Eğitim Maliyeti</span><strong>{trainingSummary.totalCost.toLocaleString('tr-TR')} Rub</strong></article>
                <article className="equipment-kpi-card"><span>Çalışan Başına Ortalama Eğitim Maliyeti</span><strong>{trainingSummary.avgCostPerEmployee.toLocaleString('tr-TR')} Rub</strong></article>
              </div>
            </section>

            <section className="panel">
              <h2>Eğitim Performans Görselleştirmesi</h2>
              <div className="training-chart-grid">
                <article className="chart-card">
                  <strong>Eğitim Türleri</strong>
                  <DashboardChart
                    type="donut"
                    values={trainingTypeDistribution.length > 0 ? trainingTypeDistribution.map((item) => item[1]) : [0]}
                    themeName="operations"
                    donutSize="large"
                  />
                </article>

                <article className="chart-card">
                  <strong>Projeye Göre Eğitim Maliyeti</strong>
                  <DashboardChart
                    type="bar"
                    values={trainingCostByProject.map((item) => item.value)}
                    xLabels={trainingCostByProject.map((item) => item.label.split(' ')[0])}
                    themeName="risk"
                    showTrend={false}
                  />
                </article>

                <article className="chart-card training-horizontal-card">
                  <strong>Projeye Göre Sertifikalı İşgücü</strong>
                  <div className="training-project-bars">
                    {trainingProjectWorkforce.map((item) => (
                      <div key={item.label} className="training-project-row">
                        <span>{item.label}</span>
                        <div className="training-project-track-wrap">
                          <div className="training-project-track training-project-track-total">
                            <div className="training-project-fill training-project-fill-total" style={{ width: `${item.workforce > 0 ? Math.max((item.workforce / trainingProjectWorkforceMax) * 100, 4) : 0}%` }} />
                          </div>
                          <div className="training-project-track training-project-track-certified">
                            <div className="training-project-fill training-project-fill-certified" style={{ width: `${item.certified > 0 ? Math.max((item.certified / trainingProjectWorkforceMax) * 100, 4) : 0}%` }} />
                          </div>
                        </div>
                        <strong>{item.certified}/{item.workforce}</strong>
                      </div>
                    ))}
                  </div>
                </article>

                <article className="chart-card">
                  <strong>Eğitim Durumu</strong>
                  <DashboardChart
                    type="donut"
                    values={[
                      trainingStatusDistribution.TAMAMLANDI,
                      trainingStatusDistribution.PLANLANDI,
                      trainingStatusDistribution.DEVAM_EDIYOR,
                      trainingStatusDistribution.SURESI_DOLDU
                    ]}
                    themeName="compliance"
                    donutSize="large"
                  />
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>Eğitim Veri Girişi</h2>
              <p className="inline-hint">
                Bu modül İK personel yönetimi değil; İSG eğitim yönetimi, sertifikasyon uyumluluğu ve proje eğitim performansının kurumsal takibi için tasarlanmıştır.
              </p>

              <div className="risk-meta-row">
                <div className="risk-meta-card"><span>Eğitim Kimliği</span><strong>{nextTrainingId}</strong></div>
                <div className="risk-meta-card"><span>Bekleyen Çalışan</span><strong>{trainingSummary.pendingEmployees}</strong></div>
                <div className="risk-meta-card"><span>Toplam Maliyet</span><strong>{trainingSummary.totalCost.toLocaleString('tr-TR')} Rub</strong></div>
              </div>

              <div className="training-form-layout">
                <article className="workforce-form-section">
                  <h3>Eğitim Bilgileri</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Proje
                      <select value={trainingForm.projectId} onChange={(event) => setTrainingForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                        {projectCatalog.map((project) => <option key={project.id} value={project.id}>{project.name}</option>)}
                      </select>
                    </label>
                    <label>
                      Eğitim Kimliği (Otomatik Oluşturulur)
                      <input value={nextTrainingId} disabled />
                    </label>
                    <label>
                      Eğitim Türü
                      <select value={trainingForm.trainingType} onChange={(event) => setTrainingForm((prev) => ({ ...prev, trainingType: event.target.value }))}>
                        {trainingTypeOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Eğitim Başlığı
                      <input value={trainingForm.trainingTitle} onChange={(event) => setTrainingForm((prev) => ({ ...prev, trainingTitle: event.target.value }))} />
                    </label>
                    <label>
                      Eğitim Kategorisi
                      <select value={trainingForm.trainingCategory} onChange={(event) => setTrainingForm((prev) => ({ ...prev, trainingCategory: event.target.value }))}>
                        {trainingCategoryOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Eğitim Tarihi
                      <input type="date" value={trainingForm.trainingDate} onChange={(event) => setTrainingForm((prev) => ({ ...prev, trainingDate: event.target.value }))} />
                    </label>
                    <label className="full-row">
                      Eğitimi Veren Kurum/Kişi
                      <input value={trainingForm.provider} onChange={(event) => setTrainingForm((prev) => ({ ...prev, provider: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>Çalışan Bilgileri</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Departman
                      <select value={trainingForm.department} onChange={(event) => setTrainingForm((prev) => ({ ...prev, department: event.target.value }))}>
                        {trainingDepartmentOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Pozisyon
                      <select value={trainingForm.position} onChange={(event) => setTrainingForm((prev) => ({ ...prev, position: event.target.value }))}>
                        {trainingPositionOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Proje Çalışan Sayısı
                      <input type="number" min={0} value={trainingForm.projectEmployeeCount} onChange={(event) => setTrainingForm((prev) => ({ ...prev, projectEmployeeCount: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                    <label>
                      Sertifikalı Çalışan Sayısı
                      <input type="number" min={0} value={trainingForm.certifiedEmployeeCount} onChange={(event) => setTrainingForm((prev) => ({ ...prev, certifiedEmployeeCount: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>Sertifikasyon</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Sertifika Gerekliliği
                      <select value={trainingForm.certificateRequired ? 'EVET' : 'HAYIR'} onChange={(event) => setTrainingForm((prev) => ({ ...prev, certificateRequired: event.target.value === 'EVET' }))}>
                        <option value="EVET">Evet</option>
                        <option value="HAYIR">Hayır</option>
                      </select>
                    </label>
                    <label>
                      Sertifika Geçerlilik Tarihi
                      <input type="date" value={trainingForm.certificateValidityDate} onChange={(event) => setTrainingForm((prev) => ({ ...prev, certificateValidityDate: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>Finansal</h3>
                  <div className="equipment-form-grid">
                    <label>
                      Toplam Eğitim Maliyeti
                      <input type="number" min={0} value={trainingForm.totalTrainingCost} onChange={(event) => setTrainingForm((prev) => ({ ...prev, totalTrainingCost: Math.max(0, Number(event.target.value) || 0) }))} />
                    </label>
                    <label>
                      Çalışan Başına Maliyet
                      <input
                        value={`${trainingForm.projectEmployeeCount > 0 ? Math.round(trainingForm.totalTrainingCost / trainingForm.projectEmployeeCount).toLocaleString('tr-TR') : 0} Rub`}
                        disabled
                      />
                    </label>
                  </div>
                </article>

                <article className="workforce-form-section">
                  <h3>{language === 'ru' ? 'Общие сведения' : 'Genel'}</h3>
                  <div className="equipment-form-grid">
                    <label>
                      {language === 'ru' ? 'Статус' : 'Durum'}
                      <select value={trainingForm.status} onChange={(event) => setTrainingForm((prev) => ({ ...prev, status: event.target.value as TrainingStatus }))}>
                        <option value="TAMAMLANDI">{language === 'ru' ? 'Завершено' : 'Tamamlandı'}</option>
                        <option value="PLANLANDI">{language === 'ru' ? 'Запланировано' : 'Planlandı'}</option>
                        <option value="DEVAM_EDIYOR">{language === 'ru' ? 'В процессе' : 'Devam Ediyor'}</option>
                        <option value="SURESI_DOLDU">{language === 'ru' ? 'Срок истек' : 'Süresi Doldu'}</option>
                      </select>
                    </label>
                    <label>
                      {language === 'ru' ? 'Вложения' : 'Ekler'}
                      {language === 'ru' ? (
                        <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " multiple onFilesChange={(files) => setTrainingForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }))} />
                      ) : (
                        <input type="file" multiple onChange={(event) => {
                          const files = Array.from(event.target.files ?? []);
                          setTrainingForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }));
                        }} />
                      )}
                    </label>
                    <label>
                      {language === 'ru' ? 'Список участников обучения' : 'Eğitim Katılımcı Listesi'}
                      {language === 'ru' ? (
                        <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " multiple onFilesChange={(files) => setTrainingForm((prev) => ({ ...prev, participantList: files.map((file) => file.name) }))} />
                      ) : (
                        <input type="file" multiple onChange={(event) => {
                          const files = Array.from(event.target.files ?? []);
                          setTrainingForm((prev) => ({ ...prev, participantList: files.map((file) => file.name) }));
                        }} />
                      )}
                    </label>
                    <label>
                      {language === 'ru' ? 'Сертификаты' : 'Sertifikalar'}
                      {language === 'ru' ? (
                        <CustomFileUpload buttonLabel="Выбрать файлы" emptyLabel="Файлы не выбраны" singleLabel="Выбран файл: " multipleLabel="Выбрано файлов: " multiple onFilesChange={(files) => setTrainingForm((prev) => ({ ...prev, certificates: files.map((file) => file.name) }))} />
                      ) : (
                        <input type="file" multiple onChange={(event) => {
                          const files = Array.from(event.target.files ?? []);
                          setTrainingForm((prev) => ({ ...prev, certificates: files.map((file) => file.name) }));
                        }} />
                      )}
                    </label>
                    <label className="full-row">
                      {language === 'ru' ? 'Примечания' : 'Notlar'}
                      <textarea rows={3} value={trainingForm.notes} onChange={(event) => setTrainingForm((prev) => ({ ...prev, notes: event.target.value }))} />
                    </label>
                  </div>
                </article>

                <div className="actions">
                  <button type="button" onClick={saveTrainingEntry}>{editingTrainingId ? (language === 'ru' ? 'Обновить' : 'Güncelle') : (language === 'ru' ? 'Сохранить' : 'Kaydet')}</button>
                  {editingTrainingId ? (
                    <button type="button" className="secondary" onClick={resetTrainingForm}>{language === 'ru' ? 'Отмена' : 'Vazgeç'}</button>
                  ) : null}
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{language === 'ru' ? 'Записи обучения' : 'Eğitim Kayıtları'}</h2>
              <table>
                <thead>
                  <tr>
                    <th>Proje</th>
                    <th>Eğitim Tarihi</th>
                    <th>Eğitim Başlığı</th>
                    <th>Eğitim Türü</th>
                    <th>Katılımcılar</th>
                    <th>Sertifikalı Çalışanlar</th>
                    <th>Toplam Maliyet</th>
                    <th>Durum</th>
                    <th>İşlem</th>
                  </tr>
                </thead>
                <tbody>
                  {trainingScopedRows.length > 0 ? (
                    trainingScopedRows.map((row) => (
                      <tr key={row.id}>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.trainingDate}</td>
                        <td>{row.trainingTitle}</td>
                        <td>{row.trainingType}</td>
                        <td>{row.projectEmployeeCount}</td>
                        <td>{row.certifiedEmployeeCount}</td>
                        <td>{row.totalTrainingCost.toLocaleString('tr-TR')} Rub</td>
                        <td>
                          <div className="training-status-stack">
                            <span className={`traffic-chip ${trainingCertificateTrafficClass(row)}`}>{trainingCertificateTrafficLabel(row)}</span>
                          </div>
                        </td>
                        <td className="actions-cell training-actions-cell">
                          <div className="table-action-group" aria-label={`${row.trainingTitle} işlemleri`}>
                            <button type="button" className="table-action-button" onClick={() => editTrainingEntry(row.id)}>Düzenle</button>
                            <button type="button" className="table-action-button danger" onClick={() => deleteTrainingEntry(row.id)}>Sil</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={9}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'equipment-management' ? (
          <>
            <section className="panel">
              <h2>{language === 'ru' ? 'Панель KPI по оборудованию' : 'Ekipman Yönetimi KPI'}</h2>
              <div className="equipment-kpi-grid">
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Всего оборудования' : 'Toplam Ekipman'}</span><strong>{equipmentSummary.total}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Активное оборудование' : 'Aktif Ekipman'}</span><strong>{equipmentSummary.active}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Предстоящая проверка' : 'Yaklaşan Muayene'}</span><strong>{equipmentSummary.upcoming}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Просроченная проверка' : 'Gecikmiş Muayene'}</span><strong>{equipmentSummary.overdue}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Сертификаты с истекающим сроком' : 'Süresi Dolacak Sertifika'}</span><strong>{equipmentSummary.expiringCertificates}</strong></article>
                <article className="equipment-kpi-card"><span>{language === 'ru' ? 'Оборудование вне эксплуатации' : 'Servis Dışı Ekipman'}</span><strong>{equipmentSummary.outOfService}</strong></article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'График соответствия проверок оборудования' : 'Ekipman Muayene Uygunluk Grafiği'}</h2>
              <div className="equipment-chart-grid">
                <article className="chart-card">
                  <strong>{language === 'ru' ? 'Распределение соответствия проверок' : 'Muayene Uygunluk Dağılımı'}</strong>
                  <DashboardChart type="donut" values={[equipmentSummary.compliant, equipmentSummary.upcoming, equipmentSummary.overdue]} themeName="compliance" />
                </article>
                <article className="chart-card">
                  <strong>{language === 'ru' ? 'Соответствие и критические состояния' : 'Uygunluk ve Kritik Durumlar'}</strong>
                  <DashboardChart
                    type="bar"
                    values={[
                      equipmentSummary.compliant,
                      equipmentSummary.upcoming,
                      equipmentSummary.overdue,
                      equipmentSummary.expiringCertificates,
                      equipmentSummary.outOfService
                    ]}
                    themeName="operations"
                    xLabels={[
                      language === 'ru' ? 'Соответствует' : 'Uygun',
                      language === 'ru' ? 'Скоро истекает' : 'Yaklaşan',
                      language === 'ru' ? 'Просрочено' : 'Gecikmiş',
                      language === 'ru' ? 'Сертификаты' : 'Sertifika',
                      language === 'ru' ? 'Вне эксплуатации' : 'Servis Dışı'
                    ]}
                  />
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>{language === 'ru' ? 'Регистрация оборудования' : 'Ekipman Veri Girişi'}</h2>
              <p className="inline-hint">
                {language === 'ru'
                  ? 'Данная форма предназначена для регистрации оборудования, контроля результатов проверок, срока действия сертификатов, соблюдения нормативных требований и управления техническим состоянием оборудования.'
                  : 'Bu form ekipman muayene uygunluğu, sertifika geçerliliği, yasal takip ve ekipman durum yönetimini tek ekranda toplar.'}
              </p>

              <div className="risk-meta-row">
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Идентификатор оборудования' : 'Ekipman ID'}</span>
                  <strong>{nextEquipmentId}</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Процент соответствия проверкам' : 'Muayene Uygunluk Oranı'}</span>
                  <strong>{equipmentSummary.complianceRate}%</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{language === 'ru' ? 'Просроченные проверки' : 'Gecikmiş Muayene'}</span>
                  <strong>{equipmentSummary.overdue}</strong>
                </div>
              </div>

              <div className="equipment-form-grid">
                <label>
                  {language === 'ru' ? 'Проект' : 'Proje'}
                  <select value={equipmentForm.projectId} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, projectId: event.target.value }))}>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Идентификатор оборудования (автоматически)' : 'Ekipman ID (Otomatik)'}
                  <input value={nextEquipmentId} disabled />
                </label>

                <label>
                  {language === 'ru' ? 'Наименование оборудования' : 'Ekipman Adı'}
                  <input value={equipmentForm.equipmentName} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, equipmentName: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Тип оборудования' : 'Ekipman Türü'}
                  <input value={equipmentForm.equipmentType} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, equipmentType: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Марка / Модель' : 'Marka / Model'}
                  <input value={equipmentForm.brandModel} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, brandModel: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Серийный номер' : 'Seri Numarası'}
                  <input value={equipmentForm.serialNumber} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, serialNumber: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Местоположение' : 'Konum'}
                  <input value={equipmentForm.location} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, location: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Ответственное лицо' : 'Sorumlu Kişi'}
                  <input value={equipmentForm.responsiblePerson} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Дата последней проверки' : 'Son Muayene Tarihi'}
                  <input type="date" value={equipmentForm.lastInspectionDate} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, lastInspectionDate: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Дата следующей проверки' : 'Sonraki Muayene Tarihi'}
                  <input type="date" value={equipmentForm.nextInspectionDate} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, nextInspectionDate: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Статус проверки' : 'Muayene Durumu'}
                  <select
                    value={equipmentForm.inspectionStatus}
                    onChange={(event) => setEquipmentForm((prev) => ({ ...prev, inspectionStatus: event.target.value as EquipmentInspectionStatus }))}
                  >
                    <option value="COMPLIANT">{language === 'ru' ? 'Соответствует' : 'Uygun'}</option>
                    <option value="UPCOMING">{language === 'ru' ? 'Скоро истекает' : 'Yaklaşan'}</option>
                    <option value="OVERDUE">{language === 'ru' ? 'Просрочено' : 'Gecikmiş'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Номер сертификата' : 'Sertifika Numarası'}
                  <input value={equipmentForm.certificateNumber} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, certificateNumber: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Дата окончания сертификата' : 'Sertifika Bitiş Tarihi'}
                  <input type="date" value={equipmentForm.certificateExpiryDate} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, certificateExpiryDate: event.target.value }))} />
                </label>

                <label>
                  {language === 'ru' ? 'Статус оборудования' : 'Ekipman Durumu'}
                  <select value={equipmentForm.equipmentStatus} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, equipmentStatus: event.target.value as EquipmentStatus }))}>
                    <option value="ACTIVE">{language === 'ru' ? 'Активное' : 'Aktif'}</option>
                    <option value="OUT_OF_SERVICE">{language === 'ru' ? 'Оборудование вне эксплуатации' : 'Servis Dışı'}</option>
                    <option value="UNDER_MAINTENANCE">{language === 'ru' ? 'На техническом обслуживании' : 'Bakımda'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Уровень риска' : 'Risk Seviyesi'}
                  <select value={equipmentForm.riskLevel} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, riskLevel: event.target.value as EquipmentRiskLevel }))}>
                    <option value="LOW">{language === 'ru' ? 'Низкий' : 'Düşük'}</option>
                    <option value="MEDIUM">{language === 'ru' ? 'Средний' : 'Orta'}</option>
                    <option value="HIGH">{language === 'ru' ? 'Высокий' : 'Yüksek'}</option>
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Вложения' : 'Ekler'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      multiple
                      onFilesChange={(files) => setEquipmentForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEquipmentForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label>
                  {language === 'ru' ? 'Отчёт о проверке' : 'Muayene Raporu'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      multiple
                      onFilesChange={(files) => setEquipmentForm((prev) => ({ ...prev, inspectionReports: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEquipmentForm((prev) => ({ ...prev, inspectionReports: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label>
                  {language === 'ru' ? 'Фотография оборудования' : 'Ekipman Fotoğrafı'}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      accept="image/*"
                      multiple
                      onFilesChange={(files) => setEquipmentForm((prev) => ({ ...prev, equipmentPhotos: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      accept="image/*"
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setEquipmentForm((prev) => ({ ...prev, equipmentPhotos: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label className="full-row">
                  {language === 'ru' ? 'Примечания' : 'Notlar'}
                  <textarea rows={3} value={equipmentForm.notes} onChange={(event) => setEquipmentForm((prev) => ({ ...prev, notes: event.target.value }))} />
                </label>

                <div className="full-row actions">
                  <button type="button" onClick={saveEquipmentEntry}>{language === 'ru' ? 'Сохранить запись' : 'Kaydet'}</button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{language === 'ru' ? 'Реестр оборудования' : 'Ekipman Kayıtları'}</h2>
              <table>
                <thead>
                  <tr>
                    <th>{language === 'ru' ? 'Идентификатор' : 'Ekipman ID'}</th>
                    <th>{language === 'ru' ? 'Проект' : 'Proje'}</th>
                    <th>{language === 'ru' ? 'Наименование оборудования' : 'Ekipman Adı'}</th>
                    <th>{language === 'ru' ? 'Тип' : 'Tür'}</th>
                    <th>{language === 'ru' ? 'Следующая проверка' : 'Sonraki Muayene'}</th>
                    <th>{language === 'ru' ? 'Статус проверки' : 'Muayene Durumu'}</th>
                    <th>{language === 'ru' ? 'Срок действия сертификата' : 'Sertifika Bitişi'}</th>
                    <th>{language === 'ru' ? 'Статус оборудования' : 'Ekipman Durumu'}</th>
                    <th>{language === 'ru' ? 'Уровень риска' : 'Risk'}</th>
                    <th>{language === 'ru' ? 'История изменений' : 'Trafik'}</th>
                    <th>{language === 'ru' ? 'Ответственное лицо' : 'Sorumlu'}</th>
                  </tr>
                </thead>
                <tbody>
                  {equipmentScopedRows.length > 0 ? (
                    equipmentScopedRows.map((row) => (
                      <tr key={row.id} className={selectedEquipmentId === row.id ? 'risk-row-selected' : ''} onClick={() => setSelectedEquipmentId(row.id)}>
                        <td>{row.equipmentId}</td>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.equipmentName}</td>
                        <td>{row.equipmentType}</td>
                        <td>{row.nextInspectionDate}</td>
                        <td>{equipmentInspectionStatusLabel(row.inspectionStatus)}</td>
                        <td>{row.certificateExpiryDate}</td>
                        <td>
                          <span className={`equipment-status-chip ${equipmentStatusClassName(row.equipmentStatus)}`}>
                            {equipmentStatusLabel(row.equipmentStatus)}
                          </span>
                        </td>
                        <td>{equipmentRiskLabel(row.riskLevel)}</td>
                        <td><span className={`traffic-chip ${getEquipmentTrafficClass(row)}`}>{getEquipmentTrafficLabel(row)}</span></td>
                        <td>{row.responsiblePerson}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={11}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            {selectedEquipmentRecord ? (
              <section className="panel risk-detail-panel">
                <div className="risk-detail-head">
                  <div>
                    <h2>{language === 'ru' ? 'Карточка оборудования' : 'Ekipman Detayı'} - {selectedEquipmentRecord.equipmentId}</h2>
                    <p className="risk-detail-subtitle">
                      {selectedEquipmentRecord.equipmentName} - {selectedEquipmentRecord.equipmentType}
                    </p>
                  </div>
                  <div className="risk-detail-badges">
                    <span className={`status-badge ${getEquipmentTrafficClass(selectedEquipmentRecord)}`}>{equipmentInspectionStatusLabel(selectedEquipmentRecord.inspectionStatus)}</span>
                    <span className={`status-badge ${equipmentStatusBadgeClassName(selectedEquipmentRecord)} ${equipmentStatusClassName(selectedEquipmentRecord.equipmentStatus)}`}>{equipmentStatusLabel(selectedEquipmentRecord.equipmentStatus)}</span>
                    <span className="risk-band-badge risk-band-medium">{language === 'ru' ? 'Уровень риска' : 'Risk Seviyesi'}: {equipmentRiskLabel(selectedEquipmentRecord.riskLevel)}</span>
                    <span className={`traffic-chip ${getEquipmentTrafficClass(selectedEquipmentRecord)}`}>
                      {language === 'ru' ? 'Светофор статуса' : 'Trafik Işığı'}: {getEquipmentTrafficLabel(selectedEquipmentRecord)}
                    </span>
                  </div>
                </div>

                <div className="risk-detail-grid">
                  <article>
                    <h3>{language === 'ru' ? 'Идентификация и местоположение' : 'Kimlik ve Lokasyon'}</h3>
                    <p><strong>{language === 'ru' ? 'Марка / Модель' : 'Marka / Model'}:</strong> {selectedEquipmentRecord.brandModel || '-'}</p>
                    <p><strong>{language === 'ru' ? 'Серийный номер' : 'Seri Numarası'}:</strong> {selectedEquipmentRecord.serialNumber || '-'}</p>
                    <p><strong>{language === 'ru' ? 'Местоположение' : 'Konum'}:</strong> {selectedEquipmentRecord.location || '-'}</p>
                  </article>
                  <article>
                    <h3>{language === 'ru' ? 'Проверка и соответствие' : 'Muayene ve Uyum'}</h3>
                    <p><strong>{language === 'ru' ? 'Дата последней проверки' : 'Son Muayene'}:</strong> {selectedEquipmentRecord.lastInspectionDate}</p>
                    <p><strong>{language === 'ru' ? 'Дата следующей проверки' : 'Sonraki Muayene'}:</strong> {selectedEquipmentRecord.nextInspectionDate}</p>
                    <p><strong>{language === 'ru' ? 'Статус проверки' : 'Muayene Durumu'}:</strong> {equipmentInspectionStatusLabel(selectedEquipmentRecord.inspectionStatus)}</p>
                  </article>
                  <article>
                    <h3>{language === 'ru' ? 'Сертификация' : 'Sertifikasyon'}</h3>
                    <p><strong>{language === 'ru' ? 'Номер сертификата' : 'Sertifika No'}:</strong> {selectedEquipmentRecord.certificateNumber || '-'}</p>
                    <p><strong>{language === 'ru' ? 'Дата окончания сертификата' : 'Bitiş Tarihi'}:</strong> {selectedEquipmentRecord.certificateExpiryDate}</p>
                    <p><strong>{language === 'ru' ? 'Осталось дней' : 'Kalan Gün'}:</strong> {daysUntil(selectedEquipmentRecord.certificateExpiryDate)}</p>
                  </article>
                  <article>
                    <h3>{language === 'ru' ? 'Ответственность' : 'Sorumluluk'}</h3>
                    <p><strong>{language === 'ru' ? 'Ответственное лицо' : 'Sorumlu'}:</strong> {selectedEquipmentRecord.responsiblePerson}</p>
                    <p><strong>{language === 'ru' ? 'Статус оборудования' : 'Ekipman Durumu'}:</strong> <span className={equipmentStatusClassName(selectedEquipmentRecord.equipmentStatus)}>{equipmentStatusLabel(selectedEquipmentRecord.equipmentStatus)}</span></p>
                    <p><strong>{language === 'ru' ? 'Уровень риска' : 'Risk Seviyesi'}:</strong> {equipmentRiskLabel(selectedEquipmentRecord.riskLevel)}</p>
                  </article>
                </div>

                <div className="risk-detail-files">
                  <div>
                    <h3>{language === 'ru' ? 'Вложения' : 'Ekler'}</h3>
                    {selectedEquipmentRecord.attachments.length > 0 ? (
                      <ul>{selectedEquipmentRecord.attachments.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Вложения отсутствуют.' : 'Ek bulunmuyor.'}</p>
                    )}
                  </div>
                  <div>
                    <h3>{language === 'ru' ? 'Отчёты о проверке' : 'Muayene Raporları'}</h3>
                    {selectedEquipmentRecord.inspectionReports.length > 0 ? (
                      <ul>{selectedEquipmentRecord.inspectionReports.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Отчёты отсутствуют.' : 'Rapor bulunmuyor.'}</p>
                    )}
                  </div>
                  <div>
                    <h3>{language === 'ru' ? 'Фотографии оборудования' : 'Ekipman Fotoğrafları'}</h3>
                    {selectedEquipmentRecord.equipmentPhotos.length > 0 ? (
                      <ul>{selectedEquipmentRecord.equipmentPhotos.map((fileName) => <li key={fileName}>{fileName}</li>)}</ul>
                    ) : (
                      <p>{language === 'ru' ? 'Фотографии отсутствуют.' : 'Fotoğraf bulunmuyor.'}</p>
                    )}
                  </div>
                </div>

                <div>
                  <h3>{language === 'ru' ? 'Примечания' : 'Notlar'}</h3>
                  <p>{selectedEquipmentRecord.notes || '-'}</p>
                </div>
              </section>
            ) : null}
          </>
        ) : null}

        {activeModule === 'incidents' ? (
          <>
            <section className="panel">
              <h2>{language === 'ru' ? 'Регистрация инцидентов' : 'Olaylar Veri Girişi'}</h2>
              <p className="inline-hint">
                {language === 'ru'
                  ? 'Количество инцидентов показывает общее число зарегистрированных инцидентов на объекте. Потерянные рабочие дни показывают количество рабочих дней, утраченных вследствие инцидентов.'
                  : 'Olay Sayısı sahaya ait toplam olay adedini, Kayıp İş Günü ise olaydan kaynaklı iş gücü kaybını gösterir.'}
              </p>

              <div className="incident-overview-grid">
                <article className="incident-overview-card">
                  <span>{language === 'ru' ? 'Количество инцидентов' : 'Toplam Olay'}</span>
                  <strong>{incidentSummary.totalIncidents}</strong>
                </article>
                <article className="incident-overview-card">
                  <span>{language === 'ru' ? 'Потерянные рабочие дни' : 'Kayıp İş Günü'}</span>
                  <strong>{incidentSummary.totalLostDays}</strong>
                </article>
                <article className="incident-overview-card">
                  <span>{language === 'ru' ? 'Открыто / В работе' : 'Açık / Devam'}</span>
                  <strong>{incidentSummary.openCases} / {incidentSummary.inProgressCases}</strong>
                </article>
                <article className="incident-overview-card">
                  <span>{language === 'ru' ? 'Закрытые записи' : 'Kapalı Kayıt'}</span>
                  <strong>{incidentSummary.closedCases}</strong>
                </article>
              </div>

              <div className="incident-form-grid">
                <label>
                  {language === 'ru' ? 'Проект' : 'Proje'}
                  <select
                    value={incidentForm.projectId}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, projectId: event.target.value }))}
                  >
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>

                <label>
                  {language === 'ru' ? 'Дата' : 'Tarih'}
                  <input
                    type="date"
                    value={incidentForm.date}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, date: event.target.value }))}
                  />
                </label>

                <label className="full-row">
                  {language === 'ru' ? 'Заголовок / Описание' : 'Başlık / Açıklama'}
                  <input
                    value={incidentForm.title}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, title: event.target.value }))}
                    placeholder={language === 'ru' ? 'Например: травма руки в зоне погрузки' : 'Orn: Yükleme alanında el yaralanması'}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Количество инцидентов' : 'Olay Sayısı'}
                  <input
                    type="number"
                    min={1}
                    value={incidentForm.incidentCount}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, incidentCount: Math.max(1, Number(event.target.value) || 1) }))}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Потерянные рабочие дни' : 'Kayıp İş Günü'}
                  <input
                    type="number"
                    min={0}
                    value={incidentForm.lostWorkDays}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, lostWorkDays: Math.max(0, Number(event.target.value) || 0) }))}
                  />
                </label>

                <label>
                  {language === 'ru' ? 'Статус' : 'Durum'}
                  <select
                    value={incidentForm.status}
                    onChange={(event) => setIncidentForm((prev) => ({ ...prev, status: event.target.value as Status }))}
                  >
                    <option value="OPEN">{language === 'ru' ? 'Открыто' : 'Açık'}</option>
                    <option value="IN_PROGRESS">{language === 'ru' ? 'В работе' : 'Devam Ediyor'}</option>
                    <option value="CLOSED">{language === 'ru' ? 'Закрыто' : 'Kapalı'}</option>
                  </select>
                </label>

                <div className="full-row actions">
                  <button type="button" onClick={saveIncidentEntry}>
                    {language === 'ru'
                      ? 'Сохранить запись'
                      : editingIncidentIndex !== null
                        ? 'Güncelle'
                        : 'Kaydet'}
                  </button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{language === 'ru' ? 'Реестр инцидентов' : 'Olaylar Kayıtlar'}</h2>
              <table>
                <thead>
                  <tr>
                    <th>{language === 'ru' ? 'Проект' : 'Proje'}</th>
                    <th>{language === 'ru' ? 'Дата' : 'Tarih'}</th>
                    <th>{language === 'ru' ? 'Заголовок' : 'Başlık'}</th>
                    <th>{language === 'ru' ? 'Количество инцидентов' : 'Olay Sayısı'}</th>
                    <th>{language === 'ru' ? 'Потерянные рабочие дни' : 'Kayıp İş Günü'}</th>
                    <th>{language === 'ru' ? 'Статус' : 'Durum'}</th>
                    <th>{language === 'ru' ? 'Действия' : 'İşlemler'}</th>
                  </tr>
                </thead>
                <tbody>
                  {incidentScopedRecords.length > 0 ? (
                    incidentScopedRecords.map(({ row, index }) => (
                      <tr key={`${row.title}-${row.date}-${index}`}>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.date}</td>
                        <td>{localizeText(row.title, language)}</td>
                        <td>{row.valueA}</td>
                        <td>{row.valueB}</td>
                        <td>
                          <span className={`status-badge ${getStatusBadgeClass(row.status)}`}>{localizeStatus(row.status)}</span>
                        </td>
                        <td className="actions-cell table-actions-cell">
                          <div className="table-action-group" aria-label={language === 'ru' ? `Действия по записи ${localizeText(row.title, language)}` : `${localizeText(row.title, language)} işlemleri`}>
                            <button type="button" className="table-action-button" onClick={() => editIncidentRecord(index)}>{language === 'ru' ? 'Редактировать' : 'Düzenle'}</button>
                            <button type="button" className="table-action-button danger" onClick={() => deleteIncidentRecord(index)}>{language === 'ru' ? 'Удалить' : 'Sil'}</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={7}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>
          </>
        ) : null}

        {activeModule === 'risk-assessments' ? (
          <>
            <section className="panel">
              <h2>{riskCopy.formTitle}</h2>
              <p className="inline-hint">
                {riskCopy.introLine1}
                {' '}
                {riskCopy.introLine2}
              </p>
              <div className="risk-meta-row">
                <div className="risk-meta-card">
                  <span>{riskCopy.riskId}</span>
                  <strong>{nextRiskId}</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{riskCopy.initialRiskScore}</span>
                  <strong>{riskForm.likelihood * riskForm.severity}</strong>
                </div>
                <div className="risk-meta-card">
                  <span>{riskCopy.residualRiskScore}</span>
                  <strong>{riskForm.residualLikelihood * riskForm.residualSeverity}</strong>
                </div>
              </div>

              <div className="risk-form-grid">
                <label>
                  {riskCopy.project}
                  <select
                    value={riskForm.projectId}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, projectId: event.target.value }))}
                  >
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>

                <label>
                  {riskCopy.departmentActivity}
                  <input
                    value={riskForm.departmentActivity}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, departmentActivity: event.target.value }))}
                    placeholder={riskCopy.departmentActivityPlaceholder}
                  />
                </label>

                <label>
                  {riskCopy.assessmentDate}
                  <input
                    type="date"
                    value={riskForm.assessmentDate}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, assessmentDate: event.target.value }))}
                  />
                </label>

                <label>
                  {riskCopy.targetDate}
                  <input
                    type="date"
                    value={riskForm.targetCompletionDate}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, targetCompletionDate: event.target.value }))}
                  />
                </label>

                <label className="full-row">
                  {riskCopy.hazard}
                  <textarea
                    rows={2}
                    value={riskForm.hazard}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, hazard: event.target.value }))}
                    placeholder={riskCopy.hazardPlaceholder}
                  />
                </label>

                <label className="full-row">
                  {riskCopy.potentialConsequence}
                  <textarea
                    rows={2}
                    value={riskForm.potentialConsequence}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, potentialConsequence: event.target.value }))}
                  />
                </label>

                <label className="full-row">
                  {riskCopy.existingControls}
                  <textarea
                    rows={2}
                    value={riskForm.existingControls}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, existingControls: event.target.value }))}
                  />
                </label>

                <label className="full-row">
                  {riskCopy.recommendedControls}
                  <textarea
                    rows={2}
                    value={riskForm.recommendedControls}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, recommendedControls: event.target.value }))}
                  />
                </label>

                <label>
                  {riskCopy.likelihood}
                  <input
                    type="number"
                    min={1}
                    max={5}
                    value={riskForm.likelihood}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, likelihood: Math.max(1, Math.min(5, Number(event.target.value) || 1)) }))}
                  />
                </label>

                <label>
                  {riskCopy.severity}
                  <input
                    type="number"
                    min={1}
                    max={5}
                    value={riskForm.severity}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, severity: Math.max(1, Math.min(5, Number(event.target.value) || 1)) }))}
                  />
                </label>

                <label>
                  {riskCopy.residualLikelihood}
                  <input
                    type="number"
                    min={1}
                    max={5}
                    value={riskForm.residualLikelihood}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, residualLikelihood: Math.max(1, Math.min(5, Number(event.target.value) || 1)) }))}
                  />
                </label>

                <label>
                  {riskCopy.residualSeverity}
                  <input
                    type="number"
                    min={1}
                    max={5}
                    value={riskForm.residualSeverity}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, residualSeverity: Math.max(1, Math.min(5, Number(event.target.value) || 1)) }))}
                  />
                </label>

                <label>
                  {riskCopy.responsiblePerson}
                  <input
                    value={riskForm.responsiblePerson}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))}
                  />
                </label>

                <label>
                  {riskCopy.status}
                  <select
                    value={riskForm.status}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, status: event.target.value as Status }))}
                  >
                    <option value="OPEN">{language === 'ru' ? 'Открыто' : 'Açık'}</option>
                    <option value="IN_PROGRESS">{language === 'ru' ? 'В работе' : 'Devam Ediyor'}</option>
                    <option value="CLOSED">{language === 'ru' ? 'Закрыто' : 'Kapalı'}</option>
                  </select>
                </label>

                <label>
                  {riskCopy.attachments}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      multiple
                      onFilesChange={(files) => setRiskForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setRiskForm((prev) => ({ ...prev, attachments: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label>
                  {riskCopy.photos}
                  {language === 'ru' ? (
                    <CustomFileUpload
                      buttonLabel="Выбрать файлы"
                      emptyLabel="Файлы не выбраны"
                      singleLabel="Выбран файл: "
                      multipleLabel="Выбрано файлов: "
                      accept="image/*"
                      multiple
                      onFilesChange={(files) => setRiskForm((prev) => ({ ...prev, photos: files.map((file) => file.name) }))}
                    />
                  ) : (
                    <input
                      type="file"
                      multiple
                      accept="image/*"
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        setRiskForm((prev) => ({ ...prev, photos: files.map((file) => file.name) }));
                      }}
                    />
                  )}
                </label>

                <label className="full-row">
                  {riskCopy.notes}
                  <textarea
                    rows={3}
                    value={riskForm.notes}
                    onChange={(event) => setRiskForm((prev) => ({ ...prev, notes: event.target.value }))}
                  />
                </label>

                <div className="full-row actions">
                  <button type="button" onClick={saveRiskEntry}>{riskCopy.save}</button>
                </div>
              </div>
            </section>

            <section className="panel table-wrap">
              <h2>{riskCopy.tableTitle}</h2>
              <table>
                <thead>
                  <tr>
                    <th>{riskCopy.riskId}</th>
                    <th>{riskCopy.project}</th>
                    <th>{riskCopy.activity}</th>
                    <th>{riskCopy.hazard}</th>
                    <th>{riskCopy.initialRiskScore}</th>
                    <th>{riskCopy.residualRiskScore}</th>
                    <th>{riskCopy.status}</th>
                    <th>{riskCopy.responsiblePerson}</th>
                  </tr>
                </thead>
                <tbody>
                  {scopedRiskRows.length > 0 ? (
                    scopedRiskRows.map((row) => (
                      <tr key={row.id} className={selectedRiskId === row.id ? 'risk-row-selected' : ''} onClick={() => setSelectedRiskId(row.id)}>
                        <td>{row.riskId}</td>
                        <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                        <td>{row.departmentActivity}</td>
                        <td>{row.hazard}</td>
                        <td>{row.initialRiskScore}</td>
                        <td>{row.residualRiskScore}</td>
                        <td>{localizeStatus(row.status)}</td>
                        <td>{row.responsiblePerson}</td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={8}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            {selectedRiskRecord ? (
              <section className="panel risk-detail-panel">
                <div className="risk-detail-head">
                  <div>
                    <h2>{riskCopy.detailTitle} - {selectedRiskRecord.riskId}</h2>
                    <p className="risk-detail-subtitle">
                      {projectCatalog.find((project) => project.id === selectedRiskRecord.projectId)?.name ?? selectedRiskRecord.projectId} - {selectedRiskRecord.departmentActivity}
                    </p>
                  </div>
                  <div className="risk-detail-badges">
                    <span className={`status-badge ${getStatusBadgeClass(selectedRiskRecord.status)}`}>{localizeStatus(selectedRiskRecord.status)}</span>
                    <span className={`risk-band-badge ${getRiskBand(selectedRiskRecord.initialRiskScore).className}`}>
                      {riskCopy.initialLevel}: {getRiskBand(selectedRiskRecord.initialRiskScore).label} ({getRiskBand(selectedRiskRecord.initialRiskScore).range})
                    </span>
                    <span className={`risk-band-badge ${getRiskBand(selectedRiskRecord.residualRiskScore).className}`}>
                      {riskCopy.residualLevel}: {getRiskBand(selectedRiskRecord.residualRiskScore).label} ({getRiskBand(selectedRiskRecord.residualRiskScore).range})
                    </span>
                  </div>
                </div>

                <div className="risk-score-strip">
                  <article>
                    <span>{riskCopy.initialRiskScore}</span>
                    <strong>{selectedRiskRecord.initialRiskScore}</strong>
                    <small>{selectedRiskRecord.likelihood} x {selectedRiskRecord.severity}</small>
                  </article>
                  <article>
                    <span>{riskCopy.residualRiskScore}</span>
                    <strong>{selectedRiskRecord.residualRiskScore}</strong>
                    <small>{selectedRiskRecord.residualLikelihood} x {selectedRiskRecord.residualSeverity}</small>
                  </article>
                  <article>
                    <span>{riskCopy.responsiblePerson}</span>
                    <strong>{selectedRiskRecord.responsiblePerson}</strong>
                    <small>{riskCopy.due}: {selectedRiskRecord.targetCompletionDate}</small>
                  </article>
                </div>

                <div className="risk-detail-grid">
                  <article>
                    <h3>{riskCopy.hazardAndConsequence}</h3>
                    <p><strong>{riskCopy.hazard}:</strong> {selectedRiskRecord.hazard}</p>
                    <p><strong>{riskCopy.potentialConsequence}:</strong> {selectedRiskRecord.potentialConsequence || '-'}</p>
                  </article>
                  <article>
                    <h3>{riskCopy.controlPlan}</h3>
                    <p><strong>{riskCopy.existingControls}:</strong> {selectedRiskRecord.existingControls || '-'}</p>
                    <p><strong>{riskCopy.recommendedControls}:</strong> {selectedRiskRecord.recommendedControls || '-'}</p>
                  </article>
                  <article>
                    <h3>{riskCopy.scores}</h3>
                    <p><strong>{riskCopy.initialRiskScore}:</strong> {selectedRiskRecord.likelihood} x {selectedRiskRecord.severity} = {selectedRiskRecord.initialRiskScore}</p>
                    <p><strong>{riskCopy.residualRiskScore}:</strong> {selectedRiskRecord.residualLikelihood} x {selectedRiskRecord.residualSeverity} = {selectedRiskRecord.residualRiskScore}</p>
                  </article>
                  <article>
                    <h3>{riskCopy.responsibilityAndDue}</h3>
                    <p><strong>{riskCopy.responsiblePerson}:</strong> {selectedRiskRecord.responsiblePerson}</p>
                    <p><strong>{riskCopy.due}:</strong> {selectedRiskRecord.targetCompletionDate}</p>
                    <p><strong>{riskCopy.status}:</strong> {localizeStatus(selectedRiskRecord.status)}</p>
                  </article>
                </div>

                <div className="risk-detail-files">
                  <div>
                    <h3>{riskCopy.attachments}</h3>
                    {selectedRiskRecord.attachments.length > 0 ? (
                      <ul>
                        {selectedRiskRecord.attachments.map((fileName) => (
                          <li key={fileName}>{fileName}</li>
                        ))}
                      </ul>
                    ) : (
                      <p>{riskCopy.noAttachment}</p>
                    )}
                  </div>
                  <div>
                    <h3>{riskCopy.photos}</h3>
                    {selectedRiskRecord.photos.length > 0 ? (
                      <ul>
                        {selectedRiskRecord.photos.map((fileName) => (
                          <li key={fileName}>{fileName}</li>
                        ))}
                      </ul>
                    ) : (
                      <p>{riskCopy.noPhoto}</p>
                    )}
                  </div>
                </div>

                <div>
                  <h3>{riskCopy.notes}</h3>
                  <p>{selectedRiskRecord.notes || '-'}</p>
                </div>
              </section>
            ) : null}
          </>
        ) : null}

        {activeModule === 'projects' ? (
          <section className="panel project-master-panel">
            <h2>Proje Ana Veri Yönetimi</h2>
            <div className="project-form-grid">
              <label>
                Proje Adı
                <input
                  value={projectForm.name}
                  onChange={(event) => setProjectForm((prev) => ({ ...prev, name: event.target.value }))}
                  placeholder="Proje adını girin"
                />
              </label>
              <label>
                Ülke
                <select
                  value={projectForm.country}
                  onChange={(event) => setProjectForm((prev) => ({ ...prev, country: event.target.value, city: '' }))}
                >
                  <option value="">Ülke seçin</option>
                  {projectForm.country && !countryList.includes(projectForm.country) ? (
                    <option value={projectForm.country}>{projectForm.country}</option>
                  ) : null}
                  {countryList.map((country) => (
                    <option key={country} value={country}>{country}</option>
                  ))}
                </select>
              </label>
              <label>
                Şehir
                <select
                  value={projectForm.city}
                  onChange={(event) => setProjectForm((prev) => ({ ...prev, city: event.target.value }))}
                  disabled={!projectForm.country}
                >
                  <option value="">{projectForm.country ? 'Şehir seçin' : 'Önce ülke seçin'}</option>
                  {projectForm.city && !getCitiesForCountry(projectForm.country).includes(projectForm.city) ? (
                    <option value={projectForm.city}>{projectForm.city}</option>
                  ) : null}
                  {getCitiesForCountry(projectForm.country).map((city) => (
                    <option key={city} value={city}>{city}</option>
                  ))}
                </select>
              </label>
              <label>
                Konum
                <input
                  value={projectForm.address}
                  onChange={(event) => setProjectForm((prev) => ({ ...prev, address: event.target.value }))}
                  placeholder="Saha / lokasyon"
                />
              </label>
              <label className="full-row">
                Sözleşme / İş Kapsamı
                <textarea
                  rows={3}
                  value={projectForm.contractScope}
                  onChange={(event) => setProjectForm((prev) => ({ ...prev, contractScope: event.target.value }))}
                  placeholder="Sözleşme kapsamını kısa şekilde yazın"
                />
              </label>
              <div className="full-row actions">
                <button type="button" onClick={saveProject}>{editingProjectId ? 'Güncelle' : 'Kaydet'}</button>
                {editingProjectId ? (
                  <button type="button" className="secondary" onClick={resetProjectForm}>İptal</button>
                ) : null}
              </div>
              {projectSaveFeedback ? (
                <p className={`project-save-feedback ${projectSaveFeedback.type}`}>{projectSaveFeedback.text}</p>
              ) : null}
            </div>
          </section>
        ) : null}

        {activeModule === 'departments' ? (
          <section className="panel project-master-panel">
            <h2>Departman Ana Veri Yönetimi</h2>
            <div className="project-form-grid">
              <label>
                Departman Adı
                <input
                  value={departmentForm.name}
                  onChange={(event) => setDepartmentForm((prev) => ({ ...prev, name: event.target.value }))}
                  placeholder="Departman adını girin"
                />
              </label>
              <label>
                Departman Kodu
                <input
                  value={departmentForm.code}
                  onChange={(event) => setDepartmentForm((prev) => ({ ...prev, code: event.target.value }))}
                  placeholder="Örn. HSE, OPS, PMO"
                />
              </label>
              <label className="full-row">
                Açıklama
                <textarea
                  rows={3}
                  value={departmentForm.description}
                  onChange={(event) => setDepartmentForm((prev) => ({ ...prev, description: event.target.value }))}
                  placeholder="Departmanın belge ve süreç sorumluluğunu kısaca yazın"
                />
              </label>
              <div className="full-row actions">
                <button type="button" onClick={saveDepartment}>{editingDepartmentId ? 'Güncelle' : 'Kaydet'}</button>
                {editingDepartmentId ? (
                  <button type="button" className="secondary" onClick={resetDepartmentForm}>İptal</button>
                ) : null}
              </div>
            </div>

            <div className="table-wrap">
              <table className="project-master-table">
                <thead>
                  <tr>
                    <th>Departman Adı</th>
                    <th>Kod</th>
                    <th>Açıklama</th>
                    <th>İşlem</th>
                  </tr>
                </thead>
                <tbody>
                  {departmentRecords.length > 0 ? (
                    departmentRecords.map((department) => (
                      <tr key={department.id}>
                        <td>{department.name}</td>
                        <td>{department.code}</td>
                        <td>{department.description || '-'}</td>
                        <td className="actions-cell table-actions-cell">
                          <div className="table-action-group" aria-label={`${department.name} işlemleri`}>
                            <button type="button" className="table-action-button" onClick={() => editDepartment(department.id)}>Düzenle</button>
                            <button type="button" className="table-action-button danger" onClick={() => deleteDepartment(department.id)}>Sil</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={4}>Henüz departman kaydı yok.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        ) : null}

        {activeModule === 'contractors' ? (
          <section className="panel contractor-master-panel">
            <h2>Alt Yüklenici Ana Veri Yönetimi</h2>
            <div className="contractor-form-grid">
              <label>
                Şirket Adı
                <input
                  value={contractorForm.companyName}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, companyName: event.target.value }))}
                  placeholder="Alt yüklenici şirket adı"
                />
              </label>
              <label>
                Proje Adı
                <select
                  value={contractorForm.projectId}
                  onChange={(event) => fillContractorFormFromProject(event.target.value)}
                >
                  {projectCatalog.map((project) => (
                    <option key={project.id} value={project.id}>{project.name}</option>
                  ))}
                </select>
              </label>
              <label>
                Ülke
                <select
                  value={contractorForm.country}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, country: event.target.value, city: '' }))}
                >
                  <option value="">Ülke seçin</option>
                  {contractorForm.country && !countryList.includes(contractorForm.country) ? (
                    <option value={contractorForm.country}>{contractorForm.country}</option>
                  ) : null}
                  {countryList.map((country) => (
                    <option key={country} value={country}>{country}</option>
                  ))}
                </select>
              </label>
              <label>
                Şehir
                <select
                  value={contractorForm.city}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, city: event.target.value }))}
                  disabled={!contractorForm.country}
                >
                  <option value="">{contractorForm.country ? 'Şehir seçin' : 'Önce ülke seçin'}</option>
                  {contractorForm.city && !getCitiesForCountry(contractorForm.country).includes(contractorForm.city) ? (
                    <option value={contractorForm.city}>{contractorForm.city}</option>
                  ) : null}
                  {getCitiesForCountry(contractorForm.country).map((city) => (
                    <option key={city} value={city}>{city}</option>
                  ))}
                </select>
              </label>
              <label>
                Proje Konumu
                <input
                  value={contractorForm.projectLocation}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, projectLocation: event.target.value }))}
                />
              </label>
              <label>
                Sözleşme / İş Kapsamı
                <input
                  value={contractorForm.contractScope}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, contractScope: event.target.value }))}
                />
              </label>

              <label>
                İSG Uyarı Sayısı
                <input
                  type="number"
                  min={0}
                  value={contractorForm.hseWarningCount}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, hseWarningCount: Math.max(0, Number(event.target.value) || 0) }))}
                />
              </label>
              <label>
                İSG Uyarı Tarihi
                <input
                  type="date"
                  value={contractorForm.hseWarningDate}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, hseWarningDate: event.target.value }))}
                />
              </label>

              <label>
                Yangın Güvenliği Uyarı Sayısı
                <input
                  type="number"
                  min={0}
                  value={contractorForm.fireWarningCount}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, fireWarningCount: Math.max(0, Number(event.target.value) || 0) }))}
                />
              </label>
              <label>
                Yangın Güvenliği Uyarı Tarihi
                <input
                  type="date"
                  value={contractorForm.fireWarningDate}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, fireWarningDate: event.target.value }))}
                />
              </label>

              <label>
                Çevre Uyarı Sayısı
                <input
                  type="number"
                  min={0}
                  value={contractorForm.environmentWarningCount}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, environmentWarningCount: Math.max(0, Number(event.target.value) || 0) }))}
                />
              </label>
              <label>
                Çevre Uyarı Tarihi
                <input
                  type="date"
                  value={contractorForm.environmentWarningDate}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, environmentWarningDate: event.target.value }))}
                />
              </label>

              <label>
                Ceza Sayısı
                <input
                  type="number"
                  min={0}
                  value={contractorForm.penaltyCount}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, penaltyCount: Math.max(0, Number(event.target.value) || 0) }))}
                />
              </label>
              <label>
                Toplam Ceza Tutarı
                <input
                  type="number"
                  min={0}
                  step={100}
                  value={contractorForm.totalPenaltyAmount}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, totalPenaltyAmount: Math.max(0, Number(event.target.value) || 0) }))}
                />
              </label>
              <label className="full-row">
                Ceza Nedeni Yasal Madde No
                <input
                  value={contractorForm.penaltyLegalClause}
                  onChange={(event) => setContractorForm((prev) => ({ ...prev, penaltyLegalClause: event.target.value }))}
                  placeholder="Örn: 6331/4 - Sözleşme Ek-HSE Madde 12"
                />
              </label>

              <div className="full-row actions">
                <button type="button" onClick={saveContractor}>{editingContractorId ? 'Güncelle' : 'Kaydet'}</button>
                {editingContractorId ? (
                  <button type="button" className="secondary" onClick={resetContractorForm}>İptal</button>
                ) : null}
              </div>
            </div>
          </section>
        ) : null}

        {activeModule === 'inspections' ? (
          <>
            <section className="panel">
              <h2>{inspectionCopy.dataEntryTitle}</h2>
              <div className="form-grid">
                <label>
                  {inspectionCopy.titleLabel}
                  <input
                    value={inspectionForm.title}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, title: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.businessUnitLabel}
                  <input
                    value={inspectionForm.businessUnit}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, businessUnit: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.siteAreaLabel}
                  <input
                    value={inspectionForm.siteArea}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, siteArea: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.locationTypeLabel}
                  <select
                    value={inspectionForm.locationType}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, locationType: event.target.value }))}
                  >
                    <option value="">{inspectionCopy.selectOption}</option>
                    {locationTypesLocalized.map((locationType) => (
                      <option key={locationType} value={locationType}>{locationType}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {inspectionCopy.inspectionDateLabel}
                  <input
                    type="date"
                    value={inspectionForm.inspectionDate}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, inspectionDate: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.inspectorLabel}
                  <input
                    value={inspectionForm.inspectorName}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, inspectorName: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.departmentLabel}
                  <input
                    value={inspectionForm.department}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, department: event.target.value }))}
                  />
                </label>
                <label>
                  {inspectionCopy.positiveObservationsLabel}
                  <input
                    value={inspectionForm.positiveObservations}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, positiveObservations: event.target.value }))}
                  />
                </label>
                <label className="full-row">
                  {inspectionCopy.inspectionNotesLabel}
                  <textarea
                    rows={3}
                    value={inspectionForm.inspectionNotes}
                    onChange={(event) => setInspectionForm((prev) => ({ ...prev, inspectionNotes: event.target.value }))}
                  />
                </label>
              </div>
            </section>

            <section className="panel">
              <h2>{inspectionCopy.summaryTitle}</h2>
              <div className="inspection-summary-grid">
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionUrgencyLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.complianceScore}%</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.progressLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.answeredProgress}</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.totalFindingsLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.totalFindings}</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.openActionsLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.openActions}</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.answeredProgressLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.answeredProgress}</strong>
                </article>
              </div>
              <p className="summary-line">
                <span className="compliance-score-title">{inspectionCopy.complianceScoreLabel}</span>
                <br />
                {inspectionSummary.complianceScore}% - {inspectionUrgencyLabel}
              </p>
              <div className="inspection-summary-grid">
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.highFindingsLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.highFindings}</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.mediumFindingsLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.mediumFindings}</strong>
                </article>
                <article className="kpi-card">
                  <span className="kpi-label">{inspectionCopy.lowFindingsLabel}</span>
                  <strong className="kpi-value">{inspectionSummary.lowFindings}</strong>
                </article>
              </div>
            </section>

            <section className="panel">
              <h2>{inspectionCopy.checklistTitle}</h2>
              {inspectionSections.map((section, sectionIndex) => {
                const progress = getChecklistProgress(section.id, section.questions.length);
                return (
                  <article key={section.id} className="checklist-section">
                    <h3>{sectionIndex + 1}. {section.title}</h3>
                    <p className="checklist-description">{section.description}</p>
                    <p className="checklist-progress">{progress.answered}/{section.questions.length} - {progress.percent}%</p>
                    <label className="skip-toggle">
                      <input
                        type="checkbox"
                        checked={Boolean(skippedChecklistSections[section.id])}
                        onChange={(event) =>
                          setSkippedChecklistSections((prev) => ({
                            ...prev,
                            [section.id]: event.target.checked
                          }))
                        }
                      />
                      {inspectionCopy.skipChecklistLabel}
                    </label>

                    {!skippedChecklistSections[section.id] ? (
                      <div className="question-list">
                        {section.questions.map((question, questionIndex) => {
                          const key = `${section.id}-${questionIndex}`;
                          const selected = checklistAnswers[key] ?? '';
                          return (
                            <div className="question-item" key={key}>
                              <strong>{questionIndex + 1}. {question}</strong>
                              <div className="answer-row">
                                <label>
                                  <input
                                    type="radio"
                                    name={key}
                                    checked={selected === 'YES'}
                                    onChange={() => setChecklistAnswer(section.id, questionIndex, 'YES')}
                                  />
                                  {inspectionCopy.yesLabel}
                                </label>
                                <label>
                                  <input
                                    type="radio"
                                    name={key}
                                    checked={selected === 'NO'}
                                    onChange={() => setChecklistAnswer(section.id, questionIndex, 'NO')}
                                  />
                                  {inspectionCopy.noLabel}
                                </label>
                                <label className="na-option-label">
                                  <input
                                    type="radio"
                                    name={key}
                                    checked={selected === 'NA'}
                                    onChange={() => setChecklistAnswer(section.id, questionIndex, 'NA')}
                                  />
                                  {inspectionCopy.naLabel}
                                </label>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : null}
                  </article>
                );
              })}
            </section>

            <section className="panel">
              <h2>{inspectionCopy.correctiveActionsTitle}</h2>
              <p>{inspectionCopy.correctiveHint}</p>
              <p>
                {inspectionCopy.statusOpen}: {inspectionSummary.openActions} - {inspectionCopy.statusInProgress}: {inspectionSummary.inProgressActions} - {inspectionCopy.statusClosed}: {inspectionSummary.closedActions}
              </p>
              <div className="actions-row">
                <input
                  placeholder={inspectionCopy.manualActionPlaceholder}
                  value={manualActionText}
                  onChange={(event) => setManualActionText(event.target.value)}
                />
                <button type="button" onClick={addManualAction}>{inspectionCopy.addLabel}</button>
              </div>
              {manualActions.length === 0 ? (
                <p>{inspectionCopy.noActionsYet}</p>
              ) : (
                <div className="manual-actions-list">
                  {manualActions.map((action, actionIndex) => (
                    <div className="manual-action-item" key={`${action.text}-${actionIndex}`}>
                      <span>{action.text}</span>
                      <select
                        value={action.status}
                        onChange={(event) => updateManualActionStatus(actionIndex, event.target.value as ActionStatus)}
                      >
                        <option value="OPEN">{inspectionCopy.statusOpen}</option>
                        <option value="IN_PROGRESS">{inspectionCopy.statusInProgress}</option>
                        <option value="CLOSED">{inspectionCopy.statusClosed}</option>
                      </select>
                    </div>
                  ))}
                </div>
              )}
            </section>

            <section className="panel">
              <h2>{inspectionCopy.reportsTitle}</h2>
              <p>{inspectionCopy.reportHint}</p>
              <div className="report-action-buttons">
                <button type="button">{inspectionCopy.exportDocx}</button>
                <button type="button">{inspectionCopy.exportActionCsv}</button>
                <button type="button">{inspectionCopy.exportFullCsv}</button>
                <button type="button">{inspectionCopy.sendEmailSummary}</button>
                <button type="button">{inspectionCopy.printPdf}</button>
                <button type="button" onClick={resetInspection}>{inspectionCopy.resetInspection}</button>
              </div>
              <p className="creator-note">{inspectionCopy.preparedBy}</p>
            </section>
          </>
        ) : null}

        {activeModule === 'observations' ? (
          <>
            <section className="panel observation-panel">
              <h2>{observationCopy.formTitle}</h2>
              <p>{observationCopy.moduleSubtitle}</p>
              <div className="observation-form-grid">
                <label>
                  {observationCopy.observationNo}
                  <input value={editingObservationId ? (observationRecords.find((record) => record.id === editingObservationId)?.observationNo ?? nextObservationNo) : nextObservationNo} readOnly />
                </label>
                <label>
                  {observationCopy.project}
                  <select value={observationForm.projectId} onChange={(event) => setObservationProject(event.target.value)}>
                    <option value="">{observationCopy.selectProject}</option>
                    {projectCatalog.map((project) => (
                      <option key={project.id} value={project.id}>{project.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {observationCopy.projectLocation}
                  <input value={observationForm.projectLocation} onChange={(event) => setObservationForm((prev) => ({ ...prev, projectLocation: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.inspectionDate}
                  <input type="date" value={observationForm.inspectionDate} onChange={(event) => setObservationForm((prev) => ({ ...prev, inspectionDate: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.inspectionTime}
                  <input type="time" value={observationForm.inspectionTime} onChange={(event) => setObservationForm((prev) => ({ ...prev, inspectionTime: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.inspectorName}
                  <input value={observationForm.inspectorName} onChange={(event) => setObservationForm((prev) => ({ ...prev, inspectorName: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.contractor}
                  <input value={observationForm.contractor} onChange={(event) => setObservationForm((prev) => ({ ...prev, contractor: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.subcontractor}
                  <input value={observationForm.subcontractor} onChange={(event) => setObservationForm((prev) => ({ ...prev, subcontractor: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.responsiblePerson}
                  <input value={observationForm.responsiblePerson} onChange={(event) => setObservationForm((prev) => ({ ...prev, responsiblePerson: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.category}
                  <select value={observationForm.category} onChange={(event) => setObservationForm((prev) => ({ ...prev, category: event.target.value }))}>
                    {observationCategoryOptions.map((category) => (
                      <option key={category} value={category}>{language === 'ru' ? (observationCategoryOptionsRu[category] ?? category) : category}</option>
                    ))}
                  </select>
                </label>
                <label>
                  {observationCopy.subject}
                  <input value={observationForm.subject} onChange={(event) => setObservationForm((prev) => ({ ...prev, subject: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.observationLocation}
                  <input value={observationForm.observationLocation} onChange={(event) => setObservationForm((prev) => ({ ...prev, observationLocation: event.target.value }))} />
                </label>
                <label className="full-row">
                  {observationCopy.violatedRequirement}
                  <textarea rows={2} value={observationForm.violatedRequirement} onChange={(event) => setObservationForm((prev) => ({ ...prev, violatedRequirement: event.target.value }))} />
                </label>
                <label className="full-row">
                  {observationCopy.observationDescription}
                  <textarea rows={3} value={observationForm.observationDescription} onChange={(event) => setObservationForm((prev) => ({ ...prev, observationDescription: event.target.value }))} />
                </label>
                <label className="full-row">
                  {observationCopy.correctiveAction}
                  <textarea rows={3} value={observationForm.correctiveAction} onChange={(event) => setObservationForm((prev) => ({ ...prev, correctiveAction: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.dueDate}
                  <input type="date" value={observationForm.dueDate} onChange={(event) => setObservationForm((prev) => ({ ...prev, dueDate: event.target.value }))} />
                </label>
                <label>
                  {observationCopy.priority}
                  <select value={observationForm.priority} onChange={(event) => setObservationForm((prev) => ({ ...prev, priority: event.target.value as ObservationPriority }))}>
                    <option value="DUSUK">{language === 'ru' ? observationPriorityLabelRu.DUSUK : 'Düşük'}</option>
                    <option value="ORTA">{language === 'ru' ? observationPriorityLabelRu.ORTA : 'Orta'}</option>
                    <option value="YUKSEK">{language === 'ru' ? observationPriorityLabelRu.YUKSEK : 'Yüksek'}</option>
                    <option value="KRITIK">{language === 'ru' ? observationPriorityLabelRu.KRITIK : 'Kritik'}</option>
                  </select>
                </label>
                <label>
                  {observationCopy.status}
                  <select value={observationForm.status} onChange={(event) => setObservationForm((prev) => ({ ...prev, status: event.target.value as ObservationWorkflowStatus }))}>
                    <option value="ACIK">{language === 'ru' ? observationStatusLabelRu.ACIK : 'Açık'}</option>
                    <option value="DEVAM_EDIYOR">{language === 'ru' ? observationStatusLabelRu.DEVAM_EDIYOR : 'Devam Ediyor'}</option>
                    <option value="KAPALI">{language === 'ru' ? observationStatusLabelRu.KAPALI : 'Kapalı'}</option>
                  </select>
                </label>
                <label className="full-row">
                  {observationCopy.comment}
                  <textarea rows={2} value={observationForm.comment} onChange={(event) => setObservationForm((prev) => ({ ...prev, comment: event.target.value }))} />
                </label>
              </div>

              <section className="observation-attachments">
                <h3>{observationCopy.attachmentsTitle}</h3>
                <div
                  className={`observation-dropzone ${observationDragActive ? 'active' : ''}`}
                  onDragOver={(event) => {
                    event.preventDefault();
                    setObservationDragActive(true);
                  }}
                  onDragLeave={() => setObservationDragActive(false)}
                  onDrop={(event) => {
                    event.preventDefault();
                    setObservationDragActive(false);
                    const files = Array.from(event.dataTransfer.files ?? []);
                    void addObservationFiles(files);
                  }}
                >
                  <p>{observationCopy.dragDropHint}</p>
                  <label className="browse-button">
                    {observationCopy.browseFiles}
                    <input
                      type="file"
                      multiple
                      accept=".jpg,.jpeg,.png,.pdf"
                      onChange={(event) => {
                        const files = Array.from(event.target.files ?? []);
                        void addObservationFiles(files);
                        event.currentTarget.value = '';
                      }}
                    />
                  </label>
                </div>

                {observationDraftAttachments.length > 0 ? (
                  <div className="observation-attachment-grid">
                    {observationDraftAttachments.map((attachment) => (
                      <article key={attachment.id} className="observation-attachment-card">
                        {attachment.mimeType.startsWith('image/') ? (
                          <img src={attachment.dataUrl} alt={attachment.name} />
                        ) : (
                          <div className="attachment-pdf">PDF</div>
                        )}
                        <p>{attachment.name}</p>
                        <button type="button" className="danger" onClick={() => removeObservationAttachment(attachment.id)}>{observationCopy.removeImage}</button>
                      </article>
                    ))}
                  </div>
                ) : null}
              </section>

              <div className="full-row actions observation-actions-row">
                <button type="button" onClick={saveObservationRecord}>{editingObservationId ? observationCopy.updateObservation : observationCopy.saveObservation}</button>
                {editingObservationId ? <button type="button" className="secondary" onClick={resetObservationForm}>{observationCopy.cancel}</button> : null}
              </div>
            </section>

            <section className="panel table-wrap observation-table-panel">
              <h2>{observationCopy.registryTitle}</h2>
              <table className="compact-actions-table">
                <thead>
                  <tr>
                    <th>{observationCopy.observationNo}</th>
                    <th>{observationCopy.project}</th>
                    <th>{observationCopy.tableCategory}</th>
                    <th>{observationCopy.tableSubject}</th>
                    <th>{observationCopy.responsiblePerson}</th>
                    <th>{observationCopy.dueDate}</th>
                    <th>{observationCopy.priority}</th>
                    <th>{observationCopy.status}</th>
                    <th>{observationCopy.tableActions}</th>
                  </tr>
                </thead>
                <tbody>
                  {observationRecords.length > 0 ? (
                    observationRecords.map((record) => (
                      <tr key={record.id}>
                        <td>{record.observationNo}</td>
                        <td>{projectCatalog.find((project) => project.id === record.projectId)?.name ?? record.projectId}</td>
                        <td>{language === 'ru' ? (observationCategoryOptionsRu[record.category] ?? record.category) : record.category}</td>
                        <td>{record.subject}</td>
                        <td>{record.responsiblePerson}</td>
                        <td>{record.dueDate}</td>
                        <td>{language === 'ru' ? observationPriorityLabelRu[record.priority] : observationPriorityLabel[record.priority]}</td>
                        <td>{language === 'ru' ? observationStatusLabelRu[record.status] : observationStatusLabel[record.status]}</td>
                        <td>
                          <div className="table-actions compact">
                            <button type="button" className="secondary" onClick={() => setViewingObservationId(record.id)}>{observationCopy.view}</button>
                            <button type="button" className="secondary" onClick={() => editObservationRecord(record.id)}>{observationCopy.edit}</button>
                            <button type="button" className="danger" onClick={() => deleteObservationRecord(record.id)}>{observationCopy.delete}</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={9}>{t.noData}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </section>

            {selectedObservationRecord ? (
              <section className="panel observation-report-panel">
                <h2>{observationCopy.reportActions} - {selectedObservationRecord.observationNo}</h2>
                <div className="report-action-buttons">
                  <button type="button" onClick={() => openObservationReport(selectedObservationRecord, false)}>{observationCopy.generatePdf}</button>
                  <button type="button" onClick={() => openObservationReport(selectedObservationRecord, true)}>{observationCopy.printReport}</button>
                </div>
              </section>
            ) : null}
          </>
        ) : null}

        {activeModule === 'occupational-health' ? (
          <section className="panel">
            <h2>İş Sağlığı Veri Girişi</h2>
            <div className="form-grid">
              <label>
                Çalışan Adı
                <input
                  value={healthForm.employeeName}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, employeeName: event.target.value }))}
                />
              </label>
              <label>
                Çalışan ID
                <input
                  value={healthForm.employeeId}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, employeeId: event.target.value }))}
                />
              </label>
              <label>
                Departman
                <input
                  value={healthForm.department}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, department: event.target.value }))}
                />
              </label>
              <label>
                Pozisyon
                <input
                  value={healthForm.position}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, position: event.target.value }))}
                />
              </label>
              <label>
                Kan Grubu
                <input
                  value={healthForm.bloodGroup}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, bloodGroup: event.target.value }))}
                />
              </label>
              <label>
                Alerjiler
                <input
                  value={healthForm.allergies}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, allergies: event.target.value }))}
                />
              </label>
              <label>
                Kronik Hastalık
                <input
                  value={healthForm.chronicDisease}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, chronicDisease: event.target.value }))}
                />
              </label>
              <label>
                Bulaşıcı Hastalık
                <input
                  value={healthForm.communicableDisease}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, communicableDisease: event.target.value }))}
                />
              </label>
              <label>
                İlaç Kullanımı
                <input
                  value={healthForm.medication}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, medication: event.target.value }))}
                />
              </label>
              <label>
                Engellilik Durumu
                <input
                  value={healthForm.disabilityStatus}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, disabilityStatus: event.target.value }))}
                />
              </label>
              <label>
                İşe Uygunluk (Evet / Hayır)
                <select
                  value={healthForm.fitForWork}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, fitForWork: event.target.value as 'Yes' | 'No' }))}
                >
                  <option value="Yes">Evet</option>
                  <option value="No">Hayır</option>
                </select>
              </label>
              <label>
                Kısıtlı Çalışma
                <input
                  value={healthForm.restrictedWork}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, restrictedWork: event.target.value }))}
                />
              </label>
              <label>
                Muayene Tarihi
                <input
                  type="date"
                  value={healthForm.medicalExaminationDate}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, medicalExaminationDate: event.target.value }))}
                />
              </label>
              <label>
                Sonraki Muayene
                <input
                  type="date"
                  value={healthForm.nextMedicalExamination}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, nextMedicalExamination: event.target.value }))}
                />
              </label>
              <label>
                Aşı Durumu
                <input
                  value={healthForm.vaccinationStatus}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, vaccinationStatus: event.target.value }))}
                />
              </label>
              <label className="full-row">
                Notlar
                <textarea
                  rows={3}
                  value={healthForm.remarks}
                  onChange={(event) => setHealthForm((prev) => ({ ...prev, remarks: event.target.value }))}
                />
              </label>
              <div className="full-row actions health-entry-actions-row">
                <button type="button" onClick={saveHealthEntry}>{editingHealthEmployeeId ? 'Güncelle' : t.save}</button>
                {editingHealthEmployeeId ? (
                  <button type="button" className="secondary" onClick={resetHealthForm}>Vazgeç</button>
                ) : null}
                <div className="report-action-buttons health-entry-export-buttons">
                  <button type="button" onClick={exportHealthExecutivePdf}>PDF Dışa Aktar</button>
                  <button type="button" onClick={exportHealthRecordsCsv}>Excel Dışa Aktar</button>
                </div>
              </div>
            </div>
          </section>
        ) : null}

        {activeModule !== 'dashboard' && activeModule !== 'occupational-health' && activeModule !== 'legal-register' && activeModule !== 'documents' && activeModule !== 'reports' && activeModule !== 'export-center' && activeModule !== 'projects' && activeModule !== 'departments' && activeModule !== 'contractors' && activeModule !== 'settings' && activeModule !== 'inspections' && activeModule !== 'permit-to-work' && activeModule !== 'risk-assessments' && activeModule !== 'incidents' && activeModule !== 'equipment-management' && activeModule !== 'emergency-preparedness' && activeModule !== 'employees' && activeModule !== 'trainings' && activeModule !== 'ppe-stocks' && activeModule !== 'kpis-analytics' && activeModule !== 'observations' ? (
          <section className="panel">
            <h2>{activeModule === 'environmental' ? 'Экология Записи' : `${moduleLabels[activeModule]} ${t.records}`}</h2>
            <table>
              <thead>
                <tr>
                  <th>{activeModule === 'environmental' ? 'Проект' : t.project}</th>
                  <th>{language === 'ru' || activeModule === 'environmental' ? 'Дата' : 'Дата'}</th>
                  <th>{language === 'ru' || activeModule === 'environmental' ? 'Заголовок' : 'Заголовок'}</th>
                  <th>{localizeText(metricFieldLabels[activeModule]?.primary ?? t.primaryMetric, language)}</th>
                  <th>{localizeText(metricFieldLabels[activeModule]?.secondary ?? t.secondaryMetric, language)}</th>
                  <th>{language === 'ru' || activeModule === 'environmental' ? 'Статус' : 'Статус'}</th>
                  <th>{language === 'ru' || activeModule === 'environmental' ? 'Действия' : 'Действия'}</th>
                </tr>
              </thead>
              <tbody>
                {scopedModuleRecords.length > 0 ? (
                  scopedModuleRecords.map(({ row, index }) => (
                    <tr key={`${row.title}-${index}`}>
                      <td>{projectCatalog.find((project) => project.id === row.projectId)?.name ?? row.projectId}</td>
                      <td>{row.date}</td>
                      <td>{localizeText(row.title, language)}</td>
                      <td>{row.valueA}</td>
                      <td>{row.valueB}</td>
                      <td>{localizeStatus(row.status)}</td>
                      <td className="actions-cell table-actions-cell">
                        <div className="table-action-group" aria-label={language === 'ru' || activeModule === 'environmental' ? `Действия по записи ${localizeText(row.title, language)}` : `Действия по записи ${localizeText(row.title, language)}`}>
                          <button type="button" className="table-action-button" onClick={() => editModuleRecord(index)}>{language === 'ru' || activeModule === 'environmental' ? 'Редактировать' : 'Редактировать'}</button>
                          <button type="button" className="table-action-button danger" onClick={() => deleteModuleRecord(index)}>{language === 'ru' || activeModule === 'environmental' ? 'Удалить' : 'Удалить'}</button>
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={7}>{activeModule === 'environmental' ? 'Нет данных для выбранной области.' : t.noData}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </section>
        ) : null}

        {activeModule === 'projects' ? (
          <section className="panel table-wrap project-master-panel">
            <h2>Proje Listesi</h2>
            <table className="project-master-table">
              <thead>
                <tr>
                  <th>Proje Adı</th>
                  <th>Ülke</th>
                  <th>Şehir</th>
                  <th>Konum</th>
                  <th>Sözleşme / İş Kapsamı</th>
                  <th>İşlemler</th>
                </tr>
              </thead>
              <tbody>
                {projectCatalog.length > 0 ? (
                  projectCatalog.map((project) => (
                    <tr key={project.id}>
                      <td>{project.name}</td>
                      <td>{project.country}</td>
                      <td>{project.city}</td>
                      <td>{project.address}</td>
                      <td>{project.contractScope}</td>
                      <td>
                        <div className="project-row-actions">
                          <button type="button" onClick={() => editProject(project.id)}>Düzenle</button>
                          <button type="button" className="danger" onClick={() => deleteProject(project.id)}>Sil</button>
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={6}>{t.noData}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </section>
        ) : null}

        {activeModule === 'contractors' ? (
          <section className="panel table-wrap contractor-master-panel">
            <h2>Alt Yüklenici Listesi</h2>
            <table className="contractor-master-table">
              <thead>
                <tr>
                  <th>Şirket Adı</th>
                  <th>Proje Adı</th>
                  <th>Ülke</th>
                  <th>Şehir</th>
                  <th>Proje Konumu</th>
                  <th>Sözleşme / İş Kapsamı</th>
                  <th>İSG Uyarıları</th>
                  <th>Yangın Uyarıları</th>
                  <th>Çevre Uyarıları</th>
                  <th>Ceza Sayısı</th>
                  <th>Toplam Ceza Tutarı</th>
                  <th>İşlemler</th>
                </tr>
              </thead>
              <tbody>
                {contractorRecords.length > 0 ? (
                  contractorRecords.map((record) => (
                    <tr key={record.id}>
                      <td>{record.companyName}</td>
                      <td>{record.projectName}</td>
                      <td>{record.country}</td>
                      <td>{record.city}</td>
                      <td>{record.projectLocation}</td>
                      <td>{record.contractScope}</td>
                      <td>{record.hseWarningCount}</td>
                      <td>{record.fireWarningCount}</td>
                      <td>{record.environmentWarningCount}</td>
                      <td>{record.penaltyCount}</td>
                      <td>{record.totalPenaltyAmount.toLocaleString('tr-TR')} Rub</td>
                      <td>
                        <div className="contractor-row-actions">
                          <button type="button" onClick={() => editContractor(record.id)}>Düzenle</button>
                          <button type="button" className="danger" onClick={() => deleteContractor(record.id)}>Sil</button>
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={12}>{t.noData}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </section>
        ) : null}

        {activeModule === 'settings' ? (
          <section className="panel settings-config-panel">
            <h2>Sistem Yapılandırma Merkezi</h2>

            <div className="settings-tab-bar" role="tablist" aria-label="Ayar Sekmeleri">
              {settingsTabOptions.map((tab) => (
                <button
                  key={tab.key}
                  type="button"
                  role="tab"
                  aria-selected={activeSettingsTab === tab.key}
                  className={activeSettingsTab === tab.key ? 'settings-tab-button active' : 'settings-tab-button'}
                  onClick={() => setActiveSettingsTab(tab.key)}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            <div className="settings-tab-content" role="tabpanel">
              {activeSettingsTab === 'company-info' ? (
                <div className="settings-grid">
                  <label>
                    Şirket Adı
                    <input value={settingsConfig.companyName} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, companyName: event.target.value }))} />
                  </label>
                  <label>
                    Vergi No
                    <input value={settingsConfig.taxNumber} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, taxNumber: event.target.value }))} />
                  </label>
                  <label className="full-row">
                    Merkez Adresi
                    <input value={settingsConfig.headquartersAddress} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, headquartersAddress: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'general-settings' ? (
                <div className="settings-grid">
                  <label>
                    Zaman Dilimi
                    <input value={settingsConfig.timezone} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, timezone: event.target.value }))} />
                  </label>
                  <label>
                    Varsayılan Dil
                    <input value={settingsConfig.defaultLanguage} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, defaultLanguage: event.target.value }))} />
                  </label>
                  <label>
                    Tarih Formatı
                    <input value={settingsConfig.dateFormat} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, dateFormat: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'roles-permissions' ? (
                <div className="settings-grid">
                  <label className="full-row">
                    Rol Yetki Matrisi
                    <textarea rows={4} value={settingsConfig.roleMatrixSummary} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, roleMatrixSummary: event.target.value }))} />
                  </label>
                  <label className="full-row">
                    Onay Akışı
                    <textarea rows={3} value={settingsConfig.approvalFlow} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, approvalFlow: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'notifications' ? (
                <div className="settings-grid">
                  <label className="toggle-row">
                    <input type="checkbox" checked={settingsConfig.emailNotifications} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, emailNotifications: event.target.checked }))} />
                    E-posta Bildirimleri Aktif
                  </label>
                  <label className="toggle-row">
                    <input type="checkbox" checked={settingsConfig.smsNotifications} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, smsNotifications: event.target.checked }))} />
                    SMS Bildirimleri Aktif
                  </label>
                  <label>
                    Eskalasyon Süresi (Saat)
                    <input type="number" min={1} value={settingsConfig.escalationHours} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, escalationHours: Math.max(1, Number(event.target.value) || 1) }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'risk-matrix' ? (
                <div className="settings-grid">
                  <label>
                    Olasılık Skalası
                    <input value={settingsConfig.riskLikelihoodScale} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, riskLikelihoodScale: event.target.value }))} />
                  </label>
                  <label>
                    Şiddet Skalası
                    <input value={settingsConfig.riskSeverityScale} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, riskSeverityScale: event.target.value }))} />
                  </label>
                  <label className="full-row">
                    Risk Eşikleri
                    <textarea rows={3} value={settingsConfig.riskThresholds} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, riskThresholds: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'hse-categories' ? (
                <div className="settings-grid">
                  <label className="full-row">
                    İSG Kategorileri
                    <textarea rows={3} value={settingsConfig.hseCategories} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, hseCategories: event.target.value }))} />
                  </label>
                  <label className="full-row">
                    Çevre Kategorileri
                    <textarea rows={3} value={settingsConfig.environmentalCategories} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, environmentalCategories: event.target.value }))} />
                  </label>
                  <label className="full-row">
                    Yangın Güvenliği Kategorileri
                    <textarea rows={3} value={settingsConfig.fireSafetyCategories} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, fireSafetyCategories: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'penalty-settings' ? (
                <div className="settings-grid">
                  <label className="full-row">
                    Uyarı / Ceza Kuralı
                    <textarea rows={3} value={settingsConfig.warningPenaltyRule} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, warningPenaltyRule: event.target.value }))} />
                  </label>
                  <label>
                    Para Birimi
                    <input value={settingsConfig.currency} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, currency: event.target.value }))} />
                  </label>
                  <label>
                    Varsayılan Ceza Tutarı
                    <input type="number" min={0} step={100} value={settingsConfig.defaultPenaltyAmount} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, defaultPenaltyAmount: Math.max(0, Number(event.target.value) || 0) }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'document-numbering' ? (
                <div className="settings-grid">
                  <label>
                    Belge Prefix
                    <input value={settingsConfig.documentPrefix} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, documentPrefix: event.target.value }))} />
                  </label>
                  <label>
                    Sıra Başlangıcı
                    <input type="number" min={1} value={settingsConfig.documentSequenceStart} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, documentSequenceStart: Math.max(1, Number(event.target.value) || 1) }))} />
                  </label>
                  <label>
                    Revizyon Formatı
                    <input value={settingsConfig.revisionFormat} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, revisionFormat: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'backup-restore' ? (
                <div className="settings-grid">
                  <label>
                    Yedekleme Sıklığı
                    <input value={settingsConfig.backupFrequency} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, backupFrequency: event.target.value }))} />
                  </label>
                  <label>
                    Saklama Süresi (Gün)
                    <input type="number" min={30} value={settingsConfig.retentionDays} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, retentionDays: Math.max(30, Number(event.target.value) || 30) }))} />
                  </label>
                  <label className="full-row">
                    Geri Yükleme Onay Rolü
                    <input value={settingsConfig.restoreApprovalRole} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, restoreApprovalRole: event.target.value }))} />
                  </label>
                </div>
              ) : null}

              {activeSettingsTab === 'audit-log' ? (
                <div className="settings-grid">
                  <label className="full-row">
                    Denetim Günlüğü Politikası
                    <textarea rows={4} value={settingsConfig.auditLogPolicy} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, auditLogPolicy: event.target.value }))} />
                  </label>
                  <label>
                    Kayıt Saklama Süresi (Ay)
                    <input type="number" min={6} value={settingsConfig.auditLogRetentionMonths} onChange={(event) => setSettingsConfig((prev) => ({ ...prev, auditLogRetentionMonths: Math.max(6, Number(event.target.value) || 6) }))} />
                  </label>
                </div>
              ) : null}
            </div>

            <div className="settings-actions">
              <button type="button" onClick={saveSettingsConfig}>Kaydet</button>
              <button type="button" className="secondary" onClick={resetSettingsConfig}>Sıfırla</button>
            </div>
          </section>
        ) : null}

        {activeModule === 'occupational-health' ? (
          <section className="panel table-wrap">
            <h2>İş Sağlığı Kayıtları</h2>
            <table>
              <thead>
                <tr>
                  <th>Çalışan Adı</th>
                  <th>Çalışan ID</th>
                  <th>Departman</th>
                  <th>Pozisyon</th>
                  <th>Kan Grubu</th>
                  <th>Alerjiler</th>
                  <th>Kronik Hastalık</th>
                  <th>Bulaşıcı Hastalık</th>
                  <th>İlaç Kullanımı</th>
                  <th>Engellilik Durumu</th>
                  <th>İşe Uygunluk</th>
                  <th>Kısıtlı Çalışma</th>
                  <th>Muayene Tarihi</th>
                  <th>Sonraki Muayene</th>
                  <th>Aşı Durumu</th>
                  <th>Notlar</th>
                  <th>İşlem</th>
                </tr>
              </thead>
              <tbody>
                {healthRecords.length > 0 ? (
                  healthRecords.map((record, index) => (
                    <tr key={`${record.employeeId}-${index}`}>
                      <td>{record.employeeName}</td>
                      <td>{record.employeeId}</td>
                      <td>{record.department}</td>
                      <td>{record.position}</td>
                      <td>{record.bloodGroup}</td>
                      <td>{record.allergies}</td>
                      <td>{record.chronicDisease}</td>
                      <td>{record.communicableDisease}</td>
                      <td>{record.medication}</td>
                      <td>{record.disabilityStatus}</td>
                      <td>{record.fitForWork === 'Yes' ? 'Evet' : 'Hayır'}</td>
                      <td>{record.restrictedWork}</td>
                      <td>{record.medicalExaminationDate}</td>
                      <td>{record.nextMedicalExamination}</td>
                      <td>{record.vaccinationStatus}</td>
                      <td>{record.remarks}</td>
                      <td className="actions-cell table-actions-cell">
                        <div className="table-action-group" aria-label={`${record.employeeName} işlemleri`}>
                          <button type="button" className="table-action-button" onClick={() => editHealthEntry(record.employeeId)}>Düzenle</button>
                          <button type="button" className="table-action-button danger" onClick={() => deleteHealthEntry(record.employeeId)}>Sil</button>
                        </div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={16}>{t.noHealth}</td>
                    <td className="actions-cell table-actions-cell">
                      <div className="table-action-group" aria-label="İş sağlığı kayıt işlemleri">
                        <button type="button" className="table-action-button" disabled>Düzenle</button>
                        <button type="button" className="table-action-button danger" disabled>Sil</button>
                      </div>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </section>
        ) : null}

        <footer className="global-footer" role="contentinfo">
          <p>
            {localizedFooterText}
          </p>
        </footer>
      </main>
    </div>
  );
}
