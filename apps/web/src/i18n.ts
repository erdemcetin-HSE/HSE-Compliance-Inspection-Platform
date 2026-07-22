export type LanguageCode = 'tr' | 'en' | 'ru';

type Dictionary = {
  language: string;
  title: string;
  subtitle: string;
  loading: string;
  dataUnavailable: string;
  totalInspections: string;
  openPtw: string;
  closedPtw: string;
  openObservations: string;
  highCriticalRisks: string;
  trainingCompliance: string;
  ppeUsage: string;
  reportActions: string;
  print: string;
  downloadPdf: string;
  downloadExcel: string;
  emailPdf: string;
  emailExcel: string;
  emailPlaceholder: string;
  reportEmailRequired: string;
  reportEmailTriggered: string;
  inspectionEntry: string;
  employee: string;
  select: string;
  contractor: string;
  region: string;
  workNature: string;
  riskLevel: string;
  findings: string;
  correctiveAction: string;
  attachmentFileName: string;
  attachmentUrl: string;
  attachmentType: string;
  saveInspection: string;
  inspectionSaved: string;
  ptwStatus: string;
  permitNo: string;
  owner: string;
  status: string;
  observationStatus: string;
  type: string;
  location: string;
  riskRegisterSnapshot: string;
  risk: string;
  score: string;
  level: string;
  ppeStockUsage: string;
  item: string;
  lastIn: string;
  lastUsed: string;
  currentLeft: string;
  photoJpeg: string;
  photoPng: string;
  certificatePdf: string;
  sidebarDashboard: string;
  sidebarAudit: string;
  sidebarKpis: string;
  sidebarInspection: string;
  sidebarObservation: string;
  sidebarTraining: string;
  sidebarPpeStock: string;
  sidebarReport: string;
  uploadFile: string;
  uploading: string;
  attachmentUploaded: string;
  trainingName: string;
  validUntil: string;
  certificateNo: string;
  projectFilter: string;
  allProjects: string;
  projectName: string;
  periodLabel: string;
  kpiWorkforce: string;
  kpiManHours: string;
  kpiLti: string;
  kpiNearMiss: string;
  kpiPpeLeft: string;
  kpiPenaltyOpen: string;
  saveKpi: string;
  inspectionDate: string;
  firstName: string;
  lastName: string;
  roleTitle: string;
  certificateExpiryDate: string;
  saveTraining: string;
  quantityIn: string;
  quantityUsed: string;
  quantityLeft: string;
  entryDate: string;
  savePpe: string;
  penaltySource: string;
  customer: string;
  government: string;
  penaltyTitle: string;
  penaltyAmount: string;
  statusOpen: string;
  statusInProgress: string;
  statusClosed: string;
  savePenalty: string;
  reportScopeInfo: string;
  projectHealth: string;
  healthScore: string;
  daily: string;
  weekly: string;
  monthly: string;
  hseSpecialistFirstName: string;
  hseSpecialistLastName: string;
  kpiToolboxCount: string;
  kpiToolboxPeople: string;
  kpiJobInduction: string;
};

export const dictionaries: Record<LanguageCode, Dictionary> = {
  tr: {
    language: 'Dil',
    title: 'HSE Yonetici ve Operasyon Platformu',
    subtitle:
      'Haftalik/aylik KPI izleme, PTW/Observation kapanis takibi, risk gorunurlugu ve PPE operasyon kontrolu tek ekranda.',
    loading: 'Dashboard verileri yukleniyor...',
    dataUnavailable: 'Veri kaynagi su anda kullanilamiyor.',
    totalInspections: 'Toplam Inspection',
    openPtw: 'Acik PTW',
    closedPtw: 'Kapali PTW',
    openObservations: 'Acik Observation',
    highCriticalRisks: 'Yuksek + Kritik Risk',
    trainingCompliance: 'Egitim Uyumu',
    ppeUsage: 'PPE Kullanim',
    reportActions: 'Rapor Islemleri',
    print: 'Yazdir',
    downloadPdf: 'PDF Indir',
    downloadExcel: 'Excel Indir',
    emailPdf: 'PDF E-posta',
    emailExcel: 'Excel E-posta',
    emailPlaceholder: 'rapor@firma.com',
    reportEmailRequired: 'Rapor e-posta adresini giriniz.',
    reportEmailTriggered: 'Rapor e-posta gonderimi baslatildi',
    inspectionEntry: 'Inspection Girisi',
    employee: 'Personel',
    select: 'Seciniz',
    contractor: 'Taseron',
    region: 'Bolge',
    workNature: 'Isin Niteligi',
    riskLevel: 'Risk Seviyesi',
    findings: 'Bulgu',
    correctiveAction: 'Duzeltici Aksiyon',
    attachmentFileName: 'Ek Dosya Adi',
    attachmentUrl: 'Ek Dosya URL',
    attachmentType: 'Ek Dosya Tipi',
    saveInspection: 'Inspection Kaydet',
    inspectionSaved: 'Inspection kaydi basariyla olusturuldu.',
    ptwStatus: 'PTW Durumu',
    permitNo: 'Permit No',
    owner: 'Sorumlu',
    status: 'Durum',
    observationStatus: 'Observation Durumu',
    type: 'Tip',
    location: 'Konum',
    riskRegisterSnapshot: 'Risk Kayit Ozeti',
    risk: 'Risk',
    score: 'Skor',
    level: 'Seviye',
    ppeStockUsage: 'PPE Stok ve Kullanim',
    item: 'Ekipman',
    lastIn: 'Son Giren',
    lastUsed: 'Son Kullanilan',
    currentLeft: 'Kalan',
    photoJpeg: 'Fotograf (JPEG)',
    photoPng: 'Fotograf (PNG)',
    certificatePdf: 'Sertifika (PDF)',
    sidebarDashboard: 'Gosterge Paneli',
    sidebarAudit: 'Audit',
    sidebarKpis: 'KPIs',
    sidebarInspection: 'Inspection',
    sidebarObservation: 'Observation',
    sidebarTraining: 'Training',
    sidebarPpeStock: 'PPE Stock',
    sidebarReport: 'Report',
    uploadFile: 'Dosya Yukle',
    uploading: 'Yukleniyor...',
    attachmentUploaded: 'Ek dosya basariyla yuklendi.',
    trainingName: 'Egitim Adi',
    validUntil: 'Gecerlilik',
    certificateNo: 'Sertifika No',
    projectFilter: 'Proje Secimi',
    allProjects: 'All Projects',
    projectName: 'Proje Adi',
    periodLabel: 'Periyot',
    kpiWorkforce: 'Isgucu Sayisi',
    kpiManHours: 'Kazasiz Adam-Saat',
    kpiLti: 'LTI Sayisi',
    kpiNearMiss: 'Near Miss Sayisi',
    kpiPpeLeft: 'Kalan PPE',
    kpiPenaltyOpen: 'Acik Ceza/Uyari',
    saveKpi: 'KPI Kaydet',
    inspectionDate: 'Inspection Tarihi',
    firstName: 'Ad',
    lastName: 'Soyad',
    roleTitle: 'Gorev',
    certificateExpiryDate: 'Sertifika Bitis Tarihi',
    saveTraining: 'Egitim Kaydet',
    quantityIn: 'Giren',
    quantityUsed: 'Kullanilan',
    quantityLeft: 'Kalan',
    entryDate: 'Giris Tarihi',
    savePpe: 'PPE Kaydet',
    penaltySource: 'Ceza/Uyari Kaynagi',
    customer: 'Musteri',
    government: 'Devlet',
    penaltyTitle: 'Ceza/Uyari Basligi',
    penaltyAmount: 'Tutar',
    statusOpen: 'Acik',
    statusInProgress: 'In Progress',
    statusClosed: 'Kapali',
    savePenalty: 'Ceza/Uyari Kaydet',
    reportScopeInfo: 'Secilen proje kapsaminda rapor ciktisi olusturulur. All Projects secildiginde tum projeler birlesik raporlanir.',
    projectHealth: 'Proje Durum Karsilastirmasi',
    healthScore: 'Saglik Skoru',
    daily: 'Gunluk',
    weekly: 'Haftalik',
    monthly: 'Aylik',
    hseSpecialistFirstName: 'HSE Uzmani Adi',
    hseSpecialistLastName: 'HSE Uzmani Soyadi',
    kpiToolboxCount: 'Toolbox Sayisi',
    kpiToolboxPeople: 'Toolbox Kisi Sayisi',
    kpiJobInduction: 'Is Giris Egitimi Alan'
  },
  en: {
    language: 'Language',
    title: 'HSE Executive and Operations Platform',
    subtitle:
      'Weekly/monthly KPI monitoring, PTW/Observation closure tracking, risk visibility, and PPE operational control in one view.',
    loading: 'Loading dashboard data...',
    dataUnavailable: 'Data source is currently unavailable.',
    totalInspections: 'Total Inspections',
    openPtw: 'Open PTW',
    closedPtw: 'Closed PTW',
    openObservations: 'Open Observations',
    highCriticalRisks: 'High + Critical Risks',
    trainingCompliance: 'Training Compliance',
    ppeUsage: 'PPE Usage',
    reportActions: 'Report Actions',
    print: 'Print',
    downloadPdf: 'Download PDF',
    downloadExcel: 'Download Excel',
    emailPdf: 'Email PDF',
    emailExcel: 'Email Excel',
    emailPlaceholder: 'report@company.com',
    reportEmailRequired: 'Please enter report email address.',
    reportEmailTriggered: 'Report email dispatch started',
    inspectionEntry: 'Inspection Entry',
    employee: 'Employee',
    select: 'Select',
    contractor: 'Contractor',
    region: 'Region',
    workNature: 'Work Nature',
    riskLevel: 'Risk Level',
    findings: 'Findings',
    correctiveAction: 'Corrective Action',
    attachmentFileName: 'Attachment File Name',
    attachmentUrl: 'Attachment URL',
    attachmentType: 'Attachment Type',
    saveInspection: 'Save Inspection',
    inspectionSaved: 'Inspection record created successfully.',
    ptwStatus: 'PTW Status',
    permitNo: 'Permit No',
    owner: 'Owner',
    status: 'Status',
    observationStatus: 'Observation Status',
    type: 'Type',
    location: 'Location',
    riskRegisterSnapshot: 'Risk Register Snapshot',
    risk: 'Risk',
    score: 'Score',
    level: 'Level',
    ppeStockUsage: 'PPE Stock and Usage',
    item: 'Item',
    lastIn: 'Last In',
    lastUsed: 'Last Used',
    currentLeft: 'Current Left',
    photoJpeg: 'Photo (JPEG)',
    photoPng: 'Photo (PNG)',
    certificatePdf: 'Certificate (PDF)',
    sidebarDashboard: 'Dashboard',
    sidebarAudit: 'Audit',
    sidebarKpis: 'KPIs',
    sidebarInspection: 'Inspection',
    sidebarObservation: 'Observation',
    sidebarTraining: 'Training',
    sidebarPpeStock: 'PPE Stock',
    sidebarReport: 'Report',
    uploadFile: 'Upload File',
    uploading: 'Uploading...',
    attachmentUploaded: 'Attachment uploaded successfully.',
    trainingName: 'Training Name',
    validUntil: 'Valid Until',
    certificateNo: 'Certificate No',
    projectFilter: 'Project Filter',
    allProjects: 'All Projects',
    projectName: 'Project Name',
    periodLabel: 'Period',
    kpiWorkforce: 'Workforce Count',
    kpiManHours: 'Accident-Free Man-Hours',
    kpiLti: 'LTI Count',
    kpiNearMiss: 'Near Miss Count',
    kpiPpeLeft: 'PPE Left',
    kpiPenaltyOpen: 'Open Penalty/Warning',
    saveKpi: 'Save KPI',
    inspectionDate: 'Inspection Date',
    firstName: 'First Name',
    lastName: 'Last Name',
    roleTitle: 'Role Title',
    certificateExpiryDate: 'Certificate Expiry Date',
    saveTraining: 'Save Training',
    quantityIn: 'Quantity In',
    quantityUsed: 'Quantity Used',
    quantityLeft: 'Quantity Left',
    entryDate: 'Entry Date',
    savePpe: 'Save PPE',
    penaltySource: 'Penalty/Warning Source',
    customer: 'Customer',
    government: 'Government',
    penaltyTitle: 'Penalty/Warning Title',
    penaltyAmount: 'Amount',
    statusOpen: 'Open',
    statusInProgress: 'In Progress',
    statusClosed: 'Closed',
    savePenalty: 'Save Penalty/Warning',
    reportScopeInfo: 'Reports are generated by selected project scope. Selecting All Projects combines all project data.',
    projectHealth: 'Project Health Comparison',
    healthScore: 'Health Score',
    daily: 'Daily',
    weekly: 'Weekly',
    monthly: 'Monthly',
    hseSpecialistFirstName: 'HSE Specialist First Name',
    hseSpecialistLastName: 'HSE Specialist Last Name',
    kpiToolboxCount: 'Daily Toolbox Count',
    kpiToolboxPeople: 'Toolbox Participants',
    kpiJobInduction: 'Job Induction Participants'
  },
  ru: {
    language: 'Язык',
    title: 'Платформа HSE для руководства и операций',
    subtitle:
      'Еженедельный/ежемесячный контроль KPI, отслеживание закрытия PTW/предписаний, видимость рисков и контроль PPE в одном интерфейсе.',
    loading: 'Загрузка данных панели...',
    dataUnavailable: 'Источник данных временно недоступен.',
    totalInspections: 'Всего инспекций',
    openPtw: 'Открытые PTW',
    closedPtw: 'Закрытые PTW',
    openObservations: 'Открытые предписания',
    highCriticalRisks: 'Высокие + Критические риски',
    trainingCompliance: 'Соответствие обучению',
    ppeUsage: 'Использование PPE',
    reportActions: 'Действия с отчетами',
    print: 'Печать',
    downloadPdf: 'Скачать PDF',
    downloadExcel: 'Скачать Excel',
    emailPdf: 'Отправить PDF',
    emailExcel: 'Отправить Excel',
    emailPlaceholder: 'report@company.com',
    reportEmailRequired: 'Введите email для отправки отчета.',
    reportEmailTriggered: 'Отправка отчета по email запущена',
    inspectionEntry: 'Ввод инспекции',
    employee: 'Сотрудник',
    select: 'Выберите',
    contractor: 'Подрядчик',
    region: 'Регион',
    workNature: 'Характер работ',
    riskLevel: 'Уровень риска',
    findings: 'Предписания',
    correctiveAction: 'Корректирующее действие',
    attachmentFileName: 'Имя вложения',
    attachmentUrl: 'URL вложения',
    attachmentType: 'Тип вложения',
    saveInspection: 'Сохранить инспекцию',
    inspectionSaved: 'Запись инспекции успешно создана.',
    ptwStatus: 'Статус PTW',
    permitNo: 'Номер Permit',
    owner: 'Ответственный',
    status: 'Статус',
    observationStatus: 'Статус предписания',
    type: 'Тип',
    location: 'Локация',
    riskRegisterSnapshot: 'Сводка реестра рисков',
    risk: 'Риск',
    score: 'Оценка',
    level: 'Уровень',
    ppeStockUsage: 'Склад и расход PPE',
    item: 'Номенклатура',
    lastIn: 'Последнее поступление',
    lastUsed: 'Последний расход',
    currentLeft: 'Остаток',
    photoJpeg: 'Фото (JPEG)',
    photoPng: 'Фото (PNG)',
    certificatePdf: 'Сертификат (PDF)',
    sidebarDashboard: 'Панель управления',
    sidebarAudit: 'Аудит',
    sidebarKpis: 'KPI',
    sidebarInspection: 'Инспекция',
    sidebarObservation: 'Предписание',
    sidebarTraining: 'Обучение',
    sidebarPpeStock: 'Склад PPE',
    sidebarReport: 'Отчет',
    uploadFile: 'Загрузить файл',
    uploading: 'Загрузка...',
    attachmentUploaded: 'Вложение успешно загружено.',
    trainingName: 'Название обучения',
    validUntil: 'Действует до',
    certificateNo: 'Номер сертификата',
    projectFilter: 'Фильтр проекта',
    allProjects: 'Все проекты',
    projectName: 'Название проекта',
    periodLabel: 'Период',
    kpiWorkforce: 'Численность персонала',
    kpiManHours: 'Человеко-часы без аварий',
    kpiLti: 'Количество LTI',
    kpiNearMiss: 'Количество Near Miss',
    kpiPpeLeft: 'Остаток PPE',
    kpiPenaltyOpen: 'Открытые штрафы/предупреждения',
    saveKpi: 'Сохранить KPI',
    inspectionDate: 'Дата инспекции',
    firstName: 'Имя',
    lastName: 'Фамилия',
    roleTitle: 'Должность',
    certificateExpiryDate: 'Срок действия сертификата',
    saveTraining: 'Сохранить обучение',
    quantityIn: 'Поступило',
    quantityUsed: 'Использовано',
    quantityLeft: 'Остаток',
    entryDate: 'Дата ввода',
    savePpe: 'Сохранить PPE',
    penaltySource: 'Источник штрафа/предупреждения',
    customer: 'Заказчик',
    government: 'Государство',
    penaltyTitle: 'Заголовок штрафа/предупреждения',
    penaltyAmount: 'Сумма',
    statusOpen: 'Открыто',
    statusInProgress: 'В процессе',
    statusClosed: 'Закрыто',
    savePenalty: 'Сохранить штраф/предупреждение',
    reportScopeInfo: 'Отчет формируется в рамках выбранного проекта. При выборе Все проекты данные объединяются.',
    projectHealth: 'Сравнение состояния проектов',
    healthScore: 'Индекс состояния',
    daily: 'Ежедневно',
    weekly: 'Еженедельно',
    monthly: 'Ежемесячно',
    hseSpecialistFirstName: 'Имя специалиста HSE',
    hseSpecialistLastName: 'Фамилия специалиста HSE',
    kpiToolboxCount: 'Количество Toolbox в день',
    kpiToolboxPeople: 'Участники Toolbox',
    kpiJobInduction: 'Прошли вводный инструктаж'
  }
};
