export type RiskLevel = 'LOW' | 'MEDIUM' | 'HIGH' | 'CRITICAL';

export type StatusOpenClosed = 'OPEN' | 'CLOSED';

export interface WeeklyMonthlyPeriod {
  periodType: 'WEEKLY' | 'MONTHLY';
  periodStart: string;
  periodEnd: string;
}
