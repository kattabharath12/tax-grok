
// Enhanced tax document type definitions with all missing fields

// W2 Box 12 Codes (A-EE) - Complete IRS list
export enum W2Box12Code {
  A = 'A', // Uncollected social security or RRTA tax on tips
  B = 'B', // Uncollected Medicare tax on tips
  C = 'C', // Taxable cost of group-term life insurance over $50,000
  D = 'D', // Elective deferrals under a section 401(k) cash or deferred arrangement
  E = 'E', // Elective deferrals under a section 403(b) salary reduction agreement
  F = 'F', // Elective deferrals under a section 408(k)(6) salary reduction SEP
  G = 'G', // Elective deferrals and employer contributions to a section 457(b) deferred compensation plan
  H = 'H', // Elective deferrals under a section 501(c)(18)(D) tax-exempt organization plan
  J = 'J', // Nontaxable sick pay
  K = 'K', // 20% excise tax on excess golden parachute payments
  L = 'L', // Substantiated employee business expense reimbursements
  M = 'M', // Uncollected social security or RRTA tax on taxable cost of group-term life insurance over $50,000
  N = 'N', // Uncollected Medicare tax on taxable cost of group-term life insurance over $50,000
  P = 'P', // Excludable moving expense reimbursements paid directly to member of Armed Forces
  Q = 'Q', // Nontaxable combat pay
  R = 'R', // Employer contributions to an Archer MSA
  S = 'S', // Employee salary reduction contributions under a section 408(p) SIMPLE
  T = 'T', // Adoption benefits
  V = 'V', // Income from the exercise of nonstatutory stock options
  W = 'W', // Employer contributions to an employee health savings account (HSA)
  Y = 'Y', // Deferrals under a section 409A nonqualified deferred compensation plan
  Z = 'Z', // Income under a section 409A nonqualified deferred compensation plan
  AA = 'AA', // Designated Roth contributions under a section 401(k) plan
  BB = 'BB', // Designated Roth contributions under a section 403(b) plan
  DD = 'DD', // Cost of employer-sponsored health coverage
  EE = 'EE', // Designated Roth contributions under a governmental section 457(b) plan
  FF = 'FF', // Permitted benefits under a qualified small employer health reimbursement arrangement
  GG = 'GG', // Income from qualified equity grants under section 83(i)
  HH = 'HH' // Aggregate deferrals for section 83(i) elections as of the close of the calendar year
}

// W2 Box 13 Checkboxes
export interface W2Box13Checkboxes {
  retirementPlan?: boolean; // Retirement plan checkbox
  thirdPartySickPay?: boolean; // Third-party sick pay checkbox
  statutoryEmployee?: boolean; // Statutory employee checkbox
}

// Enhanced W2 Data Interface with all missing fields
export interface EnhancedW2Data {
  // Existing fields
  employeeName?: string;
  employeeSSN?: string;
  employeeAddress?: string;
  employerName?: string;
  employerEIN?: string;
  employerAddress?: string;
  wages?: number; // Box 1
  federalTaxWithheld?: number; // Box 2
  socialSecurityWages?: number; // Box 3
  socialSecurityTaxWithheld?: number; // Box 4
  medicareWages?: number; // Box 5
  medicareTaxWithheld?: number; // Box 6
  
  // MISSING FIELDS - NOW ADDED
  socialSecurityTips?: number; // Box 7
  allocatedTips?: number; // Box 8
  advanceEIC?: number; // Box 9
  dependentCareBenefits?: number; // Box 10 - CRITICAL MISSING FIELD
  nonqualifiedPlans?: number; // Box 11
  
  // Box 12: Deferred compensation codes (CRITICAL MISSING FIELD)
  box12Codes?: Array<{
    code: W2Box12Code;
    amount: number;
  }>;
  
  // Box 13: Checkboxes (MISSING FIELD)
  box13Checkboxes?: W2Box13Checkboxes;
  
  otherTaxInfo?: string; // Box 14
  
  // Enhanced state/local information (MISSING FIELDS)
  stateEmployerID?: string; // Box 15 - State employer's state ID number
  stateWages?: number; // Box 16 - State wages, tips, etc.
  stateTaxWithheld?: number; // Box 17 - State income tax
  localWages?: number; // Box 18 - Local wages, tips, etc.
  localTaxWithheld?: number; // Box 19 - Local income tax
  localityName?: string; // Box 20 - Locality name
  
  // Document tracking
  documentId?: string;
}

// Enhanced 1099-DIV Data Interface with all missing fields
export interface Enhanced1099DivData {
  // Existing fields
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  ordinaryDividends?: number; // Box 1a
  qualifiedDividends?: number; // Box 1b
  totalCapitalGain?: number; // Box 2a
  nondividendDistributions?: number; // Box 3
  federalTaxWithheld?: number; // Box 4
  section199ADividends?: number; // Box 5
  
  // MISSING FIELDS - NOW ADDED
  unrecapturedSection1250Gain?: number; // Box 2b - CRITICAL MISSING FIELD
  section1202Gain?: number; // Box 2c - CRITICAL MISSING FIELD
  collectiblesGain?: number; // Box 2d - CRITICAL MISSING FIELD
  section897OrdinaryDividends?: number; // Box 2e - CRITICAL MISSING FIELD
  section897CapitalGain?: number; // Box 2f - CRITICAL MISSING FIELD
  exemptInterestDividends?: number; // Box 6 - CRITICAL MISSING FIELD
  foreignTaxPaid?: number; // Box 7 - CRITICAL MISSING FIELD
  foreignCountry?: string; // Box 8 - CRITICAL MISSING FIELD
  cashLiquidationDistributions?: number; // Box 9 - CRITICAL MISSING FIELD
  noncashLiquidationDistributions?: number; // Box 10 - CRITICAL MISSING FIELD
  fatcaFilingRequirement?: boolean; // Box 11 - CRITICAL MISSING FIELD
  investmentExpenses?: number; // Box 13 - CRITICAL MISSING FIELD
  
  // State information
  stateTaxWithheld?: number; // Box 14
  statePayerNumber?: string; // Box 15
  stateIncome?: number; // Box 16
  
  // Document tracking
  documentId?: string;
}

// Type guards for enhanced data
export function isEnhancedW2Data(data: any): data is EnhancedW2Data {
  return data && typeof data === 'object';
}

export function isEnhanced1099DivData(data: any): data is Enhanced1099DivData {
  return data && typeof data === 'object';
}

// Utility functions for Box 12 codes
export function parseW2Box12Codes(box12String?: string): Array<{ code: W2Box12Code; amount: number }> {
  if (!box12String) return [];
  
  const codes: Array<{ code: W2Box12Code; amount: number }> = [];
  
  // Parse format like "D 5000.00 W 2500.00" or "D$5000 W$2500"
  const codePattern = /([A-Z]{1,2})\s*\$?(\d+(?:\.\d{2})?)/g;
  let match;
  
  while ((match = codePattern.exec(box12String)) !== null) {
    const code = match[1] as W2Box12Code;
    const amount = parseFloat(match[2]);
    
    if (Object.values(W2Box12Code).includes(code) && !isNaN(amount)) {
      codes.push({ code, amount });
    }
  }
  
  return codes;
}

export function formatW2Box12Codes(codes: Array<{ code: W2Box12Code; amount: number }>): string {
  return codes.map(({ code, amount }) => `${code} $${amount.toFixed(2)}`).join(' ');
}

// Tax impact calculations for new fields
export function calculateDependentCareBenefitTaxImpact(dependentCareBenefits: number): {
  excludableAmount: number;
  taxableAmount: number;
} {
  const maxExcludable = 5000; // 2023 limit
  const excludableAmount = Math.min(dependentCareBenefits, maxExcludable);
  const taxableAmount = Math.max(0, dependentCareBenefits - maxExcludable);
  
  return { excludableAmount, taxableAmount };
}

export function getBox12CodeDescription(code: W2Box12Code): string {
  const descriptions: Record<W2Box12Code, string> = {
    [W2Box12Code.A]: 'Uncollected social security or RRTA tax on tips',
    [W2Box12Code.B]: 'Uncollected Medicare tax on tips',
    [W2Box12Code.C]: 'Taxable cost of group-term life insurance over $50,000',
    [W2Box12Code.D]: 'Elective deferrals under a section 401(k) plan',
    [W2Box12Code.E]: 'Elective deferrals under a section 403(b) plan',
    [W2Box12Code.F]: 'Elective deferrals under a section 408(k)(6) salary reduction SEP',
    [W2Box12Code.G]: 'Elective deferrals under a section 457(b) plan',
    [W2Box12Code.H]: 'Elective deferrals under a section 501(c)(18)(D) plan',
    [W2Box12Code.J]: 'Nontaxable sick pay',
    [W2Box12Code.K]: '20% excise tax on excess golden parachute payments',
    [W2Box12Code.L]: 'Substantiated employee business expense reimbursements',
    [W2Box12Code.M]: 'Uncollected social security tax on group-term life insurance',
    [W2Box12Code.N]: 'Uncollected Medicare tax on group-term life insurance',
    [W2Box12Code.P]: 'Excludable moving expense reimbursements (Armed Forces)',
    [W2Box12Code.Q]: 'Nontaxable combat pay',
    [W2Box12Code.R]: 'Employer contributions to an Archer MSA',
    [W2Box12Code.S]: 'Employee salary reduction contributions under a SIMPLE plan',
    [W2Box12Code.T]: 'Adoption benefits',
    [W2Box12Code.V]: 'Income from nonstatutory stock options',
    [W2Box12Code.W]: 'Employer contributions to employee HSA',
    [W2Box12Code.Y]: 'Deferrals under a section 409A plan',
    [W2Box12Code.Z]: 'Income under a section 409A plan',
    [W2Box12Code.AA]: 'Designated Roth contributions under a 401(k) plan',
    [W2Box12Code.BB]: 'Designated Roth contributions under a 403(b) plan',
    [W2Box12Code.DD]: 'Cost of employer-sponsored health coverage',
    [W2Box12Code.EE]: 'Designated Roth contributions under a governmental 457(b) plan',
    [W2Box12Code.FF]: 'Qualified small employer health reimbursement arrangement',
    [W2Box12Code.GG]: 'Income from qualified equity grants under section 83(i)',
    [W2Box12Code.HH]: 'Aggregate deferrals for section 83(i) elections'
  };
  
  return descriptions[code] || 'Unknown code';
}
