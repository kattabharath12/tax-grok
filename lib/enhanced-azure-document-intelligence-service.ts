
// Enhanced Azure Document Intelligence Service with all missing fields

import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";
import { readFile } from "fs/promises";
import { 
  EnhancedW2Data, 
  Enhanced1099DivData, 
  W2Box12Code, 
  W2Box13Checkboxes,
  parseW2Box12Codes 
} from './tax-document-types';

export interface ExtractedFieldData {
  [key: string]: string | number | DocumentType | number[] | boolean |Array<{ code: W2Box12Code; amount: number }> | W2Box13Checkboxes | undefined;
  correctedDocumentType?: DocumentType;
  fullText?: string;
}

export class EnhancedAzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: { endpoint: string; apiKey: string };

  constructor(config: { endpoint: string; apiKey: string }) {
    this.config = config;
    this.client = new DocumentAnalysisClient(
      this.config.endpoint,
      new AzureKeyCredential(this.config.apiKey)
    );
  }

  // Enhanced W2 field processing with all missing fields
  private processEnhancedW2Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const w2Data = { ...baseData };
    
    // Enhanced W2 field mappings with ALL missing fields
    const enhancedW2FieldMappings = {
      // Existing fields
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN',
      'Employee.Address': 'employeeAddress',
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'WagesAndTips': 'wages',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'SocialSecurityWages': 'socialSecurityWages',
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'MedicareWagesAndTips': 'medicareWages',
      'MedicareTaxWithheld': 'medicareTaxWithheld',
      
      // MISSING FIELDS - NOW ADDED
      'SocialSecurityTips': 'socialSecurityTips', // Box 7
      'AllocatedTips': 'allocatedTips', // Box 8
      'AdvanceEIC': 'advanceEIC', // Box 9
      'DependentCareBenefits': 'dependentCareBenefits', // Box 10 - CRITICAL
      'NonqualifiedPlans': 'nonqualifiedPlans', // Box 11
      'DeferredCompensation': 'box12Raw', // Box 12 - will be parsed
      'OtherTaxInfo': 'otherTaxInfo', // Box 14
      
      // Enhanced state/local fields
      'StateEmployerID': 'stateEmployerID', // Box 15
      'StateWagesTipsEtc': 'stateWages', // Box 16
      'StateIncomeTax': 'stateTaxWithheld', // Box 17
      'LocalWagesTipsEtc': 'localWages', // Box 18
      'LocalIncomeTax': 'localTaxWithheld', // Box 19
      'LocalityName': 'localityName', // Box 20
      
      // Alternative field names
      'Box7': 'socialSecurityTips',
      'Box8': 'allocatedTips',
      'Box9': 'advanceEIC',
      'Box10': 'dependentCareBenefits',
      'Box11': 'nonqualifiedPlans',
      'Box12': 'box12Raw',
      'Box14': 'otherTaxInfo',
      'Box15': 'stateEmployerID',
      'Box16': 'stateWages',
      'Box17': 'stateTaxWithheld',
      'Box18': 'localWages',
      'Box19': 'localTaxWithheld',
      'Box20': 'localityName'
    };
    
    // Process standard field mappings
    for (const [azureFieldName, mappedFieldName] of Object.entries(enhancedW2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        if (mappedFieldName === 'box12Raw' || mappedFieldName === 'otherTaxInfo' || 
            mappedFieldName === 'stateEmployerID' || mappedFieldName === 'localityName') {
          // Text fields
          w2Data[mappedFieldName] = String(value).trim();
        } else {
          // Numeric fields
          w2Data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // ENHANCED: Parse Box 12 codes from raw text
    if (w2Data.box12Raw) {
      const box12Codes = parseW2Box12Codes(w2Data.box12Raw as string);
      if (box12Codes.length > 0) {
        w2Data.box12Codes = box12Codes;
        console.log('✅ [Enhanced Azure DI] Parsed Box 12 codes:', box12Codes);
      }
    }
    
    // ENHANCED: Extract Box 13 checkboxes
    w2Data.box13Checkboxes = this.extractW2Box13Checkboxes(fields, baseData.fullText as string);
    
    // OCR fallback for missing fields
    if (baseData.fullText) {
      this.enhanceW2WithOCRFallback(w2Data, baseData.fullText as string);
    }
    
    return w2Data;
  }

  // Enhanced 1099-DIV field processing with all missing fields
  private processEnhanced1099DivFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    // Enhanced 1099-DIV field mappings with ALL missing fields
    const enhanced1099DivFieldMappings = {
      // Existing fields
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'OrdinaryDividends': 'ordinaryDividends', // Box 1a
      'QualifiedDividends': 'qualifiedDividends', // Box 1b
      'TotalCapitalGainDistributions': 'totalCapitalGain', // Box 2a
      'NondividendDistributions': 'nondividendDistributions', // Box 3
      'FederalIncomeTaxWithheld': 'federalTaxWithheld', // Box 4
      'Section199ADividends': 'section199ADividends', // Box 5
      
      // MISSING FIELDS - NOW ADDED
      'UnrecapturedSection1250Gain': 'unrecapturedSection1250Gain', // Box 2b - CRITICAL
      'Section1202Gain': 'section1202Gain', // Box 2c - CRITICAL
      'CollectiblesGain': 'collectiblesGain', // Box 2d - CRITICAL
      'Section897OrdinaryDividends': 'section897OrdinaryDividends', // Box 2e - CRITICAL
      'Section897CapitalGain': 'section897CapitalGain', // Box 2f - CRITICAL
      'ExemptInterestDividends': 'exemptInterestDividends', // Box 6 - CRITICAL
      'ForeignTaxPaid': 'foreignTaxPaid', // Box 7 - CRITICAL
      'ForeignCountry': 'foreignCountry', // Box 8 - CRITICAL
      'CashLiquidationDistributions': 'cashLiquidationDistributions', // Box 9 - CRITICAL
      'NoncashLiquidationDistributions': 'noncashLiquidationDistributions', // Box 10 - CRITICAL
      'FATCAFilingRequirement': 'fatcaFilingRequirement', // Box 11 - CRITICAL
      'InvestmentExpenses': 'investmentExpenses', // Box 13 - CRITICAL
      
      // Alternative field names
      'Box1a': 'ordinaryDividends',
      'Box1b': 'qualifiedDividends',
      'Box2a': 'totalCapitalGain',
      'Box2b': 'unrecapturedSection1250Gain',
      'Box2c': 'section1202Gain',
      'Box2d': 'collectiblesGain',
      'Box2e': 'section897OrdinaryDividends',
      'Box2f': 'section897CapitalGain',
      'Box3': 'nondividendDistributions',
      'Box4': 'federalTaxWithheld',
      'Box5': 'section199ADividends',
      'Box6': 'exemptInterestDividends',
      'Box7': 'foreignTaxPaid',
      'Box8': 'foreignCountry',
      'Box9': 'cashLiquidationDistributions',
      'Box10': 'noncashLiquidationDistributions',
      'Box11': 'fatcaFilingRequirement',
      'Box13': 'investmentExpenses',
      
      // State information
      'StateTaxWithheld': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',
      'StateIncome': 'stateIncome'
    };
    
    // Process field mappings
    for (const [azureFieldName, mappedFieldName] of Object.entries(enhanced1099DivFieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        if (mappedFieldName === 'foreignCountry' || mappedFieldName === 'statePayerNumber') {
          // Text fields
          data[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'fatcaFilingRequirement') {
          // Boolean field
          data[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields
          data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // OCR fallback for missing fields
    if (baseData.fullText) {
      this.enhance1099DivWithOCRFallback(data, baseData.fullText as string);
    }
    
    return data;
  }

  // Extract W2 Box 13 checkboxes
  private extractW2Box13Checkboxes(fields: any, ocrText: string): W2Box13Checkboxes {
    const checkboxes: W2Box13Checkboxes = {};
    
    // Try structured fields first
    if (fields['RetirementPlan']?.value !== undefined) {
      checkboxes.retirementPlan = this.parseBoolean(fields['RetirementPlan'].value);
    }
    if (fields['ThirdPartySickPay']?.value !== undefined) {
      checkboxes.thirdPartySickPay = this.parseBoolean(fields['ThirdPartySickPay'].value);
    }
    if (fields['StatutoryEmployee']?.value !== undefined) {
      checkboxes.statutoryEmployee = this.parseBoolean(fields['StatutoryEmployee'].value);
    }
    
    // OCR fallback for checkboxes
    if (ocrText && (!checkboxes.retirementPlan && !checkboxes.thirdPartySickPay && !checkboxes.statutoryEmployee)) {
      const ocrCheckboxes = this.extractBox13CheckboxesFromOCR(ocrText);
      Object.assign(checkboxes, ocrCheckboxes);
    }
    
    return checkboxes;
  }

  // OCR fallback for W2 Box 13 checkboxes
  private extractBox13CheckboxesFromOCR(ocrText: string): W2Box13Checkboxes {
    const checkboxes: W2Box13Checkboxes = {};
    const text = ocrText.toLowerCase();
    
    // Patterns for retirement plan checkbox
    const retirementPlanPatterns = [
      /retirement\s+plan\s*[:\s]*(?:x|✓|checked|yes)/i,
      /13\s*retirement\s+plan\s*[:\s]*(?:x|✓|checked|yes)/i,
      /box\s*13.*retirement.*(?:x|✓|checked|yes)/i
    ];
    
    // Patterns for third-party sick pay
    const thirdPartySickPayPatterns = [
      /third.party\s+sick\s+pay\s*[:\s]*(?:x|✓|checked|yes)/i,
      /13\s*third.party\s+sick\s+pay\s*[:\s]*(?:x|✓|checked|yes)/i
    ];
    
    // Patterns for statutory employee
    const statutoryEmployeePatterns = [
      /statutory\s+employee\s*[:\s]*(?:x|✓|checked|yes)/i,
      /13\s*statutory\s+employee\s*[:\s]*(?:x|✓|checked|yes)/i
    ];
    
    // Check patterns
    checkboxes.retirementPlan = retirementPlanPatterns.some(pattern => pattern.test(text));
    checkboxes.thirdPartySickPay = thirdPartySickPayPatterns.some(pattern => pattern.test(text));
    checkboxes.statutoryEmployee = statutoryEmployeePatterns.some(pattern => pattern.test(text));
    
    return checkboxes;
  }

  // Enhanced OCR fallback for W2 missing fields
  private enhanceW2WithOCRFallback(w2Data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Enhanced Azure DI] Applying OCR fallback for missing W2 fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'socialSecurityTips', patterns: [/box\s*7[:\s]*\$?(\d+(?:\.\d{2})?)/i, /social\s+security\s+tips[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'allocatedTips', patterns: [/box\s*8[:\s]*\$?(\d+(?:\.\d{2})?)/i, /allocated\s+tips[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'advanceEIC', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /advance\s+eic[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'dependentCareBenefits', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /dependent\s+care\s+benefits[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'nonqualifiedPlans', patterns: [/box\s*11[:\s]*\$?(\d+(?:\.\d{2})?)/i, /nonqualified\s+plans[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'stateWages', patterns: [/box\s*16[:\s]*\$?(\d+(?:\.\d{2})?)/i, /state\s+wages[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'localWages', patterns: [/box\s*18[:\s]*\$?(\d+(?:\.\d{2})?)/i, /local\s+wages[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'localTaxWithheld', patterns: [/box\s*19[:\s]*\$?(\d+(?:\.\d{2})?)/i, /local\s+income\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!w2Data[field] || w2Data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              w2Data[field] = amount;
              console.log(`✅ [Enhanced Azure DI] Extracted ${field} from OCR: $${amount}`);
              break;
            }
          }
        }
      }
    }
    
    // Extract Box 12 codes if not already parsed
    if (!w2Data.box12Codes && !w2Data.box12Raw) {
      const box12Pattern = /box\s*12[:\s]*([A-Z]{1,2}\s*\$?\d+(?:\.\d{2})?(?:\s+[A-Z]{1,2}\s*\$?\d+(?:\.\d{2})?)*)/i;
      const match = ocrText.match(box12Pattern);
      if (match && match[1]) {
        const box12Codes = parseW2Box12Codes(match[1]);
        if (box12Codes.length > 0) {
          w2Data.box12Codes = box12Codes;
          console.log('✅ [Enhanced Azure DI] Extracted Box 12 codes from OCR:', box12Codes);
        }
      }
    }
    
    // Extract text fields
    if (!w2Data.stateEmployerID) {
      const stateIdPattern = /box\s*15[:\s]*([A-Z0-9\-]+)/i;
      const match = ocrText.match(stateIdPattern);
      if (match && match[1]) {
        w2Data.stateEmployerID = match[1].trim();
        console.log(`✅ [Enhanced Azure DI] Extracted state employer ID from OCR: ${w2Data.stateEmployerID}`);
      }
    }
    
    if (!w2Data.localityName) {
      const localityPattern = /box\s*20[:\s]*([A-Za-z\s]+)/i;
      const match = ocrText.match(localityPattern);
      if (match && match[1]) {
        w2Data.localityName = match[1].trim();
        console.log(`✅ [Enhanced Azure DI] Extracted locality name from OCR: ${w2Data.localityName}`);
      }
    }
  }

  // Enhanced OCR fallback for 1099-DIV missing fields
  private enhance1099DivWithOCRFallback(data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Enhanced Azure DI] Applying OCR fallback for missing 1099-DIV fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'unrecapturedSection1250Gain', patterns: [/box\s*2b[:\s]*\$?(\d+(?:\.\d{2})?)/i, /unrecap.*sec.*1250[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section1202Gain', patterns: [/box\s*2c[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*1202[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'collectiblesGain', patterns: [/box\s*2d[:\s]*\$?(\d+(?:\.\d{2})?)/i, /collectibles.*28%[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section897OrdinaryDividends', patterns: [/box\s*2e[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*897.*ordinary[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section897CapitalGain', patterns: [/box\s*2f[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*897.*capital[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'exemptInterestDividends', patterns: [/box\s*6[:\s]*\$?(\d+(?:\.\d{2})?)/i, /exempt.interest\s+dividends[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'foreignTaxPaid', patterns: [/box\s*7[:\s]*\$?(\d+(?:\.\d{2})?)/i, /foreign\s+tax\s+paid[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'cashLiquidationDistributions', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /cash\s+liquidation[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'noncashLiquidationDistributions', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /noncash\s+liquidation[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'investmentExpenses', patterns: [/box\s*13[:\s]*\$?(\d+(?:\.\d{2})?)/i, /investment\s+expenses[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!data[field] || data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              data[field] = amount;
              console.log(`✅ [Enhanced Azure DI] Extracted ${field} from OCR: $${amount}`);
              break;
            }
          }
        }
      }
    }
    
    // Extract foreign country
    if (!data.foreignCountry) {
      const foreignCountryPattern = /box\s*8[:\s]*([A-Za-z\s]+)/i;
      const match = ocrText.match(foreignCountryPattern);
      if (match && match[1]) {
        data.foreignCountry = match[1].trim();
        console.log(`✅ [Enhanced Azure DI] Extracted foreign country from OCR: ${data.foreignCountry}`);
      }
    }
    
    // Extract FATCA filing requirement
    if (data.fatcaFilingRequirement === undefined) {
      const fatcaPattern = /box\s*11[:\s]*(?:x|✓|checked|yes)/i;
      data.fatcaFilingRequirement = fatcaPattern.test(ocrText);
      if (data.fatcaFilingRequirement) {
        console.log('✅ [Enhanced Azure DI] Extracted FATCA filing requirement from OCR: true');
      }
    }
  }

  // Utility methods
  private parseAmount(value: any): number {
    if (value === null || value === undefined) return 0;
    if (typeof value === 'number') return isNaN(value) ? 0 : value;
    if (typeof value === 'string') {
      const cleaned = value.replace(/[$,\s]/g, '');
      const parsed = parseFloat(cleaned);
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }

  private parseBoolean(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const lower = value.toLowerCase().trim();
      return lower === 'true' || lower === 'yes' || lower === 'x' || lower === '✓' || lower === 'checked';
    }
    return false;
  }
}
