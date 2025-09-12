
// Enhanced 1099-DIV to Form 1040 mapping with all missing fields

import { Form1040Data } from './form-1040-types';
import { Enhanced1099DivData } from './tax-document-types';

export class Enhanced1099DivToForm1040Mapper {
  /**
   * Maps enhanced 1099-DIV data with all missing fields to Form 1040
   */
  static mapEnhanced1099DivToForm1040(divData: Enhanced1099DivData, existingForm1040?: Partial<Form1040Data>): Partial<Form1040Data> {
    console.log('🔍 [Enhanced 1099-DIV MAPPER] Starting enhanced 1099-DIV to 1040 mapping...');
    
    const form1040Data: Partial<Form1040Data> = {
      ...existingForm1040,
    };

    // Personal Information Mapping (existing logic)
    if (divData.recipientName && (!form1040Data.personalInfo?.sourceDocument?.includes('W2'))) {
      const nameParts = divData.recipientName.trim().split(/\s+/);
      form1040Data.firstName = nameParts[0] || '';
      form1040Data.lastName = nameParts.slice(1).join(' ') || '';
    }

    if (divData.recipientTIN && (!form1040Data.personalInfo?.sourceDocument?.includes('W2'))) {
      form1040Data.ssn = this.formatSSN(divData.recipientTIN);
    }

    // Enhanced address mapping
    if (divData.recipientAddress && (!form1040Data.personalInfo?.sourceDocument?.includes('W2'))) {
      const addressParts = this.parseAddress(divData.recipientAddress);
      form1040Data.address = addressParts.street;
      form1040Data.city = addressParts.city;
      form1040Data.state = addressParts.state;
      form1040Data.zipCode = addressParts.zipCode;
    }

    // EXISTING: Basic dividend income mapping
    if (divData.ordinaryDividends && divData.ordinaryDividends > 0) {
      form1040Data.line3b = (form1040Data.line3b || 0) + divData.ordinaryDividends;
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped ordinary dividends to Line 3b: $${divData.ordinaryDividends}`);
    }

    if (divData.qualifiedDividends && divData.qualifiedDividends > 0) {
      form1040Data.line3a = (form1040Data.line3a || 0) + divData.qualifiedDividends;
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped qualified dividends to Line 3a: $${divData.qualifiedDividends}`);
    }

    if (divData.totalCapitalGain && divData.totalCapitalGain > 0) {
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.totalCapitalGain;
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped capital gain distributions to Line 7: $${divData.totalCapitalGain}`);
    }

    if (divData.federalTaxWithheld && divData.federalTaxWithheld > 0) {
      form1040Data.line25a = (form1040Data.line25a || 0) + divData.federalTaxWithheld;
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped federal tax withheld to Line 25a: $${divData.federalTaxWithheld}`);
    }

    // ENHANCED: Process Box 2b - Unrecaptured Section 1250 Gain (25% tax rate)
    if (divData.unrecapturedSection1250Gain && divData.unrecapturedSection1250Gain > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing unrecaptured Section 1250 gain:', divData.unrecapturedSection1250Gain);
      
      // This goes to Schedule D and affects capital gains tax calculation
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.unrecapturedSection1250Gain = (form1040Data.scheduleA.unrecapturedSection1250Gain || 0) + divData.unrecapturedSection1250Gain;
      
      // Also add to Line 7 for total capital gains
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.unrecapturedSection1250Gain;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped unrecaptured Section 1250 gain (25% rate): $${divData.unrecapturedSection1250Gain}`);
    }

    // ENHANCED: Process Box 2c - Section 1202 Gain (small business stock exclusion)
    if (divData.section1202Gain && divData.section1202Gain > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing Section 1202 gain:', divData.section1202Gain);
      
      // Section 1202 gain may be partially excludable (up to $10M or 10x basis)
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.section1202Gain = (form1040Data.scheduleA.section1202Gain || 0) + divData.section1202Gain;
      
      // Add to Line 7 for total capital gains
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.section1202Gain;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped Section 1202 gain (small business stock): $${divData.section1202Gain}`);
    }

    // ENHANCED: Process Box 2d - Collectibles Gain (28% tax rate)
    if (divData.collectiblesGain && divData.collectiblesGain > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing collectibles gain:', divData.collectiblesGain);
      
      // Collectibles are taxed at 28% maximum rate
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.collectiblesGain = (form1040Data.scheduleA.collectiblesGain || 0) + divData.collectiblesGain;
      
      // Add to Line 7 for total capital gains
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.collectiblesGain;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped collectibles gain (28% rate): $${divData.collectiblesGain}`);
    }

    // ENHANCED: Process Box 2e - Section 897 Ordinary Dividends
    if (divData.section897OrdinaryDividends && divData.section897OrdinaryDividends > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing Section 897 ordinary dividends:', divData.section897OrdinaryDividends);
      
      // These are treated as ordinary dividends but may have special reporting requirements
      form1040Data.line3b = (form1040Data.line3b || 0) + divData.section897OrdinaryDividends;
      
      // Track separately for FIRPTA reporting
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.section897OrdinaryDividends = (form1040Data.scheduleA.section897OrdinaryDividends || 0) + divData.section897OrdinaryDividends;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped Section 897 ordinary dividends: $${divData.section897OrdinaryDividends}`);
    }

    // ENHANCED: Process Box 2f - Section 897 Capital Gain
    if (divData.section897CapitalGain && divData.section897CapitalGain > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing Section 897 capital gain:', divData.section897CapitalGain);
      
      // These are treated as capital gains but may have special reporting requirements
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.section897CapitalGain;
      
      // Track separately for FIRPTA reporting
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.section897CapitalGain = (form1040Data.scheduleA.section897CapitalGain || 0) + divData.section897CapitalGain;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped Section 897 capital gain: $${divData.section897CapitalGain}`);
    }

    // ENHANCED: Process Box 6 - Exempt-Interest Dividends
    if (divData.exemptInterestDividends && divData.exemptInterestDividends > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing exempt-interest dividends:', divData.exemptInterestDividends);
      
      // These go to Line 2a (tax-exempt interest)
      form1040Data.line2a = (form1040Data.line2a || 0) + divData.exemptInterestDividends;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped exempt-interest dividends to Line 2a: $${divData.exemptInterestDividends}`);
    }

    // ENHANCED: Process Box 7 - Foreign Tax Paid
    if (divData.foreignTaxPaid && divData.foreignTaxPaid > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing foreign tax paid:', divData.foreignTaxPaid);
      
      // Foreign tax paid can be claimed as credit or deduction
      if (!form1040Data.foreignTaxCredit) form1040Data.foreignTaxCredit = {};
      form1040Data.foreignTaxCredit.foreignTaxPaid = (form1040Data.foreignTaxCredit.foreignTaxPaid || 0) + divData.foreignTaxPaid;
      
      if (divData.foreignCountry) {
        form1040Data.foreignTaxCredit.foreignCountry = divData.foreignCountry;
      }
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped foreign tax paid for credit: $${divData.foreignTaxPaid}`);
    }

    // ENHANCED: Process Box 9 - Cash Liquidation Distributions
    if (divData.cashLiquidationDistributions && divData.cashLiquidationDistributions > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing cash liquidation distributions:', divData.cashLiquidationDistributions);
      
      // Liquidation distributions are typically treated as capital gains
      form1040Data.line7 = (form1040Data.line7 || 0) + divData.cashLiquidationDistributions;
      
      // Track separately for basis adjustment
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.cashLiquidationDistributions = (form1040Data.scheduleA.cashLiquidationDistributions || 0) + divData.cashLiquidationDistributions;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped cash liquidation distributions to Line 7: $${divData.cashLiquidationDistributions}`);
    }

    // ENHANCED: Process Box 10 - Noncash Liquidation Distributions
    if (divData.noncashLiquidationDistributions && divData.noncashLiquidationDistributions > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing noncash liquidation distributions:', divData.noncashLiquidationDistributions);
      
      // Noncash liquidations require special handling - typically not immediately taxable
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.noncashLiquidationDistributions = (form1040Data.scheduleA.noncashLiquidationDistributions || 0) + divData.noncashLiquidationDistributions;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Recorded noncash liquidation distributions (special handling required): $${divData.noncashLiquidationDistributions}`);
    }

    // ENHANCED: Process Box 11 - FATCA Filing Requirement
    if (divData.fatcaFilingRequirement) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing FATCA filing requirement');
      
      // FATCA filing requirement affects Form 8938 reporting
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.fatcaFilingRequired = true;
      
      console.log('✅ [Enhanced 1099-DIV MAPPER] Marked FATCA filing requirement (Form 8938 may be required)');
    }

    // ENHANCED: Process Box 13 - Investment Expenses
    if (divData.investmentExpenses && divData.investmentExpenses > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing investment expenses:', divData.investmentExpenses);
      
      // Investment expenses go to Schedule A (itemized deductions)
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.investmentExpenses = (form1040Data.scheduleA.investmentExpenses || 0) + divData.investmentExpenses;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped investment expenses to Schedule A: $${divData.investmentExpenses}`);
    }

    // ENHANCED: Process nondividend distributions (return of capital)
    if (divData.nondividendDistributions && divData.nondividendDistributions > 0) {
      console.log('🔍 [Enhanced 1099-DIV MAPPER] Processing nondividend distributions:', divData.nondividendDistributions);
      
      // Nondividend distributions reduce basis and are not immediately taxable
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.nondividendDistributions = (form1040Data.scheduleA.nondividendDistributions || 0) + divData.nondividendDistributions;
      
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Recorded nondividend distributions (reduces basis): $${divData.nondividendDistributions}`);
    }

    // Enhanced state tax information
    if (divData.stateTaxWithheld && divData.stateTaxWithheld > 0) {
      if (!form1040Data.stateData) form1040Data.stateData = {};
      form1040Data.stateData.taxWithheld = (form1040Data.stateData.taxWithheld || 0) + divData.stateTaxWithheld;
      console.log(`✅ [Enhanced 1099-DIV MAPPER] Mapped state tax withheld: $${divData.stateTaxWithheld}`);
    }

    // Create enhanced personal info object
    if (!form1040Data.personalInfo?.sourceDocument?.includes('W2')) {
      const personalInfo = {
        firstName: form1040Data.firstName ?? '',
        lastName: form1040Data.lastName ?? '',
        ssn: form1040Data.ssn ?? '',
        address: form1040Data.address ?? '',
        city: form1040Data.city ?? '',
        state: form1040Data.state ?? '',
        zipCode: form1040Data.zipCode ?? '',
        sourceDocument: form1040Data.personalInfo?.sourceDocument ? 
          `${form1040Data.personalInfo.sourceDocument}, Enhanced 1099-DIV` : 'Enhanced 1099-DIV',
        sourceDocumentId: String(divData.documentId || 'unknown')
      };

      form1040Data.personalInfo = personalInfo;
    }

    console.log('✅ [Enhanced 1099-DIV MAPPER] Enhanced mapping completed successfully');
    return form1040Data;
  }

  /**
   * Create enhanced mapping summary for all new fields
   */
  static createEnhancedMappingSummary(divData: Enhanced1099DivData): Array<{
    divField: string;
    divValue: any;
    form1040Line: string;
    form1040Value: any;
    description: string;
    taxTreatment: string;
  }> {
    const mappings = [];

    // Add mappings for all enhanced fields
    if (divData.unrecapturedSection1250Gain) {
      mappings.push({
        divField: 'Box 2b - Unrecaptured Section 1250 Gain',
        divValue: divData.unrecapturedSection1250Gain,
        form1040Line: 'Line 7 / Schedule D',
        form1040Value: divData.unrecapturedSection1250Gain,
        description: 'Depreciation recapture on real estate',
        taxTreatment: 'Taxed at 25% maximum rate'
      });
    }

    if (divData.section1202Gain) {
      mappings.push({
        divField: 'Box 2c - Section 1202 Gain',
        divValue: divData.section1202Gain,
        form1040Line: 'Line 7 / Schedule D',
        form1040Value: divData.section1202Gain,
        description: 'Small business stock gain',
        taxTreatment: 'May be partially excludable (up to $10M or 10x basis)'
      });
    }

    if (divData.collectiblesGain) {
      mappings.push({
        divField: 'Box 2d - Collectibles Gain',
        divValue: divData.collectiblesGain,
        form1040Line: 'Line 7 / Schedule D',
        form1040Value: divData.collectiblesGain,
        description: 'Gain from collectibles',
        taxTreatment: 'Taxed at 28% maximum rate'
      });
    }

    if (divData.section897OrdinaryDividends) {
      mappings.push({
        divField: 'Box 2e - Section 897 Ordinary Dividends',
        divValue: divData.section897OrdinaryDividends,
        form1040Line: 'Line 3b',
        form1040Value: divData.section897OrdinaryDividends,
        description: 'FIRPTA ordinary dividends',
        taxTreatment: 'Taxed as ordinary income'
      });
    }

    if (divData.section897CapitalGain) {
      mappings.push({
        divField: 'Box 2f - Section 897 Capital Gain',
        divValue: divData.section897CapitalGain,
        form1040Line: 'Line 7',
        form1040Value: divData.section897CapitalGain,
        description: 'FIRPTA capital gain',
        taxTreatment: 'Taxed as capital gain'
      });
    }

    if (divData.exemptInterestDividends) {
      mappings.push({
        divField: 'Box 6 - Exempt-Interest Dividends',
        divValue: divData.exemptInterestDividends,
        form1040Line: 'Line 2a',
        form1040Value: divData.exemptInterestDividends,
        description: 'Tax-exempt interest dividends',
        taxTreatment: 'Not taxable for federal income tax'
      });
    }

    if (divData.foreignTaxPaid) {
      mappings.push({
        divField: 'Box 7 - Foreign Tax Paid',
        divValue: divData.foreignTaxPaid,
        form1040Line: 'Form 1116 / Schedule 3',
        form1040Value: divData.foreignTaxPaid,
        description: 'Foreign tax paid on dividends',
        taxTreatment: 'Eligible for foreign tax credit or deduction'
      });
    }

    if (divData.cashLiquidationDistributions) {
      mappings.push({
        divField: 'Box 9 - Cash Liquidation Distributions',
        divValue: divData.cashLiquidationDistributions,
        form1040Line: 'Line 7',
        form1040Value: divData.cashLiquidationDistributions,
        description: 'Cash received in corporate liquidation',
        taxTreatment: 'Treated as capital gain/loss'
      });
    }

    if (divData.noncashLiquidationDistributions) {
      mappings.push({
        divField: 'Box 10 - Noncash Liquidation Distributions',
        divValue: divData.noncashLiquidationDistributions,
        form1040Line: 'Special handling required',
        form1040Value: divData.noncashLiquidationDistributions,
        description: 'Property received in corporate liquidation',
        taxTreatment: 'Complex basis and recognition rules apply'
      });
    }

    if (divData.fatcaFilingRequirement) {
      mappings.push({
        divField: 'Box 11 - FATCA Filing Requirement',
        divValue: 'Required',
        form1040Line: 'Form 8938',
        form1040Value: 'Filing required',
        description: 'Foreign account reporting requirement',
        taxTreatment: 'Informational - affects Form 8938 filing'
      });
    }

    if (divData.investmentExpenses) {
      mappings.push({
        divField: 'Box 13 - Investment Expenses',
        divValue: divData.investmentExpenses,
        form1040Line: 'Schedule A',
        form1040Value: divData.investmentExpenses,
        description: 'Investment-related expenses',
        taxTreatment: 'Itemized deduction (subject to limitations)'
      });
    }

    return mappings;
  }

  /**
   * Calculate tax impact of enhanced 1099-DIV fields
   */
  static calculateEnhancedTaxImpact(divData: Enhanced1099DivData): {
    ordinaryIncome: number;
    qualifiedDividends: number;
    capitalGains: number;
    specialRateGains: number;
    taxExemptIncome: number;
    foreignTaxCredit: number;
    deductions: number;
  } {
    return {
      ordinaryIncome: (divData.ordinaryDividends || 0) + (divData.section897OrdinaryDividends || 0),
      qualifiedDividends: divData.qualifiedDividends || 0,
      capitalGains: (divData.totalCapitalGain || 0) + (divData.section897CapitalGain || 0) + 
                   (divData.cashLiquidationDistributions || 0),
      specialRateGains: (divData.unrecapturedSection1250Gain || 0) + (divData.collectiblesGain || 0) + 
                       (divData.section1202Gain || 0),
      taxExemptIncome: divData.exemptInterestDividends || 0,
      foreignTaxCredit: divData.foreignTaxPaid || 0,
      deductions: divData.investmentExpenses || 0
    };
  }

  // Utility methods
  private static formatSSN(ssn: string): string {
    if (!ssn) return '';
    const cleaned = ssn.replace(/\D/g, '');
    if (cleaned.length === 9) {
      return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 5)}-${cleaned.slice(5)}`;
    }
    return cleaned;
  }

  private static parseAddress(address: string): {
    street: string;
    city: string;
    state: string;
    zipCode: string;
  } {
    const commaParts = address.split(',').map(part => part.trim());
    
    if (commaParts.length >= 3) {
      const street = commaParts.slice(0, -2).join(', ');
      const city = commaParts[commaParts.length - 2];
      const stateZip = commaParts[commaParts.length - 1];
      const stateZipMatch = stateZip.match(/^([A-Z]{2})\s*(\d{5}(-\d{4})?)$/);
      
      if (stateZipMatch) {
        return {
          street,
          city,
          state: stateZipMatch[1],
          zipCode: stateZipMatch[2]
        };
      }
    }
    
    return {
      street: address,
      city: '',
      state: '',
      zipCode: ''
    };
  }
}
