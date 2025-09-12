
// Enhanced W2 to Form 1040 mapping with all missing fields

import { Form1040Data } from './form-1040-types';
import { EnhancedW2Data, W2Box12Code, getBox12CodeDescription } from './tax-document-types';

export class EnhancedW2ToForm1040Mapper {
  /**
   * Maps enhanced W2 data with all missing fields to Form 1040
   */
  static mapEnhancedW2ToForm1040(w2Data: EnhancedW2Data, existingForm1040?: Partial<Form1040Data>): Partial<Form1040Data> {
    console.log('🔍 [Enhanced W2 MAPPER] Starting enhanced W2 to 1040 mapping...');
    
    const form1040Data: Partial<Form1040Data> = {
      ...existingForm1040,
    };

    // Personal Information Mapping (existing logic)
    if (w2Data.employeeName) {
      const nameParts = w2Data.employeeName.trim().split(/\s+/);
      form1040Data.firstName = nameParts[0] || '';
      form1040Data.lastName = nameParts.slice(1).join(' ') || '';
    }

    if (w2Data.employeeSSN) {
      form1040Data.ssn = this.formatSSN(w2Data.employeeSSN);
    }

    // Enhanced address mapping with new state/local fields
    if (w2Data.employeeAddress) {
      const addressParts = this.parseAddress(w2Data.employeeAddress);
      form1040Data.address = addressParts.street;
      form1040Data.city = addressParts.city;
      form1040Data.state = addressParts.state;
      form1040Data.zipCode = addressParts.zipCode;
    }

    // Core income mapping (existing)
    if (w2Data.wages && w2Data.wages > 0) {
      form1040Data.line1 = (form1040Data.line1 || 0) + w2Data.wages;
    }

    if (w2Data.federalTaxWithheld && w2Data.federalTaxWithheld > 0) {
      form1040Data.line25a = (form1040Data.line25a || 0) + w2Data.federalTaxWithheld;
    }

    // ENHANCED: Process Box 10 - Dependent Care Benefits
    if (w2Data.dependentCareBenefits && w2Data.dependentCareBenefits > 0) {
      console.log('🔍 [Enhanced W2 MAPPER] Processing dependent care benefits:', w2Data.dependentCareBenefits);
      
      // Dependent care benefits up to $5,000 are excludable from income
      const maxExcludable = 5000;
      const excludableAmount = Math.min(w2Data.dependentCareBenefits, maxExcludable);
      const taxableAmount = Math.max(0, w2Data.dependentCareBenefits - maxExcludable);
      
      if (taxableAmount > 0) {
        // Add taxable portion to wages (Line 1)
        form1040Data.line1 = (form1040Data.line1 || 0) + taxableAmount;
        console.log(`✅ [Enhanced W2 MAPPER] Added taxable dependent care benefits to Line 1: $${taxableAmount}`);
      }
      
      // Store for Form 2441 (Child and Dependent Care Expenses)
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.dependentCareExcludable = excludableAmount;
      form1040Data.scheduleA.dependentCareTaxable = taxableAmount;
      
      console.log(`✅ [Enhanced W2 MAPPER] Processed dependent care benefits - Excludable: $${excludableAmount}, Taxable: $${taxableAmount}`);
    }

    // ENHANCED: Process Box 12 - Deferred Compensation Codes
    if (w2Data.box12Codes && w2Data.box12Codes.length > 0) {
      console.log('🔍 [Enhanced W2 MAPPER] Processing Box 12 codes:', w2Data.box12Codes);
      
      for (const { code, amount } of w2Data.box12Codes) {
        this.processBox12Code(code, amount, form1040Data);
      }
    }

    // ENHANCED: Process Box 13 - Checkboxes
    if (w2Data.box13Checkboxes) {
      console.log('🔍 [Enhanced W2 MAPPER] Processing Box 13 checkboxes:', w2Data.box13Checkboxes);
      
      // Retirement plan participation affects IRA deduction limits
      if (w2Data.box13Checkboxes.retirementPlan) {
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.retirementPlanParticipation = true;
        console.log('✅ [Enhanced W2 MAPPER] Marked retirement plan participation');
      }
      
      // Statutory employee affects Schedule C reporting
      if (w2Data.box13Checkboxes.statutoryEmployee) {
        if (!form1040Data.schedule1) form1040Data.schedule1 = {};
        form1040Data.schedule1.statutoryEmployee = true;
        console.log('✅ [Enhanced W2 MAPPER] Marked statutory employee status');
      }
      
      // Third-party sick pay (informational)
      if (w2Data.box13Checkboxes.thirdPartySickPay) {
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.thirdPartySickPay = true;
        console.log('✅ [Enhanced W2 MAPPER] Marked third-party sick pay');
      }
    }

    // ENHANCED: Process additional missing fields
    if (w2Data.socialSecurityTips && w2Data.socialSecurityTips > 0) {
      // Social security tips are included in wages but tracked separately
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.socialSecurityTips = w2Data.socialSecurityTips;
      console.log(`✅ [Enhanced W2 MAPPER] Recorded social security tips: $${w2Data.socialSecurityTips}`);
    }

    if (w2Data.allocatedTips && w2Data.allocatedTips > 0) {
      // Allocated tips may need to be added to income if not already included
      if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
      form1040Data.scheduleA.allocatedTips = w2Data.allocatedTips;
      console.log(`✅ [Enhanced W2 MAPPER] Recorded allocated tips: $${w2Data.allocatedTips}`);
    }

    // ENHANCED: Enhanced state and local tax information
    if (w2Data.stateWages && w2Data.stateWages > 0) {
      if (!form1040Data.stateData) form1040Data.stateData = {};
      form1040Data.stateData.stateWages = w2Data.stateWages;
      form1040Data.stateData.stateEmployerID = w2Data.stateEmployerID || '';
      console.log(`✅ [Enhanced W2 MAPPER] Mapped state wages: $${w2Data.stateWages}`);
    }

    if (w2Data.stateTaxWithheld && w2Data.stateTaxWithheld > 0) {
      if (!form1040Data.stateData) form1040Data.stateData = {};
      form1040Data.stateData.taxWithheld = (form1040Data.stateData.taxWithheld || 0) + w2Data.stateTaxWithheld;
      console.log(`✅ [Enhanced W2 MAPPER] Mapped state tax withheld: $${w2Data.stateTaxWithheld}`);
    }

    if (w2Data.localWages && w2Data.localWages > 0) {
      if (!form1040Data.stateData) form1040Data.stateData = {};
      form1040Data.stateData.localWages = w2Data.localWages;
      form1040Data.stateData.localityName = w2Data.localityName || '';
      console.log(`✅ [Enhanced W2 MAPPER] Mapped local wages: $${w2Data.localWages}`);
    }

    if (w2Data.localTaxWithheld && w2Data.localTaxWithheld > 0) {
      if (!form1040Data.stateData) form1040Data.stateData = {};
      form1040Data.stateData.localTaxWithheld = w2Data.localTaxWithheld;
      console.log(`✅ [Enhanced W2 MAPPER] Mapped local tax withheld: $${w2Data.localTaxWithheld}`);
    }

    // Create enhanced personal info object
    const personalInfo = {
      firstName: form1040Data.firstName ?? '',
      lastName: form1040Data.lastName ?? '',
      ssn: form1040Data.ssn ?? '',
      address: form1040Data.address ?? '',
      city: form1040Data.city ?? '',
      state: form1040Data.state ?? '',
      zipCode: form1040Data.zipCode ?? '',
      sourceDocument: 'Enhanced W2',
      sourceDocumentId: String(w2Data.documentId || 'unknown')
    };

    form1040Data.personalInfo = personalInfo;

    console.log('✅ [Enhanced W2 MAPPER] Enhanced mapping completed successfully');
    return form1040Data;
  }

  /**
   * Process Box 12 codes and map to appropriate Form 1040 fields
   */
  private static processBox12Code(code: W2Box12Code, amount: number, form1040Data: Partial<Form1040Data>): void {
    console.log(`🔍 [Enhanced W2 MAPPER] Processing Box 12 code ${code}: $${amount} - ${getBox12CodeDescription(code)}`);
    
    switch (code) {
      case W2Box12Code.D: // 401(k) deferrals
      case W2Box12Code.E: // 403(b) deferrals
      case W2Box12Code.F: // SEP deferrals
      case W2Box12Code.G: // 457(b) deferrals
      case W2Box12Code.H: // 501(c)(18)(D) deferrals
      case W2Box12Code.S: // SIMPLE deferrals
        // These reduce current year taxable income (already excluded from Box 1)
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.retirementContributions = (form1040Data.scheduleA.retirementContributions || 0) + amount;
        console.log(`✅ [Enhanced W2 MAPPER] Added retirement contribution (${code}): $${amount}`);
        break;
        
      case W2Box12Code.W: // HSA contributions
        // HSA contributions are deductible (already excluded from Box 1)
        if (!form1040Data.schedule1) form1040Data.schedule1 = {};
        form1040Data.schedule1.hsaContributions = (form1040Data.schedule1.hsaContributions || 0) + amount;
        console.log(`✅ [Enhanced W2 MAPPER] Added HSA contribution: $${amount}`);
        break;
        
      case W2Box12Code.C: // Group-term life insurance over $50,000
        // This is taxable income (should already be included in Box 1)
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.groupTermLifeInsurance = amount;
        console.log(`✅ [Enhanced W2 MAPPER] Recorded group-term life insurance: $${amount}`);
        break;
        
      case W2Box12Code.AA: // Roth 401(k) contributions
      case W2Box12Code.BB: // Roth 403(b) contributions
      case W2Box12Code.EE: // Roth 457(b) contributions
        // Roth contributions are made with after-tax dollars (included in Box 1)
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.rothContributions = (form1040Data.scheduleA.rothContributions || 0) + amount;
        console.log(`✅ [Enhanced W2 MAPPER] Added Roth contribution (${code}): $${amount}`);
        break;
        
      case W2Box12Code.DD: // Cost of employer-sponsored health coverage
        // This is informational only (not taxable)
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.employerHealthCoverage = amount;
        console.log(`✅ [Enhanced W2 MAPPER] Recorded employer health coverage cost: $${amount}`);
        break;
        
      case W2Box12Code.A: // Uncollected social security tax on tips
      case W2Box12Code.B: // Uncollected Medicare tax on tips
        // These need to be added to tax liability
        if (!form1040Data.schedule1) form1040Data.schedule1 = {};
        form1040Data.schedule1.uncollectedTaxOnTips = (form1040Data.schedule1.uncollectedTaxOnTips || 0) + amount;
        console.log(`✅ [Enhanced W2 MAPPER] Added uncollected tax on tips (${code}): $${amount}`);
        break;
        
      case W2Box12Code.V: // Income from nonstatutory stock options
        // This should be included in Box 1 wages
        console.log(`✅ [Enhanced W2 MAPPER] Noted nonstatutory stock option income (${code}): $${amount}`);
        break;
        
      case W2Box12Code.T: // Adoption benefits
        // Up to annual limit is excludable
        const adoptionExclusionLimit = 15950; // 2023 limit
        const excludableAdoption = Math.min(amount, adoptionExclusionLimit);
        const taxableAdoption = Math.max(0, amount - adoptionExclusionLimit);
        
        if (taxableAdoption > 0) {
          form1040Data.line1 = (form1040Data.line1 || 0) + taxableAdoption;
        }
        
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        form1040Data.scheduleA.adoptionBenefits = excludableAdoption;
        console.log(`✅ [Enhanced W2 MAPPER] Processed adoption benefits - Excludable: $${excludableAdoption}, Taxable: $${taxableAdoption}`);
        break;
        
      default:
        // For other codes, store for reference
        if (!form1040Data.scheduleA) form1040Data.scheduleA = {};
        if (!form1040Data.scheduleA.otherBox12Codes) form1040Data.scheduleA.otherBox12Codes = [];
        form1040Data.scheduleA.otherBox12Codes.push({ code, amount, description: getBox12CodeDescription(code) });
        console.log(`✅ [Enhanced W2 MAPPER] Stored other Box 12 code (${code}): $${amount}`);
        break;
    }
  }

  /**
   * Create enhanced mapping summary
   */
  static createEnhancedMappingSummary(w2Data: EnhancedW2Data): Array<{
    w2Field: string;
    w2Value: any;
    form1040Line: string;
    form1040Value: any;
    description: string;
  }> {
    const mappings = [];

    // Add mappings for all new fields
    if (w2Data.dependentCareBenefits) {
      mappings.push({
        w2Field: 'Box 10 - Dependent Care Benefits',
        w2Value: w2Data.dependentCareBenefits,
        form1040Line: 'Line 1 (if over $5,000) / Form 2441',
        form1040Value: w2Data.dependentCareBenefits,
        description: 'Dependent care benefits (excludable up to $5,000)'
      });
    }

    if (w2Data.box12Codes) {
      for (const { code, amount } of w2Data.box12Codes) {
        mappings.push({
          w2Field: `Box 12${code} - ${getBox12CodeDescription(code)}`,
          w2Value: amount,
          form1040Line: this.getBox12MappingLine(code),
          form1040Value: amount,
          description: getBox12CodeDescription(code)
        });
      }
    }

    if (w2Data.box13Checkboxes?.retirementPlan) {
      mappings.push({
        w2Field: 'Box 13 - Retirement Plan',
        w2Value: 'Checked',
        form1040Line: 'IRA Deduction Limitation',
        form1040Value: 'Affects IRA deduction limits',
        description: 'Employee participated in retirement plan'
      });
    }

    return mappings;
  }

  private static getBox12MappingLine(code: W2Box12Code): string {
    switch (code) {
      case W2Box12Code.D:
      case W2Box12Code.E:
      case W2Box12Code.F:
      case W2Box12Code.G:
      case W2Box12Code.H:
      case W2Box12Code.S:
        return 'Pre-tax retirement contribution (excluded from Line 1)';
      case W2Box12Code.W:
        return 'Schedule 1 - HSA deduction';
      case W2Box12Code.C:
        return 'Line 1 - Taxable group-term life insurance';
      case W2Box12Code.AA:
      case W2Box12Code.BB:
      case W2Box12Code.EE:
        return 'After-tax Roth contribution (included in Line 1)';
      case W2Box12Code.DD:
        return 'Informational only';
      case W2Box12Code.A:
      case W2Box12Code.B:
        return 'Schedule 1 - Additional tax';
      case W2Box12Code.T:
        return 'Line 1 (if over limit) / Form 8839';
      default:
        return 'Various forms depending on code';
    }
  }

  // Utility methods (existing)
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
    // Enhanced address parsing logic (existing implementation)
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
