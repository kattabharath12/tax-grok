import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { DocumentType } from "@prisma/client";
import { readFile } from "fs/promises";

// Helper function to safely extract error messages
function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

export interface ExtractedFieldData {
  [key: string]: string | number | DocumentType | number[] | boolean | Array<{ code: string; amount: number }> | undefined;
  correctedDocumentType?: DocumentType;
  fullText?: string;
  box12Codes?: Array<{ code: string; amount: number }>;
}

// === COMPREHENSIVE TAX DOCUMENT INTERFACES ===

// ENHANCED: Complete W-2 interface with ALL 20+ fields
export interface TaxW2Result {
  // Employee Information (Boxes a, e, f)
  employeeName?: string;
  employeeSSN?: string;
  employeeAddress?: string;
  employeeAddressStreet?: string;
  employeeCity?: string;
  employeeState?: string;
  employeeZipCode?: string;
  
  // Employer Information (Boxes b, c)
  employerName?: string;
  employerEIN?: string;
  employerAddress?: string;
  controlNumber?: string; // Box d
  
  // Box 1-20: All wage and tax information
  wages?: number;                           // Box 1: Wages, tips, other compensation
  federalIncomeTaxWithheld?: number;        // Box 2: Federal income tax withheld
  socialSecurityWages?: number;             // Box 3: Social security wages
  socialSecurityTaxWithheld?: number;       // Box 4: Social security tax withheld
  medicareWages?: number;                   // Box 5: Medicare wages and tips
  medicareTaxWithheld?: number;             // Box 6: Medicare tax withheld
  socialSecurityTips?: number;              // Box 7: Social security tips
  allocatedTips?: number;                   // Box 8: Allocated tips
  advanceEIC?: number;                      // Box 9: Advance EIC payments
  dependentCareBenefits?: number;           // Box 10: Dependent care benefits
  nonqualifiedPlans?: number;               // Box 11: Nonqualified plans
  box12Codes?: Array<{ code: string; amount: number }>; // Box 12: Deferred compensation codes
  box13Checkboxes?: {                       // Box 13: Checkboxes
    statutoryEmployee?: boolean;
    retirementPlan?: boolean;
    thirdPartySickPay?: boolean;
  };
  otherInfo?: string;                       // Box 14: Other
  
  // State and Local Information (Boxes 15-20)
  stateEmployerID?: string;                 // Box 15: State/Employer's state ID no.
  stateWages?: number;                      // Box 16: State wages, tips, etc.
  stateIncomeTax?: number;                  // Box 17: State income tax
  localWages?: number;                      // Box 18: Local wages, tips, etc.
  localIncomeTax?: number;                  // Box 19: Local income tax
  localityName?: string;                    // Box 20: Locality name
  stateName?: string;                       // State name
  
  // Additional processing fields
  fullText?: string;
  correctedDocumentType?: DocumentType;
}

// ENHANCED: Complete 1099-INT interface with ALL 15+ fields
export interface Tax1099IntResult {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  payerPhone?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  secondTINNotification?: boolean;
  
  // Box 1-15: All interest income fields
  interestIncome?: number;                              // Box 1: Interest income
  earlyWithdrawalPenalty?: number;                      // Box 2: Early withdrawal penalty
  interestOnUSSavingsBonds?: number;                    // Box 3: Interest on U.S. Savings Bonds and Treasury obligations
  federalIncomeTaxWithheld?: number;                    // Box 4: Federal income tax withheld
  investmentExpenses?: number;                          // Box 5: Investment expenses
  foreignTaxPaid?: number;                              // Box 6: Foreign tax paid
  foreignCountry?: string;                              // Box 7: Foreign country or U.S. possession
  taxExemptInterest?: number;                           // Box 8: Tax-exempt interest
  specifiedPrivateActivityBondInterest?: number;        // Box 9: Specified private activity bond interest
  marketDiscount?: number;                              // Box 10: Market discount
  bondPremium?: number;                                 // Box 11: Bond premium
  bondPremiumOnTreasury?: number;                       // Box 12: Bond premium on Treasury obligations
  bondPremiumOnTaxExempt?: number;                      // Box 13: Bond premium on tax-exempt bond
  cusipNumber?: string;                                 // Box 14: Tax-exempt and tax credit bond CUSIP no.
  state?: string;                                       // Box 15: State
  stateIdentificationNumber?: string;                   // Box 16: State identification no.
  stateTaxWithheld?: number;                           // Box 17: State tax withheld
  
  // Additional processing fields
  fullText?: string;
  correctedDocumentType?: DocumentType;
}

// ENHANCED: Complete 1099-MISC interface with ALL 18 fields
export interface Tax1099MiscResult {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  payerPhone?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  secondTINNotification?: boolean;
  
  // Box 1-18: All miscellaneous income fields
  rents?: number;                                       // Box 1: Rents
  royalties?: number;                                   // Box 2: Royalties
  otherIncome?: number;                                 // Box 3: Other income
  federalIncomeTaxWithheld?: number;                    // Box 4: Federal income tax withheld
  fishingBoatProceeds?: number;                         // Box 5: Fishing boat proceeds
  medicalHealthPayments?: number;                       // Box 6: Medical and health care payments
  payerMadeDirectSales?: boolean;                       // Box 7: Payer made direct sales totaling $5,000 or more (checkbox)
  substitutePayments?: number;                          // Box 8: Substitute payments in lieu of dividends or interest
  cropInsuranceProceeds?: number;                       // Box 9: Crop insurance proceeds
  grossProceedsAttorney?: number;                       // Box 10: Gross proceeds paid to an attorney
  fishPurchasedForResale?: number;                      // Box 11: Fish purchased for resale
  section409ADeferrals?: number;                        // Box 12: Section 409A deferrals
  fatcaFilingRequirement?: boolean;                     // Box 13: FATCA filing requirement (checkbox)
  // Box 14: Reserved (not used)
  section409AIncome?: number;                           // Box 15: Nonqualified deferred compensation
  stateTaxWithheld?: number;                           // Box 16: State tax withheld
  statePayerNumber?: string;                           // Box 17: State/Payer's state no.
  stateIncome?: number;                                // Box 18: State income
  
  // State information
  stateName?: string;
  
  // Additional processing fields
  fullText?: string;
  correctedDocumentType?: DocumentType;
}

// ENHANCED: Complete 1099-DIV interface with ALL fields from IRS form (KEEP EXISTING)
export interface Tax1099DivResult {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  payerPhone?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  secondTINNotification?: boolean;
  
  // Box 1a & 1b - Dividend Income
  totalOrdinaryDividends?: number;        // Box 1a - Total ordinary dividends
  qualifiedDividends?: number;            // Box 1b - Qualified dividends
  
  // Box 2a-2f - Capital Gain Distributions
  totalCapitalGainDistributions?: number; // Box 2a - Total capital gain distributions
  unrecapturedSection1250Gain?: number;   // Box 2b - Unrecaptured Section 1250 gain
  section1202Gain?: number;               // Box 2c - Section 1202 gain
  collectiblesGain?: number;              // Box 2d - Collectibles (28%) gain
  section897OrdinaryDividends?: number;   // Box 2e - Section 897 ordinary dividends
  section897CapitalGain?: number;         // Box 2f - Section 897 capital gain
  
  // Box 3-13 - Other Distributions and Information
  nondividendDistributions?: number;      // Box 3 - Nondividend distributions
  federalIncomeTaxWithheld?: number;      // Box 4 - Federal income tax withheld
  section199ADividends?: number;          // Box 5 - Section 199A dividends
  investmentExpenses?: number;            // Box 6 - Investment expenses
  foreignTaxPaid?: number;                // Box 7 - Foreign tax paid
  foreignCountry?: string;                // Box 8 - Foreign country or U.S. possession
  cashLiquidationDistributions?: number;  // Box 9 - Cash liquidation distributions
  noncashLiquidationDistributions?: number; // Box 10 - Noncash liquidation distributions
  fatcaFilingRequirement?: boolean;       // Box 11 - FATCA filing requirement
  exemptInterestDividends?: number;       // Box 12 - Exempt-interest dividends
  specifiedPrivateActivityBondInterest?: number; // Box 13 - Specified private activity bond interest dividends
  
  // State Information (Boxes 14-16)
  state?: string;                         // Box 14 - State
  stateIdentificationNumber?: string;     // Box 15 - State identification no.
  stateTaxWithheld?: number;             // Box 16 - State tax withheld
  
  // Additional fields for enhanced processing
  fullText?: string;
  correctedDocumentType?: DocumentType;
}

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;
  private config: AzureDocumentIntelligenceConfig;

  constructor(config: AzureDocumentIntelligenceConfig) {
    this.config = config;
    this.client = new DocumentAnalysisClient(
      this.config.endpoint,
      new AzureKeyCredential(this.config.apiKey)
    );
  }

  async extractDataFromDocument(
    documentPathOrBuffer: string | Buffer,
    documentType: string
  ): Promise<ExtractedFieldData> {
    try {
      console.log('🔍 [Azure DI] Processing document with Azure Document Intelligence...');
      console.log('🔍 [Azure DI] Initial document type:', documentType);
      
      // Get document buffer - either from file path or use provided buffer
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Determine the model to use based on document type
      const modelId = this.getModelIdForDocumentType(documentType);
      console.log('🔍 [Azure DI] Using model:', modelId);
      
      let extractedData: ExtractedFieldData;
      let correctedDocumentType: DocumentType | undefined;
      
      try {
        // Analyze the document with specific tax model
        const poller = await this.client.beginAnalyzeDocument(modelId, documentBuffer);
        const result = await poller.pollUntilDone();
        
        console.log('✅ [Azure DI] Document analysis completed with tax model');
        
        // Extract the data based on document type
        extractedData = this.extractTaxDocumentFields(result, documentType);
        
        // Perform OCR-based document type correction if we have OCR text
        if (extractedData.fullText) {
          const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText as string);
          if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
            console.log(`🔄 [Azure DI] Document type correction: ${documentType} → ${ocrBasedType}`);
            
            // Convert string to DocumentType enum with validation
            if (Object.values(DocumentType).includes(ocrBasedType as DocumentType)) {
              correctedDocumentType = ocrBasedType as DocumentType;
              
              // Re-extract data with the corrected document type
              console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
              extractedData = this.extractTaxDocumentFields(result, ocrBasedType);
            } else {
              console.log(`⚠️ [Azure DI] Invalid document type detected: ${ocrBasedType}, ignoring correction`);
            }
          }
        }
        
      } catch (modelError: unknown) {
        const modelErrorMessage = getErrorMessage(modelError);
        console.warn('⚠️ [Azure DI] Tax model failed, attempting fallback to OCR model:', modelErrorMessage);
        
        // Check if it's a ModelNotFound error
        if (modelErrorMessage.includes('ModelNotFound') || 
            modelErrorMessage.includes('Resource not found') ||
            (modelError instanceof Error && 'code' in modelError && modelError.code === 'NotFound')) {
          
          console.log('🔍 [Azure DI] Falling back to prebuilt-read model for OCR extraction...');
          
          // Fallback to general OCR model
          const fallbackPoller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
          const fallbackResult = await fallbackPoller.pollUntilDone();
          
          console.log('✅ [Azure DI] Document analysis completed with OCR fallback');
          
          // Extract data using OCR-based approach
          extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, documentType);
          
          // Perform OCR-based document type correction
          if (extractedData.fullText) {
            const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText as string);
            if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
              console.log(`🔄 [Azure DI] Document type correction (OCR fallback): ${documentType} → ${ocrBasedType}`);
              
              // Convert string to DocumentType enum with validation
              if (Object.values(DocumentType).includes(ocrBasedType as DocumentType)) {
                correctedDocumentType = ocrBasedType as DocumentType;
                
                // Re-extract data with the corrected document type
                console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
                extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, ocrBasedType);
              } else {
                console.log(`⚠️ [Azure DI] Invalid document type detected: ${ocrBasedType}, ignoring correction`);
              }
            }
          }
        } else {
          // Re-throw if it's not a model availability issue
          throw modelError;
        }
      }
      
      // Add the corrected document type to the result if it was changed
      if (correctedDocumentType) {
        extractedData.correctedDocumentType = correctedDocumentType;
      }
      
      return extractedData;
    } catch (error: unknown) {
      console.error('❌ [Azure DI] Processing error:', error);
      throw new Error(`Azure Document Intelligence processing failed: ${getErrorMessage(error)}`);
    }
  }

  // === PUBLIC TAX DOCUMENT EXTRACTION METHODS ===

  /**
   * ENHANCED: Extract data from W2 tax document with ALL 20+ fields
   * @param documentPathOrBuffer - Path to document file or Buffer containing document data
   * @returns Promise<TaxW2Result> - Extracted W2 data with comprehensive field coverage
   */
  async extractW2(documentPathOrBuffer: string | Buffer): Promise<TaxW2Result> {
    try {
      console.log('🔍 [Azure DI] Extracting W2 document with comprehensive field mapping...');
      const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, 'W2');
      
      return {
        // Employee Information
        employeeName: extractedData.employeeName as string,
        employeeSSN: extractedData.employeeSSN as string,
        employeeAddress: extractedData.employeeAddress as string,
        employeeAddressStreet: extractedData.employeeAddressStreet as string,
        employeeCity: extractedData.employeeCity as string,
        employeeState: extractedData.employeeState as string,
        employeeZipCode: extractedData.employeeZipCode as string,
        
        // Employer Information
        employerName: extractedData.employerName as string,
        employerEIN: extractedData.employerEIN as string,
        employerAddress: extractedData.employerAddress as string,
        controlNumber: extractedData.controlNumber as string,
        
        // Box 1-20: All wage and tax information
        wages: extractedData.wages as number,
        federalIncomeTaxWithheld: extractedData.federalIncomeTaxWithheld as number,
        socialSecurityWages: extractedData.socialSecurityWages as number,
        socialSecurityTaxWithheld: extractedData.socialSecurityTaxWithheld as number,
        medicareWages: extractedData.medicareWages as number,
        medicareTaxWithheld: extractedData.medicareTaxWithheld as number,
        socialSecurityTips: extractedData.socialSecurityTips as number,
        allocatedTips: extractedData.allocatedTips as number,
        advanceEIC: extractedData.advanceEIC as number,
        dependentCareBenefits: extractedData.dependentCareBenefits as number,
        nonqualifiedPlans: extractedData.nonqualifiedPlans as number,
        box12Codes: extractedData.box12Codes as Array<{ code: string; amount: number }>,
        box13Checkboxes: extractedData.box13Checkboxes as {
          statutoryEmployee?: boolean;
          retirementPlan?: boolean;
          thirdPartySickPay?: boolean;
        },
        otherInfo: extractedData.otherInfo as string,
        
        // State and Local Information
        stateEmployerID: extractedData.stateEmployerID as string,
        stateWages: extractedData.stateWages as number,
        stateIncomeTax: extractedData.stateIncomeTax as number,
        localWages: extractedData.localWages as number,
        localIncomeTax: extractedData.localIncomeTax as number,
        localityName: extractedData.localityName as string,
        stateName: extractedData.stateName as string,
        
        // Additional processing fields
        fullText: extractedData.fullText as string,
        correctedDocumentType: extractedData.correctedDocumentType
      };
    } catch (error: unknown) {
      console.error('❌ [Azure DI] W2 extraction failed:', error);
      throw new Error(`W2 extraction failed: ${getErrorMessage(error)}`);
    }
  }

  /**
   * ENHANCED: Extract data from 1099-INT tax document with ALL 15+ fields
   * @param documentPathOrBuffer - Path to document file or Buffer containing document data
   * @returns Promise<Tax1099IntResult> - Extracted 1099-INT data with comprehensive field coverage
   */
  async extract1099Int(documentPathOrBuffer: string | Buffer): Promise<Tax1099IntResult> {
    try {
      console.log('🔍 [Azure DI] Extracting 1099-INT document with comprehensive field mapping...');
      const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, 'FORM_1099_INT');
      
      return {
        // Payer Information
        payerName: extractedData.payerName as string,
        payerTIN: extractedData.payerTIN as string,
        payerAddress: extractedData.payerAddress as string,
        payerPhone: extractedData.payerPhone as string,
        
        // Recipient Information
        recipientName: extractedData.recipientName as string,
        recipientTIN: extractedData.recipientTIN as string,
        recipientAddress: extractedData.recipientAddress as string,
        accountNumber: extractedData.accountNumber as string,
        secondTINNotification: extractedData.secondTINNotification as boolean,
        
        // Box 1-15: All interest income fields
        interestIncome: extractedData.interestIncome as number,
        earlyWithdrawalPenalty: extractedData.earlyWithdrawalPenalty as number,
        interestOnUSSavingsBonds: extractedData.interestOnUSSavingsBonds as number,
        federalIncomeTaxWithheld: extractedData.federalIncomeTaxWithheld as number,
        investmentExpenses: extractedData.investmentExpenses as number,
        foreignTaxPaid: extractedData.foreignTaxPaid as number,
        foreignCountry: extractedData.foreignCountry as string,
        taxExemptInterest: extractedData.taxExemptInterest as number,
        specifiedPrivateActivityBondInterest: extractedData.specifiedPrivateActivityBondInterest as number,
        marketDiscount: extractedData.marketDiscount as number,
        bondPremium: extractedData.bondPremium as number,
        bondPremiumOnTreasury: extractedData.bondPremiumOnTreasury as number,
        bondPremiumOnTaxExempt: extractedData.bondPremiumOnTaxExempt as number,
        cusipNumber: extractedData.cusipNumber as string,
        state: extractedData.state as string,
        stateIdentificationNumber: extractedData.stateIdentificationNumber as string,
        stateTaxWithheld: extractedData.stateTaxWithheld as number,
        
        // Additional processing fields
        fullText: extractedData.fullText as string,
        correctedDocumentType: extractedData.correctedDocumentType
      };
    } catch (error: unknown) {
      console.error('❌ [Azure DI] 1099-INT extraction failed:', error);
      throw new Error(`1099-INT extraction failed: ${getErrorMessage(error)}`);
    }
  }

  /**
   * ENHANCED: Extract data from 1099-MISC tax document with ALL 18 fields
   * @param documentPathOrBuffer - Path to document file or Buffer containing document data
   * @returns Promise<Tax1099MiscResult> - Extracted 1099-MISC data with comprehensive field coverage
   */
  async extract1099Misc(documentPathOrBuffer: string | Buffer): Promise<Tax1099MiscResult> {
    try {
      console.log('🔍 [Azure DI] Extracting 1099-MISC document with comprehensive field mapping...');
      const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, 'FORM_1099_MISC');
      
      return {
        // Payer Information
        payerName: extractedData.payerName as string,
        payerTIN: extractedData.payerTIN as string,
        payerAddress: extractedData.payerAddress as string,
        payerPhone: extractedData.payerPhone as string,
        
        // Recipient Information
        recipientName: extractedData.recipientName as string,
        recipientTIN: extractedData.recipientTIN as string,
        recipientAddress: extractedData.recipientAddress as string,
        accountNumber: extractedData.accountNumber as string,
        secondTINNotification: extractedData.secondTINNotification as boolean,
        
        // Box 1-18: All miscellaneous income fields
        rents: extractedData.rents as number,
        royalties: extractedData.royalties as number,
        otherIncome: extractedData.otherIncome as number,
        federalIncomeTaxWithheld: extractedData.federalIncomeTaxWithheld as number,
        fishingBoatProceeds: extractedData.fishingBoatProceeds as number,
        medicalHealthPayments: extractedData.medicalHealthPayments as number,
        payerMadeDirectSales: extractedData.payerMadeDirectSales as boolean,
        substitutePayments: extractedData.substitutePayments as number,
        cropInsuranceProceeds: extractedData.cropInsuranceProceeds as number,
        grossProceedsAttorney: extractedData.grossProceedsAttorney as number,
        fishPurchasedForResale: extractedData.fishPurchasedForResale as number,
        section409ADeferrals: extractedData.section409ADeferrals as number,
        fatcaFilingRequirement: extractedData.fatcaFilingRequirement as boolean,
        section409AIncome: extractedData.section409AIncome as number,
        stateTaxWithheld: extractedData.stateTaxWithheld as number,
        statePayerNumber: extractedData.statePayerNumber as string,
        stateIncome: extractedData.stateIncome as number,
        stateName: extractedData.stateName as string,
        
        // Additional processing fields
        fullText: extractedData.fullText as string,
        correctedDocumentType: extractedData.correctedDocumentType
      };
    } catch (error: unknown) {
      console.error('❌ [Azure DI] 1099-MISC extraction failed:', error);
      throw new Error(`1099-MISC extraction failed: ${getErrorMessage(error)}`);
    }
  }

  /**
   * ENHANCED: Extract data from 1099-DIV tax document with ALL fields (KEEP EXISTING IMPLEMENTATION)
   * @param documentPathOrBuffer - Path to document file or Buffer containing document data
   * @returns Promise<Tax1099DivResult> - Extracted 1099-DIV data with all fields
   */
  async extract1099Div(documentPathOrBuffer: string | Buffer): Promise<Tax1099DivResult> {
    try {
      console.log('🔍 [Azure DI] Extracting 1099-DIV document with comprehensive field mapping...');
      const extractedData = await this.extractDataFromDocument(documentPathOrBuffer, 'FORM_1099_DIV');
      
      return {
        // Payer Information
        payerName: extractedData.payerName as string,
        payerTIN: extractedData.payerTIN as string,
        payerAddress: extractedData.payerAddress as string,
        payerPhone: extractedData.payerPhone as string,
        
        // Recipient Information
        recipientName: extractedData.recipientName as string,
        recipientTIN: extractedData.recipientTIN as string,
        recipientAddress: extractedData.recipientAddress as string,
        accountNumber: extractedData.accountNumber as string,
        secondTINNotification: extractedData.secondTINNotification as boolean,
        
        // Box 1a & 1b - Dividend Income
        totalOrdinaryDividends: extractedData.totalOrdinaryDividends as number,
        qualifiedDividends: extractedData.qualifiedDividends as number,
        
        // Box 2a-2f - Capital Gain Distributions
        totalCapitalGainDistributions: extractedData.totalCapitalGainDistributions as number,
        unrecapturedSection1250Gain: extractedData.unrecapturedSection1250Gain as number,
        section1202Gain: extractedData.section1202Gain as number,
        collectiblesGain: extractedData.collectiblesGain as number,
        section897OrdinaryDividends: extractedData.section897OrdinaryDividends as number,
        section897CapitalGain: extractedData.section897CapitalGain as number,
        
        // Box 3-13 - Other Distributions and Information
        nondividendDistributions: extractedData.nondividendDistributions as number,
        federalIncomeTaxWithheld: extractedData.federalIncomeTaxWithheld as number,
        section199ADividends: extractedData.section199ADividends as number,
        investmentExpenses: extractedData.investmentExpenses as number,
        foreignTaxPaid: extractedData.foreignTaxPaid as number,
        foreignCountry: extractedData.foreignCountry as string,
        cashLiquidationDistributions: extractedData.cashLiquidationDistributions as number,
        noncashLiquidationDistributions: extractedData.noncashLiquidationDistributions as number,
        fatcaFilingRequirement: extractedData.fatcaFilingRequirement as boolean,
        exemptInterestDividends: extractedData.exemptInterestDividends as number,
        specifiedPrivateActivityBondInterest: extractedData.specifiedPrivateActivityBondInterest as number,
        
        // State Information (Boxes 14-16)
        state: extractedData.state as string,
        stateIdentificationNumber: extractedData.stateIdentificationNumber as string,
        stateTaxWithheld: extractedData.stateTaxWithheld as number,
        
        // Additional fields
        fullText: extractedData.fullText as string,
        correctedDocumentType: extractedData.correctedDocumentType
      };
    } catch (error: unknown) {
      console.error('❌ [Azure DI] 1099-DIV extraction failed:', error);
      throw new Error(`1099-DIV extraction failed: ${getErrorMessage(error)}`);
    }
  }

  private getModelIdForDocumentType(documentType: string): string {
    switch (documentType) {
      case 'W2':
        return 'prebuilt-tax.us.w2';
      case 'FORM_1099_INT':
      case 'FORM_1099_DIV':
      case 'FORM_1099_MISC':
      case 'FORM_1099_NEC':
        // All 1099 variants use the unified 1099 model
        return 'prebuilt-tax.us.1099';
      default:
        // Use general document model for other types
        return 'prebuilt-document';
    }
  }

  private extractTaxDocumentFieldsFromOCR(result: any, documentType: string): ExtractedFieldData {
    console.log('🔍 [Azure DI] Extracting tax document fields using OCR fallback...');
    
    const extractedData: ExtractedFieldData = {};
    
    // Extract text content from OCR result
    extractedData.fullText = result.content || '';
    
    // Use OCR-based extraction methods for different document types
    switch (documentType) {
      case 'W2':
        return this.extractW2FieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_INT':
        return this.extract1099IntFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_DIV':
        return this.extract1099DivFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_MISC':
        return this.extract1099MiscFieldsFromOCR(extractedData.fullText as string, extractedData);
      case 'FORM_1099_NEC':
        return this.extract1099NecFieldsFromOCR(extractedData.fullText as string, extractedData);
      default:
        console.log('🔍 [Azure DI] Using generic OCR extraction for document type:', documentType);
        return this.extractGenericFieldsFromOCR(extractedData.fullText as string, extractedData);
    }
  }

  private extractTaxDocumentFields(result: any, documentType: string): ExtractedFieldData {
    const extractedData: ExtractedFieldData = {};
    
    // Extract text content
    extractedData.fullText = result.content || '';
    
    // Extract form fields
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      if (document.fields) {
        // Process fields based on document type
        switch (documentType) {
          case 'W2':
            return this.processW2Fields(document.fields, extractedData);
          case 'FORM_1099_INT':
            return this.process1099IntFields(document.fields, extractedData);
          case 'FORM_1099_DIV':
            return this.process1099DivFields(document.fields, extractedData);
          case 'FORM_1099_MISC':
            return this.process1099MiscFields(document.fields, extractedData);
          case 'FORM_1099_NEC':
            return this.process1099NecFields(document.fields, extractedData);
          default:
            return this.processGenericFields(document.fields, extractedData);
        }
      }
    }
    
    // Extract key-value pairs from tables if available
    if (result.keyValuePairs) {
      for (const kvp of result.keyValuePairs) {
        const key = kvp.key?.content?.trim();
        const value = kvp.value?.content?.trim();
        if (key && value) {
          extractedData[key] = value;
        }
      }
    }
    
    return extractedData;
  }

  // ENHANCED: Complete W2 field processing with ALL 20+ fields
  private processW2Fields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const w2Data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing W2 fields with comprehensive mapping...');
    
    // COMPREHENSIVE: All W2 field mappings based on IRS form structure
    const w2FieldMappings = {
      // Employee Information (Boxes a, e, f)
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN',
      'Employee.Address': 'employeeAddress',
      'EmployeeName': 'employeeName',
      'EmployeeSSN': 'employeeSSN',
      'EmployeeAddress': 'employeeAddress',
      'RecipientName': 'employeeName',
      'RecipientTIN': 'employeeSSN',
      'RecipientAddress': 'employeeAddress',
      
      // Employer Information (Boxes b, c, d)
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'EmployerName': 'employerName',
      'EmployerEIN': 'employerEIN',
      'EmployerAddress': 'employerAddress',
      'PayerName': 'employerName',
      'PayerTIN': 'employerEIN',
      'PayerAddress': 'employerAddress',
      'ControlNumber': 'controlNumber',
      
      // Box 1-20: All wage and tax information
      'WagesAndTips': 'wages',                              // Box 1
      'Wages': 'wages',                                     // Box 1 alternative
      'Box1': 'wages',
      'FederalIncomeTaxWithheld': 'federalIncomeTaxWithheld', // Box 2
      'FederalTaxWithheld': 'federalIncomeTaxWithheld',     // Box 2 alternative
      'Box2': 'federalIncomeTaxWithheld',
      'SocialSecurityWages': 'socialSecurityWages',         // Box 3
      'SSWages': 'socialSecurityWages',                     // Box 3 alternative
      'Box3': 'socialSecurityWages',
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld', // Box 4
      'SSTaxWithheld': 'socialSecurityTaxWithheld',         // Box 4 alternative
      'Box4': 'socialSecurityTaxWithheld',
      'MedicareWagesAndTips': 'medicareWages',              // Box 5
      'MedicareWages': 'medicareWages',                     // Box 5 alternative
      'Box5': 'medicareWages',
      'MedicareTaxWithheld': 'medicareTaxWithheld',         // Box 6
      'Box6': 'medicareTaxWithheld',
      'SocialSecurityTips': 'socialSecurityTips',           // Box 7
      'SSTips': 'socialSecurityTips',                       // Box 7 alternative
      'Box7': 'socialSecurityTips',
      'AllocatedTips': 'allocatedTips',                     // Box 8
      'Box8': 'allocatedTips',
      'AdvanceEIC': 'advanceEIC',                           // Box 9
      'AdvanceEICPayments': 'advanceEIC',                   // Box 9 alternative
      'Box9': 'advanceEIC',
      'DependentCareBenefits': 'dependentCareBenefits',     // Box 10
      'DepCareBenefits': 'dependentCareBenefits',           // Box 10 alternative
      'Box10': 'dependentCareBenefits',
      'NonqualifiedPlans': 'nonqualifiedPlans',             // Box 11
      'NonqualPlans': 'nonqualifiedPlans',                  // Box 11 alternative
      'Box11': 'nonqualifiedPlans',
      'DeferredCompensation': 'box12Raw',                   // Box 12 - will be parsed
      'Box12': 'box12Raw',
      'OtherInfo': 'otherInfo',                             // Box 14
      'Other': 'otherInfo',                                 // Box 14 alternative
      'Box14': 'otherInfo',
      
      // State and Local Information (Boxes 15-20)
      'StateEmployerID': 'stateEmployerID',                 // Box 15
      'StateID': 'stateEmployerID',                         // Box 15 alternative
      'Box15': 'stateEmployerID',
      'StateWagesTipsEtc': 'stateWages',                    // Box 16
      'StateWages': 'stateWages',                           // Box 16 alternative
      'Box16': 'stateWages',
      'StateIncomeTax': 'stateIncomeTax',                   // Box 17
      'StateTaxWithheld': 'stateIncomeTax',                 // Box 17 alternative
      'Box17': 'stateIncomeTax',
      'LocalWagesTipsEtc': 'localWages',                    // Box 18
      'LocalWages': 'localWages',                           // Box 18 alternative
      'Box18': 'localWages',
      'LocalIncomeTax': 'localIncomeTax',                   // Box 19
      'LocalTaxWithheld': 'localIncomeTax',                 // Box 19 alternative
      'Box19': 'localIncomeTax',
      'LocalityName': 'localityName',                       // Box 20
      'Locality': 'localityName',                           // Box 20 alternative
      'Box20': 'localityName',
      
      // State information
      'State': 'stateName',
      'StateName': 'stateName',
      
      // Retirement plan checkbox (Box 13)
      'RetirementPlan': 'retirementPlan',
      'ThirdPartySickPay': 'thirdPartySickPay',
      'StatutoryEmployee': 'statutoryEmployee'
    };
    
    // Process all field mappings
    for (const [azureFieldName, mappedFieldName] of Object.entries(w2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        // Handle different field types appropriately
        if (mappedFieldName === 'box12Raw' || mappedFieldName === 'otherInfo' || 
            mappedFieldName === 'stateEmployerID' || mappedFieldName === 'localityName' ||
            mappedFieldName === 'stateName' || mappedFieldName === 'controlNumber') {
          // Text fields
          w2Data[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'retirementPlan' || mappedFieldName === 'thirdPartySickPay' || 
                   mappedFieldName === 'statutoryEmployee') {
          // Boolean fields for Box 13 checkboxes
          w2Data[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields
          w2Data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // ENHANCED: Parse Box 12 codes from raw text
    if (w2Data.box12Raw) {
      const box12Codes = this.parseW2Box12Codes(w2Data.box12Raw as string);
      if (box12Codes.length > 0) {
        w2Data.box12Codes = box12Codes;
        console.log('✅ [Azure DI] Parsed Box 12 codes:', box12Codes);
      }
    }
    
    // ENHANCED: Extract Box 13 checkboxes
    w2Data.box13Checkboxes = {
      statutoryEmployee: w2Data.statutoryEmployee as boolean || false,
      retirementPlan: w2Data.retirementPlan as boolean || false,
      thirdPartySickPay: w2Data.thirdPartySickPay as boolean || false
    };
    
    // Clean up individual checkbox fields
    delete w2Data.statutoryEmployee;
    delete w2Data.retirementPlan;
    delete w2Data.thirdPartySickPay;
    
    // Enhanced personal info extraction with better fallback handling
    if ((!w2Data.employeeName || !w2Data.employeeSSN || !w2Data.employeeAddress || 
         !w2Data.employerName || !w2Data.employerAddress) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some W2 info missing from structured fields, attempting OCR extraction...');
      
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string, w2Data.employeeName as string);
      
      if (!w2Data.employeeName && personalInfoFromOCR.name) {
        w2Data.employeeName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted employee name from OCR:', w2Data.employeeName);
      }
      
      if (!w2Data.employeeSSN && personalInfoFromOCR.ssn) {
        w2Data.employeeSSN = personalInfoFromOCR.ssn;
        console.log('✅ [Azure DI] Extracted employee SSN from OCR:', w2Data.employeeSSN);
      }
      
      if (!w2Data.employeeAddress && personalInfoFromOCR.address) {
        w2Data.employeeAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted employee address from OCR:', w2Data.employeeAddress);
      }
      
      if (!w2Data.employerName && personalInfoFromOCR.employerName) {
        w2Data.employerName = personalInfoFromOCR.employerName;
        console.log('✅ [Azure DI] Extracted employer name from OCR:', w2Data.employerName);
      }
      
      if (!w2Data.employerAddress && personalInfoFromOCR.employerAddress) {
        w2Data.employerAddress = personalInfoFromOCR.employerAddress;
        console.log('✅ [Azure DI] Extracted employer address from OCR:', w2Data.employerAddress);
      }
    }

    // Enhanced address parsing - extract city, state, and zipCode from full address
    if (w2Data.employeeAddress && typeof w2Data.employeeAddress === 'string') {
      console.log('🔍 [Azure DI] Parsing address components from:', w2Data.employeeAddress);
      const ocrText = typeof baseData.fullText === 'string' ? baseData.fullText : '';
      const addressParts = this.extractAddressParts(w2Data.employeeAddress, ocrText);
      
      // Add parsed address components to W2 data
      w2Data.employeeAddressStreet = addressParts.street;
      w2Data.employeeCity = addressParts.city;
      w2Data.employeeState = addressParts.state;
      w2Data.employeeZipCode = addressParts.zipCode;
      
      console.log('✅ [Azure DI] Parsed address components:', {
        street: w2Data.employeeAddressStreet,
        city: w2Data.employeeCity,
        state: w2Data.employeeState,
        zipCode: w2Data.employeeZipCode
      });
    }
    
    // Enhanced OCR fallback for missing W2 fields
    if (baseData.fullText) {
      this.enhanceW2WithOCRFallback(w2Data, baseData.fullText as string);
    }
    
    return w2Data;
  }

  // ENHANCED: Complete 1099-INT field processing with ALL 15+ fields
  private process1099IntFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-INT fields with comprehensive mapping...');
    
    // COMPREHENSIVE: All 1099-INT field mappings based on IRS form structure
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Payer.Phone': 'payerPhone',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerPhone': 'payerPhone',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      'SecondTINNotification': 'secondTINNotification',
      '2ndTINNot': 'secondTINNotification',
      
      // Box 1-15 mappings (complete 1099-INT form)
      'InterestIncome': 'interestIncome',                                    // Box 1
      'Interest': 'interestIncome',                                          // Box 1 alternative
      'Box1': 'interestIncome',
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',                   // Box 2
      'EarlyWithdrawal': 'earlyWithdrawalPenalty',                          // Box 2 alternative
      'Box2': 'earlyWithdrawalPenalty',
      'InterestOnUSTreasuryObligations': 'interestOnUSSavingsBonds',        // Box 3
      'InterestOnUSavingsBonds': 'interestOnUSSavingsBonds',                // Box 3 alternative
      'USavingsBondsInterest': 'interestOnUSSavingsBonds',                  // Box 3 alternative
      'TreasuryInterest': 'interestOnUSSavingsBonds',                       // Box 3 alternative
      'Box3': 'interestOnUSSavingsBonds',
      'FederalIncomeTaxWithheld': 'federalIncomeTaxWithheld',               // Box 4
      'FederalTaxWithheld': 'federalIncomeTaxWithheld',                     // Box 4 alternative
      'Box4': 'federalIncomeTaxWithheld',
      'InvestmentExpenses': 'investmentExpenses',                           // Box 5
      'Investment': 'investmentExpenses',                                   // Box 5 alternative
      'Box5': 'investmentExpenses',
      'ForeignTaxPaid': 'foreignTaxPaid',                                   // Box 6
      'ForeignTax': 'foreignTaxPaid',                                       // Box 6 alternative
      'Box6': 'foreignTaxPaid',
      'ForeignCountry': 'foreignCountry',                                   // Box 7
      'ForeignCountryOrUSPossession': 'foreignCountry',                     // Box 7 alternative
      'Box7': 'foreignCountry',
      'TaxExemptInterest': 'taxExemptInterest',                            // Box 8
      'TaxExempt': 'taxExemptInterest',                                     // Box 8 alternative
      'Box8': 'taxExemptInterest',
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest', // Box 9
      'PrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest', // Box 9 alternative
      'PABInterest': 'specifiedPrivateActivityBondInterest',                // Box 9 alternative
      'Box9': 'specifiedPrivateActivityBondInterest',
      'MarketDiscount': 'marketDiscount',                                   // Box 10
      'Market': 'marketDiscount',                                           // Box 10 alternative
      'Box10': 'marketDiscount',
      'BondPremium': 'bondPremium',                                         // Box 11
      'Premium': 'bondPremium',                                             // Box 11 alternative
      'Box11': 'bondPremium',
      'BondPremiumOnTreasury': 'bondPremiumOnTreasury',                     // Box 12
      'TreasuryPremium': 'bondPremiumOnTreasury',                           // Box 12 alternative
      'Box12': 'bondPremiumOnTreasury',
      'BondPremiumOnTaxExempt': 'bondPremiumOnTaxExempt',                   // Box 13
      'TaxExemptPremium': 'bondPremiumOnTaxExempt',                         // Box 13 alternative
      'Box13': 'bondPremiumOnTaxExempt',
      'CUSIPNumber': 'cusipNumber',                                         // Box 14
      'CUSIP': 'cusipNumber',                                               // Box 14 alternative
      'TaxExemptAndTaxCreditBondCUSIP': 'cusipNumber',                      // Box 14 alternative
      'Box14': 'cusipNumber',
      'State': 'state',                                                     // Box 15
      'StateCode': 'state',                                                 // Box 15 alternative
      'Box15': 'state',
      'StateIdentificationNumber': 'stateIdentificationNumber',             // Box 16
      'StateID': 'stateIdentificationNumber',                               // Box 16 alternative
      'StatePayerNumber': 'stateIdentificationNumber',                      // Box 16 alternative
      'Box16': 'stateIdentificationNumber',
      'StateTaxWithheld': 'stateTaxWithheld',                              // Box 17
      'StateWithholding': 'stateTaxWithheld',                              // Box 17 alternative
      'Box17': 'stateTaxWithheld'
    };
    
    // Process all field mappings
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        // Handle text fields vs numeric fields appropriately
        if (mappedFieldName === 'foreignCountry' || 
            mappedFieldName === 'stateIdentificationNumber' || 
            mappedFieldName === 'accountNumber' ||
            mappedFieldName === 'cusipNumber' ||
            mappedFieldName === 'state' ||
            mappedFieldName === 'payerPhone') {
          // Text fields - store as string
          data[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'secondTINNotification') {
          // Boolean field
          data[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields - parse as amount
          data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || 
         !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-INT info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    // Enhanced OCR fallback for missing 1099-INT fields
    if (baseData.fullText) {
      this.enhance1099IntWithOCRFallback(data, baseData.fullText as string);
    }
    
    return data;
  }

  // ENHANCED: Complete 1099-MISC field processing with ALL 18 fields
  private process1099MiscFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-MISC fields with comprehensive mapping...');
    
    // COMPREHENSIVE: All 1099-MISC field mappings based on IRS form structure
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Payer.Phone': 'payerPhone',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerPhone': 'payerPhone',
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      'SecondTINNotification': 'secondTINNotification',
      '2ndTINNot': 'secondTINNotification',
      
      // Box 1-18 mappings (complete 1099-MISC form)
      'Rents': 'rents',                                                     // Box 1
      'Rent': 'rents',                                                      // Box 1 alternative
      'Box1': 'rents',
      'Royalties': 'royalties',                                             // Box 2
      'Royalty': 'royalties',                                               // Box 2 alternative
      'Box2': 'royalties',
      'OtherIncome': 'otherIncome',                                         // Box 3
      'Other': 'otherIncome',                                               // Box 3 alternative
      'Box3': 'otherIncome',
      'FederalIncomeTaxWithheld': 'federalIncomeTaxWithheld',               // Box 4
      'FederalTaxWithheld': 'federalIncomeTaxWithheld',                     // Box 4 alternative
      'Box4': 'federalIncomeTaxWithheld',
      'FishingBoatProceeds': 'fishingBoatProceeds',                         // Box 5
      'FishingBoat': 'fishingBoatProceeds',                                 // Box 5 alternative
      'Box5': 'fishingBoatProceeds',
      'MedicalAndHealthCarePayments': 'medicalHealthPayments',              // Box 6
      'MedicalHealthPayments': 'medicalHealthPayments',                     // Box 6 alternative
      'Medical': 'medicalHealthPayments',                                   // Box 6 alternative
      'Box6': 'medicalHealthPayments',
      'PayerMadeDirectSales': 'payerMadeDirectSales',                       // Box 7 (checkbox)
      'DirectSales': 'payerMadeDirectSales',                                // Box 7 alternative
      'Box7': 'payerMadeDirectSales',
      'SubstitutePayments': 'substitutePayments',                           // Box 8
      'Substitute': 'substitutePayments',                                   // Box 8 alternative
      'SubstitutePaymentsInLieuOfDividends': 'substitutePayments',          // Box 8 alternative
      'Box8': 'substitutePayments',
      'CropInsuranceProceeds': 'cropInsuranceProceeds',                     // Box 9
      'CropInsurance': 'cropInsuranceProceeds',                             // Box 9 alternative
      'Box9': 'cropInsuranceProceeds',
      'GrossProceedsPaidToAttorney': 'grossProceedsAttorney',               // Box 10
      'GrossProceeds': 'grossProceedsAttorney',                             // Box 10 alternative
      'AttorneyProceeds': 'grossProceedsAttorney',                          // Box 10 alternative
      'Box10': 'grossProceedsAttorney',
      'FishPurchasedForResale': 'fishPurchasedForResale',                   // Box 11
      'Fish': 'fishPurchasedForResale',                                     // Box 11 alternative
      'Box11': 'fishPurchasedForResale',
      'Section409ADeferrals': 'section409ADeferrals',                       // Box 12
      '409ADeferrals': 'section409ADeferrals',                              // Box 12 alternative
      'Box12': 'section409ADeferrals',
      'FATCAFilingRequirement': 'fatcaFilingRequirement',                   // Box 13 (checkbox)
      'FATCA': 'fatcaFilingRequirement',                                    // Box 13 alternative
      'Box13': 'fatcaFilingRequirement',
      // Box 14: Reserved (not used)
      'NonqualifiedDeferredCompensation': 'section409AIncome',              // Box 15
      'Section409AIncome': 'section409AIncome',                            // Box 15 alternative
      '409AIncome': 'section409AIncome',                                    // Box 15 alternative
      'Box15': 'section409AIncome',
      'StateTaxWithheld': 'stateTaxWithheld',                              // Box 16
      'StateWithholding': 'stateTaxWithheld',                              // Box 16 alternative
      'Box16': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',                              // Box 17
      'StateNumber': 'statePayerNumber',                                    // Box 17 alternative
      'Box17': 'statePayerNumber',
      'StateIncome': 'stateIncome',                                         // Box 18
      'Box18': 'stateIncome',
      
      // State information
      'State': 'stateName',
      'StateName': 'stateName'
    };
    
    // Process all field mappings
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        // Handle different field types appropriately
        if (mappedFieldName === 'statePayerNumber' || 
            mappedFieldName === 'accountNumber' ||
            mappedFieldName === 'stateName' ||
            mappedFieldName === 'payerPhone') {
          // Text fields - store as string
          data[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'secondTINNotification' ||
                   mappedFieldName === 'payerMadeDirectSales' ||
                   mappedFieldName === 'fatcaFilingRequirement') {
          // Boolean fields (checkboxes)
          data[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields - parse as amount
          data[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || 
         !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-MISC info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    // Enhanced OCR fallback for missing 1099-MISC fields
    if (baseData.fullText) {
      this.enhance1099MiscWithOCRFallback(data, baseData.fullText as string);
    }
    
    return data;
  }

  // ENHANCED: Complete 1099-DIV field processing with ALL 16 boxes (KEEP EXISTING)
  private process1099DivFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    const data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-DIV fields with comprehensive mapping...');
    
    // COMPREHENSIVE: All 1099-DIV field mappings based on IRS form structure
    const azureToLocalFieldMap = {
      // Payer Information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'Payer.Phone': 'payerPhone',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerPhone': 'payerPhone',
      
      // Recipient Information
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      'SecondTINNotification': 'secondTINNotification',
      '2ndTINNot': 'secondTINNotification',
      
      // Box 1a & 1b - Dividend Income
      'TotalOrdinaryDividends': 'totalOrdinaryDividends',        // Box 1a
      'OrdinaryDividends': 'totalOrdinaryDividends',             // Box 1a alternative
      'QualifiedDividends': 'qualifiedDividends',                // Box 1b
      'Box1a': 'totalOrdinaryDividends',
      'Box1b': 'qualifiedDividends',
      
      // Box 2a-2f - Capital Gain Distributions (CRITICAL - these were missing!)
      'TotalCapitalGainDistributions': 'totalCapitalGainDistributions', // Box 2a
      'CapitalGainDistributions': 'totalCapitalGainDistributions',      // Box 2a alternative
      'UnrecapturedSection1250Gain': 'unrecapturedSection1250Gain',     // Box 2b - CRITICAL
      'Unrecap1250Gain': 'unrecapturedSection1250Gain',                 // Box 2b alternative
      'Section1202Gain': 'section1202Gain',                             // Box 2c - CRITICAL
      'CollectiblesGain': 'collectiblesGain',                           // Box 2d - CRITICAL
      'Collectibles28Gain': 'collectiblesGain',                         // Box 2d alternative
      'Section897OrdinaryDividends': 'section897OrdinaryDividends',     // Box 2e - CRITICAL
      'Section897CapitalGain': 'section897CapitalGain',                 // Box 2f - CRITICAL
      'Box2a': 'totalCapitalGainDistributions',
      'Box2b': 'unrecapturedSection1250Gain',
      'Box2c': 'section1202Gain',
      'Box2d': 'collectiblesGain',
      'Box2e': 'section897OrdinaryDividends',
      'Box2f': 'section897CapitalGain',
      
      // Box 3-13 - Other Distributions and Information
      'NondividendDistributions': 'nondividendDistributions',           // Box 3
      'ReturnOfCapital': 'nondividendDistributions',                    // Box 3 alternative
      'FederalIncomeTaxWithheld': 'federalIncomeTaxWithheld',          // Box 4
      'FederalTaxWithheld': 'federalIncomeTaxWithheld',                // Box 4 alternative
      'Section199ADividends': 'section199ADividends',                   // Box 5
      'QBIDividends': 'section199ADividends',                          // Box 5 alternative
      'InvestmentExpenses': 'investmentExpenses',                       // Box 6
      'ForeignTaxPaid': 'foreignTaxPaid',                              // Box 7
      'ForeignCountry': 'foreignCountry',                              // Box 8
      'ForeignCountryOrUSPossession': 'foreignCountry',                // Box 8 alternative
      'CashLiquidationDistributions': 'cashLiquidationDistributions',   // Box 9
      'CashLiquidation': 'cashLiquidationDistributions',               // Box 9 alternative
      'NoncashLiquidationDistributions': 'noncashLiquidationDistributions', // Box 10
      'NoncashLiquidation': 'noncashLiquidationDistributions',         // Box 10 alternative
      'FATCAFilingRequirement': 'fatcaFilingRequirement',              // Box 11
      'FATCA': 'fatcaFilingRequirement',                               // Box 11 alternative
      'ExemptInterestDividends': 'exemptInterestDividends',            // Box 12
      'TaxExemptInterest': 'exemptInterestDividends',                  // Box 12 alternative
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest', // Box 13
      'PABInterest': 'specifiedPrivateActivityBondInterest',           // Box 13 alternative
      'Box3': 'nondividendDistributions',
      'Box4': 'federalIncomeTaxWithheld',
      'Box5': 'section199ADividends',
      'Box6': 'investmentExpenses',
      'Box7': 'foreignTaxPaid',
      'Box8': 'foreignCountry',
      'Box9': 'cashLiquidationDistributions',
      'Box10': 'noncashLiquidationDistributions',
      'Box11': 'fatcaFilingRequirement',
      'Box12': 'exemptInterestDividends',
      'Box13': 'specifiedPrivateActivityBondInterest',
      
      // State Information (Boxes 14-16)
      'State': 'state',                                                // Box 14
      'StateCode': 'state',                                            // Box 14 alternative
      'StateIdentificationNumber': 'stateIdentificationNumber',        // Box 15
      'StateID': 'stateIdentificationNumber',                          // Box 15 alternative
      'StatePayerNumber': 'stateIdentificationNumber',                 // Box 15 alternative
      'StateTaxWithheld': 'stateTaxWithheld',                         // Box 16
      'StateWithholding': 'stateTaxWithheld',                         // Box 16 alternative
      'Box14': 'state',
      'Box15': 'stateIdentificationNumber',
      'Box16': 'stateTaxWithheld'
    };
    
    // Process all field mappings with enhanced logic
    for (const [azureFieldName, mappedFieldName] of Object.entries(azureToLocalFieldMap)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        console.log(`🔍 [Azure DI] Processing field: ${azureFieldName} -> ${mappedFieldName} = ${value}`);
        
        // Handle different field types appropriately
        if (mappedFieldName === 'foreignCountry' || 
            mappedFieldName === 'state' || 
            mappedFieldName === 'stateIdentificationNumber' ||
            mappedFieldName === 'accountNumber' ||
            mappedFieldName === 'payerPhone') {
          // Text fields - store as string
          data[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'secondTINNotification' || 
                   mappedFieldName === 'fatcaFilingRequirement') {
          // Boolean fields
          data[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields - parse as amount
          const numericValue = typeof value === 'number' ? value : this.parseAmount(value);
          if (numericValue > 0) {
            data[mappedFieldName] = numericValue;
            console.log(`✅ [Azure DI] Successfully mapped ${mappedFieldName}: $${numericValue}`);
          }
        }
      }
    }
    
    // OCR fallback for personal info if not found in structured fields
    if ((!data.recipientName || !data.recipientTIN || !data.recipientAddress || 
         !data.payerName || !data.payerTIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some 1099-DIV info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText as string);
      
      if (!data.recipientName && personalInfoFromOCR.name) {
        data.recipientName = personalInfoFromOCR.name;
        console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
      }
      
      if (!data.recipientTIN && personalInfoFromOCR.tin) {
        data.recipientTIN = personalInfoFromOCR.tin;
        console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
      }
      
      if (!data.recipientAddress && personalInfoFromOCR.address) {
        data.recipientAddress = personalInfoFromOCR.address;
        console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
      }
      
      if (!data.payerName && personalInfoFromOCR.payerName) {
        data.payerName = personalInfoFromOCR.payerName;
        console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
      }
      
      if (!data.payerTIN && personalInfoFromOCR.payerTIN) {
        data.payerTIN = personalInfoFromOCR.payerTIN;
        console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
      }
    }
    
    // Enhanced OCR fallback for missing 1099-DIV fields
    if (baseData.fullText) {
      this.enhance1099DivWithOCRFallback(data, baseData.fullText as string);
    }
    
    return data;
  }

  // === ENHANCED OCR FALLBACK METHODS ===

  // ENHANCED: Parse W2 Box 12 codes
  private parseW2Box12Codes(box12String: string): Array<{ code: string; amount: number }> {
    if (!box12String) return [];
    
    const codes: Array<{ code: string; amount: number }> = [];
    
    // Parse format like "D 5000.00 W 2500.00" or "D$5000 W$2500"
    const codePattern = /([A-Z]{1,2})\s*\$?(\d+(?:\.\d{2})?)/g;
    let match;
    
    while ((match = codePattern.exec(box12String)) !== null) {
      const code = match[1];
      const amount = parseFloat(match[2]);
      
      if (!isNaN(amount)) {
        codes.push({ code, amount });
      }
    }
    
    return codes;
  }

  // ENHANCED: Parse boolean values
  private parseBoolean(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const lower = value.toLowerCase().trim();
      return lower === 'true' || lower === 'yes' || lower === 'x' || lower === '✓' || lower === 'checked';
    }
    return false;
  }

  // ENHANCED: OCR fallback for missing W2 fields
  private enhanceW2WithOCRFallback(w2Data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Azure DI] Applying OCR fallback for missing W2 fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'advanceEIC', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /advance\s+eic[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'dependentCareBenefits', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /dependent\s+care\s+benefits[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'nonqualifiedPlans', patterns: [/box\s*11[:\s]*\$?(\d+(?:\.\d{2})?)/i, /nonqualified\s+plans[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'localWages', patterns: [/box\s*18[:\s]*\$?(\d+(?:\.\d{2})?)/i, /local\s+wages[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'localIncomeTax', patterns: [/box\s*19[:\s]*\$?(\d+(?:\.\d{2})?)/i, /local\s+income\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!w2Data[field] || w2Data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              w2Data[field] = amount;
              console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
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
        const box12Codes = this.parseW2Box12Codes(match[1]);
        if (box12Codes.length > 0) {
          w2Data.box12Codes = box12Codes;
          console.log('✅ [Azure DI] Extracted Box 12 codes from OCR:', box12Codes);
        }
      }
    }
    
    // Extract text fields
    if (!w2Data.stateEmployerID) {
      const stateIdPattern = /box\s*15[:\s]*([A-Z0-9\-]+)/i;
      const match = ocrText.match(stateIdPattern);
      if (match && match[1]) {
        w2Data.stateEmployerID = match[1].trim();
        console.log(`✅ [Azure DI] Extracted state employer ID from OCR: ${w2Data.stateEmployerID}`);
      }
    }
    
    if (!w2Data.localityName) {
      const localityPattern = /box\s*20[:\s]*([A-Za-z\s]+)/i;
      const match = ocrText.match(localityPattern);
      if (match && match[1]) {
        w2Data.localityName = match[1].trim();
        console.log(`✅ [Azure DI] Extracted locality name from OCR: ${w2Data.localityName}`);
      }
    }
  }

  // ENHANCED: OCR fallback for missing 1099-INT fields
  private enhance1099IntWithOCRFallback(data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Azure DI] Applying OCR fallback for missing 1099-INT fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'interestIncome', patterns: [/box\s*1[:\s]*\$?(\d+(?:\.\d{2})?)/i, /interest\s+income[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'earlyWithdrawalPenalty', patterns: [/box\s*2[:\s]*\$?(\d+(?:\.\d{2})?)/i, /early\s+withdrawal[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'interestOnUSSavingsBonds', patterns: [/box\s*3[:\s]*\$?(\d+(?:\.\d{2})?)/i, /u\.?s\.?\s+savings\s+bonds[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'federalIncomeTaxWithheld', patterns: [/box\s*4[:\s]*\$?(\d+(?:\.\d{2})?)/i, /federal\s+income\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'investmentExpenses', patterns: [/box\s*5[:\s]*\$?(\d+(?:\.\d{2})?)/i, /investment\s+expenses[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'foreignTaxPaid', patterns: [/box\s*6[:\s]*\$?(\d+(?:\.\d{2})?)/i, /foreign\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'taxExemptInterest', patterns: [/box\s*8[:\s]*\$?(\d+(?:\.\d{2})?)/i, /tax.exempt\s+interest[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'specifiedPrivateActivityBondInterest', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /private\s+activity\s+bond[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'marketDiscount', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /market\s+discount[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'bondPremium', patterns: [/box\s*11[:\s]*\$?(\d+(?:\.\d{2})?)/i, /bond\s+premium[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'bondPremiumOnTreasury', patterns: [/box\s*12[:\s]*\$?(\d+(?:\.\d{2})?)/i, /treasury.*premium[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'bondPremiumOnTaxExempt', patterns: [/box\s*13[:\s]*\$?(\d+(?:\.\d{2})?)/i, /tax.exempt.*premium[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'stateTaxWithheld', patterns: [/box\s*17[:\s]*\$?(\d+(?:\.\d{2})?)/i, /state\s+tax\s+withheld[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!data[field] || data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              data[field] = amount;
              console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
              break;
            }
          }
        }
      }
    }
    
    // Extract text fields
    if (!data.foreignCountry) {
      const foreignCountryPattern = /box\s*7[:\s]*([A-Za-z\s]+)/i;
      const match = ocrText.match(foreignCountryPattern);
      if (match && match[1]) {
        data.foreignCountry = match[1].trim();
        console.log(`✅ [Azure DI] Extracted foreign country from OCR: ${data.foreignCountry}`);
      }
    }
    
    if (!data.cusipNumber) {
      const cusipPattern = /box\s*14[:\s]*([A-Z0-9]+)/i;
      const match = ocrText.match(cusipPattern);
      if (match && match[1]) {
        data.cusipNumber = match[1].trim();
        console.log(`✅ [Azure DI] Extracted CUSIP number from OCR: ${data.cusipNumber}`);
      }
    }
    
    if (!data.state) {
      const statePattern = /box\s*15[:\s]*([A-Z]{2})/i;
      const match = ocrText.match(statePattern);
      if (match && match[1]) {
        data.state = match[1].trim();
        console.log(`✅ [Azure DI] Extracted state from OCR: ${data.state}`);
      }
    }
  }

  // ENHANCED: OCR fallback for missing 1099-MISC fields
  private enhance1099MiscWithOCRFallback(data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Azure DI] Applying OCR fallback for missing 1099-MISC fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'rents', patterns: [/box\s*1[:\s]*\$?(\d+(?:\.\d{2})?)/i, /rents[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'royalties', patterns: [/box\s*2[:\s]*\$?(\d+(?:\.\d{2})?)/i, /royalties[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'otherIncome', patterns: [/box\s*3[:\s]*\$?(\d+(?:\.\d{2})?)/i, /other\s+income[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'federalIncomeTaxWithheld', patterns: [/box\s*4[:\s]*\$?(\d+(?:\.\d{2})?)/i, /federal\s+income\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'fishingBoatProceeds', patterns: [/box\s*5[:\s]*\$?(\d+(?:\.\d{2})?)/i, /fishing\s+boat[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'medicalHealthPayments', patterns: [/box\s*6[:\s]*\$?(\d+(?:\.\d{2})?)/i, /medical.*health[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'substitutePayments', patterns: [/box\s*8[:\s]*\$?(\d+(?:\.\d{2})?)/i, /substitute\s+payments[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'cropInsuranceProceeds', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /crop\s+insurance[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'grossProceedsAttorney', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /attorney.*proceeds[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'fishPurchasedForResale', patterns: [/box\s*11[:\s]*\$?(\d+(?:\.\d{2})?)/i, /fish.*resale[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section409ADeferrals', patterns: [/box\s*12[:\s]*\$?(\d+(?:\.\d{2})?)/i, /409a.*deferrals[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section409AIncome', patterns: [/box\s*15[:\s]*\$?(\d+(?:\.\d{2})?)/i, /409a.*income[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'stateTaxWithheld', patterns: [/box\s*16[:\s]*\$?(\d+(?:\.\d{2})?)/i, /state\s+tax\s+withheld[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'stateIncome', patterns: [/box\s*18[:\s]*\$?(\d+(?:\.\d{2})?)/i, /state\s+income[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!data[field] || data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              data[field] = amount;
              console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
              break;
            }
          }
        }
      }
    }
    
    // Extract checkbox fields
    if (!data.payerMadeDirectSales) {
      const directSalesPattern = /box\s*7.*(?:x|✓|checked|yes)/i;
      if (directSalesPattern.test(ocrText)) {
        data.payerMadeDirectSales = true;
        console.log('✅ [Azure DI] Extracted direct sales checkbox from OCR: true');
      }
    }
    
    if (!data.fatcaFilingRequirement) {
      const fatcaPattern = /box\s*13.*(?:x|✓|checked|yes)/i;
      if (fatcaPattern.test(ocrText)) {
        data.fatcaFilingRequirement = true;
        console.log('✅ [Azure DI] Extracted FATCA checkbox from OCR: true');
      }
    }
    
    // Extract text fields
    if (!data.statePayerNumber) {
      const statePayerPattern = /box\s*17[:\s]*([A-Z0-9\-]+)/i;
      const match = ocrText.match(statePayerPattern);
      if (match && match[1]) {
        data.statePayerNumber = match[1].trim();
        console.log(`✅ [Azure DI] Extracted state payer number from OCR: ${data.statePayerNumber}`);
      }
    }
  }

  // ENHANCED: OCR fallback for missing 1099-DIV fields (KEEP EXISTING)
  private enhance1099DivWithOCRFallback(data: ExtractedFieldData, ocrText: string): void {
    console.log('🔍 [Azure DI] Applying OCR fallback for missing 1099-DIV fields...');
    
    // Extract missing numeric fields using OCR patterns
    const missingFields = [
      { field: 'totalOrdinaryDividends', patterns: [/box\s*1a[:\s]*\$?(\d+(?:\.\d{2})?)/i, /ordinary\s+dividends[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'qualifiedDividends', patterns: [/box\s*1b[:\s]*\$?(\d+(?:\.\d{2})?)/i, /qualified\s+dividends[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'totalCapitalGainDistributions', patterns: [/box\s*2a[:\s]*\$?(\d+(?:\.\d{2})?)/i, /capital\s+gain\s+distributions[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'unrecapturedSection1250Gain', patterns: [/box\s*2b[:\s]*\$?(\d+(?:\.\d{2})?)/i, /unrecaptured.*1250[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section1202Gain', patterns: [/box\s*2c[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*1202[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'collectiblesGain', patterns: [/box\s*2d[:\s]*\$?(\d+(?:\.\d{2})?)/i, /collectibles[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section897OrdinaryDividends', patterns: [/box\s*2e[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*897.*ordinary[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section897CapitalGain', patterns: [/box\s*2f[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*897.*capital[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'nondividendDistributions', patterns: [/box\s*3[:\s]*\$?(\d+(?:\.\d{2})?)/i, /nondividend[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'federalIncomeTaxWithheld', patterns: [/box\s*4[:\s]*\$?(\d+(?:\.\d{2})?)/i, /federal\s+income\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'section199ADividends', patterns: [/box\s*5[:\s]*\$?(\d+(?:\.\d{2})?)/i, /section\s*199a[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'investmentExpenses', patterns: [/box\s*6[:\s]*\$?(\d+(?:\.\d{2})?)/i, /investment\s+expenses[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'foreignTaxPaid', patterns: [/box\s*7[:\s]*\$?(\d+(?:\.\d{2})?)/i, /foreign\s+tax[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'cashLiquidationDistributions', patterns: [/box\s*9[:\s]*\$?(\d+(?:\.\d{2})?)/i, /cash\s+liquidation[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'noncashLiquidationDistributions', patterns: [/box\s*10[:\s]*\$?(\d+(?:\.\d{2})?)/i, /noncash\s+liquidation[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'exemptInterestDividends', patterns: [/box\s*12[:\s]*\$?(\d+(?:\.\d{2})?)/i, /exempt.interest[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'specifiedPrivateActivityBondInterest', patterns: [/box\s*13[:\s]*\$?(\d+(?:\.\d{2})?)/i, /private\s+activity\s+bond[:\s]*\$?(\d+(?:\.\d{2})?)/i] },
      { field: 'stateTaxWithheld', patterns: [/box\s*16[:\s]*\$?(\d+(?:\.\d{2})?)/i, /state\s+tax\s+withheld[:\s]*\$?(\d+(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of missingFields) {
      if (!data[field] || data[field] === 0) {
        for (const pattern of patterns) {
          const match = ocrText.match(pattern);
          if (match && match[1]) {
            const amount = this.parseAmount(match[1]);
            if (amount > 0) {
              data[field] = amount;
              console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
              break;
            }
          }
        }
      }
    }
    
    // Extract checkbox fields
    if (!data.fatcaFilingRequirement) {
      const fatcaPattern = /box\s*11.*(?:x|✓|checked|yes)/i;
      if (fatcaPattern.test(ocrText)) {
        data.fatcaFilingRequirement = true;
        console.log('✅ [Azure DI] Extracted FATCA checkbox from OCR: true');
      }
    }
    
    // Extract text fields
    if (!data.foreignCountry) {
      const foreignCountryPattern = /box\s*8[:\s]*([A-Za-z\s]+)/i;
      const match = ocrText.match(foreignCountryPattern);
      if (match && match[1]) {
        data.foreignCountry = match[1].trim();
        console.log(`✅ [Azure DI] Extracted foreign country from OCR: ${data.foreignCountry}`);
      }
    }
    
    if (!data.state) {
      const statePattern = /box\s*14[:\s]*([A-Z]{2})/i;
      const match = ocrText.match(statePattern);
      if (match && match[1]) {
        data.state = match[1].trim();
        console.log(`✅ [Azure DI] Extracted state from OCR: ${data.state}`);
      }
    }
    
    if (!data.stateIdentificationNumber) {
      const stateIdPattern = /box\s*15[:\s]*([A-Z0-9\-]+)/i;
      const match = ocrText.match(stateIdPattern);
      if (match && match[1]) {
        data.stateIdentificationNumber = match[1].trim();
        console.log(`✅ [Azure DI] Extracted state ID from OCR: ${data.stateIdentificationNumber}`);
      }
    }
  }

  // === HELPER METHODS (KEEP EXISTING) ===

  private parseAmount(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and whitespace
      const cleanValue = value.replace(/[$,\s]/g, '');
      const parsed = parseFloat(cleanValue);
      return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
  }

  private extractPersonalInfoFromOCR(ocrText: string, targetEmployeeName?: string): any {
    const personalInfo: any = {};
    
    // Extract names, SSNs, addresses, etc. using OCR patterns
    // This is a simplified implementation - you would expand this based on your needs
    
    const ssnPattern = /\b\d{3}-?\d{2}-?\d{4}\b/g;
    const ssnMatches = ocrText.match(ssnPattern);
    if (ssnMatches && ssnMatches.length > 0) {
      personalInfo.ssn = ssnMatches[0];
      personalInfo.tin = ssnMatches[0];
    }
    
    // Extract names (this is a basic implementation)
    const namePattern = /([A-Z][a-z]+ [A-Z][a-z]+(?:\s[A-Z][a-z]+)?)/g;
    const nameMatches = ocrText.match(namePattern);
    if (nameMatches && nameMatches.length > 0) {
      personalInfo.name = nameMatches[0];
    }
    
    return personalInfo;
  }

  private extractAddressParts(fullAddress: string, ocrText: string): any {
    const addressParts: any = {};
    
    // Basic address parsing - you would expand this based on your needs
    const parts = fullAddress.split(',').map(part => part.trim());
    
    if (parts.length >= 3) {
      addressParts.street = parts[0];
      addressParts.city = parts[1];
      
      // Extract state and zip from last part
      const lastPart = parts[parts.length - 1];
      const stateZipPattern = /([A-Z]{2})\s+(\d{5}(?:-\d{4})?)/;
      const match = lastPart.match(stateZipPattern);
      if (match) {
        addressParts.state = match[1];
        addressParts.zipCode = match[2];
      }
    }
    
    return addressParts;
  }

  private extractWagesFromOCR(ocrText: string): number {
    // Extract wages from OCR text using patterns
    const wagesPattern = /(?:wages|box\s*1)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i;
    const match = ocrText.match(wagesPattern);
    if (match && match[1]) {
      return this.parseAmount(match[1]);
    }
    return 0;
  }

  private analyzeDocumentTypeFromOCR(ocrText: string): string {
    const text = ocrText.toLowerCase();
    
    if (text.includes('form w-2') || text.includes('wage and tax statement')) {
      return 'W2';
    } else if (text.includes('form 1099-int') || text.includes('interest income')) {
      return 'FORM_1099_INT';
    } else if (text.includes('form 1099-div') || text.includes('dividends and distributions')) {
      return 'FORM_1099_DIV';
    } else if (text.includes('form 1099-misc') || text.includes('miscellaneous income')) {
      return 'FORM_1099_MISC';
    } else if (text.includes('form 1099-nec') || text.includes('nonemployee compensation')) {
      return 'FORM_1099_NEC';
    }
    
    return 'UNKNOWN';
  }

  // Placeholder methods for OCR extraction (you would implement these based on your needs)
  private extractW2FieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for W2 OCR extraction
    return baseData;
  }

  private extract1099IntFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-INT OCR extraction
    return baseData;
  }

  private extract1099DivFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-DIV OCR extraction
    return baseData;
  }

  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-MISC OCR extraction
    return baseData;
  }

  private extract1099NecFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-NEC OCR extraction
    return baseData;
  }

  private extractGenericFieldsFromOCR(ocrText: string, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for generic OCR extraction
    return baseData;
  }

  private process1099NecFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for 1099-NEC field processing
    return baseData;
  }

  private processGenericFields(fields: any, baseData: ExtractedFieldData): ExtractedFieldData {
    // Implementation for generic field processing
    return baseData;
  }
}
