import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import { readFile } from "fs/promises";

export interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  apiKey: string;
}

// Base interface for all tax documents
export interface BaseTaxDocument {
  fullText?: string;
  correctedDocumentType?: TaxDocumentType;
}

// Tax document types
export type TaxDocumentType = 'W2' | 'FORM_1099_INT' | 'FORM_1099_DIV' | 'FORM_1099_MISC' | 'FORM_1099_NEC';

// W2 specific interface
export interface W2Data extends BaseTaxDocument {
  // Employee Information
  employeeName?: string;
  employeeSSN?: string;
  employeeAddress?: string;
  employeeAddressStreet?: string;
  employeeCity?: string;
  employeeState?: string;
  employeeZipCode?: string;
  
  // Employer Information
  employerName?: string;
  employerEIN?: string;
  employerAddress?: string;
  
  // Box 1-6: Core wage and tax information
  wages?: number;                    // Box 1
  federalTaxWithheld?: number;       // Box 2
  socialSecurityWages?: number;      // Box 3
  socialSecurityTaxWithheld?: number; // Box 4
  medicareWages?: number;            // Box 5
  medicareTaxWithheld?: number;      // Box 6
  
  // Box 7-11: Additional compensation
  socialSecurityTips?: number;       // Box 7
  allocatedTips?: number;            // Box 8
  advanceEIC?: number;               // Box 9
  dependentCareBenefits?: number;    // Box 10
  nonqualifiedPlans?: number;        // Box 11
  
  // Box 12: Deferred compensation codes
  box12Raw?: string;
  box12Codes?: Array<{ code: string; amount: number }>;
  
  // Box 13: Checkboxes
  box13Checkboxes?: {
    retirementPlan?: boolean;
    thirdPartySickPay?: boolean;
    statutoryEmployee?: boolean;
  };
  
  // Box 14: Other
  otherTaxInfo?: string;             // Box 14
  
  // Box 15-20: State and local information
  stateEmployerID?: string;          // Box 15
  stateWages?: number;               // Box 16
  stateTaxWithheld?: number;         // Box 17
  localWages?: number;               // Box 18
  localTaxWithheld?: number;         // Box 19
  localityName?: string;             // Box 20
}

// 1099-INT specific interface
export interface Form1099IntData extends BaseTaxDocument {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  
  // Box 1-15: Interest income details
  interestIncome?: number;                           // Box 1
  earlyWithdrawalPenalty?: number;                   // Box 2
  interestOnUSavingsBonds?: number;                  // Box 3
  federalTaxWithheld?: number;                       // Box 4
  investmentExpenses?: number;                       // Box 5
  foreignTaxPaid?: number;                           // Box 6
  foreignCountry?: string;                           // Box 7
  taxExemptInterest?: number;                        // Box 8
  specifiedPrivateActivityBondInterest?: number;     // Box 9
  marketDiscount?: number;                           // Box 10
  bondPremium?: number;                              // Box 11
  stateTaxWithheld?: number;                         // Box 13
  statePayerNumber?: string;                         // Box 14
  stateInterest?: number;                            // Box 15
}

// 1099-DIV specific interface
export interface Form1099DivData extends BaseTaxDocument {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  
  // Box 1a-1b: Dividend information
  ordinaryDividends?: number;                        // Box 1a
  qualifiedDividends?: number;                       // Box 1b
  
  // Box 2a-2f: Capital gain distributions
  totalCapitalGain?: number;                         // Box 2a
  unrecapturedSection1250Gain?: number;              // Box 2b
  section1202Gain?: number;                          // Box 2c
  collectiblesGain?: number;                         // Box 2d
  section897OrdinaryDividends?: number;              // Box 2e
  section897CapitalGain?: number;                    // Box 2f
  
  // Box 3-13: Other dividend information
  nondividendDistributions?: number;                 // Box 3
  federalTaxWithheld?: number;                       // Box 4
  section199ADividends?: number;                     // Box 5
  exemptInterestDividends?: number;                  // Box 6
  foreignTaxPaid?: number;                           // Box 7
  foreignCountry?: string;                           // Box 8
  cashLiquidationDistributions?: number;             // Box 9
  noncashLiquidationDistributions?: number;          // Box 10
  fatcaFilingRequirement?: boolean;                  // Box 11
  investmentExpenses?: number;                       // Box 13
  
  // State information
  stateTaxWithheld?: number;
  statePayerNumber?: string;
  stateIncome?: number;
}

// 1099-MISC specific interface
export interface Form1099MiscData extends BaseTaxDocument {
  // Payer Information
  payerName?: string;
  payerTIN?: string;
  payerAddress?: string;
  
  // Recipient Information
  recipientName?: string;
  recipientTIN?: string;
  recipientAddress?: string;
  accountNumber?: string;
  
  // Box 1-18: Miscellaneous income
  rents?: number;                                    // Box 1
  royalties?: number;                                // Box 2
  otherIncome?: number;                              // Box 3
  federalTaxWithheld?: number;                       // Box 4
  fishingBoatProceeds?: number;                      // Box 5
  medicalHealthPayments?: number;                    // Box 6
  nonemployeeCompensation?: number;                  // Box 7 (deprecated)
  substitutePayments?: number;                       // Box 8
  cropInsuranceProceeds?: number;                    // Box 9
  grossProceedsAttorney?: number;                    // Box 10
  fishPurchases?: number;                            // Box 11
  section409ADeferrals?: number;                     // Box 12
  excessGoldenParachutePayments?: number;            // Box 13
  nonqualifiedDeferredCompensation?: number;         // Box 14
  section409AIncome?: number;                        // Box 15a
  stateTaxWithheld?: number;                         // Box 16
  statePayerNumber?: string;                         // Box 17
  stateIncome?: number;                              // Box 18
}

// Union type for all tax document data
export type TaxDocumentData = W2Data | Form1099IntData | Form1099DivData | Form1099MiscData;

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

  /**
   * Extract data from tax documents (W2, 1099-DIV, 1099-INT, 1099-MISC)
   */
  async extractTaxDocumentData(
    documentPathOrBuffer: string | Buffer,
    documentType: TaxDocumentType
  ): Promise<TaxDocumentData> {
    try {
      console.log('🔍 [Azure DI] Processing tax document with Azure Document Intelligence...');
      console.log('🔍 [Azure DI] Document type:', documentType);
      
      // Get document buffer
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Determine the model to use based on document type
      const modelId = this.getModelIdForTaxDocument(documentType);
      console.log('🔍 [Azure DI] Using model:', modelId);
      
      let extractedData: TaxDocumentData;
      let correctedDocumentType: TaxDocumentType | undefined;
      
      try {
        // Analyze the document with specific tax model
        const poller = await this.client.beginAnalyzeDocument(modelId, documentBuffer);
        const result = await poller.pollUntilDone();
        
        console.log('✅ [Azure DI] Document analysis completed with tax model');
        console.log('🔍 [Azure DI] Available fields from Azure:', Object.keys(result.documents?.[0]?.fields || {}));
        
        // Extract the data based on document type
        extractedData = this.extractTaxDocumentFields(result, documentType);
        
        // Perform OCR-based document type correction if we have OCR text
        if (extractedData.fullText) {
          const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText);
          if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
            console.log(`🔄 [Azure DI] Document type correction: ${documentType} → ${ocrBasedType}`);
            
            // Validate the corrected document type
            if (this.isValidTaxDocumentType(ocrBasedType)) {
              correctedDocumentType = ocrBasedType as TaxDocumentType;
              
              // Re-extract data with the corrected document type
              console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
              extractedData = this.extractTaxDocumentFields(result, correctedDocumentType);
            } else {
              console.log(`⚠️ [Azure DI] Invalid document type detected: ${ocrBasedType}, ignoring correction`);
            }
          }
        }
        
      } catch (modelError: any) {
        console.warn('⚠️ [Azure DI] Tax model failed, attempting fallback to OCR model:', modelError?.message);
        
        // Check if it's a ModelNotFound error
        if (this.isModelNotFoundError(modelError)) {
          console.log('🔍 [Azure DI] Falling back to prebuilt-read model for OCR extraction...');
          
          // Fallback to general OCR model
          const fallbackPoller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
          const fallbackResult = await fallbackPoller.pollUntilDone();
          
          console.log('✅ [Azure DI] Document analysis completed with OCR fallback');
          
          // Extract data using OCR-based approach
          extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, documentType);
          
          // Perform OCR-based document type correction
          if (extractedData.fullText) {
            const ocrBasedType = this.analyzeDocumentTypeFromOCR(extractedData.fullText);
            if (ocrBasedType !== 'UNKNOWN' && ocrBasedType !== documentType) {
              console.log(`🔄 [Azure DI] Document type correction (OCR fallback): ${documentType} → ${ocrBasedType}`);
              
              if (this.isValidTaxDocumentType(ocrBasedType)) {
                correctedDocumentType = ocrBasedType as TaxDocumentType;
                
                // Re-extract data with the corrected document type
                console.log('🔍 [Azure DI] Re-extracting data with corrected document type...');
                extractedData = this.extractTaxDocumentFieldsFromOCR(fallbackResult, correctedDocumentType);
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
    } catch (error: any) {
      console.error('❌ [Azure DI] Processing error:', error);
      throw new Error(`Azure Document Intelligence processing failed: ${error?.message || 'Unknown error'}`);
    }
  }

  /**
   * Get the appropriate model ID for tax document types
   */
  private getModelIdForTaxDocument(documentType: TaxDocumentType): string {
    switch (documentType) {
      case 'W2':
        return 'prebuilt-tax.us.w2';
      case 'FORM_1099_INT':
        return 'prebuilt-tax.us.1099INT';
      case 'FORM_1099_DIV':
        return 'prebuilt-tax.us.1099DIV';
      case 'FORM_1099_MISC':
        return 'prebuilt-tax.us.1099MISC';
      case 'FORM_1099_NEC':
        return 'prebuilt-tax.us.1099NEC';
      default:
        // Use unified tax model as fallback
        return 'prebuilt-tax.us';
    }
  }

  /**
   * Extract fields from tax documents using structured analysis
   */
  private extractTaxDocumentFields(result: any, documentType: TaxDocumentType): TaxDocumentData {
    const baseData: BaseTaxDocument = {
      fullText: result.content || ''
    };
    
    console.log('🔍 [Azure DI] Extracting structured fields for document type:', documentType);
    
    // Extract form fields
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      if (document.fields) {
        console.log('🔍 [Azure DI] Available Azure fields:', Object.keys(document.fields));
        
        // Process fields based on document type
        switch (documentType) {
          case 'W2':
            return this.processW2Fields(document.fields, baseData);
          case 'FORM_1099_INT':
            return this.process1099IntFields(document.fields, baseData);
          case 'FORM_1099_DIV':
            return this.process1099DivFields(document.fields, baseData);
          case 'FORM_1099_MISC':
            return this.process1099MiscFields(document.fields, baseData);
          case 'FORM_1099_NEC':
            return this.process1099NecFields(document.fields, baseData);
          default:
            return this.processGenericTaxFields(document.fields, baseData);
        }
      }
    }
    
    // Extract key-value pairs from tables if available
    if (result.keyValuePairs) {
      console.log('🔍 [Azure DI] Processing key-value pairs...');
      const genericData = { ...baseData } as any;
      for (const kvp of result.keyValuePairs) {
        const key = kvp.key?.content?.trim();
        const value = kvp.value?.content?.trim();
        if (key && value) {
          genericData[key] = value;
        }
      }
      return genericData;
    }
    
    console.log('⚠️ [Azure DI] No structured fields found, returning base data only');
    return baseData as TaxDocumentData;
  }

  /**
   * Process W2 fields from structured analysis with comprehensive field mapping
   */
  private processW2Fields(fields: any, baseData: BaseTaxDocument): W2Data {
    const w2Data: W2Data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing W2 fields with comprehensive mapping...');
    
    // Comprehensive W2 field mappings covering all possible Azure field names
    const w2FieldMappings = {
      // Employee information - all possible variations
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN', 
      'Employee.Address': 'employeeAddress',
      'EmployeeName': 'employeeName',
      'EmployeeSSN': 'employeeSSN',
      'EmployeeAddress': 'employeeAddress',
      'Employee': 'employeeName',
      'employeeName': 'employeeName',
      'employeeSSN': 'employeeSSN',
      'employeeAddress': 'employeeAddress',
      'name': 'employeeName',
      'ssn': 'employeeSSN',
      'address': 'employeeAddress',
      'streetAddress': 'employeeAddress',
      'employeeStreetAddress': 'employeeAddress',
      'recipientName': 'employeeName',
      'recipientSSN': 'employeeSSN',
      'recipientAddress': 'employeeAddress',
      
      // Employer information - all possible variations
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'EmployerName': 'employerName',
      'EmployerEIN': 'employerEIN',
      'EmployerAddress': 'employerAddress',
      'Employer': 'employerName',
      'employerName': 'employerName',
      'employerEIN': 'employerEIN',
      'employerAddress': 'employerAddress',
      'payerName': 'employerName',
      'payerEIN': 'employerEIN',
      'payerAddress': 'employerAddress',
      'companyName': 'employerName',
      'companyEIN': 'employerEIN',
      'companyAddress': 'employerAddress',
      
      // Box 1-6: Core wage and tax information - all variations
      'WagesAndTips': 'wages',
      'Wages': 'wages',
      'wages': 'wages',
      'wagesAndTips': 'wages',
      'totalWages': 'wages',
      'grossWages': 'wages',
      'Box1': 'wages',
      'box1': 'wages',
      '1': 'wages',
      
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'federalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalTaxWithheld': 'federalTaxWithheld',
      'fedTaxWithheld': 'federalTaxWithheld',
      'Box2': 'federalTaxWithheld',
      'box2': 'federalTaxWithheld',
      '2': 'federalTaxWithheld',
      
      'SocialSecurityWages': 'socialSecurityWages',
      'socialSecurityWages': 'socialSecurityWages',
      'ssWages': 'socialSecurityWages',
      'Box3': 'socialSecurityWages',
      'box3': 'socialSecurityWages',
      '3': 'socialSecurityWages',
      
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'socialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'ssTaxWithheld': 'socialSecurityTaxWithheld',
      'Box4': 'socialSecurityTaxWithheld',
      'box4': 'socialSecurityTaxWithheld',
      '4': 'socialSecurityTaxWithheld',
      
      'MedicareWagesAndTips': 'medicareWages',
      'MedicareWages': 'medicareWages',
      'medicareWagesAndTips': 'medicareWages',
      'medicareWages': 'medicareWages',
      'Box5': 'medicareWages',
      'box5': 'medicareWages',
      '5': 'medicareWages',
      
      'MedicareTaxWithheld': 'medicareTaxWithheld',
      'medicareTaxWithheld': 'medicareTaxWithheld',
      'Box6': 'medicareTaxWithheld',
      'box6': 'medicareTaxWithheld',
      '6': 'medicareTaxWithheld',
      
      // Box 7-11: Additional compensation - all variations
      'SocialSecurityTips': 'socialSecurityTips',
      'socialSecurityTips': 'socialSecurityTips',
      'ssTips': 'socialSecurityTips',
      'Box7': 'socialSecurityTips',
      'box7': 'socialSecurityTips',
      '7': 'socialSecurityTips',
      
      'AllocatedTips': 'allocatedTips',
      'allocatedTips': 'allocatedTips',
      'Box8': 'allocatedTips',
      'box8': 'allocatedTips',
      '8': 'allocatedTips',
      
      'AdvanceEIC': 'advanceEIC',
      'advanceEIC': 'advanceEIC',
      'advanceEarnedIncomeCredit': 'advanceEIC',
      'Box9': 'advanceEIC',
      'box9': 'advanceEIC',
      '9': 'advanceEIC',
      
      'DependentCareBenefits': 'dependentCareBenefits',
      'dependentCareBenefits': 'dependentCareBenefits',
      'Box10': 'dependentCareBenefits',
      'box10': 'dependentCareBenefits',
      '10': 'dependentCareBenefits',
      
      'NonqualifiedPlans': 'nonqualifiedPlans',
      'nonqualifiedPlans': 'nonqualifiedPlans',
      'Box11': 'nonqualifiedPlans',
      'box11': 'nonqualifiedPlans',
      '11': 'nonqualifiedPlans',
      
      // Box 12: Deferred compensation - all variations
      'DeferredCompensation': 'box12Raw',
      'deferredCompensation': 'box12Raw',
      'Box12': 'box12Raw',
      'box12': 'box12Raw',
      '12': 'box12Raw',
      'codes': 'box12Raw',
      'compensationCodes': 'box12Raw',
      
      // Box 14: Other - all variations
      'OtherTaxInfo': 'otherTaxInfo',
      'otherTaxInfo': 'otherTaxInfo',
      'other': 'otherTaxInfo',
      'Box14': 'otherTaxInfo',
      'box14': 'otherTaxInfo',
      '14': 'otherTaxInfo',
      
      // Box 15-20: State and local information - all variations
      'StateEmployerID': 'stateEmployerID',
      'stateEmployerID': 'stateEmployerID',
      'EmployerStateIdNumber': 'stateEmployerID',
      'employerStateIdNumber': 'stateEmployerID',
      'stateId': 'stateEmployerID',
      'Box15': 'stateEmployerID',
      'box15': 'stateEmployerID',
      '15': 'stateEmployerID',
      
      'StateWagesTipsEtc': 'stateWages',
      'StateWages': 'stateWages',
      'stateWagesTipsEtc': 'stateWages',
      'stateWages': 'stateWages',
      'Box16': 'stateWages',
      'box16': 'stateWages',
      '16': 'stateWages',
      
      'StateIncomeTax': 'stateTaxWithheld',
      'StateTaxWithheld': 'stateTaxWithheld',
      'stateIncomeTax': 'stateTaxWithheld',
      'stateTaxWithheld': 'stateTaxWithheld',
      'Box17': 'stateTaxWithheld',
      'box17': 'stateTaxWithheld',
      '17': 'stateTaxWithheld',
      
      'LocalWagesTipsEtc': 'localWages',
      'LocalWages': 'localWages',
      'localWagesTipsEtc': 'localWages',
      'localWages': 'localWages',
      'Box18': 'localWages',
      'box18': 'localWages',
      '18': 'localWages',
      
      'LocalIncomeTax': 'localTaxWithheld',
      'LocalTaxWithheld': 'localTaxWithheld',
      'localIncomeTax': 'localTaxWithheld',
      'localTaxWithheld': 'localTaxWithheld',
      'Box19': 'localTaxWithheld',
      'box19': 'localTaxWithheld',
      '19': 'localTaxWithheld',
      
      'LocalityName': 'localityName',
      'localityName': 'localityName',
      'locality': 'localityName',
      'Box20': 'localityName',
      'box20': 'localityName',
      '20': 'localityName'
    };
    
    let fieldsExtracted = 0;
    
    // Map fields using comprehensive mapping
    for (const [azureFieldName, mappedFieldName] of Object.entries(w2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        console.log(`✅ [Azure DI] Mapping ${azureFieldName} → ${mappedFieldName}:`, value);
        
        if (mappedFieldName === 'box12Raw' || mappedFieldName === 'otherTaxInfo' || 
            mappedFieldName === 'stateEmployerID' || mappedFieldName === 'localityName' ||
            mappedFieldName === 'employeeName' || mappedFieldName === 'employerName' ||
            mappedFieldName === 'employeeSSN' || mappedFieldName === 'employerEIN' ||
            mappedFieldName === 'employeeAddress' || mappedFieldName === 'employerAddress') {
          // Text fields
          (w2Data as any)[mappedFieldName] = String(value).trim();
        } else {
          // Numeric fields
          const numericValue = this.parseAmount(value);
          if (numericValue > 0) {
            (w2Data as any)[mappedFieldName] = numericValue;
          }
        }
        fieldsExtracted++;
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} W2 fields from structured data`);
    
    // Parse Box 12 codes if available
    if (w2Data.box12Raw) {
      const box12Codes = this.parseW2Box12Codes(w2Data.box12Raw);
      if (box12Codes.length > 0) {
        w2Data.box12Codes = box12Codes;
        console.log('✅ [Azure DI] Parsed Box 12 codes:', box12Codes);
      }
    }
    
    // Extract Box 13 checkboxes
    w2Data.box13Checkboxes = this.extractW2Box13Checkboxes(fields, baseData.fullText || '');
    
    // OCR fallback for missing personal info
    if ((!w2Data.employeeName || !w2Data.employeeSSN || !w2Data.employeeAddress || 
         !w2Data.employerName || !w2Data.employerEIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some W2 info missing from structured fields, attempting OCR extraction...');
      const personalInfoFromOCR = this.extractPersonalInfoFromOCR(baseData.fullText, w2Data.employeeName);
      
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
      
      if (!w2Data.employerEIN && personalInfoFromOCR.employerEIN) {
        w2Data.employerEIN = personalInfoFromOCR.employerEIN;
        console.log('✅ [Azure DI] Extracted employer EIN from OCR:', w2Data.employerEIN);
      }
    }
    
    // Parse address components if full address is available
    if (w2Data.employeeAddress && typeof w2Data.employeeAddress === 'string') {
      const addressParts = this.extractAddressParts(w2Data.employeeAddress, baseData.fullText || '');
      w2Data.employeeAddressStreet = addressParts.street;
      w2Data.employeeCity = addressParts.city;
      w2Data.employeeState = addressParts.state;
      w2Data.employeeZipCode = addressParts.zipCode;
    }
    
    return w2Data;
  }

  /**
   * Process 1099-INT fields from structured analysis
   */
  private process1099IntFields(fields: any, baseData: BaseTaxDocument): Form1099IntData {
    const data: Form1099IntData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-INT fields...');
    
    const fieldMappings = {
      // Payer and recipient information - all variations
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'payerName': 'payerName',
      'payerTIN': 'payerTIN',
      'payerAddress': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'recipientName': 'recipientName',
      'recipientTIN': 'recipientTIN',
      'recipientAddress': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'accountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Box 1-15 mappings - all variations
      'InterestIncome': 'interestIncome',
      'interestIncome': 'interestIncome',
      'Interest': 'interestIncome',
      'Box1': 'interestIncome',
      'box1': 'interestIncome',
      '1': 'interestIncome',
      
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
      'earlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
      'Box2': 'earlyWithdrawalPenalty',
      'box2': 'earlyWithdrawalPenalty',
      '2': 'earlyWithdrawalPenalty',
      
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',
      'InterestOnUSavingsBonds': 'interestOnUSavingsBonds',
      'interestOnUSavingsBonds': 'interestOnUSavingsBonds',
      'Box3': 'interestOnUSavingsBonds',
      'box3': 'interestOnUSavingsBonds',
      '3': 'interestOnUSavingsBonds',
      
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalTaxWithheld': 'federalTaxWithheld',
      'Box4': 'federalTaxWithheld',
      'box4': 'federalTaxWithheld',
      '4': 'federalTaxWithheld',
      
      'InvestmentExpenses': 'investmentExpenses',
      'investmentExpenses': 'investmentExpenses',
      'Box5': 'investmentExpenses',
      'box5': 'investmentExpenses',
      '5': 'investmentExpenses',
      
      'ForeignTaxPaid': 'foreignTaxPaid',
      'foreignTaxPaid': 'foreignTaxPaid',
      'Box6': 'foreignTaxPaid',
      'box6': 'foreignTaxPaid',
      '6': 'foreignTaxPaid',
      
      'ForeignCountry': 'foreignCountry',
      'foreignCountry': 'foreignCountry',
      'Box7': 'foreignCountry',
      'box7': 'foreignCountry',
      '7': 'foreignCountry',
      
      'TaxExemptInterest': 'taxExemptInterest',
      'taxExemptInterest': 'taxExemptInterest',
      'Box8': 'taxExemptInterest',
      'box8': 'taxExemptInterest',
      '8': 'taxExemptInterest',
      
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest',
      'specifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest',
      'Box9': 'specifiedPrivateActivityBondInterest',
      'box9': 'specifiedPrivateActivityBondInterest',
      '9': 'specifiedPrivateActivityBondInterest',
      
      'MarketDiscount': 'marketDiscount',
      'marketDiscount': 'marketDiscount',
      'Box10': 'marketDiscount',
      'box10': 'marketDiscount',
      '10': 'marketDiscount',
      
      'BondPremium': 'bondPremium',
      'bondPremium': 'bondPremium',
      'Box11': 'bondPremium',
      'box11': 'bondPremium',
      '11': 'bondPremium',
      
      'StateTaxWithheld': 'stateTaxWithheld',
      'stateTaxWithheld': 'stateTaxWithheld',
      'Box13': 'stateTaxWithheld',
      'box13': 'stateTaxWithheld',
      '13': 'stateTaxWithheld',
      
      'StatePayerNumber': 'statePayerNumber',
      'statePayerNumber': 'statePayerNumber',
      'Box14': 'statePayerNumber',
      'box14': 'statePayerNumber',
      '14': 'statePayerNumber',
      
      'StateInterest': 'stateInterest',
      'stateInterest': 'stateInterest',
      'Box15': 'stateInterest',
      'box15': 'stateInterest',
      '15': 'stateInterest'
    };
    
    let fieldsExtracted = 0;
    
    // Map fields
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        console.log(`✅ [Azure DI] Mapping ${azureFieldName} → ${mappedFieldName}:`, value);
        
        if (mappedFieldName === 'foreignCountry' || mappedFieldName === 'statePayerNumber' || 
            mappedFieldName === 'accountNumber' || mappedFieldName === 'payerName' ||
            mappedFieldName === 'recipientName' || mappedFieldName === 'payerTIN' ||
            mappedFieldName === 'recipientTIN' || mappedFieldName === 'payerAddress' ||
            mappedFieldName === 'recipientAddress') {
          // Text fields
          (data as any)[mappedFieldName] = String(value).trim();
        } else {
          // Numeric fields
          const numericValue = this.parseAmount(value);
          if (numericValue > 0) {
            (data as any)[mappedFieldName] = numericValue;
          }
        }
        fieldsExtracted++;
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-INT fields from structured data`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-DIV fields from structured analysis
   */
  private process1099DivFields(fields: any, baseData: BaseTaxDocument): Form1099DivData {
    const data: Form1099DivData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-DIV fields...');
    
    const fieldMappings = {
      // Payer and recipient information - all variations
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'payerName': 'payerName',
      'payerTIN': 'payerTIN',
      'payerAddress': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'recipientName': 'recipientName',
      'recipientTIN': 'recipientTIN',
      'recipientAddress': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'accountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Dividend fields - all variations
      'OrdinaryDividends': 'ordinaryDividends',
      'ordinaryDividends': 'ordinaryDividends',
      'Box1a': 'ordinaryDividends',
      'box1a': 'ordinaryDividends',
      '1a': 'ordinaryDividends',
      
      'QualifiedDividends': 'qualifiedDividends',
      'qualifiedDividends': 'qualifiedDividends',
      'Box1b': 'qualifiedDividends',
      'box1b': 'qualifiedDividends',
      '1b': 'qualifiedDividends',
      
      'TotalCapitalGainDistributions': 'totalCapitalGain',
      'totalCapitalGainDistributions': 'totalCapitalGain',
      'totalCapitalGain': 'totalCapitalGain',
      'capitalGainDistributions': 'totalCapitalGain',
      'Box2a': 'totalCapitalGain',
      'box2a': 'totalCapitalGain',
      '2a': 'totalCapitalGain',
      
      'UnrecapturedSection1250Gain': 'unrecapturedSection1250Gain',
      'unrecapturedSection1250Gain': 'unrecapturedSection1250Gain',
      'Box2b': 'unrecapturedSection1250Gain',
      'box2b': 'unrecapturedSection1250Gain',
      '2b': 'unrecapturedSection1250Gain',
      
      'Section1202Gain': 'section1202Gain',
      'section1202Gain': 'section1202Gain',
      'Box2c': 'section1202Gain',
      'box2c': 'section1202Gain',
      '2c': 'section1202Gain',
      
      'CollectiblesGain': 'collectiblesGain',
      'collectiblesGain': 'collectiblesGain',
      'Box2d': 'collectiblesGain',
      'box2d': 'collectiblesGain',
      '2d': 'collectiblesGain',
      
      'Section897OrdinaryDividends': 'section897OrdinaryDividends',
      'section897OrdinaryDividends': 'section897OrdinaryDividends',
      'Box2e': 'section897OrdinaryDividends',
      'box2e': 'section897OrdinaryDividends',
      '2e': 'section897OrdinaryDividends',
      
      'Section897CapitalGain': 'section897CapitalGain',
      'section897CapitalGain': 'section897CapitalGain',
      'Box2f': 'section897CapitalGain',
      'box2f': 'section897CapitalGain',
      '2f': 'section897CapitalGain',
      
      'NondividendDistributions': 'nondividendDistributions',
      'nondividendDistributions': 'nondividendDistributions',
      'Box3': 'nondividendDistributions',
      'box3': 'nondividendDistributions',
      '3': 'nondividendDistributions',
      
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalTaxWithheld': 'federalTaxWithheld',
      'Box4': 'federalTaxWithheld',
      'box4': 'federalTaxWithheld',
      '4': 'federalTaxWithheld',
      
      'Section199ADividends': 'section199ADividends',
      'section199ADividends': 'section199ADividends',
      'Box5': 'section199ADividends',
      'box5': 'section199ADividends',
      '5': 'section199ADividends',
      
      'ExemptInterestDividends': 'exemptInterestDividends',
      'exemptInterestDividends': 'exemptInterestDividends',
      'Box6': 'exemptInterestDividends',
      'box6': 'exemptInterestDividends',
      '6': 'exemptInterestDividends',
      
      'ForeignTaxPaid': 'foreignTaxPaid',
      'foreignTaxPaid': 'foreignTaxPaid',
      'Box7': 'foreignTaxPaid',
      'box7': 'foreignTaxPaid',
      '7': 'foreignTaxPaid',
      
      'ForeignCountry': 'foreignCountry',
      'foreignCountry': 'foreignCountry',
      'Box8': 'foreignCountry',
      'box8': 'foreignCountry',
      '8': 'foreignCountry',
      
      'CashLiquidationDistributions': 'cashLiquidationDistributions',
      'cashLiquidationDistributions': 'cashLiquidationDistributions',
      'Box9': 'cashLiquidationDistributions',
      'box9': 'cashLiquidationDistributions',
      '9': 'cashLiquidationDistributions',
      
      'NoncashLiquidationDistributions': 'noncashLiquidationDistributions',
      'noncashLiquidationDistributions': 'noncashLiquidationDistributions',
      'Box10': 'noncashLiquidationDistributions',
      'box10': 'noncashLiquidationDistributions',
      '10': 'noncashLiquidationDistributions',
      
      'FATCAFilingRequirement': 'fatcaFilingRequirement',
      'fatcaFilingRequirement': 'fatcaFilingRequirement',
      'Box11': 'fatcaFilingRequirement',
      'box11': 'fatcaFilingRequirement',
      '11': 'fatcaFilingRequirement',
      
      'InvestmentExpenses': 'investmentExpenses',
      'investmentExpenses': 'investmentExpenses',
      'Box13': 'investmentExpenses',
      'box13': 'investmentExpenses',
      '13': 'investmentExpenses',
      
      'StateTaxWithheld': 'stateTaxWithheld',
      'stateTaxWithheld': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',
      'statePayerNumber': 'statePayerNumber',
      'StateIncome': 'stateIncome',
      'stateIncome': 'stateIncome'
    };
    
    let fieldsExtracted = 0;
    
    // Map fields
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        console.log(`✅ [Azure DI] Mapping ${azureFieldName} → ${mappedFieldName}:`, value);
        
        if (mappedFieldName === 'foreignCountry' || mappedFieldName === 'statePayerNumber' || 
            mappedFieldName === 'accountNumber' || mappedFieldName === 'payerName' ||
            mappedFieldName === 'recipientName' || mappedFieldName === 'payerTIN' ||
            mappedFieldName === 'recipientTIN' || mappedFieldName === 'payerAddress' ||
            mappedFieldName === 'recipientAddress') {
          // Text fields
          (data as any)[mappedFieldName] = String(value).trim();
        } else if (mappedFieldName === 'fatcaFilingRequirement') {
          // Boolean field
          (data as any)[mappedFieldName] = this.parseBoolean(value);
        } else {
          // Numeric fields
          const numericValue = this.parseAmount(value);
          if (numericValue > 0) {
            (data as any)[mappedFieldName] = numericValue;
          }
        }
        fieldsExtracted++;
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-DIV fields from structured data`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-MISC fields from structured analysis
   */
  private process1099MiscFields(fields: any, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-MISC fields...');
    
    const fieldMappings = {
      // Payer and recipient information - all variations
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'payerName': 'payerName',
      'payerTIN': 'payerTIN',
      'payerAddress': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'recipientName': 'recipientName',
      'recipientTIN': 'recipientTIN',
      'recipientAddress': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'accountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Box 1-18 mappings - all variations
      'Rents': 'rents',
      'rents': 'rents',
      'Box1': 'rents',
      'box1': 'rents',
      '1': 'rents',
      
      'Royalties': 'royalties',
      'royalties': 'royalties',
      'Box2': 'royalties',
      'box2': 'royalties',
      '2': 'royalties',
      
      'OtherIncome': 'otherIncome',
      'otherIncome': 'otherIncome',
      'Box3': 'otherIncome',
      'box3': 'otherIncome',
      '3': 'otherIncome',
      
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalTaxWithheld': 'federalTaxWithheld',
      'Box4': 'federalTaxWithheld',
      'box4': 'federalTaxWithheld',
      '4': 'federalTaxWithheld',
      
      'FishingBoatProceeds': 'fishingBoatProceeds',
      'fishingBoatProceeds': 'fishingBoatProceeds',
      'Box5': 'fishingBoatProceeds',
      'box5': 'fishingBoatProceeds',
      '5': 'fishingBoatProceeds',
      
      'MedicalAndHealthCarePayments': 'medicalHealthPayments',
      'medicalAndHealthCarePayments': 'medicalHealthPayments',
      'medicalHealthPayments': 'medicalHealthPayments',
      'Box6': 'medicalHealthPayments',
      'box6': 'medicalHealthPayments',
      '6': 'medicalHealthPayments',
      
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'nonemployeeCompensation': 'nonemployeeCompensation',
      'Box7': 'nonemployeeCompensation',
      'box7': 'nonemployeeCompensation',
      '7': 'nonemployeeCompensation',
      
      'SubstitutePayments': 'substitutePayments',
      'substitutePayments': 'substitutePayments',
      'Box8': 'substitutePayments',
      'box8': 'substitutePayments',
      '8': 'substitutePayments',
      
      'CropInsuranceProceeds': 'cropInsuranceProceeds',
      'cropInsuranceProceeds': 'cropInsuranceProceeds',
      'Box9': 'cropInsuranceProceeds',
      'box9': 'cropInsuranceProceeds',
      '9': 'cropInsuranceProceeds',
      
      'GrossProceedsPaidToAttorney': 'grossProceedsAttorney',
      'grossProceedsPaidToAttorney': 'grossProceedsAttorney',
      'grossProceedsAttorney': 'grossProceedsAttorney',
      'Box10': 'grossProceedsAttorney',
      'box10': 'grossProceedsAttorney',
      '10': 'grossProceedsAttorney',
      
      'FishPurchasedForResale': 'fishPurchases',
      'fishPurchasedForResale': 'fishPurchases',
      'fishPurchases': 'fishPurchases',
      'Box11': 'fishPurchases',
      'box11': 'fishPurchases',
      '11': 'fishPurchases',
      
      'Section409ADeferrals': 'section409ADeferrals',
      'section409ADeferrals': 'section409ADeferrals',
      'Box12': 'section409ADeferrals',
      'box12': 'section409ADeferrals',
      '12': 'section409ADeferrals',
      
      'ExcessGoldenParachutePayments': 'excessGoldenParachutePayments',
      'excessGoldenParachutePayments': 'excessGoldenParachutePayments',
      'Box13': 'excessGoldenParachutePayments',
      'box13': 'excessGoldenParachutePayments',
      '13': 'excessGoldenParachutePayments',
      
      'NonqualifiedDeferredCompensation': 'nonqualifiedDeferredCompensation',
      'nonqualifiedDeferredCompensation': 'nonqualifiedDeferredCompensation',
      'Box14': 'nonqualifiedDeferredCompensation',
      'box14': 'nonqualifiedDeferredCompensation',
      '14': 'nonqualifiedDeferredCompensation',
      
      'Section409AIncome': 'section409AIncome',
      'section409AIncome': 'section409AIncome',
      'Box15a': 'section409AIncome',
      'box15a': 'section409AIncome',
      '15a': 'section409AIncome',
      
      'StateTaxWithheld': 'stateTaxWithheld',
      'stateTaxWithheld': 'stateTaxWithheld',
      'Box16': 'stateTaxWithheld',
      'box16': 'stateTaxWithheld',
      '16': 'stateTaxWithheld',
      
      'StatePayerNumber': 'statePayerNumber',
      'statePayerNumber': 'statePayerNumber',
      'Box17': 'statePayerNumber',
      'box17': 'statePayerNumber',
      '17': 'statePayerNumber',
      
      'StateIncome': 'stateIncome',
      'stateIncome': 'stateIncome',
      'Box18': 'stateIncome',
      'box18': 'stateIncome',
      '18': 'stateIncome'
    };
    
    let fieldsExtracted = 0;
    
    // Map fields
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        console.log(`✅ [Azure DI] Mapping ${azureFieldName} → ${mappedFieldName}:`, value);
        
        if (mappedFieldName === 'statePayerNumber' || mappedFieldName === 'accountNumber' ||
            mappedFieldName === 'payerName' || mappedFieldName === 'recipientName' ||
            mappedFieldName === 'payerTIN' || mappedFieldName === 'recipientTIN' ||
            mappedFieldName === 'payerAddress' || mappedFieldName === 'recipientAddress') {
          // Text fields
          (data as any)[mappedFieldName] = String(value).trim();
        } else {
          // Numeric fields
          const numericValue = this.parseAmount(value);
          if (numericValue > 0) {
            (data as any)[mappedFieldName] = numericValue;
          }
        }
        fieldsExtracted++;
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-MISC fields from structured data`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    // Validate and correct field mappings using OCR
    if (baseData.fullText) {
      return this.validateAndCorrect1099MiscFields(data, baseData.fullText);
    }
    
    return data;
  }

  /**
   * Process 1099-NEC fields from structured analysis
   */
  private process1099NecFields(fields: any, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-NEC fields...');
    
    const fieldMappings = {
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'payerName': 'payerName',
      'payerTIN': 'payerTIN',
      'payerAddress': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'recipientName': 'recipientName',
      'recipientTIN': 'recipientTIN',
      'recipientAddress': 'recipientAddress',
      
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'nonemployeeCompensation': 'nonemployeeCompensation',
      'Box1': 'nonemployeeCompensation',
      'box1': 'nonemployeeCompensation',
      '1': 'nonemployeeCompensation',
      
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalIncomeTaxWithheld': 'federalTaxWithheld',
      'federalTaxWithheld': 'federalTaxWithheld',
      'Box4': 'federalTaxWithheld',
      'box4': 'federalTaxWithheld',
      '4': 'federalTaxWithheld'
    };
    
    let fieldsExtracted = 0;
    
    // Map fields
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        console.log(`✅ [Azure DI] Mapping ${azureFieldName} → ${mappedFieldName}:`, value);
        
        if (mappedFieldName === 'payerName' || mappedFieldName === 'recipientName' ||
            mappedFieldName === 'payerTIN' || mappedFieldName === 'recipientTIN' ||
            mappedFieldName === 'payerAddress' || mappedFieldName === 'recipientAddress') {
          // Text fields
          (data as any)[mappedFieldName] = String(value).trim();
        } else {
          // Numeric fields
          const numericValue = this.parseAmount(value);
          if (numericValue > 0) {
            (data as any)[mappedFieldName] = numericValue;
          }
        }
        fieldsExtracted++;
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-NEC fields from structured data`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process generic tax fields
   */
  private processGenericTaxFields(fields: any, baseData: BaseTaxDocument): TaxDocumentData {
    const data = { ...baseData } as any;
    
    console.log('🔍 [Azure DI] Processing generic tax fields...');
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [fieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && 'value' in fieldData) {
        const value = (fieldData as any).value;
        if (value !== undefined && value !== null && value !== '') {
          console.log(`✅ [Azure DI] Generic field ${fieldName}:`, value);
          data[fieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          fieldsExtracted++;
        }
      }
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} generic fields from structured data`);
    
    return data;
  }

  // OCR-based extraction methods (simplified versions for fallback)
  private extractTaxDocumentFieldsFromOCR(result: any, documentType: TaxDocumentType): TaxDocumentData {
    console.log('🔍 [Azure DI] Extracting tax document fields using OCR fallback...');
    
    const baseData: BaseTaxDocument = {
      fullText: result.content || ''
    };
    
    // Use OCR-based extraction methods for different document types
    switch (documentType) {
      case 'W2':
        return this.extractW2FieldsFromOCR(baseData.fullText!, baseData);
      case 'FORM_1099_INT':
        return this.extract1099IntFieldsFromOCR(baseData.fullText!, baseData);
      case 'FORM_1099_DIV':
        return this.extract1099DivFieldsFromOCR(baseData.fullText!, baseData);
      case 'FORM_1099_MISC':
        return this.extract1099MiscFieldsFromOCR(baseData.fullText!, baseData);
      case 'FORM_1099_NEC':
        return this.extract1099NecFieldsFromOCR(baseData.fullText!, baseData);
      default:
        console.log('🔍 [Azure DI] Using generic OCR extraction for document type:', documentType);
        return this.extractGenericTaxFieldsFromOCR(baseData.fullText!, baseData);
    }
  }

  private extractW2FieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): W2Data {
    const w2Data: W2Data = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting W2 fields from OCR text...');
    
    // Extract wages (Box 1)
    const wagesMatch = ocrText.match(/(?:box\s*1|wages.*tips.*compensation)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (wagesMatch) {
      w2Data.wages = this.parseAmount(wagesMatch[1]);
      console.log('✅ [Azure DI] OCR extracted wages:', w2Data.wages);
    }
    
    // Extract federal tax withheld (Box 2)
    const fedTaxMatch = ocrText.match(/(?:box\s*2|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      w2Data.federalTaxWithheld = this.parseAmount(fedTaxMatch[1]);
      console.log('✅ [Azure DI] OCR extracted federal tax withheld:', w2Data.federalTaxWithheld);
    }
    
    // Extract personal info using existing method
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      w2Data.employeeName = personalInfo.name;
      console.log('✅ [Azure DI] OCR extracted employee name:', w2Data.employeeName);
    }
    if (personalInfo.ssn) {
      w2Data.employeeSSN = personalInfo.ssn;
      console.log('✅ [Azure DI] OCR extracted employee SSN:', w2Data.employeeSSN);
    }
    if (personalInfo.address) {
      w2Data.employeeAddress = personalInfo.address;
      console.log('✅ [Azure DI] OCR extracted employee address:', w2Data.employeeAddress);
    }
    if (personalInfo.employerName) {
      w2Data.employerName = personalInfo.employerName;
      console.log('✅ [Azure DI] OCR extracted employer name:', w2Data.employerName);
    }
    if (personalInfo.employerEIN) {
      w2Data.employerEIN = personalInfo.employerEIN;
      console.log('✅ [Azure DI] OCR extracted employer EIN:', w2Data.employerEIN);
    }
    
    return w2Data;
  }

  private extract1099IntFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099IntData {
    const data: Form1099IntData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-INT fields from OCR text...');
    
    // Extract interest income (Box 1)
    const interestMatch = ocrText.match(/(?:box\s*1|interest.*income)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (interestMatch) {
      data.interestIncome = this.parseAmount(interestMatch[1]);
      console.log('✅ [Azure DI] OCR extracted interest income:', data.interestIncome);
    }
    
    // Extract federal tax withheld (Box 4)
    const fedTaxMatch = ocrText.match(/(?:box\s*4|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      data.federalTaxWithheld = this.parseAmount(fedTaxMatch[1]);
      console.log('✅ [Azure DI] OCR extracted federal tax withheld:', data.federalTaxWithheld);
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    return data;
  }

  private extract1099DivFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099DivData {
    const data: Form1099DivData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-DIV fields from OCR text...');
    
    // Extract ordinary dividends (Box 1a)
    const ordinaryDivMatch = ocrText.match(/(?:box\s*1a|ordinary.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (ordinaryDivMatch) {
      data.ordinaryDividends = this.parseAmount(ordinaryDivMatch[1]);
      console.log('✅ [Azure DI] OCR extracted ordinary dividends:', data.ordinaryDividends);
    }
    
    // Extract qualified dividends (Box 1b)
    const qualifiedDivMatch = ocrText.match(/(?:box\s*1b|qualified.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (qualifiedDivMatch) {
      data.qualifiedDividends = this.parseAmount(qualifiedDivMatch[1]);
      console.log('✅ [Azure DI] OCR extracted qualified dividends:', data.qualifiedDividends);
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    return data;
  }

  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-MISC fields from OCR text...');
    
    // Extract key fields using OCR patterns
    const fieldPatterns = [
      { field: 'rents', patterns: [/(?:box\s*1|rents)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] },
      { field: 'royalties', patterns: [/(?:box\s*2|royalties)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] },
      { field: 'otherIncome', patterns: [/(?:box\s*3|other.*income)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] },
      { field: 'federalTaxWithheld', patterns: [/(?:box\s*4|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] },
      { field: 'fishingBoatProceeds', patterns: [/(?:box\s*5|fishing.*boat.*proceeds)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] },
      { field: 'medicalHealthPayments', patterns: [/(?:box\s*6|medical.*health.*care.*payments)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i] }
    ];
    
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
            console.log(`✅ [Azure DI] OCR extracted ${field}:`, amount);
            break;
          }
        }
      }
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    return data;
  }

  private extract1099NecFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-NEC fields from OCR text...');
    
    // Extract nonemployee compensation
    const necMatch = ocrText.match(/(?:nonemployee.*compensation|box\s*1)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (necMatch) {
      data.nonemployeeCompensation = this.parseAmount(necMatch[1]);
      console.log('✅ [Azure DI] OCR extracted nonemployee compensation:', data.nonemployeeCompensation);
    }
    
    // Extract federal tax withheld
    const fedTaxMatch = ocrText.match(/(?:federal.*income.*tax.*withheld|box\s*4)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      data.federalTaxWithheld = this.parseAmount(fedTaxMatch[1]);
      console.log('✅ [Azure DI] OCR extracted federal tax withheld:', data.federalTaxWithheld);
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    return data;
  }

  private extractGenericTaxFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): TaxDocumentData {
    const data = { ...baseData } as any;
    
    console.log('🔍 [Azure DI] Extracting generic tax fields from OCR text...');
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) data.recipientName = personalInfo.name;
    if (personalInfo.tin) data.recipientTIN = personalInfo.tin;
    if (personalInfo.address) data.recipientAddress = personalInfo.address;
    if (personalInfo.payerName) data.payerName = personalInfo.payerName;
    if (personalInfo.payerTIN) data.payerTIN = personalInfo.payerTIN;
    
    return data;
  }

  // Utility methods
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

  private parseBoolean(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'string') {
      const lower = value.toLowerCase().trim();
      return lower === 'true' || lower === 'yes' || lower === 'x' || lower === '✓' || lower === 'checked';
    }
    return false;
  }

  private parseW2Box12Codes(box12String: string): Array<{ code: string; amount: number }> {
    if (!box12String) return [];
    
    const codes: Array<{ code: string; amount: number }> = [];
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

  private extractW2Box13Checkboxes(fields: any, ocrText: string): any {
    const checkboxes: any = {};
    
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
      const text = ocrText.toLowerCase();
      checkboxes.retirementPlan = /retirement\s+plan\s*[:\s]*(?:x|✓|checked|yes)/i.test(text);
      checkboxes.thirdPartySickPay = /third.party\s+sick\s+pay\s*[:\s]*(?:x|✓|checked|yes)/i.test(text);
      checkboxes.statutoryEmployee = /statutory\s+employee\s*[:\s]*(?:x|✓|checked|yes)/i.test(text);
    }
    
    return checkboxes;
  }

  private extractPersonalInfoFromOCR(ocrText: string, targetEmployeeName?: string): any {
    const personalInfo: any = {};
    
    console.log('🔍 [Azure DI] Extracting personal info from OCR...');
    
    // Extract SSN/TIN patterns
    const ssnPattern = /\b(\d{3}[-\s]?\d{2}[-\s]?\d{4})\b/g;
    const ssnMatches = Array.from(ocrText.matchAll(ssnPattern));
    if (ssnMatches.length > 0) {
      personalInfo.ssn = ssnMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.tin = personalInfo.ssn;
      console.log('✅ [Azure DI] OCR extracted SSN/TIN:', personalInfo.ssn);
    }
    
    // Extract EIN patterns
    const einPattern = /\b(\d{2}[-\s]?\d{7})\b/g;
    const einMatches = Array.from(ocrText.matchAll(einPattern));
    if (einMatches.length > 0) {
      personalInfo.employerEIN = einMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.payerTIN = personalInfo.employerEIN;
      console.log('✅ [Azure DI] OCR extracted EIN:', personalInfo.employerEIN);
    }
    
    // Extract names (simplified approach)
    const lines = ocrText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
    for (const line of lines) {
      // Look for name patterns (capitalized words)
      if (/^[A-Z][a-z]+\s+[A-Z][a-z]+/.test(line) && line.length < 50) {
        if (!personalInfo.name) {
          personalInfo.name = line;
          personalInfo.employerName = line; // Could be either
          personalInfo.payerName = line;
          console.log('✅ [Azure DI] OCR extracted name:', personalInfo.name);
        }
      }
    }
    
    // Extract addresses (simplified approach)
    for (const line of lines) {
      if (/\d+.*(?:street|st|avenue|ave|road|rd|drive|dr|lane|ln|way|blvd|boulevard)/i.test(line)) {
        if (!personalInfo.address) {
          personalInfo.address = line;
          console.log('✅ [Azure DI] OCR extracted address:', personalInfo.address);
        }
      }
    }
    
    return personalInfo;
  }

  private extractAddressParts(fullAddress: string, ocrText: string): any {
    const addressParts: any = {};
    
    // Extract ZIP code
    const zipMatch = fullAddress.match(/\b(\d{5}(?:-\d{4})?)\b/);
    if (zipMatch) {
      addressParts.zipCode = zipMatch[1];
    }
    
    // Extract state (2-letter abbreviation before ZIP)
    const stateMatch = fullAddress.match(/\b([A-Z]{2})\s+\d{5}/);
    if (stateMatch) {
      addressParts.state = stateMatch[1];
    }
    
    // Extract city (word(s) before state)
    const cityMatch = fullAddress.match(/([A-Za-z\s]+)\s+[A-Z]{2}\s+\d{5}/);
    if (cityMatch) {
      addressParts.city = cityMatch[1].trim();
    }
    
    // Extract street (everything before city)
    const streetMatch = fullAddress.match(/^(.+?)(?:\s+[A-Za-z\s]+\s+[A-Z]{2}\s+\d{5})/);
    if (streetMatch) {
      addressParts.street = streetMatch[1].trim();
    }
    
    return addressParts;
  }

  private applyPersonalInfoOCRFallback(data: any, fullText?: string): void {
    if (!fullText) return;
    
    const personalInfo = this.extractPersonalInfoFromOCR(fullText);
    
    if (!data.recipientName && personalInfo.name) {
      data.recipientName = personalInfo.name;
      console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
    }
    
    if (!data.recipientTIN && personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
    }
    
    if (!data.recipientAddress && personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
    }
    
    if (!data.payerName && personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
    }
    
    if (!data.payerTIN && personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
    }
  }

  private validateAndCorrect1099MiscFields(structuredData: Form1099MiscData, ocrText: string): Form1099MiscData {
    console.log('🔍 [Azure DI] Validating 1099-MISC field mappings...');
    
    // Extract data using OCR as ground truth
    const ocrData = this.extract1099MiscFieldsFromOCR(ocrText, { fullText: ocrText });
    
    const correctedData = { ...structuredData };
    let correctionsMade = 0;
    
    // Define validation rules for critical fields
    const criticalFields = ['otherIncome', 'fishingBoatProceeds', 'medicalHealthPayments', 'rents', 'royalties', 'federalTaxWithheld'];
    
    for (const field of criticalFields) {
      const structuredValue = this.parseAmount((structuredData as any)[field]) || 0;
      const ocrValue = this.parseAmount((ocrData as any)[field]) || 0;
      
      // If values differ significantly, trust OCR
      if (Math.abs(structuredValue - ocrValue) > 100) {
        console.log(`🔧 [Azure DI] Correcting ${field}: $${structuredValue} → $${ocrValue} (OCR)`);
        (correctedData as any)[field] = ocrValue;
        correctionsMade++;
      }
      // If structured field is empty but OCR found a value, use OCR
      else if ((structuredValue === 0 || !(structuredData as any)[field]) && ocrValue > 0) {
        console.log(`🔧 [Azure DI] Filling missing ${field}: $0 → $${ocrValue} (OCR)`);
        (correctedData as any)[field] = ocrValue;
        correctionsMade++;
      }
    }
    
    if (correctionsMade > 0) {
      console.log(`✅ [Azure DI] Made ${correctionsMade} field corrections using OCR validation`);
    } else {
      console.log('✅ [Azure DI] No field corrections needed - structured extraction appears accurate');
    }
    
    return correctedData;
  }

  private analyzeDocumentTypeFromOCR(ocrText: string): string {
    console.log('🔍 [Azure DI] Analyzing document type from OCR content...');
    
    const text = ocrText.toLowerCase();
    
    // Check for W2 indicators
    if (text.includes('wage and tax statement') || 
        text.includes('form w-2') || 
        text.includes('w-2') ||
        (text.includes('wages') && text.includes('social security'))) {
      return 'W2';
    }
    
    // Check for 1099-INT indicators
    if (text.includes('1099-int') || 
        text.includes('interest income') ||
        (text.includes('1099') && text.includes('interest'))) {
      return 'FORM_1099_INT';
    }
    
    // Check for 1099-DIV indicators
    if (text.includes('1099-div') || 
        text.includes('dividends and distributions') ||
        (text.includes('1099') && text.includes('dividend'))) {
      return 'FORM_1099_DIV';
    }
    
    // Check for 1099-MISC indicators
    if (text.includes('1099-misc') || 
        text.includes('miscellaneous income') ||
        (text.includes('1099') && (text.includes('rents') || text.includes('royalties')))) {
      return 'FORM_1099_MISC';
    }
    
    // Check for 1099-NEC indicators
    if (text.includes('1099-nec') || 
        text.includes('nonemployee compensation')) {
      return 'FORM_1099_NEC';
    }
    
    return 'UNKNOWN';
  }

  private isValidTaxDocumentType(docType: string): boolean {
    const validTypes: TaxDocumentType[] = ['W2', 'FORM_1099_INT', 'FORM_1099_DIV', 'FORM_1099_MISC', 'FORM_1099_NEC'];
    return validTypes.includes(docType as TaxDocumentType);
  }

  private isModelNotFoundError(error: any): boolean {
    return error?.message?.includes('ModelNotFound') || 
           error?.message?.includes('Resource not found') ||
           error?.code === 'NotFound';
  }

  // Legacy methods for backward compatibility
  async extractDataFromDocument(documentPathOrBuffer: string | Buffer): Promise<any> {
    console.log('⚠️ [Azure DI] Using legacy extractDataFromDocument method - consider using extractTaxDocumentData instead');
    
    try {
      // Try to determine document type from OCR first
      const documentBuffer = typeof documentPathOrBuffer === 'string' 
        ? await readFile(documentPathOrBuffer)
        : documentPathOrBuffer;
      
      // Use OCR to analyze document type
      const ocrPoller = await this.client.beginAnalyzeDocument('prebuilt-read', documentBuffer);
      const ocrResult = await ocrPoller.pollUntilDone();
      
      const documentType = this.analyzeDocumentTypeFromOCR(ocrResult.content || '');
      
      if (this.isValidTaxDocumentType(documentType)) {
        console.log(`🔍 [Azure DI] Detected document type: ${documentType}, using tax extraction`);
        return await this.extractTaxDocumentData(documentPathOrBuffer, documentType as TaxDocumentType);
      } else {
        console.log('🔍 [Azure DI] Unknown document type, using generic extraction');
        return {
          fullText: ocrResult.content || '',
          extractedData: {}
        };
      }
    } catch (error: any) {
      console.error('❌ [Azure DI] Legacy extraction error:', error);
      throw new Error(`Document extraction failed: ${error?.message || 'Unknown error'}`);
    }
  }

  async extractW2(documentPathOrBuffer: string | Buffer): Promise<W2Data> {
    return await this.extractTaxDocumentData(documentPathOrBuffer, 'W2') as W2Data;
  }

  async extract1099Int(documentPathOrBuffer: string | Buffer): Promise<Form1099IntData> {
    return await this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_INT') as Form1099IntData;
  }

  async extract1099Div(documentPathOrBuffer: string | Buffer): Promise<Form1099DivData> {
    return await this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_DIV') as Form1099DivData;
  }

  async extract1099Misc(documentPathOrBuffer: string | Buffer): Promise<Form1099MiscData> {
    return await this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_MISC') as Form1099MiscData;
  }
}

// Export the service and interfaces
export default AzureDocumentIntelligenceService;

// Factory function to create and configure the Azure Document Intelligence service
export function getAzureDocumentIntelligenceService(): AzureDocumentIntelligenceService {
  const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;
  const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY;
  
  if (!endpoint || !apiKey) {
    throw new Error(
      'Azure Document Intelligence configuration missing. Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_API_KEY environment variables.'
    );
  }
  
  const config: AzureDocumentIntelligenceConfig = {
    endpoint,
    apiKey
  };
  
  return new AzureDocumentIntelligenceService(config);
}
