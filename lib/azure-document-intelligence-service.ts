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

// Type alias for extracted field data (used by route handlers)
export type ExtractedFieldData = TaxDocumentData;

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
    
    // Extract form fields
    if (result.documents && result.documents.length > 0) {
      const document = result.documents[0];
      
      if (document.fields) {
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
    
    return baseData as TaxDocumentData;
  }

  /**
   * Extract tax document fields using OCR fallback
   */
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

  /**
   * Process W2 fields from structured analysis
   */
  private processW2Fields(fields: any, baseData: BaseTaxDocument): W2Data {
    const w2Data: W2Data = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing W2 fields from structured analysis...');
    
    // Enhanced W2 field mappings based on Azure Document Intelligence schema
    const w2FieldMappings = {
      // Employee information - multiple possible field names
      'Employee.Name': 'employeeName',
      'Employee.SSN': 'employeeSSN',
      'Employee.Address': 'employeeAddress',
      'EmployeeName': 'employeeName',
      'EmployeeSSN': 'employeeSSN',
      'EmployeeAddress': 'employeeAddress',
      'Employee': 'employeeName',
      'EmployeeInformation.Name': 'employeeName',
      'EmployeeInformation.SSN': 'employeeSSN',
      'EmployeeInformation.Address': 'employeeAddress',
      
      // Employer information
      'Employer.Name': 'employerName',
      'Employer.EIN': 'employerEIN',
      'Employer.Address': 'employerAddress',
      'EmployerName': 'employerName',
      'EmployerEIN': 'employerEIN',
      'EmployerAddress': 'employerAddress',
      'EmployerInformation.Name': 'employerName',
      'EmployerInformation.EIN': 'employerEIN',
      'EmployerInformation.Address': 'employerAddress',
      
      // Box 1-6: Core wage and tax information
      'WagesAndTips': 'wages',
      'Wages': 'wages',
      'WagesTipsOtherComp': 'wages',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'SocialSecurityWages': 'socialSecurityWages',
      'SocialSecurityTaxWithheld': 'socialSecurityTaxWithheld',
      'MedicareWagesAndTips': 'medicareWages',
      'MedicareWages': 'medicareWages',
      'MedicareTaxWithheld': 'medicareTaxWithheld',
      
      // Box 7-11: Additional compensation
      'SocialSecurityTips': 'socialSecurityTips',
      'AllocatedTips': 'allocatedTips',
      'AdvanceEIC': 'advanceEIC',
      'AdvanceEICPayment': 'advanceEIC',
      'DependentCareBenefits': 'dependentCareBenefits',
      'NonqualifiedPlans': 'nonqualifiedPlans',
      'NonqualifiedDeferredComp': 'nonqualifiedPlans',
      
      // Box 12: Deferred compensation
      'DeferredCompensation': 'box12Raw',
      'Box12': 'box12Raw',
      'Box12Codes': 'box12Raw',
      
      // Box 14: Other
      'OtherTaxInfo': 'otherTaxInfo',
      'Box14': 'otherTaxInfo',
      'Other': 'otherTaxInfo',
      
      // Box 15-20: State and local information
      'StateEmployerID': 'stateEmployerID',
      'StateWagesTipsEtc': 'stateWages',
      'StateWages': 'stateWages',
      'StateIncomeTax': 'stateTaxWithheld',
      'StateTaxWithheld': 'stateTaxWithheld',
      'LocalWagesTipsEtc': 'localWages',
      'LocalWages': 'localWages',
      'LocalIncomeTax': 'localTaxWithheld',
      'LocalTaxWithheld': 'localTaxWithheld',
      'LocalityName': 'localityName',
      
      // Alternative box number field names
      'Box1': 'wages',
      'Box2': 'federalTaxWithheld',
      'Box3': 'socialSecurityWages',
      'Box4': 'socialSecurityTaxWithheld',
      'Box5': 'medicareWages',
      'Box6': 'medicareTaxWithheld',
      'Box7': 'socialSecurityTips',
      'Box8': 'allocatedTips',
      'Box9': 'advanceEIC',
      'Box10': 'dependentCareBenefits',
      'Box11': 'nonqualifiedPlans',
      'Box15': 'stateEmployerID',
      'Box16': 'stateWages',
      'Box17': 'stateTaxWithheld',
      'Box18': 'localWages',
      'Box19': 'localTaxWithheld',
      'Box20': 'localityName'
    };
    
    // Map fields with enhanced error handling
    let fieldsProcessed = 0;
    for (const [azureFieldName, mappedFieldName] of Object.entries(w2FieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        try {
          if (mappedFieldName === 'box12Raw' || mappedFieldName === 'otherTaxInfo' || 
              mappedFieldName === 'stateEmployerID' || mappedFieldName === 'localityName' ||
              mappedFieldName === 'employeeName' || mappedFieldName === 'employerName' ||
              mappedFieldName === 'employeeSSN' || mappedFieldName === 'employerEIN' ||
              mappedFieldName === 'employeeAddress' || mappedFieldName === 'employerAddress') {
            // Text fields
            (w2Data as any)[mappedFieldName] = String(value).trim();
          } else {
            // Numeric fields
            (w2Data as any)[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          }
          fieldsProcessed++;
          console.log(`✅ [Azure DI] Mapped ${azureFieldName} → ${mappedFieldName}: ${value}`);
        } catch (error) {
          console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} W2 fields from structured analysis`);
    
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
    
    console.log('🔍 [Azure DI] Processing 1099-INT fields from structured analysis...');
    
    const fieldMappings = {
      // Payer and recipient information - multiple possible field names
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerInformation.Name': 'payerName',
      'PayerInformation.TIN': 'payerTIN',
      'PayerInformation.Address': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'RecipientInformation.Name': 'recipientName',
      'RecipientInformation.TIN': 'recipientTIN',
      'RecipientInformation.Address': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Box 1-15 mappings with enhanced field names
      'InterestIncome': 'interestIncome',
      'Interest': 'interestIncome',
      'TotalInterestIncome': 'interestIncome',
      'EarlyWithdrawalPenalty': 'earlyWithdrawalPenalty',
      'EarlyWithdrawal': 'earlyWithdrawalPenalty',
      'InterestOnUSTreasuryObligations': 'interestOnUSavingsBonds',
      'InterestOnUSavingsBonds': 'interestOnUSavingsBonds',
      'USTreasuryInterest': 'interestOnUSavingsBonds',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'InvestmentExpenses': 'investmentExpenses',
      'ForeignTaxPaid': 'foreignTaxPaid',
      'ForeignTax': 'foreignTaxPaid',
      'ForeignCountry': 'foreignCountry',
      'TaxExemptInterest': 'taxExemptInterest',
      'ExemptInterest': 'taxExemptInterest',
      'SpecifiedPrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest',
      'PrivateActivityBondInterest': 'specifiedPrivateActivityBondInterest',
      'MarketDiscount': 'marketDiscount',
      'BondPremium': 'bondPremium',
      'StateTaxWithheld': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',
      'StateInterest': 'stateInterest',
      
      // Alternative field names with box numbers
      'Box1': 'interestIncome',
      'Box2': 'earlyWithdrawalPenalty',
      'Box3': 'interestOnUSavingsBonds',
      'Box4': 'federalTaxWithheld',
      'Box5': 'investmentExpenses',
      'Box6': 'foreignTaxPaid',
      'Box7': 'foreignCountry',
      'Box8': 'taxExemptInterest',
      'Box9': 'specifiedPrivateActivityBondInterest',
      'Box10': 'marketDiscount',
      'Box11': 'bondPremium',
      'Box13': 'stateTaxWithheld',
      'Box14': 'statePayerNumber',
      'Box15': 'stateInterest'
    };
    
    // Map fields with enhanced error handling
    let fieldsProcessed = 0;
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        try {
          if (mappedFieldName === 'foreignCountry' || mappedFieldName === 'statePayerNumber' || 
              mappedFieldName === 'accountNumber' || mappedFieldName === 'payerName' ||
              mappedFieldName === 'recipientName' || mappedFieldName === 'payerTIN' ||
              mappedFieldName === 'recipientTIN' || mappedFieldName === 'payerAddress' ||
              mappedFieldName === 'recipientAddress') {
            // Text fields
            (data as any)[mappedFieldName] = String(value).trim();
          } else {
            // Numeric fields
            (data as any)[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          }
          fieldsProcessed++;
          console.log(`✅ [Azure DI] Mapped ${azureFieldName} → ${mappedFieldName}: ${value}`);
        } catch (error) {
          console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} 1099-INT fields from structured analysis`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-DIV fields from structured analysis
   */
  private process1099DivFields(fields: any, baseData: BaseTaxDocument): Form1099DivData {
    const data: Form1099DivData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-DIV fields from structured analysis...');
    
    const fieldMappings = {
      // Payer and recipient information - multiple possible field names
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerInformation.Name': 'payerName',
      'PayerInformation.TIN': 'payerTIN',
      'PayerInformation.Address': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'RecipientInformation.Name': 'recipientName',
      'RecipientInformation.TIN': 'recipientTIN',
      'RecipientInformation.Address': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Dividend fields with enhanced mappings
      'OrdinaryDividends': 'ordinaryDividends',
      'Dividends': 'ordinaryDividends',
      'TotalOrdinaryDividends': 'ordinaryDividends',
      'QualifiedDividends': 'qualifiedDividends',
      'Qualified': 'qualifiedDividends',
      'TotalCapitalGainDistributions': 'totalCapitalGain',
      'CapitalGainDistributions': 'totalCapitalGain',
      'CapitalGain': 'totalCapitalGain',
      'UnrecapturedSection1250Gain': 'unrecapturedSection1250Gain',
      'Section1250Gain': 'unrecapturedSection1250Gain',
      'Section1202Gain': 'section1202Gain',
      'CollectiblesGain': 'collectiblesGain',
      'Section897OrdinaryDividends': 'section897OrdinaryDividends',
      'Section897CapitalGain': 'section897CapitalGain',
      'NondividendDistributions': 'nondividendDistributions',
      'NonDividend': 'nondividendDistributions',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'Section199ADividends': 'section199ADividends',
      'ExemptInterestDividends': 'exemptInterestDividends',
      'ExemptInterest': 'exemptInterestDividends',
      'ForeignTaxPaid': 'foreignTaxPaid',
      'ForeignTax': 'foreignTaxPaid',
      'ForeignCountry': 'foreignCountry',
      'CashLiquidationDistributions': 'cashLiquidationDistributions',
      'CashLiquidation': 'cashLiquidationDistributions',
      'NoncashLiquidationDistributions': 'noncashLiquidationDistributions',
      'NoncashLiquidation': 'noncashLiquidationDistributions',
      'FATCAFilingRequirement': 'fatcaFilingRequirement',
      'FATCA': 'fatcaFilingRequirement',
      'InvestmentExpenses': 'investmentExpenses',
      'StateTaxWithheld': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',
      'StateIncome': 'stateIncome',
      
      // Alternative field names with box numbers
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
      'Box13': 'investmentExpenses'
    };
    
    // Map fields with enhanced error handling
    let fieldsProcessed = 0;
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        try {
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
            (data as any)[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          }
          fieldsProcessed++;
          console.log(`✅ [Azure DI] Mapped ${azureFieldName} → ${mappedFieldName}: ${value}`);
        } catch (error) {
          console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} 1099-DIV fields from structured analysis`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-MISC fields from structured analysis
   */
  private process1099MiscFields(fields: any, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Processing 1099-MISC fields from structured analysis...');
    
    const fieldMappings = {
      // Payer and recipient information - multiple possible field names
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerInformation.Name': 'payerName',
      'PayerInformation.TIN': 'payerTIN',
      'PayerInformation.Address': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'RecipientInformation.Name': 'recipientName',
      'RecipientInformation.TIN': 'recipientTIN',
      'RecipientInformation.Address': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // Box 1-18 mappings with enhanced field names
      'Rents': 'rents',
      'RentIncome': 'rents',
      'Royalties': 'royalties',
      'RoyaltyIncome': 'royalties',
      'OtherIncome': 'otherIncome',
      'Other': 'otherIncome',
      'MiscellaneousIncome': 'otherIncome',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      'FishingBoatProceeds': 'fishingBoatProceeds',
      'FishingProceeds': 'fishingBoatProceeds',
      'MedicalAndHealthCarePayments': 'medicalHealthPayments',
      'MedicalPayments': 'medicalHealthPayments',
      'HealthCarePayments': 'medicalHealthPayments',
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'NonEmployeeComp': 'nonemployeeCompensation',
      'SubstitutePayments': 'substitutePayments',
      'SubstitutePaymentsInLieuOfDividends': 'substitutePayments',
      'CropInsuranceProceeds': 'cropInsuranceProceeds',
      'CropInsurance': 'cropInsuranceProceeds',
      'GrossProceedsPaidToAttorney': 'grossProceedsAttorney',
      'AttorneyProceeds': 'grossProceedsAttorney',
      'FishPurchasedForResale': 'fishPurchases',
      'FishPurchases': 'fishPurchases',
      'Section409ADeferrals': 'section409ADeferrals',
      'ExcessGoldenParachutePayments': 'excessGoldenParachutePayments',
      'GoldenParachute': 'excessGoldenParachutePayments',
      'NonqualifiedDeferredCompensation': 'nonqualifiedDeferredCompensation',
      'DeferredCompensation': 'nonqualifiedDeferredCompensation',
      'Section409AIncome': 'section409AIncome',
      'StateTaxWithheld': 'stateTaxWithheld',
      'StatePayerNumber': 'statePayerNumber',
      'StateIncome': 'stateIncome',
      
      // Alternative field names with box numbers
      'Box1': 'rents',
      'Box2': 'royalties',
      'Box3': 'otherIncome',
      'Box4': 'federalTaxWithheld',
      'Box5': 'fishingBoatProceeds',
      'Box6': 'medicalHealthPayments',
      'Box7': 'nonemployeeCompensation',
      'Box8': 'substitutePayments',
      'Box9': 'cropInsuranceProceeds',
      'Box10': 'grossProceedsAttorney',
      'Box11': 'fishPurchases',
      'Box12': 'section409ADeferrals',
      'Box13': 'excessGoldenParachutePayments',
      'Box14': 'nonqualifiedDeferredCompensation',
      'Box15a': 'section409AIncome',
      'Box16': 'stateTaxWithheld',
      'Box17': 'statePayerNumber',
      'Box18': 'stateIncome'
    };
    
    // Map fields with enhanced error handling
    let fieldsProcessed = 0;
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        try {
          if (mappedFieldName === 'statePayerNumber' || mappedFieldName === 'accountNumber' ||
              mappedFieldName === 'payerName' || mappedFieldName === 'recipientName' ||
              mappedFieldName === 'payerTIN' || mappedFieldName === 'recipientTIN' ||
              mappedFieldName === 'payerAddress' || mappedFieldName === 'recipientAddress') {
            // Text fields
            (data as any)[mappedFieldName] = String(value).trim();
          } else {
            // Numeric fields
            (data as any)[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          }
          fieldsProcessed++;
          console.log(`✅ [Azure DI] Mapped ${azureFieldName} → ${mappedFieldName}: ${value}`);
        } catch (error) {
          console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} 1099-MISC fields from structured analysis`);
    
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
    
    console.log('🔍 [Azure DI] Processing 1099-NEC fields from structured analysis...');
    
    const fieldMappings = {
      // Payer and recipient information
      'Payer.Name': 'payerName',
      'Payer.TIN': 'payerTIN',
      'Payer.Address': 'payerAddress',
      'PayerName': 'payerName',
      'PayerTIN': 'payerTIN',
      'PayerAddress': 'payerAddress',
      'PayerInformation.Name': 'payerName',
      'PayerInformation.TIN': 'payerTIN',
      'PayerInformation.Address': 'payerAddress',
      
      'Recipient.Name': 'recipientName',
      'Recipient.TIN': 'recipientTIN',
      'Recipient.Address': 'recipientAddress',
      'RecipientName': 'recipientName',
      'RecipientTIN': 'recipientTIN',
      'RecipientAddress': 'recipientAddress',
      'RecipientInformation.Name': 'recipientName',
      'RecipientInformation.TIN': 'recipientTIN',
      'RecipientInformation.Address': 'recipientAddress',
      
      'AccountNumber': 'accountNumber',
      'Account': 'accountNumber',
      
      // NEC specific fields
      'NonemployeeCompensation': 'nonemployeeCompensation',
      'NonEmployeeCompensation': 'nonemployeeCompensation',
      'NonEmployeeComp': 'nonemployeeCompensation',
      'FederalIncomeTaxWithheld': 'federalTaxWithheld',
      'FederalTaxWithheld': 'federalTaxWithheld',
      
      // Box number alternatives
      'Box1': 'nonemployeeCompensation',
      'Box4': 'federalTaxWithheld'
    };
    
    // Map fields with enhanced error handling
    let fieldsProcessed = 0;
    for (const [azureFieldName, mappedFieldName] of Object.entries(fieldMappings)) {
      if (fields[azureFieldName]?.value !== undefined) {
        const value = fields[azureFieldName].value;
        
        try {
          if (mappedFieldName === 'accountNumber' || mappedFieldName === 'payerName' ||
              mappedFieldName === 'recipientName' || mappedFieldName === 'payerTIN' ||
              mappedFieldName === 'recipientTIN' || mappedFieldName === 'payerAddress' ||
              mappedFieldName === 'recipientAddress') {
            // Text fields
            (data as any)[mappedFieldName] = String(value).trim();
          } else {
            // Numeric fields
            (data as any)[mappedFieldName] = typeof value === 'number' ? value : this.parseAmount(value);
          }
          fieldsProcessed++;
          console.log(`✅ [Azure DI] Mapped ${azureFieldName} → ${mappedFieldName}: ${value}`);
        } catch (error) {
          console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} 1099-NEC fields from structured analysis`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process generic tax fields
   */
  private processGenericTaxFields(fields: any, baseData: BaseTaxDocument): TaxDocumentData {
    const data = { ...baseData } as any;
    
    console.log('🔍 [Azure DI] Processing generic tax fields from structured analysis...');
    
    // Process all available fields
    let fieldsProcessed = 0;
    for (const [fieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && 'value' in fieldData) {
        const value = (fieldData as any).value;
        if (value !== undefined && value !== null && value !== '') {
          try {
            data[fieldName] = typeof value === 'number' ? value : this.parseAmount(value);
            fieldsProcessed++;
            console.log(`✅ [Azure DI] Mapped generic field ${fieldName}: ${value}`);
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing generic field ${fieldName}:`, error);
          }
        }
      }
    }
    
    console.log(`✅ [Azure DI] Processed ${fieldsProcessed} generic tax fields from structured analysis`);
    
    return data;
  }

  // Enhanced OCR-based extraction methods
  private extractW2FieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): W2Data {
    const w2Data: W2Data = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting W2 fields using enhanced OCR patterns...');
    
    // Enhanced patterns for W2 fields
    const fieldPatterns = [
      // Box 1: Wages, tips, other compensation
      { 
        field: 'wages', 
        patterns: [
          /(?:box\s*1|wages.*tips.*other.*compensation|wages.*tips.*compensation)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1\s*wages.*tips.*other.*compensation[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /wages[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 2: Federal income tax withheld
      { 
        field: 'federalTaxWithheld', 
        patterns: [
          /(?:box\s*2|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /2\s*federal.*income.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /federal.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 3: Social security wages
      { 
        field: 'socialSecurityWages', 
        patterns: [
          /(?:box\s*3|social.*security.*wages)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /3\s*social.*security.*wages[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 4: Social security tax withheld
      { 
        field: 'socialSecurityTaxWithheld', 
        patterns: [
          /(?:box\s*4|social.*security.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /4\s*social.*security.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 5: Medicare wages and tips
      { 
        field: 'medicareWages', 
        patterns: [
          /(?:box\s*5|medicare.*wages.*tips|medicare.*wages)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /5\s*medicare.*wages.*tips[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 6: Medicare tax withheld
      { 
        field: 'medicareTaxWithheld', 
        patterns: [
          /(?:box\s*6|medicare.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /6\s*medicare.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 7: Social security tips
      { 
        field: 'socialSecurityTips', 
        patterns: [
          /(?:box\s*7|social.*security.*tips)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /7\s*social.*security.*tips[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 8: Allocated tips
      { 
        field: 'allocatedTips', 
        patterns: [
          /(?:box\s*8|allocated.*tips)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /8\s*allocated.*tips[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 10: Dependent care benefits
      { 
        field: 'dependentCareBenefits', 
        patterns: [
          /(?:box\s*10|dependent.*care.*benefits)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /10\s*dependent.*care.*benefits[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      }
    ];
    
    // Extract numeric fields using patterns
    let fieldsExtracted = 0;
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (w2Data as any)[field] = amount;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
            break;
          }
        }
      }
    }
    
    // Extract Box 12 codes
    const box12Pattern = /(?:box\s*12|12)[:\s]*([A-Z]\s*\$?\d+(?:\.\d{2})?(?:\s*[A-Z]\s*\$?\d+(?:\.\d{2})?)*)/i;
    const box12Match = ocrText.match(box12Pattern);
    if (box12Match) {
      w2Data.box12Raw = box12Match[1];
      const box12Codes = this.parseW2Box12Codes(box12Match[1]);
      if (box12Codes.length > 0) {
        w2Data.box12Codes = box12Codes;
        fieldsExtracted++;
        console.log(`✅ [Azure DI] Extracted Box 12 codes from OCR:`, box12Codes);
      }
    }
    
    // Extract personal info using enhanced method
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      w2Data.employeeName = personalInfo.name;
      fieldsExtracted++;
      console.log(`✅ [Azure DI] Extracted employee name from OCR: ${personalInfo.name}`);
    }
    if (personalInfo.ssn) {
      w2Data.employeeSSN = personalInfo.ssn;
      fieldsExtracted++;
      console.log(`✅ [Azure DI] Extracted employee SSN from OCR: ${personalInfo.ssn}`);
    }
    if (personalInfo.address) {
      w2Data.employeeAddress = personalInfo.address;
      fieldsExtracted++;
      console.log(`✅ [Azure DI] Extracted employee address from OCR: ${personalInfo.address}`);
    }
    if (personalInfo.employerName) {
      w2Data.employerName = personalInfo.employerName;
      fieldsExtracted++;
      console.log(`✅ [Azure DI] Extracted employer name from OCR: ${personalInfo.employerName}`);
    }
    if (personalInfo.employerEIN) {
      w2Data.employerEIN = personalInfo.employerEIN;
      fieldsExtracted++;
      console.log(`✅ [Azure DI] Extracted employer EIN from OCR: ${personalInfo.employerEIN}`);
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} W2 fields using OCR`);
    
    return w2Data;
  }

  private extract1099IntFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099IntData {
    const data: Form1099IntData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-INT fields using enhanced OCR patterns...');
    
    // Enhanced patterns for 1099-INT fields
    const fieldPatterns = [
      // Box 1: Interest income
      { 
        field: 'interestIncome', 
        patterns: [
          /(?:box\s*1|interest.*income|total.*interest)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1\s*interest.*income[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 2: Early withdrawal penalty
      { 
        field: 'earlyWithdrawalPenalty', 
        patterns: [
          /(?:box\s*2|early.*withdrawal.*penalty)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /2\s*early.*withdrawal.*penalty[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 3: Interest on U.S. Savings Bonds
      { 
        field: 'interestOnUSavingsBonds', 
        patterns: [
          /(?:box\s*3|interest.*u\.?s\.?.*savings.*bonds|interest.*treasury)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /3\s*interest.*u\.?s\.?.*savings.*bonds[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 4: Federal income tax withheld
      { 
        field: 'federalTaxWithheld', 
        patterns: [
          /(?:box\s*4|federal.*income.*tax.*withheld|federal.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /4\s*federal.*income.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 6: Foreign tax paid
      { 
        field: 'foreignTaxPaid', 
        patterns: [
          /(?:box\s*6|foreign.*tax.*paid)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /6\s*foreign.*tax.*paid[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 8: Tax-exempt interest
      { 
        field: 'taxExemptInterest', 
        patterns: [
          /(?:box\s*8|tax.*exempt.*interest)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /8\s*tax.*exempt.*interest[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      }
    ];
    
    // Extract numeric fields using patterns
    let fieldsExtracted = 0;
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
            break;
          }
        }
      }
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      data.recipientName = personalInfo.name;
      fieldsExtracted++;
    }
    if (personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      fieldsExtracted++;
    }
    if (personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      fieldsExtracted++;
    }
    if (personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      fieldsExtracted++;
    }
    if (personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      fieldsExtracted++;
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-INT fields using OCR`);
    
    return data;
  }

  private extract1099DivFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099DivData {
    const data: Form1099DivData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-DIV fields using enhanced OCR patterns...');
    
    // Enhanced patterns for 1099-DIV fields
    const fieldPatterns = [
      // Box 1a: Total ordinary dividends
      { 
        field: 'ordinaryDividends', 
        patterns: [
          /(?:box\s*1a|total.*ordinary.*dividends|ordinary.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1a\s*total.*ordinary.*dividends[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 1b: Qualified dividends
      { 
        field: 'qualifiedDividends', 
        patterns: [
          /(?:box\s*1b|qualified.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1b\s*qualified.*dividends[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 2a: Total capital gain distributions
      { 
        field: 'totalCapitalGain', 
        patterns: [
          /(?:box\s*2a|total.*capital.*gain.*distributions|capital.*gain.*distributions)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /2a\s*total.*capital.*gain.*distributions[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 3: Nondividend distributions
      { 
        field: 'nondividendDistributions', 
        patterns: [
          /(?:box\s*3|nondividend.*distributions)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /3\s*nondividend.*distributions[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 4: Federal income tax withheld
      { 
        field: 'federalTaxWithheld', 
        patterns: [
          /(?:box\s*4|federal.*income.*tax.*withheld|federal.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /4\s*federal.*income.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 7: Foreign tax paid
      { 
        field: 'foreignTaxPaid', 
        patterns: [
          /(?:box\s*7|foreign.*tax.*paid)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /7\s*foreign.*tax.*paid[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      }
    ];
    
    // Extract numeric fields using patterns
    let fieldsExtracted = 0;
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
            break;
          }
        }
      }
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      data.recipientName = personalInfo.name;
      fieldsExtracted++;
    }
    if (personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      fieldsExtracted++;
    }
    if (personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      fieldsExtracted++;
    }
    if (personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      fieldsExtracted++;
    }
    if (personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      fieldsExtracted++;
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-DIV fields using OCR`);
    
    return data;
  }

  private extract1099MiscFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-MISC fields using enhanced OCR patterns...');
    
    // Enhanced patterns for 1099-MISC fields
    const fieldPatterns = [
      // Box 1: Rents
      { 
        field: 'rents', 
        patterns: [
          /(?:box\s*1|rents)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1\s*rents[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 2: Royalties
      { 
        field: 'royalties', 
        patterns: [
          /(?:box\s*2|royalties)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /2\s*royalties[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 3: Other income
      { 
        field: 'otherIncome', 
        patterns: [
          /(?:box\s*3|other.*income)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /3\s*other.*income[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 4: Federal income tax withheld
      { 
        field: 'federalTaxWithheld', 
        patterns: [
          /(?:box\s*4|federal.*income.*tax.*withheld|federal.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /4\s*federal.*income.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 5: Fishing boat proceeds
      { 
        field: 'fishingBoatProceeds', 
        patterns: [
          /(?:box\s*5|fishing.*boat.*proceeds)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /5\s*fishing.*boat.*proceeds[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 6: Medical and health care payments
      { 
        field: 'medicalHealthPayments', 
        patterns: [
          /(?:box\s*6|medical.*health.*care.*payments|medical.*payments)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /6\s*medical.*health.*care.*payments[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 7: Nonemployee compensation (deprecated but still used)
      { 
        field: 'nonemployeeCompensation', 
        patterns: [
          /(?:box\s*7|nonemployee.*compensation)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /7\s*nonemployee.*compensation[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 10: Gross proceeds paid to an attorney
      { 
        field: 'grossProceedsAttorney', 
        patterns: [
          /(?:box\s*10|gross.*proceeds.*paid.*attorney|attorney.*proceeds)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /10\s*gross.*proceeds.*paid.*attorney[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      }
    ];
    
    // Extract numeric fields using patterns
    let fieldsExtracted = 0;
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
            break;
          }
        }
      }
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      data.recipientName = personalInfo.name;
      fieldsExtracted++;
    }
    if (personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      fieldsExtracted++;
    }
    if (personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      fieldsExtracted++;
    }
    if (personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      fieldsExtracted++;
    }
    if (personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      fieldsExtracted++;
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-MISC fields using OCR`);
    
    return data;
  }

  private extract1099NecFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099MiscData {
    const data: Form1099MiscData = { ...baseData };
    
    console.log('🔍 [Azure DI] Extracting 1099-NEC fields using enhanced OCR patterns...');
    
    // Enhanced patterns for 1099-NEC fields
    const fieldPatterns = [
      // Box 1: Nonemployee compensation
      { 
        field: 'nonemployeeCompensation', 
        patterns: [
          /(?:box\s*1|nonemployee.*compensation)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /1\s*nonemployee.*compensation[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      },
      // Box 4: Federal income tax withheld
      { 
        field: 'federalTaxWithheld', 
        patterns: [
          /(?:box\s*4|federal.*income.*tax.*withheld|federal.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i,
          /4\s*federal.*income.*tax.*withheld[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i
        ]
      }
    ];
    
    // Extract numeric fields using patterns
    let fieldsExtracted = 0;
    for (const { field, patterns } of fieldPatterns) {
      for (const pattern of patterns) {
        const match = ocrText.match(pattern);
        if (match && match[1]) {
          const amount = this.parseAmount(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted ${field} from OCR: $${amount}`);
            break;
          }
        }
      }
    }
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) {
      data.recipientName = personalInfo.name;
      fieldsExtracted++;
    }
    if (personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      fieldsExtracted++;
    }
    if (personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      fieldsExtracted++;
    }
    if (personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      fieldsExtracted++;
    }
    if (personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      fieldsExtracted++;
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} 1099-NEC fields using OCR`);
    
    return data;
  }

  private extractGenericTaxFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): TaxDocumentData {
    const data = { ...baseData } as any;
    
    console.log('🔍 [Azure DI] Extracting generic tax fields using OCR patterns...');
    
    // Extract personal info
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    let fieldsExtracted = 0;
    
    if (personalInfo.name) {
      data.recipientName = personalInfo.name;
      fieldsExtracted++;
    }
    if (personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      fieldsExtracted++;
    }
    if (personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      fieldsExtracted++;
    }
    if (personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      fieldsExtracted++;
    }
    if (personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      fieldsExtracted++;
    }
    
    console.log(`✅ [Azure DI] Extracted ${fieldsExtracted} generic tax fields using OCR`);
    
    return data;
  }

  // Enhanced utility methods
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
    // Enhanced pattern to handle various formats
    const codePattern = /([A-Z]{1,2})\s*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/g;
    let match;
    
    while ((match = codePattern.exec(box12String)) !== null) {
      const code = match[1];
      const amount = parseFloat(match[2].replace(/,/g, ''));
      
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
    
    // Enhanced OCR fallback for checkboxes
    if (ocrText && (!checkboxes.retirementPlan && !checkboxes.thirdPartySickPay && !checkboxes.statutoryEmployee)) {
      const text = ocrText.toLowerCase();
      
      // More comprehensive patterns for checkbox detection
      checkboxes.retirementPlan = /(?:retirement\s+plan|13.*retirement)[:\s]*(?:x|✓|checked|yes|\[x\])/i.test(text);
      checkboxes.thirdPartySickPay = /(?:third.party\s+sick\s+pay|13.*third.party)[:\s]*(?:x|✓|checked|yes|\[x\])/i.test(text);
      checkboxes.statutoryEmployee = /(?:statutory\s+employee|13.*statutory)[:\s]*(?:x|✓|checked|yes|\[x\])/i.test(text);
    }
    
    return checkboxes;
  }

  private extractPersonalInfoFromOCR(ocrText: string, targetEmployeeName?: string): any {
    const personalInfo: any = {};
    
    console.log('🔍 [Azure DI] Extracting personal information from OCR...');
    
    // Enhanced SSN/TIN patterns
    const ssnPattern = /\b(\d{3}[-\s]?\d{2}[-\s]?\d{4})\b/g;
    const ssnMatches = Array.from(ocrText.matchAll(ssnPattern));
    if (ssnMatches.length > 0) {
      personalInfo.ssn = ssnMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.tin = personalInfo.ssn;
      console.log('✅ [Azure DI] Found SSN/TIN pattern');
    }
    
    // Enhanced EIN patterns
    const einPattern = /\b(\d{2}[-\s]?\d{7})\b/g;
    const einMatches = Array.from(ocrText.matchAll(einPattern));
    if (einMatches.length > 0) {
      personalInfo.employerEIN = einMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.payerTIN = personalInfo.employerEIN;
      console.log('✅ [Azure DI] Found EIN pattern');
    }
    
    // Enhanced name extraction
    const lines = ocrText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
    const namePatterns = [
      /^([A-Z][a-z]+(?:\s+[A-Z][a-z]*)*(?:\s+[A-Z][a-z]+)+)$/,  // Full name pattern
      /^([A-Z][A-Z\s]+)$/,  // All caps name pattern
      /employee[:\s]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/i,  // Employee: Name pattern
      /recipient[:\s]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)/i,  // Recipient: Name pattern
    ];
    
    for (const line of lines) {
      if (line.length > 50 || line.length < 5) continue; // Skip very long or short lines
      
      for (const pattern of namePatterns) {
        const match = line.match(pattern);
        if (match && match[1]) {
          const name = match[1].trim();
          if (!personalInfo.name && this.isValidName(name)) {
            personalInfo.name = name;
            personalInfo.employerName = name; // Could be either
            personalInfo.payerName = name;
            console.log('✅ [Azure DI] Found name pattern:', name);
            break;
          }
        }
      }
      
      if (personalInfo.name) break;
    }
    
    // Enhanced address extraction
    const addressPatterns = [
      /\d+\s+[A-Za-z\s]+(?:street|st|avenue|ave|road|rd|drive|dr|lane|ln|way|blvd|boulevard)/i,
      /\d+\s+[A-Za-z\s]+\s+[A-Z]{2}\s+\d{5}/,  // Street City State ZIP
      /[A-Za-z\s]+,\s*[A-Z]{2}\s+\d{5}/  // City, State ZIP
    ];
    
    for (const line of lines) {
      for (const pattern of addressPatterns) {
        if (pattern.test(line)) {
          if (!personalInfo.address) {
            personalInfo.address = line;
            console.log('✅ [Azure DI] Found address pattern:', line);
            break;
          }
        }
      }
      
      if (personalInfo.address) break;
    }
    
    return personalInfo;
  }

  private isValidName(name: string): boolean {
    // Basic validation for names
    if (name.length < 3 || name.length > 50) return false;
    if (!/^[A-Za-z\s\-'\.]+$/.test(name)) return false;
    if (name.split(' ').length < 2) return false; // At least first and last name
    
    // Exclude common non-name patterns
    const excludePatterns = [
      /\d/,  // Contains numbers
      /^(box|form|tax|income|federal|state|local|employer|employee|payer|recipient)$/i,
      /^(w-?2|1099|misc|div|int|nec)$/i
    ];
    
    for (const pattern of excludePatterns) {
      if (pattern.test(name)) return false;
    }
    
    return true;
  }

  private extractAddressParts(fullAddress: string, ocrText: string): any {
    const addressParts: any = {};
    
    // Enhanced ZIP code extraction
    const zipMatch = fullAddress.match(/\b(\d{5}(?:-\d{4})?)\b/);
    if (zipMatch) {
      addressParts.zipCode = zipMatch[1];
    }
    
    // Enhanced state extraction (2-letter abbreviation before ZIP)
    const stateMatch = fullAddress.match(/\b([A-Z]{2})\s+\d{5}/);
    if (stateMatch) {
      addressParts.state = stateMatch[1];
    }
    
    // Enhanced city extraction (word(s) before state)
    const cityMatch = fullAddress.match(/([A-Za-z\s]+)\s+[A-Z]{2}\s+\d{5}/);
    if (cityMatch) {
      addressParts.city = cityMatch[1].trim();
    }
    
    // Enhanced street extraction (everything before city)
    const streetMatch = fullAddress.match(/^(.+?)(?:\s+[A-Za-z\s]+\s+[A-Z]{2}\s+\d{5})/);
    if (streetMatch) {
      addressParts.street = streetMatch[1].trim();
    }
    
    return addressParts;
  }

  private applyPersonalInfoOCRFallback(data: any, fullText?: string): void {
    if (!fullText) return;
    
    const personalInfo = this.extractPersonalInfoFromOCR(fullText);
    let infoApplied = 0;
    
    if (!data.recipientName && personalInfo.name) {
      data.recipientName = personalInfo.name;
      infoApplied++;
      console.log('✅ [Azure DI] Extracted recipient name from OCR:', data.recipientName);
    }
    
    if (!data.recipientTIN && personalInfo.tin) {
      data.recipientTIN = personalInfo.tin;
      infoApplied++;
      console.log('✅ [Azure DI] Extracted recipient TIN from OCR:', data.recipientTIN);
    }
    
    if (!data.recipientAddress && personalInfo.address) {
      data.recipientAddress = personalInfo.address;
      infoApplied++;
      console.log('✅ [Azure DI] Extracted recipient address from OCR:', data.recipientAddress);
    }
    
    if (!data.payerName && personalInfo.payerName) {
      data.payerName = personalInfo.payerName;
      infoApplied++;
      console.log('✅ [Azure DI] Extracted payer name from OCR:', data.payerName);
    }
    
    if (!data.payerTIN && personalInfo.payerTIN) {
      data.payerTIN = personalInfo.payerTIN;
      infoApplied++;
      console.log('✅ [Azure DI] Extracted payer TIN from OCR:', data.payerTIN);
    }
    
    if (infoApplied > 0) {
      console.log(`✅ [Azure DI] Applied ${infoApplied} personal info fields from OCR fallback`);
    }
  }

  private validateAndCorrect1099MiscFields(structuredData: Form1099MiscData, ocrText: string): Form1099MiscData {
    console.log('🔍 [Azure DI] Validating 1099-MISC field mappings...');
    
    // Extract data using OCR as ground truth
    const ocrData = this.extract1099MiscFieldsFromOCR(ocrText, { fullText: ocrText });
    
    const correctedData = { ...structuredData };
    let correctionsMade = 0;
    
    // Define validation rules for critical fields
    const criticalFields = ['otherIncome', 'fishingBoatProceeds', 'medicalHealthPayments', 'rents', 'royalties', 'federalTaxWithheld', 'nonemployeeCompensation', 'grossProceedsAttorney'];
    
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
    
    // Enhanced document type detection patterns
    const documentTypePatterns = [
      { type: 'W2', patterns: [
        /wage\s+and\s+tax\s+statement/,
        /form\s+w-?2/,
        /w-?2/,
        /wages.*tips.*other.*compensation/,
        /social\s+security\s+wages/,
        /medicare\s+wages/
      ]},
      { type: 'FORM_1099_INT', patterns: [
        /1099-?int/,
        /interest\s+income/,
        /form\s+1099.*int/,
        /early\s+withdrawal\s+penalty/,
        /tax-?exempt\s+interest/
      ]},
      { type: 'FORM_1099_DIV', patterns: [
        /1099-?div/,
        /dividends?\s+and\s+distributions?/,
        /form\s+1099.*div/,
        /ordinary\s+dividends?/,
        /qualified\s+dividends?/,
        /capital\s+gain\s+distributions?/
      ]},
      { type: 'FORM_1099_MISC', patterns: [
        /1099-?misc/,
        /miscellaneous\s+income/,
        /form\s+1099.*misc/,
        /rents/,
        /royalties/,
        /fishing\s+boat\s+proceeds/,
        /medical.*health.*care.*payments/
      ]},
      { type: 'FORM_1099_NEC', patterns: [
        /1099-?nec/,
        /nonemployee\s+compensation/,
        /form\s+1099.*nec/
      ]}
    ];
    
    // Score each document type based on pattern matches
    const scores: { [key: string]: number } = {};
    
    for (const { type, patterns } of documentTypePatterns) {
      scores[type] = 0;
      for (const pattern of patterns) {
        if (pattern.test(text)) {
          scores[type]++;
        }
      }
    }
    
    // Find the type with the highest score
    let bestType = 'UNKNOWN';
    let bestScore = 0;
    
    for (const [type, score] of Object.entries(scores)) {
      if (score > bestScore) {
        bestScore = score;
        bestType = type;
      }
    }
    
    if (bestScore > 0) {
      console.log(`✅ [Azure DI] Document type detected: ${bestType} (score: ${bestScore})`);
      return bestType;
    }
    
    console.log('⚠️ [Azure DI] Could not determine document type from OCR');
    return 'UNKNOWN';
  }

  private isValidTaxDocumentType(docType: string): boolean {
    const validTypes: TaxDocumentType[] = ['W2', 'FORM_1099_INT', 'FORM_1099_DIV', 'FORM_1099_MISC', 'FORM_1099_NEC'];
    return validTypes.includes(docType as TaxDocumentType);
  }

  private isModelNotFoundError(error: any): boolean {
    return error?.message?.includes('ModelNotFound') || 
           error?.message?.includes('Resource not found') ||
           error?.code === 'NotFound' ||
           error?.status === 404;
  }

  // Convenience methods for specific document types
  async extractW2(documentPathOrBuffer: string | Buffer): Promise<W2Data> {
    return this.extractTaxDocumentData(documentPathOrBuffer, 'W2') as Promise<W2Data>;
  }

  async extract1099Div(documentPathOrBuffer: string | Buffer): Promise<Form1099DivData> {
    return this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_DIV') as Promise<Form1099DivData>;
  }

  async extract1099Int(documentPathOrBuffer: string | Buffer): Promise<Form1099IntData> {
    return this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_INT') as Promise<Form1099IntData>;
  }

  async extract1099Misc(documentPathOrBuffer: string | Buffer): Promise<Form1099MiscData> {
    return this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_MISC') as Promise<Form1099MiscData>;
  }

  async extract1099Nec(documentPathOrBuffer: string | Buffer): Promise<Form1099MiscData> {
    return this.extractTaxDocumentData(documentPathOrBuffer, 'FORM_1099_NEC') as Promise<Form1099MiscData>;
  }
}

// Export the service and interfaces
export default AzureDocumentIntelligenceService;
