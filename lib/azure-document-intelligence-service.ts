
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
        console.log('🔍 [Azure DI] Raw result structure:', JSON.stringify({
          documentsCount: result.documents?.length || 0,
          hasKeyValuePairs: !!result.keyValuePairs,
          contentLength: result.content?.length || 0
        }));
        
        // Log field names for debugging
        if (result.documents && result.documents.length > 0 && result.documents[0].fields) {
          const fieldNames = Object.keys(result.documents[0].fields);
          console.log('🔍 [Azure DI] Available field names:', fieldNames);
        }
        
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
   * Process W2 fields from structured analysis - COMPLETE IMPLEMENTATION
   */
  private processW2Fields(fields: any, baseData: BaseTaxDocument): W2Data {
    console.log('🔍 [Azure DI] Processing W2 fields from structured analysis...');
    const w2Data: W2Data = { ...baseData };
    
    // Comprehensive W2 field mappings based on Azure Document Intelligence field names
    const w2FieldMappings: { [key: string]: { property: keyof W2Data; type: 'string' | 'number' | 'boolean' } } = {
      // Employee information - various possible field names
      'Employee': { property: 'employeeName', type: 'string' },
      'EmployeeName': { property: 'employeeName', type: 'string' },
      'Employee.Name': { property: 'employeeName', type: 'string' },
      'EmployeeSSN': { property: 'employeeSSN', type: 'string' },
      'Employee.SSN': { property: 'employeeSSN', type: 'string' },
      'Employee.SocialSecurityNumber': { property: 'employeeSSN', type: 'string' },
      'SocialSecurityNumber': { property: 'employeeSSN', type: 'string' },
      'EmployeeAddress': { property: 'employeeAddress', type: 'string' },
      'Employee.Address': { property: 'employeeAddress', type: 'string' },
      'streetAddress': { property: 'employeeAddressStreet', type: 'string' },
      'Employee.StreetAddress': { property: 'employeeAddressStreet', type: 'string' },
      'Employee.City': { property: 'employeeCity', type: 'string' },
      'Employee.State': { property: 'employeeState', type: 'string' },
      'Employee.ZipCode': { property: 'employeeZipCode', type: 'string' },
      'Employee.PostalCode': { property: 'employeeZipCode', type: 'string' },
      
      // Employer information
      'Employer': { property: 'employerName', type: 'string' },
      'EmployerName': { property: 'employerName', type: 'string' },
      'Employer.Name': { property: 'employerName', type: 'string' },
      'EmployerEIN': { property: 'employerEIN', type: 'string' },
      'Employer.EIN': { property: 'employerEIN', type: 'string' },
      'Employer.EmployerIdentificationNumber': { property: 'employerEIN', type: 'string' },
      'EmployerIdentificationNumber': { property: 'employerEIN', type: 'string' },
      'EmployerAddress': { property: 'employerAddress', type: 'string' },
      'Employer.Address': { property: 'employerAddress', type: 'string' },
      'EmployerStateIdNumber': { property: 'stateEmployerID', type: 'string' },
      'Employer.StateIdNumber': { property: 'stateEmployerID', type: 'string' },
      
      // Box 1-6: Core wage and tax information
      'WagesAndTips': { property: 'wages', type: 'number' },
      'Wages': { property: 'wages', type: 'number' },
      'WagesTipsOtherCompensation': { property: 'wages', type: 'number' },
      'Box1': { property: 'wages', type: 'number' },
      'W2Box1': { property: 'wages', type: 'number' },
      'FederalIncomeTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'FederalTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'Box2': { property: 'federalTaxWithheld', type: 'number' },
      'W2Box2': { property: 'federalTaxWithheld', type: 'number' },
      'SocialSecurityWages': { property: 'socialSecurityWages', type: 'number' },
      'Box3': { property: 'socialSecurityWages', type: 'number' },
      'W2Box3': { property: 'socialSecurityWages', type: 'number' },
      'SocialSecurityTaxWithheld': { property: 'socialSecurityTaxWithheld', type: 'number' },
      'Box4': { property: 'socialSecurityTaxWithheld', type: 'number' },
      'W2Box4': { property: 'socialSecurityTaxWithheld', type: 'number' },
      'MedicareWagesAndTips': { property: 'medicareWages', type: 'number' },
      'MedicareWages': { property: 'medicareWages', type: 'number' },
      'Box5': { property: 'medicareWages', type: 'number' },
      'W2Box5': { property: 'medicareWages', type: 'number' },
      'MedicareTaxWithheld': { property: 'medicareTaxWithheld', type: 'number' },
      'Box6': { property: 'medicareTaxWithheld', type: 'number' },
      'W2Box6': { property: 'medicareTaxWithheld', type: 'number' },
      
      // Box 7-11: Additional compensation
      'SocialSecurityTips': { property: 'socialSecurityTips', type: 'number' },
      'Box7': { property: 'socialSecurityTips', type: 'number' },
      'W2Box7': { property: 'socialSecurityTips', type: 'number' },
      'AllocatedTips': { property: 'allocatedTips', type: 'number' },
      'Box8': { property: 'allocatedTips', type: 'number' },
      'W2Box8': { property: 'allocatedTips', type: 'number' },
      'AdvanceEIC': { property: 'advanceEIC', type: 'number' },
      'AdvanceEICPayment': { property: 'advanceEIC', type: 'number' },
      'Box9': { property: 'advanceEIC', type: 'number' },
      'W2Box9': { property: 'advanceEIC', type: 'number' },
      'DependentCareBenefits': { property: 'dependentCareBenefits', type: 'number' },
      'Box10': { property: 'dependentCareBenefits', type: 'number' },
      'W2Box10': { property: 'dependentCareBenefits', type: 'number' },
      'NonqualifiedPlans': { property: 'nonqualifiedPlans', type: 'number' },
      'Box11': { property: 'nonqualifiedPlans', type: 'number' },
      'W2Box11': { property: 'nonqualifiedPlans', type: 'number' },
      
      // Box 12: Deferred compensation
      'DeferredCompensation': { property: 'box12Raw', type: 'string' },
      'Box12': { property: 'box12Raw', type: 'string' },
      'W2Box12': { property: 'box12Raw', type: 'string' },
      'Box12Codes': { property: 'box12Raw', type: 'string' },
      
      // Box 14: Other
      'OtherTaxInfo': { property: 'otherTaxInfo', type: 'string' },
      'Other': { property: 'otherTaxInfo', type: 'string' },
      'Box14': { property: 'otherTaxInfo', type: 'string' },
      'W2Box14': { property: 'otherTaxInfo', type: 'string' },
      
      // Box 15-20: State and local information
      'StateEmployerID': { property: 'stateEmployerID', type: 'string' },
      'StateEmployerIdNumber': { property: 'stateEmployerID', type: 'string' },
      'Box15': { property: 'stateEmployerID', type: 'string' },
      'W2Box15': { property: 'stateEmployerID', type: 'string' },
      'StateWagesTipsEtc': { property: 'stateWages', type: 'number' },
      'StateWages': { property: 'stateWages', type: 'number' },
      'Box16': { property: 'stateWages', type: 'number' },
      'W2Box16': { property: 'stateWages', type: 'number' },
      'StateIncomeTax': { property: 'stateTaxWithheld', type: 'number' },
      'StateTaxWithheld': { property: 'stateTaxWithheld', type: 'number' },
      'Box17': { property: 'stateTaxWithheld', type: 'number' },
      'W2Box17': { property: 'stateTaxWithheld', type: 'number' },
      'LocalWagesTipsEtc': { property: 'localWages', type: 'number' },
      'LocalWages': { property: 'localWages', type: 'number' },
      'Box18': { property: 'localWages', type: 'number' },
      'W2Box18': { property: 'localWages', type: 'number' },
      'LocalIncomeTax': { property: 'localTaxWithheld', type: 'number' },
      'LocalTaxWithheld': { property: 'localTaxWithheld', type: 'number' },
      'Box19': { property: 'localTaxWithheld', type: 'number' },
      'W2Box19': { property: 'localTaxWithheld', type: 'number' },
      'LocalityName': { property: 'localityName', type: 'string' },
      'Box20': { property: 'localityName', type: 'string' },
      'W2Box20': { property: 'localityName', type: 'string' }
    };
    
    let fieldsProcessed = 0;
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [azureFieldName, fieldData] of Object.entries(fields)) {
      fieldsProcessed++;
      
      if (fieldData && typeof fieldData === 'object' && fieldData.value !== undefined && fieldData.value !== null) {
        const mapping = w2FieldMappings[azureFieldName];
        
        if (mapping) {
          const { property, type } = mapping;
          let processedValue: any;
          
          try {
            switch (type) {
              case 'string':
                processedValue = this.extractStringValue(fieldData.value);
                break;
              case 'number':
                processedValue = this.extractNumericValue(fieldData.value);
                break;
              case 'boolean':
                processedValue = this.extractBooleanValue(fieldData.value);
                break;
              default:
                processedValue = fieldData.value;
            }
            
            if (processedValue !== null && processedValue !== undefined && 
                (type !== 'string' || processedValue.trim() !== '') &&
                (type !== 'number' || processedValue !== 0)) {
              (w2Data as any)[property] = processedValue;
              fieldsExtracted++;
              console.log(`✅ [Azure DI] Extracted ${azureFieldName} → ${property}: ${processedValue}`);
            }
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
          }
        } else {
          // Log unmapped fields for debugging
          console.log(`🔍 [Azure DI] Unmapped field: ${azureFieldName} = ${fieldData.value}`);
        }
      }
    }
    
    console.log(`📊 [Azure DI] W2 field processing summary: ${fieldsExtracted}/${fieldsProcessed} fields extracted`);
    
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
    
    // OCR fallback for missing critical personal info
    if ((!w2Data.employeeName || !w2Data.employeeSSN || !w2Data.employerName || !w2Data.employerEIN) && baseData.fullText) {
      console.log('🔍 [Azure DI] Some critical W2 info missing from structured fields, attempting OCR extraction...');
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
      if (addressParts.street) w2Data.employeeAddressStreet = addressParts.street;
      if (addressParts.city) w2Data.employeeCity = addressParts.city;
      if (addressParts.state) w2Data.employeeState = addressParts.state;
      if (addressParts.zipCode) w2Data.employeeZipCode = addressParts.zipCode;
    }
    
    return w2Data;
  }

  /**
   * Process 1099-INT fields from structured analysis - COMPLETE IMPLEMENTATION
   */
  private process1099IntFields(fields: any, baseData: BaseTaxDocument): Form1099IntData {
    console.log('🔍 [Azure DI] Processing 1099-INT fields from structured analysis...');
    const data: Form1099IntData = { ...baseData };
    
    const fieldMappings: { [key: string]: { property: keyof Form1099IntData; type: 'string' | 'number' | 'boolean' } } = {
      // Payer and recipient information
      'Payer': { property: 'payerName', type: 'string' },
      'PayerName': { property: 'payerName', type: 'string' },
      'Payer.Name': { property: 'payerName', type: 'string' },
      'PayerTIN': { property: 'payerTIN', type: 'string' },
      'Payer.TIN': { property: 'payerTIN', type: 'string' },
      'Payer.TaxpayerIdentificationNumber': { property: 'payerTIN', type: 'string' },
      'PayerAddress': { property: 'payerAddress', type: 'string' },
      'Payer.Address': { property: 'payerAddress', type: 'string' },
      'Recipient': { property: 'recipientName', type: 'string' },
      'RecipientName': { property: 'recipientName', type: 'string' },
      'Recipient.Name': { property: 'recipientName', type: 'string' },
      'RecipientTIN': { property: 'recipientTIN', type: 'string' },
      'Recipient.TIN': { property: 'recipientTIN', type: 'string' },
      'Recipient.TaxpayerIdentificationNumber': { property: 'recipientTIN', type: 'string' },
      'RecipientAddress': { property: 'recipientAddress', type: 'string' },
      'Recipient.Address': { property: 'recipientAddress', type: 'string' },
      'AccountNumber': { property: 'accountNumber', type: 'string' },
      'Account': { property: 'accountNumber', type: 'string' },
      
      // Box 1-15 mappings
      'InterestIncome': { property: 'interestIncome', type: 'number' },
      'Box1': { property: 'interestIncome', type: 'number' },
      '1099IntBox1': { property: 'interestIncome', type: 'number' },
      'EarlyWithdrawalPenalty': { property: 'earlyWithdrawalPenalty', type: 'number' },
      'Box2': { property: 'earlyWithdrawalPenalty', type: 'number' },
      '1099IntBox2': { property: 'earlyWithdrawalPenalty', type: 'number' },
      'InterestOnUSTreasuryObligations': { property: 'interestOnUSavingsBonds', type: 'number' },
      'InterestOnUSavingsBonds': { property: 'interestOnUSavingsBonds', type: 'number' },
      'Box3': { property: 'interestOnUSavingsBonds', type: 'number' },
      '1099IntBox3': { property: 'interestOnUSavingsBonds', type: 'number' },
      'FederalIncomeTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'FederalTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'Box4': { property: 'federalTaxWithheld', type: 'number' },
      '1099IntBox4': { property: 'federalTaxWithheld', type: 'number' },
      'InvestmentExpenses': { property: 'investmentExpenses', type: 'number' },
      'Box5': { property: 'investmentExpenses', type: 'number' },
      '1099IntBox5': { property: 'investmentExpenses', type: 'number' },
      'ForeignTaxPaid': { property: 'foreignTaxPaid', type: 'number' },
      'Box6': { property: 'foreignTaxPaid', type: 'number' },
      '1099IntBox6': { property: 'foreignTaxPaid', type: 'number' },
      'ForeignCountry': { property: 'foreignCountry', type: 'string' },
      'Box7': { property: 'foreignCountry', type: 'string' },
      '1099IntBox7': { property: 'foreignCountry', type: 'string' },
      'TaxExemptInterest': { property: 'taxExemptInterest', type: 'number' },
      'Box8': { property: 'taxExemptInterest', type: 'number' },
      '1099IntBox8': { property: 'taxExemptInterest', type: 'number' },
      'SpecifiedPrivateActivityBondInterest': { property: 'specifiedPrivateActivityBondInterest', type: 'number' },
      'Box9': { property: 'specifiedPrivateActivityBondInterest', type: 'number' },
      '1099IntBox9': { property: 'specifiedPrivateActivityBondInterest', type: 'number' },
      'MarketDiscount': { property: 'marketDiscount', type: 'number' },
      'Box10': { property: 'marketDiscount', type: 'number' },
      '1099IntBox10': { property: 'marketDiscount', type: 'number' },
      'BondPremium': { property: 'bondPremium', type: 'number' },
      'Box11': { property: 'bondPremium', type: 'number' },
      '1099IntBox11': { property: 'bondPremium', type: 'number' },
      'StateTaxWithheld': { property: 'stateTaxWithheld', type: 'number' },
      'Box13': { property: 'stateTaxWithheld', type: 'number' },
      '1099IntBox13': { property: 'stateTaxWithheld', type: 'number' },
      'StatePayerNumber': { property: 'statePayerNumber', type: 'string' },
      'Box14': { property: 'statePayerNumber', type: 'string' },
      '1099IntBox14': { property: 'statePayerNumber', type: 'string' },
      'StateInterest': { property: 'stateInterest', type: 'number' },
      'Box15': { property: 'stateInterest', type: 'number' },
      '1099IntBox15': { property: 'stateInterest', type: 'number' }
    };
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [azureFieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && fieldData.value !== undefined && fieldData.value !== null) {
        const mapping = fieldMappings[azureFieldName];
        
        if (mapping) {
          const { property, type } = mapping;
          let processedValue: any;
          
          try {
            switch (type) {
              case 'string':
                processedValue = this.extractStringValue(fieldData.value);
                break;
              case 'number':
                processedValue = this.extractNumericValue(fieldData.value);
                break;
              default:
                processedValue = fieldData.value;
            }
            
            if (processedValue !== null && processedValue !== undefined && 
                (type !== 'string' || processedValue.trim() !== '') &&
                (type !== 'number' || processedValue !== 0)) {
              (data as any)[property] = processedValue;
              fieldsExtracted++;
              console.log(`✅ [Azure DI] Extracted ${azureFieldName} → ${property}: ${processedValue}`);
            }
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
          }
        }
      }
    }
    
    console.log(`📊 [Azure DI] 1099-INT field processing summary: ${fieldsExtracted} fields extracted`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-DIV fields from structured analysis - COMPLETE IMPLEMENTATION
   */
  private process1099DivFields(fields: any, baseData: BaseTaxDocument): Form1099DivData {
    console.log('🔍 [Azure DI] Processing 1099-DIV fields from structured analysis...');
    const data: Form1099DivData = { ...baseData };
    
    const fieldMappings: { [key: string]: { property: keyof Form1099DivData; type: 'string' | 'number' | 'boolean' } } = {
      // Payer and recipient information
      'Payer': { property: 'payerName', type: 'string' },
      'PayerName': { property: 'payerName', type: 'string' },
      'Payer.Name': { property: 'payerName', type: 'string' },
      'PayerTIN': { property: 'payerTIN', type: 'string' },
      'Payer.TIN': { property: 'payerTIN', type: 'string' },
      'PayerAddress': { property: 'payerAddress', type: 'string' },
      'Payer.Address': { property: 'payerAddress', type: 'string' },
      'Recipient': { property: 'recipientName', type: 'string' },
      'RecipientName': { property: 'recipientName', type: 'string' },
      'Recipient.Name': { property: 'recipientName', type: 'string' },
      'RecipientTIN': { property: 'recipientTIN', type: 'string' },
      'Recipient.TIN': { property: 'recipientTIN', type: 'string' },
      'RecipientAddress': { property: 'recipientAddress', type: 'string' },
      'Recipient.Address': { property: 'recipientAddress', type: 'string' },
      'AccountNumber': { property: 'accountNumber', type: 'string' },
      
      // Dividend fields
      'OrdinaryDividends': { property: 'ordinaryDividends', type: 'number' },
      'Box1a': { property: 'ordinaryDividends', type: 'number' },
      '1099DivBox1a': { property: 'ordinaryDividends', type: 'number' },
      'QualifiedDividends': { property: 'qualifiedDividends', type: 'number' },
      'Box1b': { property: 'qualifiedDividends', type: 'number' },
      '1099DivBox1b': { property: 'qualifiedDividends', type: 'number' },
      'TotalCapitalGainDistributions': { property: 'totalCapitalGain', type: 'number' },
      'TotalCapitalGain': { property: 'totalCapitalGain', type: 'number' },
      'Box2a': { property: 'totalCapitalGain', type: 'number' },
      '1099DivBox2a': { property: 'totalCapitalGain', type: 'number' },
      'UnrecapturedSection1250Gain': { property: 'unrecapturedSection1250Gain', type: 'number' },
      'Box2b': { property: 'unrecapturedSection1250Gain', type: 'number' },
      '1099DivBox2b': { property: 'unrecapturedSection1250Gain', type: 'number' },
      'Section1202Gain': { property: 'section1202Gain', type: 'number' },
      'Box2c': { property: 'section1202Gain', type: 'number' },
      '1099DivBox2c': { property: 'section1202Gain', type: 'number' },
      'CollectiblesGain': { property: 'collectiblesGain', type: 'number' },
      'Box2d': { property: 'collectiblesGain', type: 'number' },
      '1099DivBox2d': { property: 'collectiblesGain', type: 'number' },
      'Section897OrdinaryDividends': { property: 'section897OrdinaryDividends', type: 'number' },
      'Box2e': { property: 'section897OrdinaryDividends', type: 'number' },
      '1099DivBox2e': { property: 'section897OrdinaryDividends', type: 'number' },
      'Section897CapitalGain': { property: 'section897CapitalGain', type: 'number' },
      'Box2f': { property: 'section897CapitalGain', type: 'number' },
      '1099DivBox2f': { property: 'section897CapitalGain', type: 'number' },
      'NondividendDistributions': { property: 'nondividendDistributions', type: 'number' },
      'Box3': { property: 'nondividendDistributions', type: 'number' },
      '1099DivBox3': { property: 'nondividendDistributions', type: 'number' },
      'FederalIncomeTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'FederalTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'Box4': { property: 'federalTaxWithheld', type: 'number' },
      '1099DivBox4': { property: 'federalTaxWithheld', type: 'number' },
      'Section199ADividends': { property: 'section199ADividends', type: 'number' },
      'Box5': { property: 'section199ADividends', type: 'number' },
      '1099DivBox5': { property: 'section199ADividends', type: 'number' },
      'ExemptInterestDividends': { property: 'exemptInterestDividends', type: 'number' },
      'Box6': { property: 'exemptInterestDividends', type: 'number' },
      '1099DivBox6': { property: 'exemptInterestDividends', type: 'number' },
      'ForeignTaxPaid': { property: 'foreignTaxPaid', type: 'number' },
      'Box7': { property: 'foreignTaxPaid', type: 'number' },
      '1099DivBox7': { property: 'foreignTaxPaid', type: 'number' },
      'ForeignCountry': { property: 'foreignCountry', type: 'string' },
      'Box8': { property: 'foreignCountry', type: 'string' },
      '1099DivBox8': { property: 'foreignCountry', type: 'string' },
      'CashLiquidationDistributions': { property: 'cashLiquidationDistributions', type: 'number' },
      'Box9': { property: 'cashLiquidationDistributions', type: 'number' },
      '1099DivBox9': { property: 'cashLiquidationDistributions', type: 'number' },
      'NoncashLiquidationDistributions': { property: 'noncashLiquidationDistributions', type: 'number' },
      'Box10': { property: 'noncashLiquidationDistributions', type: 'number' },
      '1099DivBox10': { property: 'noncashLiquidationDistributions', type: 'number' },
      'FATCAFilingRequirement': { property: 'fatcaFilingRequirement', type: 'boolean' },
      'FatcaFilingRequirement': { property: 'fatcaFilingRequirement', type: 'boolean' },
      'Box11': { property: 'fatcaFilingRequirement', type: 'boolean' },
      '1099DivBox11': { property: 'fatcaFilingRequirement', type: 'boolean' },
      'InvestmentExpenses': { property: 'investmentExpenses', type: 'number' },
      'Box13': { property: 'investmentExpenses', type: 'number' },
      '1099DivBox13': { property: 'investmentExpenses', type: 'number' },
      'StateTaxWithheld': { property: 'stateTaxWithheld', type: 'number' },
      'StatePayerNumber': { property: 'statePayerNumber', type: 'string' },
      'StateIncome': { property: 'stateIncome', type: 'number' }
    };
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [azureFieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && fieldData.value !== undefined && fieldData.value !== null) {
        const mapping = fieldMappings[azureFieldName];
        
        if (mapping) {
          const { property, type } = mapping;
          let processedValue: any;
          
          try {
            switch (type) {
              case 'string':
                processedValue = this.extractStringValue(fieldData.value);
                break;
              case 'number':
                processedValue = this.extractNumericValue(fieldData.value);
                break;
              case 'boolean':
                processedValue = this.extractBooleanValue(fieldData.value);
                break;
              default:
                processedValue = fieldData.value;
            }
            
            if (processedValue !== null && processedValue !== undefined && 
                (type !== 'string' || processedValue.trim() !== '') &&
                (type !== 'number' || processedValue !== 0)) {
              (data as any)[property] = processedValue;
              fieldsExtracted++;
              console.log(`✅ [Azure DI] Extracted ${azureFieldName} → ${property}: ${processedValue}`);
            }
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
          }
        }
      }
    }
    
    console.log(`📊 [Azure DI] 1099-DIV field processing summary: ${fieldsExtracted} fields extracted`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process 1099-MISC fields from structured analysis - COMPLETE IMPLEMENTATION
   */
  private process1099MiscFields(fields: any, baseData: BaseTaxDocument): Form1099MiscData {
    console.log('🔍 [Azure DI] Processing 1099-MISC fields from structured analysis...');
    const data: Form1099MiscData = { ...baseData };
    
    const fieldMappings: { [key: string]: { property: keyof Form1099MiscData; type: 'string' | 'number' | 'boolean' } } = {
      // Payer and recipient information
      'Payer': { property: 'payerName', type: 'string' },
      'PayerName': { property: 'payerName', type: 'string' },
      'Payer.Name': { property: 'payerName', type: 'string' },
      'PayerTIN': { property: 'payerTIN', type: 'string' },
      'Payer.TIN': { property: 'payerTIN', type: 'string' },
      'PayerAddress': { property: 'payerAddress', type: 'string' },
      'Payer.Address': { property: 'payerAddress', type: 'string' },
      'Recipient': { property: 'recipientName', type: 'string' },
      'RecipientName': { property: 'recipientName', type: 'string' },
      'Recipient.Name': { property: 'recipientName', type: 'string' },
      'RecipientTIN': { property: 'recipientTIN', type: 'string' },
      'Recipient.TIN': { property: 'recipientTIN', type: 'string' },
      'RecipientAddress': { property: 'recipientAddress', type: 'string' },
      'Recipient.Address': { property: 'recipientAddress', type: 'string' },
      'AccountNumber': { property: 'accountNumber', type: 'string' },
      
      // Box 1-18 mappings
      'Rents': { property: 'rents', type: 'number' },
      'Box1': { property: 'rents', type: 'number' },
      '1099MiscBox1': { property: 'rents', type: 'number' },
      'Royalties': { property: 'royalties', type: 'number' },
      'Box2': { property: 'royalties', type: 'number' },
      '1099MiscBox2': { property: 'royalties', type: 'number' },
      'OtherIncome': { property: 'otherIncome', type: 'number' },
      'Box3': { property: 'otherIncome', type: 'number' },
      '1099MiscBox3': { property: 'otherIncome', type: 'number' },
      'FederalIncomeTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'FederalTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'Box4': { property: 'federalTaxWithheld', type: 'number' },
      '1099MiscBox4': { property: 'federalTaxWithheld', type: 'number' },
      'FishingBoatProceeds': { property: 'fishingBoatProceeds', type: 'number' },
      'Box5': { property: 'fishingBoatProceeds', type: 'number' },
      '1099MiscBox5': { property: 'fishingBoatProceeds', type: 'number' },
      'MedicalAndHealthCarePayments': { property: 'medicalHealthPayments', type: 'number' },
      'MedicalHealthPayments': { property: 'medicalHealthPayments', type: 'number' },
      'Box6': { property: 'medicalHealthPayments', type: 'number' },
      '1099MiscBox6': { property: 'medicalHealthPayments', type: 'number' },
      'NonemployeeCompensation': { property: 'nonemployeeCompensation', type: 'number' },
      'Box7': { property: 'nonemployeeCompensation', type: 'number' },
      '1099MiscBox7': { property: 'nonemployeeCompensation', type: 'number' },
      'SubstitutePayments': { property: 'substitutePayments', type: 'number' },
      'Box8': { property: 'substitutePayments', type: 'number' },
      '1099MiscBox8': { property: 'substitutePayments', type: 'number' },
      'CropInsuranceProceeds': { property: 'cropInsuranceProceeds', type: 'number' },
      'Box9': { property: 'cropInsuranceProceeds', type: 'number' },
      '1099MiscBox9': { property: 'cropInsuranceProceeds', type: 'number' },
      'GrossProceedsPaidToAttorney': { property: 'grossProceedsAttorney', type: 'number' },
      'GrossProceedsAttorney': { property: 'grossProceedsAttorney', type: 'number' },
      'Box10': { property: 'grossProceedsAttorney', type: 'number' },
      '1099MiscBox10': { property: 'grossProceedsAttorney', type: 'number' },
      'FishPurchasedForResale': { property: 'fishPurchases', type: 'number' },
      'FishPurchases': { property: 'fishPurchases', type: 'number' },
      'Box11': { property: 'fishPurchases', type: 'number' },
      '1099MiscBox11': { property: 'fishPurchases', type: 'number' },
      'Section409ADeferrals': { property: 'section409ADeferrals', type: 'number' },
      'Box12': { property: 'section409ADeferrals', type: 'number' },
      '1099MiscBox12': { property: 'section409ADeferrals', type: 'number' },
      'ExcessGoldenParachutePayments': { property: 'excessGoldenParachutePayments', type: 'number' },
      'Box13': { property: 'excessGoldenParachutePayments', type: 'number' },
      '1099MiscBox13': { property: 'excessGoldenParachutePayments', type: 'number' },
      'NonqualifiedDeferredCompensation': { property: 'nonqualifiedDeferredCompensation', type: 'number' },
      'Box14': { property: 'nonqualifiedDeferredCompensation', type: 'number' },
      '1099MiscBox14': { property: 'nonqualifiedDeferredCompensation', type: 'number' },
      'Section409AIncome': { property: 'section409AIncome', type: 'number' },
      'Box15a': { property: 'section409AIncome', type: 'number' },
      '1099MiscBox15a': { property: 'section409AIncome', type: 'number' },
      'StateTaxWithheld': { property: 'stateTaxWithheld', type: 'number' },
      'Box16': { property: 'stateTaxWithheld', type: 'number' },
      '1099MiscBox16': { property: 'stateTaxWithheld', type: 'number' },
      'StatePayerNumber': { property: 'statePayerNumber', type: 'string' },
      'Box17': { property: 'statePayerNumber', type: 'string' },
      '1099MiscBox17': { property: 'statePayerNumber', type: 'string' },
      'StateIncome': { property: 'stateIncome', type: 'number' },
      'Box18': { property: 'stateIncome', type: 'number' },
      '1099MiscBox18': { property: 'stateIncome', type: 'number' }
    };
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [azureFieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && fieldData.value !== undefined && fieldData.value !== null) {
        const mapping = fieldMappings[azureFieldName];
        
        if (mapping) {
          const { property, type } = mapping;
          let processedValue: any;
          
          try {
            switch (type) {
              case 'string':
                processedValue = this.extractStringValue(fieldData.value);
                break;
              case 'number':
                processedValue = this.extractNumericValue(fieldData.value);
                break;
              default:
                processedValue = fieldData.value;
            }
            
            if (processedValue !== null && processedValue !== undefined && 
                (type !== 'string' || processedValue.trim() !== '') &&
                (type !== 'number' || processedValue !== 0)) {
              (data as any)[property] = processedValue;
              fieldsExtracted++;
              console.log(`✅ [Azure DI] Extracted ${azureFieldName} → ${property}: ${processedValue}`);
            }
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
          }
        }
      }
    }
    
    console.log(`📊 [Azure DI] 1099-MISC field processing summary: ${fieldsExtracted} fields extracted`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    // Validate and correct field mappings using OCR
    if (baseData.fullText) {
      return this.validateAndCorrect1099MiscFields(data, baseData.fullText);
    }
    
    return data;
  }

  /**
   * Process 1099-NEC fields from structured analysis - COMPLETE IMPLEMENTATION
   */
  private process1099NecFields(fields: any, baseData: BaseTaxDocument): Form1099MiscData {
    console.log('🔍 [Azure DI] Processing 1099-NEC fields from structured analysis...');
    const data: Form1099MiscData = { ...baseData };
    
    const fieldMappings: { [key: string]: { property: keyof Form1099MiscData; type: 'string' | 'number' | 'boolean' } } = {
      // Payer and recipient information
      'Payer': { property: 'payerName', type: 'string' },
      'PayerName': { property: 'payerName', type: 'string' },
      'Payer.Name': { property: 'payerName', type: 'string' },
      'PayerTIN': { property: 'payerTIN', type: 'string' },
      'Payer.TIN': { property: 'payerTIN', type: 'string' },
      'PayerAddress': { property: 'payerAddress', type: 'string' },
      'Payer.Address': { property: 'payerAddress', type: 'string' },
      'Recipient': { property: 'recipientName', type: 'string' },
      'RecipientName': { property: 'recipientName', type: 'string' },
      'Recipient.Name': { property: 'recipientName', type: 'string' },
      'RecipientTIN': { property: 'recipientTIN', type: 'string' },
      'Recipient.TIN': { property: 'recipientTIN', type: 'string' },
      'RecipientAddress': { property: 'recipientAddress', type: 'string' },
      'Recipient.Address': { property: 'recipientAddress', type: 'string' },
      'AccountNumber': { property: 'accountNumber', type: 'string' },
      
      // 1099-NEC specific fields
      'NonemployeeCompensation': { property: 'nonemployeeCompensation', type: 'number' },
      'Box1': { property: 'nonemployeeCompensation', type: 'number' },
      '1099NecBox1': { property: 'nonemployeeCompensation', type: 'number' },
      'FederalIncomeTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'FederalTaxWithheld': { property: 'federalTaxWithheld', type: 'number' },
      'Box4': { property: 'federalTaxWithheld', type: 'number' },
      '1099NecBox4': { property: 'federalTaxWithheld', type: 'number' },
      'StateTaxWithheld': { property: 'stateTaxWithheld', type: 'number' },
      'StatePayerNumber': { property: 'statePayerNumber', type: 'string' },
      'StateIncome': { property: 'stateIncome', type: 'number' }
    };
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [azureFieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && fieldData.value !== undefined && fieldData.value !== null) {
        const mapping = fieldMappings[azureFieldName];
        
        if (mapping) {
          const { property, type } = mapping;
          let processedValue: any;
          
          try {
            switch (type) {
              case 'string':
                processedValue = this.extractStringValue(fieldData.value);
                break;
              case 'number':
                processedValue = this.extractNumericValue(fieldData.value);
                break;
              default:
                processedValue = fieldData.value;
            }
            
            if (processedValue !== null && processedValue !== undefined && 
                (type !== 'string' || processedValue.trim() !== '') &&
                (type !== 'number' || processedValue !== 0)) {
              (data as any)[property] = processedValue;
              fieldsExtracted++;
              console.log(`✅ [Azure DI] Extracted ${azureFieldName} → ${property}: ${processedValue}`);
            }
          } catch (error) {
            console.warn(`⚠️ [Azure DI] Error processing field ${azureFieldName}:`, error);
          }
        }
      }
    }
    
    console.log(`📊 [Azure DI] 1099-NEC field processing summary: ${fieldsExtracted} fields extracted`);
    
    // OCR fallback for missing personal info
    this.applyPersonalInfoOCRFallback(data, baseData.fullText);
    
    return data;
  }

  /**
   * Process generic tax fields
   */
  private processGenericTaxFields(fields: any, baseData: BaseTaxDocument): TaxDocumentData {
    console.log('🔍 [Azure DI] Processing generic tax fields...');
    const data = { ...baseData } as any;
    
    let fieldsExtracted = 0;
    
    // Process all available fields
    for (const [fieldName, fieldData] of Object.entries(fields)) {
      if (fieldData && typeof fieldData === 'object' && 'value' in fieldData) {
        const value = (fieldData as any).value;
        if (value !== undefined && value !== null && value !== '') {
          const processedValue = typeof value === 'number' ? value : this.extractNumericValue(value) || this.extractStringValue(value);
          if (processedValue !== null && processedValue !== undefined) {
            data[fieldName] = processedValue;
            fieldsExtracted++;
            console.log(`✅ [Azure DI] Extracted generic field ${fieldName}: ${processedValue}`);
          }
        }
      }
    }
    
    console.log(`📊 [Azure DI] Generic field processing summary: ${fieldsExtracted} fields extracted`);
    
    return data;
  }

  // Value extraction utility methods
  private extractStringValue(value: any): string {
    if (typeof value === 'string') {
      return value.trim();
    }
    if (typeof value === 'object' && value.content) {
      return String(value.content).trim();
    }
    return String(value).trim();
  }

  private extractNumericValue(value: any): number {
    if (typeof value === 'number') {
      return value;
    }
    
    let stringValue: string;
    if (typeof value === 'string') {
      stringValue = value;
    } else if (typeof value === 'object' && value.content) {
      stringValue = String(value.content);
    } else {
      stringValue = String(value);
    }
    
    // Remove currency symbols, commas, and whitespace
    const cleanValue = stringValue.replace(/[$,\s]/g, '');
    const parsed = parseFloat(cleanValue);
    return isNaN(parsed) ? 0 : parsed;
  }

  private extractBooleanValue(value: any): boolean {
    if (typeof value === 'boolean') {
      return value;
    }
    
    let stringValue: string;
    if (typeof value === 'string') {
      stringValue = value;
    } else if (typeof value === 'object' && value.content) {
      stringValue = String(value.content);
    } else {
      stringValue = String(value);
    }
    
    const lower = stringValue.toLowerCase().trim();
    return lower === 'true' || lower === 'yes' || lower === 'x' || lower === '✓' || lower === 'checked' || lower === '1';
  }

  // OCR-based extraction methods (simplified versions)
  private extractW2FieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): W2Data {
    const w2Data: W2Data = { ...baseData };
    
    // Extract wages (Box 1)
    const wagesMatch = ocrText.match(/(?:box\s*1|wages.*tips.*compensation)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (wagesMatch) {
      w2Data.wages = this.extractNumericValue(wagesMatch[1]);
    }
    
    // Extract federal tax withheld (Box 2)
    const fedTaxMatch = ocrText.match(/(?:box\s*2|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      w2Data.federalTaxWithheld = this.extractNumericValue(fedTaxMatch[1]);
    }
    
    // Extract personal info using existing method
    const personalInfo = this.extractPersonalInfoFromOCR(ocrText);
    if (personalInfo.name) w2Data.employeeName = personalInfo.name;
    if (personalInfo.ssn) w2Data.employeeSSN = personalInfo.ssn;
    if (personalInfo.address) w2Data.employeeAddress = personalInfo.address;
    if (personalInfo.employerName) w2Data.employerName = personalInfo.employerName;
    if (personalInfo.employerEIN) w2Data.employerEIN = personalInfo.employerEIN;
    
    return w2Data;
  }

  private extract1099IntFieldsFromOCR(ocrText: string, baseData: BaseTaxDocument): Form1099IntData {
    const data: Form1099IntData = { ...baseData };
    
    // Extract interest income (Box 1)
    const interestMatch = ocrText.match(/(?:box\s*1|interest.*income)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (interestMatch) {
      data.interestIncome = this.extractNumericValue(interestMatch[1]);
    }
    
    // Extract federal tax withheld (Box 4)
    const fedTaxMatch = ocrText.match(/(?:box\s*4|federal.*income.*tax.*withheld)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      data.federalTaxWithheld = this.extractNumericValue(fedTaxMatch[1]);
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
    
    // Extract ordinary dividends (Box 1a)
    const ordinaryDivMatch = ocrText.match(/(?:box\s*1a|ordinary.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (ordinaryDivMatch) {
      data.ordinaryDividends = this.extractNumericValue(ordinaryDivMatch[1]);
    }
    
    // Extract qualified dividends (Box 1b)
    const qualifiedDivMatch = ocrText.match(/(?:box\s*1b|qualified.*dividends)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (qualifiedDivMatch) {
      data.qualifiedDividends = this.extractNumericValue(qualifiedDivMatch[1]);
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
          const amount = this.extractNumericValue(match[1]);
          if (amount > 0) {
            (data as any)[field] = amount;
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
    
    // Extract nonemployee compensation
    const necMatch = ocrText.match(/(?:nonemployee.*compensation|box\s*1)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (necMatch) {
      data.nonemployeeCompensation = this.extractNumericValue(necMatch[1]);
    }
    
    // Extract federal tax withheld
    const fedTaxMatch = ocrText.match(/(?:federal.*income.*tax.*withheld|box\s*4)[:\s]*\$?(\d+(?:,\d{3})*(?:\.\d{2})?)/i);
    if (fedTaxMatch) {
      data.federalTaxWithheld = this.extractNumericValue(fedTaxMatch[1]);
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
      checkboxes.retirementPlan = this.extractBooleanValue(fields['RetirementPlan'].value);
    }
    if (fields['ThirdPartySickPay']?.value !== undefined) {
      checkboxes.thirdPartySickPay = this.extractBooleanValue(fields['ThirdPartySickPay'].value);
    }
    if (fields['StatutoryEmployee']?.value !== undefined) {
      checkboxes.statutoryEmployee = this.extractBooleanValue(fields['StatutoryEmployee'].value);
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
    
    // Extract SSN/TIN patterns
    const ssnPattern = /\b(\d{3}[-\s]?\d{2}[-\s]?\d{4})\b/g;
    const ssnMatches = Array.from(ocrText.matchAll(ssnPattern));
    if (ssnMatches.length > 0) {
      personalInfo.ssn = ssnMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.tin = personalInfo.ssn;
    }
    
    // Extract EIN patterns
    const einPattern = /\b(\d{2}[-\s]?\d{7})\b/g;
    const einMatches = Array.from(ocrText.matchAll(einPattern));
    if (einMatches.length > 0) {
      personalInfo.employerEIN = einMatches[0][1].replace(/[-\s]/g, '');
      personalInfo.payerTIN = personalInfo.employerEIN;
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
        }
      }
    }
    
    // Extract addresses (simplified approach)
    for (const line of lines) {
      if (/\d+.*(?:street|st|avenue|ave|road|rd|drive|dr|lane|ln|way|blvd|boulevard)/i.test(line)) {
        if (!personalInfo.address) {
          personalInfo.address = line;
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
      const structuredValue = this.extractNumericValue((structuredData as any)[field]) || 0;
      const ocrValue = this.extractNumericValue((ocrData as any)[field]) || 0;
      
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
}

// Export the service and interfaces
export default AzureDocumentIntelligenceService;
