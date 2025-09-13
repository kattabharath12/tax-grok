import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";

// Base interfaces
export interface BaseTaxDocument {
  documentType: string;
  confidence: number;
  extractedAt: string;
}

export interface W2Data extends BaseTaxDocument {
  documentType: "W2";
  employerName?: string;
  employerAddress?: string;
  employerEIN?: string;
  employeeName?: string;
  employeeAddress?: string;
  employeeSSN?: string;
  wages?: number;
  federalTaxWithheld?: number;
  socialSecurityWages?: number;
  socialSecurityTaxWithheld?: number;
  medicareWages?: number;
  medicareTaxWithheld?: number;
  socialSecurityTips?: number;
  allocatedTips?: number;
  dependentCareBenefits?: number;
  nonqualifiedPlans?: number;
  box12a?: string;
  box12b?: string;
  box12c?: string;
  box12d?: string;
  statutoryEmployee?: boolean;
  retirementPlan?: boolean;
  thirdPartySickPay?: boolean;
  stateWages?: number;
  stateTaxWithheld?: number;
  localWages?: number;
  localTaxWithheld?: number;
  state?: string;
  locality?: string;
}

export interface Form1099IntData extends BaseTaxDocument {
  documentType: "1099-INT";
  payerName?: string;
  payerAddress?: string;
  payerTIN?: string;
  recipientName?: string;
  recipientAddress?: string;
  recipientTIN?: string;
  accountNumber?: string;
  interestIncome?: number;
  earlyWithdrawalPenalty?: number;
  interestOnUSSavingsBonds?: number;
  federalTaxWithheld?: number;
  investmentExpenses?: number;
  foreignTaxPaid?: number;
  foreignCountry?: string;
  taxExemptInterest?: number;
  specifiedPrivateActivityBondInterest?: number;
  marketDiscount?: number;
  bondPremium?: number;
  bondPremiumOnTreasury?: number;
  bondPremiumOnTaxExempt?: number;
  cusipNumber?: string;
  stateTaxWithheld?: number;
  state?: string;
  stateIdNumber?: string;
}

export interface Form1099DivData extends BaseTaxDocument {
  documentType: "1099-DIV";
  payerName?: string;
  payerAddress?: string;
  payerTIN?: string;
  recipientName?: string;
  recipientAddress?: string;
  recipientTIN?: string;
  accountNumber?: string;
  totalOrdinaryDividends?: number;
  qualifiedDividends?: number;
  totalCapitalGainDistributions?: number;
  unrecaptured1250Gain?: number;
  section1202Gain?: number;
  collectiblesGain?: number;
  nondividendDistributions?: number;
  federalTaxWithheld?: number;
  section199ADividends?: number;
  investmentExpenses?: number;
  foreignTaxPaid?: number;
  foreignCountry?: string;
  cashLiquidationDistributions?: number;
  noncashLiquidationDistributions?: number;
  exemptInterestDividends?: number;
  specifiedPrivateActivityBondInterestDividends?: number;
  stateTaxWithheld?: number;
  state?: string;
  stateIdNumber?: string;
}

export interface Form1099MiscData extends BaseTaxDocument {
  documentType: "1099-MISC";
  payerName?: string;
  payerAddress?: string;
  payerTIN?: string;
  recipientName?: string;
  recipientAddress?: string;
  recipientTIN?: string;
  accountNumber?: string;
  rents?: number;
  royalties?: number;
  otherIncome?: number;
  federalTaxWithheld?: number;
  fishingBoatProceeds?: number;
  medicalHealthcarePayments?: number;
  nonemployeeCompensation?: number;
  substitutePayments?: number;
  cropInsuranceProceeds?: number;
  grossProceedsAttorney?: number;
  section409ADeferrals?: number;
  section409AIncome?: number;
  stateTaxWithheld?: number;
  state?: string;
  stateIdNumber?: string;
  secondTINNotice?: boolean;
}

export type TaxDocumentData = W2Data | Form1099IntData | Form1099DivData | Form1099MiscData;

export class AzureDocumentIntelligenceService {
  private client: DocumentAnalysisClient;

  constructor() {
    const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;
    const apiKey = process.env.AZURE_DOCUMENT_INTELLIGENCE_API_KEY;

    if (!endpoint || !apiKey) {
      throw new Error("Azure Document Intelligence credentials not configured");
    }

    this.client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));
  }

  async extractDataFromDocument(buffer: Buffer, filename: string): Promise<TaxDocumentData | null> {
    try {
      console.log(`Starting document analysis for: ${filename}`);
      
      const poller = await this.client.beginAnalyzeDocument("prebuilt-document", buffer);
      const result = await poller.pollUntilDone();

      if (!result.documents || result.documents.length === 0) {
        console.log("No documents found in the analysis result");
        return null;
      }

      console.log(`Found ${result.documents.length} document(s)`);
      
      // Extract text content for OCR-based extraction
      const textContent = this.extractTextContent(result);
      console.log(`Extracted text content length: ${textContent.length}`);

      // Detect document type
      const documentType = this.detectDocumentType(textContent, filename);
      console.log(`Detected document type: ${documentType}`);

      // Extract structured data based on document type
      let extractedData: TaxDocumentData | null = null;

      switch (documentType) {
        case "W2":
          extractedData = await this.extractW2(result, textContent);
          break;
        case "1099-INT":
          extractedData = await this.extract1099Int(result, textContent);
          break;
        case "1099-DIV":
          extractedData = await this.extract1099Div(result, textContent);
          break;
        case "1099-MISC":
          extractedData = await this.extract1099Misc(result, textContent);
          break;
        default:
          console.log(`Unsupported document type: ${documentType}`);
          return null;
      }

      if (extractedData) {
        console.log(`Successfully extracted ${documentType} data:`, JSON.stringify(extractedData, null, 2));
      } else {
        console.log(`Failed to extract data for ${documentType}`);
      }

      return extractedData;
    } catch (error) {
      console.error("Error in document analysis:", error);
      throw error;
    }
  }

  private extractTextContent(result: any): string {
    if (!result.content) {
      return "";
    }
    return result.content;
  }

  private detectDocumentType(textContent: string, filename: string): string {
    const text = textContent.toLowerCase();
    const name = filename.toLowerCase();

    // Check filename first
    if (name.includes("w2") || name.includes("w-2")) {
      return "W2";
    }
    if (name.includes("1099-int") || name.includes("1099int")) {
      return "1099-INT";
    }
    if (name.includes("1099-div") || name.includes("1099div")) {
      return "1099-DIV";
    }
    if (name.includes("1099-misc") || name.includes("1099misc")) {
      return "1099-MISC";
    }

    // Check content patterns
    if (text.includes("form w-2") || text.includes("wage and tax statement")) {
      return "W2";
    }
    if (text.includes("form 1099-int") || text.includes("interest income")) {
      return "1099-INT";
    }
    if (text.includes("form 1099-div") || text.includes("dividends and distributions")) {
      return "1099-DIV";
    }
    if (text.includes("form 1099-misc") || text.includes("miscellaneous income")) {
      return "1099-MISC";
    }

    return "UNKNOWN";
  }

  async extractW2(result: any, textContent: string): Promise<W2Data> {
    console.log("Extracting W2 data...");
    
    const baseData: W2Data = {
      documentType: "W2",
      confidence: 0.8,
      extractedAt: new Date().toISOString(),
    };

    // Try Azure Document Intelligence structured extraction first
    const azureData = this.extractW2FromAzureFields(result);
    
    // Then try OCR-based extraction for missing fields
    const ocrData = this.extractW2FromOCR(textContent);
    
    // Merge the results, preferring Azure data when available
    const mergedData = { ...baseData, ...ocrData, ...azureData };
    
    console.log("W2 extraction completed:", JSON.stringify(mergedData, null, 2));
    return mergedData;
  }

  private extractW2FromAzureFields(result: any): Partial<W2Data> {
    const data: Partial<W2Data> = {};
    
    if (!result.documents || result.documents.length === 0) {
      return data;
    }

    const document = result.documents[0];
    const fields = document.fields || {};

    // Map Azure Document Intelligence fields to our W2 structure
    const fieldMappings = {
      'EmployerName': ['employerName', 'Employer', 'CompanyName'],
      'EmployerAddress': ['employerAddress', 'EmployerAddr'],
      'EmployerEIN': ['employerEIN', 'EIN', 'EmployerTaxId'],
      'EmployeeName': ['employeeName', 'Employee', 'RecipientName'],
      'EmployeeAddress': ['employeeAddress', 'EmployeeAddr'],
      'EmployeeSSN': ['employeeSSN', 'SSN', 'SocialSecurityNumber'],
      'Wages': ['wages', 'WagesBox1', 'Box1'],
      'FederalTaxWithheld': ['federalTaxWithheld', 'FedTaxWithheld', 'Box2'],
      'SocialSecurityWages': ['socialSecurityWages', 'SSWages', 'Box3'],
      'SocialSecurityTaxWithheld': ['socialSecurityTaxWithheld', 'SSTaxWithheld', 'Box4'],
      'MedicareWages': ['medicareWages', 'MedWages', 'Box5'],
      'MedicareTaxWithheld': ['medicareTaxWithheld', 'MedTaxWithheld', 'Box6'],
      'SocialSecurityTips': ['socialSecurityTips', 'SSTips', 'Box7'],
      'AllocatedTips': ['allocatedTips', 'Tips', 'Box8'],
      'DependentCareBenefits': ['dependentCareBenefits', 'DepCareBenefits', 'Box10'],
      'NonqualifiedPlans': ['nonqualifiedPlans', 'NonqualPlans', 'Box11'],
      'StateWages': ['stateWages', 'StateWage', 'Box15'],
      'StateTaxWithheld': ['stateTaxWithheld', 'StateTax', 'Box17'],
      'LocalWages': ['localWages', 'LocalWage', 'Box18'],
      'LocalTaxWithheld': ['localTaxWithheld', 'LocalTax', 'Box19'],
      'State': ['state', 'StateCode', 'Box15State'],
      'Locality': ['locality', 'LocalityName', 'Box20']
    };

    // Extract fields using multiple possible field names
    for (const [targetField, possibleNames] of Object.entries(fieldMappings)) {
      for (const fieldName of possibleNames) {
        if (fields[fieldName]) {
          const field = fields[fieldName];
          let value = field.content || field.value || field.valueString;
          
          // Convert numeric fields
          if (['wages', 'federalTaxWithheld', 'socialSecurityWages', 'socialSecurityTaxWithheld', 
               'medicareWages', 'medicareTaxWithheld', 'socialSecurityTips', 'allocatedTips',
               'dependentCareBenefits', 'nonqualifiedPlans', 'stateWages', 'stateTaxWithheld',
               'localWages', 'localTaxWithheld'].includes(targetField.toLowerCase())) {
            value = this.parseNumericValue(value);
          }
          
          // Convert boolean fields
          if (['statutoryEmployee', 'retirementPlan', 'thirdPartySickPay'].includes(targetField.toLowerCase())) {
            value = this.parseBooleanValue(value);
          }
          
          if (value !== null && value !== undefined) {
            (data as any)[targetField.toLowerCase()] = value;
            console.log(`Mapped Azure field ${fieldName} -> ${targetField}: ${value}`);
            break; // Use first matching field
          }
        }
      }
    }

    return data;
  }

  private extractW2FromOCR(textContent: string): Partial<W2Data> {
    console.log("Extracting W2 data from OCR text...");
    const data: Partial<W2Data> = {};

    // Enhanced regex patterns for W2 extraction
    const patterns = {
      // Box 1 - Wages, tips, other compensation
      wages: [
        /(?:box\s*1|wages[,\s]*tips[,\s]*other\s*compensation)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /wages[,\s]*tips[,\s]*other[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^1\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)1[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 2 - Federal income tax withheld
      federalTaxWithheld: [
        /(?:box\s*2|federal\s*income\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /federal\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)2[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 3 - Social security wages
      socialSecurityWages: [
        /(?:box\s*3|social\s*security\s*wages)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /social\s*security\s*wages[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^3\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)3[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 4 - Social security tax withheld
      socialSecurityTaxWithheld: [
        /(?:box\s*4|social\s*security\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /social\s*security\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^4\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)4[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 5 - Medicare wages and tips
      medicareWages: [
        /(?:box\s*5|medicare\s*wages\s*and\s*tips)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /medicare\s*wages[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^5\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)5[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 6 - Medicare tax withheld
      medicareTaxWithheld: [
        /(?:box\s*6|medicare\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /medicare\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^6\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)6[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 7 - Social security tips
      socialSecurityTips: [
        /(?:box\s*7|social\s*security\s*tips)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /social\s*security\s*tips[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^7\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)7[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 8 - Allocated tips
      allocatedTips: [
        /(?:box\s*8|allocated\s*tips)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /allocated\s*tips[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^8\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)8[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 10 - Dependent care benefits
      dependentCareBenefits: [
        /(?:box\s*10|dependent\s*care\s*benefits)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /dependent\s*care[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^10\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)10[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Box 11 - Nonqualified plans
      nonqualifiedPlans: [
        /(?:box\s*11|nonqualified\s*plans)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /nonqualified\s*plans[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^11\s*\$?([0-9,]+\.?\d*)/m,
        /(?:^|\s)11[\s\.]*([0-9,]+\.?\d*)(?:\s|$)/m
      ],
      
      // Employer information
      employerName: [
        /(?:employer|company)[:\s]*([A-Za-z0-9\s,\.&'-]+?)(?:\n|$|address|ein|tax)/i,
        /^([A-Za-z0-9\s,\.&'-]+?)(?:\n.*address|\n.*ein|\n.*tax)/im
      ],
      
      // Employee information
      employeeName: [
        /(?:employee|name)[:\s]*([A-Za-z\s,\.'-]+?)(?:\n|$|address|ssn)/i,
        /employee\s*name[:\s]*([A-Za-z\s,\.'-]+)/i
      ],
      
      // EIN
      employerEIN: [
        /(?:ein|employer\s*identification\s*number)[:\s]*([0-9-]+)/i,
        /ein[:\s]*([0-9-]+)/i,
        /\b(\d{2}-\d{7})\b/
      ],
      
      // SSN
      employeeSSN: [
        /(?:ssn|social\s*security\s*number)[:\s]*([0-9-]+)/i,
        /ssn[:\s]*([0-9-]+)/i,
        /\b(\d{3}-\d{2}-\d{4})\b/
      ],
      
      // State information
      state: [
        /(?:state|st)[:\s]*([A-Z]{2})/i,
        /\b([A-Z]{2})\s*(?:state|tax)/i
      ],
      
      // State wages (Box 15)
      stateWages: [
        /(?:box\s*15|state\s*wages)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /state\s*wages[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^15\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // State tax withheld (Box 17)
      stateTaxWithheld: [
        /(?:box\s*17|state\s*income\s*tax)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /state\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^17\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Local wages (Box 18)
      localWages: [
        /(?:box\s*18|local\s*wages)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /local\s*wages[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^18\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Local tax withheld (Box 19)
      localTaxWithheld: [
        /(?:box\s*19|local\s*income\s*tax)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /local\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^19\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Locality name (Box 20)
      locality: [
        /(?:box\s*20|locality\s*name)[:\s]*([A-Za-z\s]+)/i,
        /locality[:\s]*([A-Za-z\s]+)/i,
        /^20\s*([A-Za-z\s]+)/m
      ]
    };

    // Extract each field using multiple patterns
    for (const [field, regexList] of Object.entries(patterns)) {
      for (const regex of regexList) {
        const match = textContent.match(regex);
        if (match && match[1]) {
          let value: any = match[1].trim();
          
          // Clean and convert numeric values
          if (['wages', 'federalTaxWithheld', 'socialSecurityWages', 'socialSecurityTaxWithheld',
               'medicareWages', 'medicareTaxWithheld', 'socialSecurityTips', 'allocatedTips',
               'dependentCareBenefits', 'nonqualifiedPlans', 'stateWages', 'stateTaxWithheld',
               'localWages', 'localTaxWithheld'].includes(field)) {
            value = this.parseNumericValue(value);
            if (value !== null) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          } else {
            // String fields
            value = value.replace(/[^\w\s\-\.&',]/g, '').trim();
            if (value) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          }
        }
      }
    }

    return data;
  }

  async extract1099Int(result: any, textContent: string): Promise<Form1099IntData> {
    console.log("Extracting 1099-INT data...");
    
    const baseData: Form1099IntData = {
      documentType: "1099-INT",
      confidence: 0.8,
      extractedAt: new Date().toISOString(),
    };

    // Try Azure Document Intelligence structured extraction first
    const azureData = this.extract1099IntFromAzureFields(result);
    
    // Then try OCR-based extraction for missing fields
    const ocrData = this.extract1099IntFromOCR(textContent);
    
    // Merge the results, preferring Azure data when available
    const mergedData = { ...baseData, ...ocrData, ...azureData };
    
    console.log("1099-INT extraction completed:", JSON.stringify(mergedData, null, 2));
    return mergedData;
  }

  private extract1099IntFromAzureFields(result: any): Partial<Form1099IntData> {
    const data: Partial<Form1099IntData> = {};
    
    if (!result.documents || result.documents.length === 0) {
      return data;
    }

    const document = result.documents[0];
    const fields = document.fields || {};

    // Map Azure Document Intelligence fields to our 1099-INT structure
    const fieldMappings = {
      'PayerName': ['payerName', 'Payer', 'PayerInfo'],
      'PayerTIN': ['payerTIN', 'PayerTaxId', 'PayerEIN'],
      'RecipientName': ['recipientName', 'Recipient', 'RecipientInfo'],
      'RecipientTIN': ['recipientTIN', 'RecipientTaxId', 'RecipientSSN'],
      'InterestIncome': ['interestIncome', 'Interest', 'Box1'],
      'EarlyWithdrawalPenalty': ['earlyWithdrawalPenalty', 'Penalty', 'Box2'],
      'InterestOnUSSavingsBonds': ['interestOnUSSavingsBonds', 'USBonds', 'Box3'],
      'FederalTaxWithheld': ['federalTaxWithheld', 'FedTaxWithheld', 'Box4'],
      'InvestmentExpenses': ['investmentExpenses', 'Expenses', 'Box5'],
      'ForeignTaxPaid': ['foreignTaxPaid', 'ForeignTax', 'Box6'],
      'TaxExemptInterest': ['taxExemptInterest', 'TaxExempt', 'Box8'],
      'SpecifiedPrivateActivityBondInterest': ['specifiedPrivateActivityBondInterest', 'PrivateBond', 'Box9']
    };

    // Extract fields using multiple possible field names
    for (const [targetField, possibleNames] of Object.entries(fieldMappings)) {
      for (const fieldName of possibleNames) {
        if (fields[fieldName]) {
          const field = fields[fieldName];
          let value = field.content || field.value || field.valueString;
          
          // Convert numeric fields
          if (['interestIncome', 'earlyWithdrawalPenalty', 'interestOnUSSavingsBonds', 
               'federalTaxWithheld', 'investmentExpenses', 'foreignTaxPaid',
               'taxExemptInterest', 'specifiedPrivateActivityBondInterest'].includes(targetField.toLowerCase())) {
            value = this.parseNumericValue(value);
          }
          
          if (value !== null && value !== undefined) {
            (data as any)[targetField.toLowerCase()] = value;
            console.log(`Mapped Azure field ${fieldName} -> ${targetField}: ${value}`);
            break;
          }
        }
      }
    }

    return data;
  }

  private extract1099IntFromOCR(textContent: string): Partial<Form1099IntData> {
    console.log("Extracting 1099-INT data from OCR text...");
    const data: Partial<Form1099IntData> = {};

    // Enhanced regex patterns for 1099-INT extraction
    const patterns = {
      // Box 1 - Interest income
      interestIncome: [
        /(?:box\s*1|interest\s*income)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /interest\s*income[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^1\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2 - Early withdrawal penalty
      earlyWithdrawalPenalty: [
        /(?:box\s*2|early\s*withdrawal\s*penalty)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /early\s*withdrawal[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 3 - Interest on U.S. Savings Bonds
      interestOnUSSavingsBonds: [
        /(?:box\s*3|interest\s*on\s*u\.?s\.?\s*savings\s*bonds)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /u\.?s\.?\s*savings\s*bonds[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^3\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 4 - Federal income tax withheld
      federalTaxWithheld: [
        /(?:box\s*4|federal\s*income\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /federal\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^4\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Payer information
      payerName: [
        /(?:payer|from)[:\s]*([A-Za-z0-9\s,\.&'-]+?)(?:\n|$|address|tin)/i,
        /^([A-Za-z0-9\s,\.&'-]+?)(?:\n.*address|\n.*tin)/im
      ],
      
      // Recipient information
      recipientName: [
        /(?:recipient|to)[:\s]*([A-Za-z\s,\.'-]+?)(?:\n|$|address|tin)/i,
        /recipient[:\s]*([A-Za-z\s,\.'-]+)/i
      ],
      
      // TIN numbers
      payerTIN: [
        /(?:payer.*tin|payer.*tax.*id)[:\s]*([0-9-]+)/i,
        /payer.*ein[:\s]*([0-9-]+)/i
      ],
      
      recipientTIN: [
        /(?:recipient.*tin|recipient.*ssn)[:\s]*([0-9-]+)/i,
        /recipient.*tax.*id[:\s]*([0-9-]+)/i
      ]
    };

    // Extract each field using multiple patterns
    for (const [field, regexList] of Object.entries(patterns)) {
      for (const regex of regexList) {
        const match = textContent.match(regex);
        if (match && match[1]) {
          let value: any = match[1].trim();
          
          // Clean and convert numeric values
          if (['interestIncome', 'earlyWithdrawalPenalty', 'interestOnUSSavingsBonds',
               'federalTaxWithheld', 'investmentExpenses', 'foreignTaxPaid',
               'taxExemptInterest', 'specifiedPrivateActivityBondInterest'].includes(field)) {
            value = this.parseNumericValue(value);
            if (value !== null) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          } else {
            // String fields
            value = value.replace(/[^\w\s\-\.&',]/g, '').trim();
            if (value) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          }
        }
      }
    }

    return data;
  }

  async extract1099Div(result: any, textContent: string): Promise<Form1099DivData> {
    console.log("Extracting 1099-DIV data...");
    
    const baseData: Form1099DivData = {
      documentType: "1099-DIV",
      confidence: 0.8,
      extractedAt: new Date().toISOString(),
    };

    // Try Azure Document Intelligence structured extraction first
    const azureData = this.extract1099DivFromAzureFields(result);
    
    // Then try OCR-based extraction for missing fields
    const ocrData = this.extract1099DivFromOCR(textContent);
    
    // Merge the results, preferring Azure data when available
    const mergedData = { ...baseData, ...ocrData, ...azureData };
    
    console.log("1099-DIV extraction completed:", JSON.stringify(mergedData, null, 2));
    return mergedData;
  }

  private extract1099DivFromAzureFields(result: any): Partial<Form1099DivData> {
    const data: Partial<Form1099DivData> = {};
    
    if (!result.documents || result.documents.length === 0) {
      return data;
    }

    const document = result.documents[0];
    const fields = document.fields || {};

    // Map Azure Document Intelligence fields to our 1099-DIV structure
    const fieldMappings = {
      'PayerName': ['payerName', 'Payer', 'PayerInfo'],
      'PayerTIN': ['payerTIN', 'PayerTaxId', 'PayerEIN'],
      'RecipientName': ['recipientName', 'Recipient', 'RecipientInfo'],
      'RecipientTIN': ['recipientTIN', 'RecipientTaxId', 'RecipientSSN'],
      'TotalOrdinaryDividends': ['totalOrdinaryDividends', 'OrdinaryDividends', 'Box1a'],
      'QualifiedDividends': ['qualifiedDividends', 'Qualified', 'Box1b'],
      'TotalCapitalGainDistributions': ['totalCapitalGainDistributions', 'CapitalGains', 'Box2a'],
      'Unrecaptured1250Gain': ['unrecaptured1250Gain', 'Section1250', 'Box2b'],
      'Section1202Gain': ['section1202Gain', 'Section1202', 'Box2c'],
      'CollectiblesGain': ['collectiblesGain', 'Collectibles', 'Box2d'],
      'NondividendDistributions': ['nondividendDistributions', 'Nondividend', 'Box3'],
      'FederalTaxWithheld': ['federalTaxWithheld', 'FedTaxWithheld', 'Box4'],
      'Section199ADividends': ['section199ADividends', 'Section199A', 'Box5'],
      'InvestmentExpenses': ['investmentExpenses', 'Expenses', 'Box6'],
      'ForeignTaxPaid': ['foreignTaxPaid', 'ForeignTax', 'Box7'],
      'CashLiquidationDistributions': ['cashLiquidationDistributions', 'CashLiquidation', 'Box8'],
      'NoncashLiquidationDistributions': ['noncashLiquidationDistributions', 'NoncashLiquidation', 'Box9'],
      'ExemptInterestDividends': ['exemptInterestDividends', 'ExemptInterest', 'Box10'],
      'SpecifiedPrivateActivityBondInterestDividends': ['specifiedPrivateActivityBondInterestDividends', 'PrivateBond', 'Box11']
    };

    // Extract fields using multiple possible field names
    for (const [targetField, possibleNames] of Object.entries(fieldMappings)) {
      for (const fieldName of possibleNames) {
        if (fields[fieldName]) {
          const field = fields[fieldName];
          let value = field.content || field.value || field.valueString;
          
          // Convert numeric fields
          if (['totalOrdinaryDividends', 'qualifiedDividends', 'totalCapitalGainDistributions',
               'unrecaptured1250Gain', 'section1202Gain', 'collectiblesGain', 'nondividendDistributions',
               'federalTaxWithheld', 'section199ADividends', 'investmentExpenses', 'foreignTaxPaid',
               'cashLiquidationDistributions', 'noncashLiquidationDistributions', 'exemptInterestDividends',
               'specifiedPrivateActivityBondInterestDividends'].includes(targetField.toLowerCase())) {
            value = this.parseNumericValue(value);
          }
          
          if (value !== null && value !== undefined) {
            (data as any)[targetField.toLowerCase()] = value;
            console.log(`Mapped Azure field ${fieldName} -> ${targetField}: ${value}`);
            break;
          }
        }
      }
    }

    return data;
  }

  private extract1099DivFromOCR(textContent: string): Partial<Form1099DivData> {
    console.log("Extracting 1099-DIV data from OCR text...");
    const data: Partial<Form1099DivData> = {};

    // Enhanced regex patterns for 1099-DIV extraction
    const patterns = {
      // Box 1a - Total ordinary dividends
      totalOrdinaryDividends: [
        /(?:box\s*1a|total\s*ordinary\s*dividends)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /ordinary\s*dividends[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^1a\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 1b - Qualified dividends
      qualifiedDividends: [
        /(?:box\s*1b|qualified\s*dividends)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /qualified\s*dividends[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^1b\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2a - Total capital gain distributions
      totalCapitalGainDistributions: [
        /(?:box\s*2a|total\s*capital\s*gain\s*distributions)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /capital\s*gain\s*distributions[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2a\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2b - Unrecaptured Section 1250 gain
      unrecaptured1250Gain: [
        /(?:box\s*2b|unrecaptured\s*section\s*1250\s*gain)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /section\s*1250\s*gain[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2b\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2c - Section 1202 gain
      section1202Gain: [
        /(?:box\s*2c|section\s*1202\s*gain)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /section\s*1202[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2c\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2d - Collectibles (28%) gain
      collectiblesGain: [
        /(?:box\s*2d|collectibles.*gain)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /collectibles[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2d\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 3 - Nondividend distributions
      nondividendDistributions: [
        /(?:box\s*3|nondividend\s*distributions)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /nondividend[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^3\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 4 - Federal income tax withheld
      federalTaxWithheld: [
        /(?:box\s*4|federal\s*income\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /federal\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^4\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 5 - Section 199A dividends
      section199ADividends: [
        /(?:box\s*5|section\s*199a\s*dividends)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /section\s*199a[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^5\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Payer information
      payerName: [
        /(?:payer|from)[:\s]*([A-Za-z0-9\s,\.&'-]+?)(?:\n|$|address|tin)/i,
        /^([A-Za-z0-9\s,\.&'-]+?)(?:\n.*address|\n.*tin)/im
      ],
      
      // Recipient information
      recipientName: [
        /(?:recipient|to)[:\s]*([A-Za-z\s,\.'-]+?)(?:\n|$|address|tin)/i,
        /recipient[:\s]*([A-Za-z\s,\.'-]+)/i
      ]
    };

    // Extract each field using multiple patterns
    for (const [field, regexList] of Object.entries(patterns)) {
      for (const regex of regexList) {
        const match = textContent.match(regex);
        if (match && match[1]) {
          let value: any = match[1].trim();
          
          // Clean and convert numeric values
          if (['totalOrdinaryDividends', 'qualifiedDividends', 'totalCapitalGainDistributions',
               'unrecaptured1250Gain', 'section1202Gain', 'collectiblesGain', 'nondividendDistributions',
               'federalTaxWithheld', 'section199ADividends', 'investmentExpenses', 'foreignTaxPaid',
               'cashLiquidationDistributions', 'noncashLiquidationDistributions', 'exemptInterestDividends',
               'specifiedPrivateActivityBondInterestDividends'].includes(field)) {
            value = this.parseNumericValue(value);
            if (value !== null) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          } else {
            // String fields
            value = value.replace(/[^\w\s\-\.&',]/g, '').trim();
            if (value) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          }
        }
      }
    }

    return data;
  }

  async extract1099Misc(result: any, textContent: string): Promise<Form1099MiscData> {
    console.log("Extracting 1099-MISC data...");
    
    const baseData: Form1099MiscData = {
      documentType: "1099-MISC",
      confidence: 0.8,
      extractedAt: new Date().toISOString(),
    };

    // Try Azure Document Intelligence structured extraction first
    const azureData = this.extract1099MiscFromAzureFields(result);
    
    // Then try OCR-based extraction for missing fields
    const ocrData = this.extract1099MiscFromOCR(textContent);
    
    // Merge the results, preferring Azure data when available
    const mergedData = { ...baseData, ...ocrData, ...azureData };
    
    console.log("1099-MISC extraction completed:", JSON.stringify(mergedData, null, 2));
    return mergedData;
  }

  private extract1099MiscFromAzureFields(result: any): Partial<Form1099MiscData> {
    const data: Partial<Form1099MiscData> = {};
    
    if (!result.documents || result.documents.length === 0) {
      return data;
    }

    const document = result.documents[0];
    const fields = document.fields || {};

    // Map Azure Document Intelligence fields to our 1099-MISC structure
    const fieldMappings = {
      'PayerName': ['payerName', 'Payer', 'PayerInfo'],
      'PayerTIN': ['payerTIN', 'PayerTaxId', 'PayerEIN'],
      'RecipientName': ['recipientName', 'Recipient', 'RecipientInfo'],
      'RecipientTIN': ['recipientTIN', 'RecipientTaxId', 'RecipientSSN'],
      'Rents': ['rents', 'RentIncome', 'Box1'],
      'Royalties': ['royalties', 'RoyaltyIncome', 'Box2'],
      'OtherIncome': ['otherIncome', 'Other', 'Box3'],
      'FederalTaxWithheld': ['federalTaxWithheld', 'FedTaxWithheld', 'Box4'],
      'FishingBoatProceeds': ['fishingBoatProceeds', 'FishingBoat', 'Box5'],
      'MedicalHealthcarePayments': ['medicalHealthcarePayments', 'Medical', 'Box6'],
      'NonemployeeCompensation': ['nonemployeeCompensation', 'Nonemployee', 'Box7'],
      'SubstitutePayments': ['substitutePayments', 'Substitute', 'Box8'],
      'CropInsuranceProceeds': ['cropInsuranceProceeds', 'CropInsurance', 'Box9'],
      'GrossProceedsAttorney': ['grossProceedsAttorney', 'Attorney', 'Box10'],
      'Section409ADeferrals': ['section409ADeferrals', 'Section409A', 'Box12'],
      'Section409AIncome': ['section409AIncome', 'Section409AIncome', 'Box14']
    };

    // Extract fields using multiple possible field names
    for (const [targetField, possibleNames] of Object.entries(fieldMappings)) {
      for (const fieldName of possibleNames) {
        if (fields[fieldName]) {
          const field = fields[fieldName];
          let value = field.content || field.value || field.valueString;
          
          // Convert numeric fields
          if (['rents', 'royalties', 'otherIncome', 'federalTaxWithheld', 'fishingBoatProceeds',
               'medicalHealthcarePayments', 'nonemployeeCompensation', 'substitutePayments',
               'cropInsuranceProceeds', 'grossProceedsAttorney', 'section409ADeferrals',
               'section409AIncome'].includes(targetField.toLowerCase())) {
            value = this.parseNumericValue(value);
          }
          
          if (value !== null && value !== undefined) {
            (data as any)[targetField.toLowerCase()] = value;
            console.log(`Mapped Azure field ${fieldName} -> ${targetField}: ${value}`);
            break;
          }
        }
      }
    }

    return data;
  }

  private extract1099MiscFromOCR(textContent: string): Partial<Form1099MiscData> {
    console.log("Extracting 1099-MISC data from OCR text...");
    const data: Partial<Form1099MiscData> = {};

    // Enhanced regex patterns for 1099-MISC extraction
    const patterns = {
      // Box 1 - Rents
      rents: [
        /(?:box\s*1|rents)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /rents[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^1\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 2 - Royalties
      royalties: [
        /(?:box\s*2|royalties)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /royalties[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^2\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 3 - Other income
      otherIncome: [
        /(?:box\s*3|other\s*income)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /other\s*income[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^3\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 4 - Federal income tax withheld
      federalTaxWithheld: [
        /(?:box\s*4|federal\s*income\s*tax\s*withheld)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /federal\s*income\s*tax[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^4\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 5 - Fishing boat proceeds
      fishingBoatProceeds: [
        /(?:box\s*5|fishing\s*boat\s*proceeds)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /fishing\s*boat[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^5\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 6 - Medical and health care payments
      medicalHealthcarePayments: [
        /(?:box\s*6|medical.*health.*care.*payments)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /medical.*payments[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^6\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Box 7 - Nonemployee compensation
      nonemployeeCompensation: [
        /(?:box\s*7|nonemployee\s*compensation)[:\s]*\$?([0-9,]+\.?\d*)/i,
        /nonemployee[:\s]*\$?([0-9,]+\.?\d*)/i,
        /^7\s*\$?([0-9,]+\.?\d*)/m
      ],
      
      // Payer information
      payerName: [
        /(?:payer|from)[:\s]*([A-Za-z0-9\s,\.&'-]+?)(?:\n|$|address|tin)/i,
        /^([A-Za-z0-9\s,\.&'-]+?)(?:\n.*address|\n.*tin)/im
      ],
      
      // Recipient information
      recipientName: [
        /(?:recipient|to)[:\s]*([A-Za-z\s,\.'-]+?)(?:\n|$|address|tin)/i,
        /recipient[:\s]*([A-Za-z\s,\.'-]+)/i
      ]
    };

    // Extract each field using multiple patterns
    for (const [field, regexList] of Object.entries(patterns)) {
      for (const regex of regexList) {
        const match = textContent.match(regex);
        if (match && match[1]) {
          let value: any = match[1].trim();
          
          // Clean and convert numeric values
          if (['rents', 'royalties', 'otherIncome', 'federalTaxWithheld', 'fishingBoatProceeds',
               'medicalHealthcarePayments', 'nonemployeeCompensation', 'substitutePayments',
               'cropInsuranceProceeds', 'grossProceedsAttorney', 'section409ADeferrals',
               'section409AIncome'].includes(field)) {
            value = this.parseNumericValue(value);
            if (value !== null) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          } else {
            // String fields
            value = value.replace(/[^\w\s\-\.&',]/g, '').trim();
            if (value) {
              (data as any)[field] = value;
              console.log(`OCR extracted ${field}: ${value}`);
              break;
            }
          }
        }
      }
    }

    return data;
  }

  private parseNumericValue(value: string): number | null {
    if (!value) return null;
    
    // Remove currency symbols, commas, and extra spaces
    const cleaned = value.toString().replace(/[$,\s]/g, '');
    
    // Check if it's a valid number
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? null : parsed;
  }

  private parseBooleanValue(value: string): boolean {
    if (!value) return false;
    
    const cleaned = value.toString().toLowerCase().trim();
    return ['true', 'yes', '1', 'x', 'checked'].includes(cleaned);
  }
}

// Export the type alias for ExtractedFieldData
export type ExtractedFieldData = TaxDocumentData;

// Export the factory function
export function getAzureDocumentIntelligenceService(): AzureDocumentIntelligenceService {
  return new AzureDocumentIntelligenceService();
}
