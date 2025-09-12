import { DocumentAnalysisClient, AzureKeyCredential } from '@azure/ai-form-recognizer';
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';

// Types for structured output
interface ExtractedField {
  value: any;
  confidence?: number;
  type?: 'object' | 'array' | 'address' | 'simple';
}

export interface ExtractedFieldData {
  [key: string]: any;
  fullText?: string;
  correctedDocumentType?: string;
}

interface ProcessedDocument {
  [fieldName: string]: ExtractedField;
}

interface FileProcessingResult {
  filename: string;
  model_used: string;
  success: boolean;
  document_count: number;
  documents: ProcessedDocument[];
}

interface AzureDocumentIntelligenceConfig {
  endpoint: string;
  key: string;
  apiVersion?: string;
}

export class ComprehensiveAzureExtractionService {
  private client: DocumentAnalysisClient;
  private seenDocuments: Set<string> = new Set();

  constructor(config: AzureDocumentIntelligenceConfig) {
    const credential = new AzureKeyCredential(config.key);
    this.client = new DocumentAnalysisClient(config.endpoint, credential);
  }

  /**
   * Smart model selection based on filename patterns (mirrors Python logic)
   */
  private getModelFromFilename(filename: string): string {
    const filenameLower = filename.toLowerCase();

    // Define filename patterns and their corresponding models (exact Python mapping)
    const modelMappings: Array<[RegExp, string]> = [
      [/1099.*div/i, "prebuilt-tax.us.1099DIV"],
      [/1099.*misc/i, "prebuilt-tax.us.1099MISC"],
      [/1099.*int/i, "prebuilt-tax.us.1099INT"],
      [/1099.*nec/i, "prebuilt-tax.us.1099NEC"],
      [/w.*2/i, "prebuilt-tax.us.w2"],
      [/1099/i, "prebuilt-tax.us.1099"],  // Generic 1099 fallback
      [/1040/i, "prebuilt-tax.us.1040"]
    ];

    // Check each pattern
    for (const [pattern, model] of modelMappings) {
      if (pattern.test(filenameLower)) {
        return model;
      }
    }

    // If no specific pattern matches, return the unified tax model
    return "prebuilt-tax.us";
  }

  /**
   * Comprehensive field extraction with recursive handling (mirrors Python extract_field_value)
   */
  private extractFieldValue(field: any): ExtractedField | null {
    if (!field) {
      return null;
    }

    // Get confidence if available
    const confidence = field.confidence || undefined;

    // Handle different field types based on Azure SDK structure
    if (field.properties && typeof field.properties === 'object') {
      // Object type - recursively extract fields
      const objDict: { [key: string]: ExtractedField | null } = {};
      for (const [key, nestedField] of Object.entries(field.properties)) {
        objDict[key] = this.extractFieldValue(nestedField);
      }
      return {
        type: "object",
        value: objDict,
        confidence
      };
    }

    if (field.values && Array.isArray(field.values)) {
      // Array type - extract each item
      const arrayList: (ExtractedField | null)[] = [];
      for (const item of field.values) {
        arrayList.push(this.extractFieldValue(item));
      }
      return {
        type: "array",
        value: arrayList,
        confidence
      };
    }

    // Handle address type (check for address-specific properties)
    if (field.value && typeof field.value === 'object' && 
        (field.value.streetAddress || field.value.city || field.value.state || field.value.postalCode)) {
      const addressDict: { [key: string]: any } = {};
      for (const [key, value] of Object.entries(field.value)) {
        if (value !== undefined && value !== null) {
          addressDict[key] = value;
        }
      }
      return {
        type: "address",
        value: addressDict,
        confidence
      };
    }

    // Simple value (string, number, date, etc.)
    return {
      value: field.value || field.content || field,
      confidence
    };
  }

  /**
   * Create unique document identifier for duplicate detection
   */
  private createDocumentKey(doc: any): string | null {
    let ssn: string | null = null;
    let employerId: string | null = null;
    let taxYear: string | null = null;

    try {
      const fields = doc.fields || {};

      // Extract W-2 identifying fields
      if (fields.Employee?.properties?.SocialSecurityNumber?.value) {
        ssn = fields.Employee.properties.SocialSecurityNumber.value;
      }
      if (fields.Employer?.properties?.IdNumber?.value) {
        employerId = fields.Employer.properties.IdNumber.value;
      }
      if (fields.TaxYear?.value) {
        taxYear = fields.TaxYear.value.toString();
      }

      // Extract 1099 identifying fields
      if (fields.Recipient?.properties?.TaxIdNumber?.value) {
        ssn = fields.Recipient.properties.TaxIdNumber.value;
      }
      if (fields.Payer?.properties?.IdNumber?.value) {
        employerId = fields.Payer.properties.IdNumber.value;
      }

      // Create unique key if we have all required components
      if (ssn && employerId && taxYear) {
        return `${ssn}_${employerId}_${taxYear}`;
      }
    } catch (error) {
      console.warn('Error creating document key:', error);
    }

    return null;
  }

  /**
   * Process a single document with comprehensive model fallback strategy
   */
  private async processSingleFile(filePath: string): Promise<FileProcessingResult> {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`PROCESSING: ${path.basename(filePath)}`);
    console.log(`${'='.repeat(60)}`);

    const filename = path.basename(filePath);
    const primaryModel = this.getModelFromFilename(filename);
    console.log(`Selected primary model based on filename: ${primaryModel}`);

    // Comprehensive fallback model list (mirrors Python logic)
    const modelCandidates = [
      primaryModel,                    // Filename-based selection (first priority)
      "prebuilt-tax.us",              // unified tax model
      "prebuilt-tax.us.1099DIV",
      "prebuilt-tax.us.1099MISC",
      "prebuilt-tax.us.1099INT",
      "prebuilt-tax.us.1099NEC",
      "prebuilt-tax.us.w2",           // W-2 specific
      "prebuilt-tax.us.1099",         // 1099 family
      "prebuilt-tax.us.1040",         // 1040 family
    ];

    // Remove duplicates while preserving order
    const uniqueModels = [...new Set(modelCandidates)];

    let result: any = null;
    let usedModel: string = '';
    let lastError: string = '';

    // Read file content
    const fileContent = fs.readFileSync(filePath);

    for (const modelId of uniqueModels) {
      console.log(`\nTrying model: ${modelId} ...`);
      try {
        const poller = await this.client.beginAnalyzeDocument(modelId, fileContent);
        result = await poller.pollUntilDone();
        console.log('✓ Success — used model:', modelId);
        
        if (modelId === primaryModel) {
          console.log('✓ Used the filename-based model selection!');
        }
        
        usedModel = modelId;
        break;

      } catch (error: any) {
        if (error.code === 'ModelNotFound' || error.message?.includes('not found')) {
          console.log(`ModelNotFound for ${modelId} — not available in your resource/region.`);
          continue;
        }
        
        lastError = error.message || error.toString();
        console.log(`Error with ${modelId}: ${lastError}`);
        continue;
      }
    }

    if (!result) {
      console.log(`\nNo model succeeded for ${filename}. Last error:\n`, lastError);
      return {
        filename,
        model_used: '',
        success: false,
        document_count: 0,
        documents: []
      };
    }

    // Process documents with duplicate detection
    const fileDocs: ProcessedDocument[] = [];
    const localSeenDocuments: Set<string> = new Set();

    for (let i = 0; i < result.documents.length; i++) {
      const doc = result.documents[i];
      console.log(`\n--- Document #${i + 1} in ${filename} (documentType: ${doc.docType || 'n/a'}) ---`);

      // Check for duplicates
      const documentKey = this.createDocumentKey(doc);
      let isDuplicate = false;

      if (documentKey) {
        if (this.seenDocuments.has(documentKey) || localSeenDocuments.has(documentKey)) {
          isDuplicate = true;
          console.log('🔄 DUPLICATE DETECTED - Skipping this copy (same SSN/TIN + Employer/Payer + Tax Year)');
        } else {
          this.seenDocuments.add(documentKey);
          localSeenDocuments.add(documentKey);
        }
      }

      if (isDuplicate) {
        continue;
      }

      // Extract all fields comprehensively
      const docDict: ProcessedDocument = {};
      
      if (doc.fields) {
        for (const [fieldName, field] of Object.entries(doc.fields)) {
          const extractedField = this.extractFieldValue(field);
          const value = extractedField?.value;
          const confidence = extractedField?.confidence;

          console.log(`${fieldName}: ${JSON.stringify(value)} (confidence: ${confidence})`);
          docDict[fieldName] = extractedField || { value: null };
        }
      }

      fileDocs.push(docDict);
    }

    return {
      filename,
      model_used: usedModel,
      success: true,
      document_count: fileDocs.length,
      documents: fileDocs
    };
  }

  /**
   * Process multiple files and generate comprehensive JSON output
   */
  public async processFiles(filePaths: string[]): Promise<FileProcessingResult[]> {
    console.log(`Processing ${filePaths.length} file(s): ${filePaths.map(f => path.basename(f))}`);

    const allProcessedFiles: FileProcessingResult[] = [];

    for (const filePath of filePaths) {
      try {
        const result = await this.processSingleFile(filePath);
        
        if (result.success) {
          // Save JSON for each file
          const outputPath = `${filePath}.parsed.json`;
          const outputData = {
            filename: result.filename,
            model_used: result.model_used,
            documents: result.documents
          };

          fs.writeFileSync(outputPath, JSON.stringify(outputData, null, 2));
          console.log(`\n✓ Saved parsed JSON to ${outputPath}`);
        } else {
          console.log(`❌ Failed to process ${result.filename}`);
        }

        allProcessedFiles.push(result);

      } catch (error) {
        console.error(`Error processing ${filePath}:`, error);
        allProcessedFiles.push({
          filename: path.basename(filePath),
          model_used: '',
          success: false,
          document_count: 0,
          documents: []
        });
      }
    }

    // Print comprehensive summary
    this.printProcessingSummary(allProcessedFiles);

    return allProcessedFiles;
  }

  /**
   * Process a single file (public method for individual file processing)
   */
  public async processSingleDocument(filePath: string): Promise<FileProcessingResult> {
    return this.processSingleFile(filePath);
  }

  /**
   * Print processing summary (mirrors Python summary output)
   */
  private printProcessingSummary(results: FileProcessingResult[]): void {
    console.log(`\n${'='.repeat(60)}`);
    console.log('PROCESSING COMPLETE - SUMMARY');
    console.log(`${'='.repeat(60)}`);

    const successfulFiles = results.filter(f => f.success);
    console.log(`Successfully processed ${successfulFiles.length} out of ${results.length} files:`);

    for (const fileInfo of successfulFiles) {
      console.log(`  📄 ${fileInfo.filename} → ${fileInfo.model_used} (${fileInfo.document_count} docs)`);
    }

    if (successfulFiles.length === 0) {
      console.log('\nNo files were successfully processed. Debug steps:');
      console.log('- Confirm you created a Document Intelligence resource (not Computer Vision).');
      console.log('- Confirm the resource is in a region that supports tax models (try East US / West US2).');
      console.log('- Open Document Intelligence Studio and test "Prebuilt -> Tax (US)".');
    }
  }

  /**
   * Reset duplicate document tracking (useful for processing new batches)
   */
  public resetDuplicateTracking(): void {
    this.seenDocuments.clear();
  }

  /**
   * Get comprehensive extraction statistics
   */
  public getExtractionStats(): {
    totalDocumentsProcessed: number;
    uniqueDocumentsFound: number;
    duplicatesSkipped: number;
  } {
    return {
      totalDocumentsProcessed: this.seenDocuments.size,
      uniqueDocumentsFound: this.seenDocuments.size,
      duplicatesSkipped: 0 // Would need additional tracking for exact count
    };
  }
}

// Example usage and CLI interface
export async function main() {
  // Configuration - replace with your actual values
  const config: AzureDocumentIntelligenceConfig = {
    endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT || "https://your-resource.cognitiveservices.azure.com/",
    key: process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY || "your-api-key-here"
  };

  const extractionService = new ComprehensiveAzureExtractionService(config);

  // Example: Process files from command line arguments
  const filePaths = process.argv.slice(2);
  
  if (filePaths.length === 0) {
    console.log('Usage: npx ts-node comprehensive-azure-extraction-service.ts <file1> [file2] [file3] ...');
    console.log('Example: npx ts-node comprehensive-azure-extraction-service.ts ./w2-2023.pdf ./1099-div-2023.pdf');
    process.exit(1);
  }

  try {
    const results = await extractionService.processFiles(filePaths);
    
    // Print final statistics
    const stats = extractionService.getExtractionStats();
    console.log('\n📊 Final Statistics:');
    console.log(`   Unique documents processed: ${stats.uniqueDocumentsFound}`);
    
  } catch (error) {
    console.error('Error during processing:', error);
    process.exit(1);
  }
}
// Factory function for backward compatibility
export function getAzureDocumentIntelligenceService(): ComprehensiveAzureExtractionService {
  const config = {
    endpoint: process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT || '',
    key: process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY || ''
  };
  
  if (!config.endpoint || !config.key) {
    throw new Error('Azure Document Intelligence configuration is missing. Please set AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT and AZURE_DOCUMENT_INTELLIGENCE_KEY environment variables.');
  }
  
  return new ComprehensiveAzureExtractionService(config);
  /**
 * Backward compatibility method for existing route code
 */
async extractDataFromDocument(filePath: string, documentType: string): Promise<ExtractedFieldData & { correctedDocumentType?: string }> {
  try {
    // Use the comprehensive extraction
    const result = await this.processSingleDocument(filePath);
    
    if (!result || !result.success) {
      throw new Error('Failed to extract data from document');
    }
    
    // Extract the first document's data (assuming single document per file)
    const firstDoc = result.documents[0];
    if (!firstDoc) {
      throw new Error('No document data extracted');
    }
    
    // Flatten the structured data to match expected format
    const extractedData: ExtractedFieldData = {};
    
    // Convert the comprehensive extraction format to flat format
    for (const [key, fieldData] of Object.entries(firstDoc)) {
      if (fieldData && typeof fieldData === 'object' && 'value' in fieldData) {
        // Handle different field types
        if (fieldData.type === 'object' && fieldData.value) {
          // For object types, flatten the nested structure
          const objValue = fieldData.value as any;
          for (const [nestedKey, nestedValue] of Object.entries(objValue)) {
            if (nestedValue && typeof nestedValue === 'object' && 'value' in (nestedValue as any)) {
              extractedData[`${key}_${nestedKey}`] = (nestedValue as any).value;
            }
          }
        } else {
          // For simple types, use the value directly
          extractedData[key] = fieldData.value;
        }
      }
    }
    
    // Add fullText if available (you might need to extract this from the original result)
    // This would require storing the OCR text from the Azure response
    extractedData.fullText = ''; // You can enhance this later
    
    // Add document type correction if model changed
    const correctedDocumentType = this.mapModelToDocumentType(result.model_used);
    if (correctedDocumentType !== documentType) {
      extractedData.correctedDocumentType = correctedDocumentType;
    }
    
    return extractedData;
    
  } catch (error) {
    console.error('Error in extractDataFromDocument:', error);
    throw error;
  }
}

/**
 * Helper method to map Azure model names to document types
 */
private mapModelToDocumentType(modelId: string): string {
  const modelMappings: Record<string, string> = {
    'prebuilt-tax.us.w2': 'W2',
    'prebuilt-tax.us.1099DIV': 'FORM_1099_DIV',
    'prebuilt-tax.us.1099MISC': 'FORM_1099_MISC',
    'prebuilt-tax.us.1099INT': 'FORM_1099_INT',
    'prebuilt-tax.us.1099NEC': 'FORM_1099_NEC',
    'prebuilt-tax.us.1099': 'FORM_1099_MISC', // fallback
    'prebuilt-tax.us.1040': '1040',
    'prebuilt-tax.us': 'UNKNOWN'
  };
  
  return modelMappings[modelId] || 'UNKNOWN';
}
}

// Types are already exported above - no need to re-export

// Run main function if this file is executed directly
if (require.main === module) {
  main().catch(console.error);
}
