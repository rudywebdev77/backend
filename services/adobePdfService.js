import {
  PDFServices,
  ServicePrincipalCredentials,
  ClientConfig,
  ExportPDFJob,
  ExportPDFParams,
  ExportPDFTargetFormat,
  ExportOCRLocale,
  ExportPDFResult,
  MimeType,
  SDKError,
  ServiceUsageError,
  ServiceApiError,
} from '@adobe/pdfservices-node-sdk';
import fs from 'fs';

export const isAdobePdfConfigured = () => {
  return (
    Boolean(process.env.PDF_SERVICES_CLIENT_ID) &&
    Boolean(process.env.PDF_SERVICES_CLIENT_SECRET) &&
    process.env.PDF_SERVICES_CLIENT_ID.trim() !== '' &&
    process.env.PDF_SERVICES_CLIENT_SECRET.trim() !== ''
  );
};

export const convertPdfToDocxWithAdobe = async (pdfPath) => {
  if (!isAdobePdfConfigured()) {
    throw new Error('Adobe PDF Services API credentials are not configured.');
  }

  let readStream = null;
  try {
    const credentials = new ServicePrincipalCredentials({
      clientId: process.env.PDF_SERVICES_CLIENT_ID.trim(),
      clientSecret: process.env.PDF_SERVICES_CLIENT_SECRET.trim(),
    });

    const clientConfig = new ClientConfig({
      connectTimeout: 30000,
      readTimeout: 60000,
    });

    const pdfServices = new PDFServices({ credentials, clientConfig });

    readStream = fs.createReadStream(pdfPath);
    const inputAsset = await pdfServices.upload({
      readStream,
      mimeType: MimeType.PDF,
    });

    const params = new ExportPDFParams({
      targetFormat: ExportPDFTargetFormat.DOCX,
      ocrLocale: ExportOCRLocale.EN_US,
    });

    const job = new ExportPDFJob({ inputAsset, params });

    console.log('[Adobe PDF Services] Submitting PDF to Word export job...');
    const pollingURL = await pdfServices.submit({ job });
    const pdfServicesResponse = await pdfServices.getJobResult({
      pollingURL,
      resultType: ExportPDFResult,
    });

    const resultAsset = pdfServicesResponse.result.asset;
    const contentStream = await pdfServices.getContent({ asset: resultAsset });

    const chunks = [];
    for await (const chunk of contentStream.readStream) {
      chunks.push(chunk);
    }
    const docxBuffer = Buffer.concat(chunks);

    console.log(`[Adobe PDF Services] Successfully generated DOCX (${docxBuffer.length} bytes).`);
    return docxBuffer;
  } catch (error) {
    if (error instanceof SDKError) {
      console.error('[Adobe PDF Services] SDK Error:', error.message);
    } else if (error instanceof ServiceUsageError) {
      console.error('[Adobe PDF Services] Service Usage / Quota Error:', error.message);
    } else if (error instanceof ServiceApiError) {
      console.error('[Adobe PDF Services] API Error:', error.message, 'StatusCode:', error.statusCode);
    } else {
      console.error('[Adobe PDF Services] Unexpected Error:', error.message);
    }
    throw error;
  } finally {
    if (readStream && !readStream.destroyed) {
      readStream.destroy();
    }
  }
};