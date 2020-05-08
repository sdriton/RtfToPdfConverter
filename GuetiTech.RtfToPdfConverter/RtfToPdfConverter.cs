/*
 * Copyright © Driton Salihu, 2010
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;


/**
 * In order to use the Office automation you have to:
 *  1. Go to Administrative Tools | Component Services
 *  2. Expand Desktop or My computer
 *  4. Then expand DCOM configuration
 *  5. Open "Word 97 - 2003 Office Documents" properties.
 *     In Windows Server 2008 it would be {000C101C-0000-0000-C000-000000000046} or {00021401-0000-0000-C000-000000000046}.
 *  6. Go to tab "Identity" and set an administrator acount to run the COM application (Word)
 *  7. In the security tab change the authorizations to custom by adding users that are allowed to use this application (word DCOM)
 * 
 */


namespace GuetiTech
{
    public class RtfToPdfConverter
    {

        public static string ConvertRtfToPdf(byte[] rtfFileContent, bool deleteTmpFile = true)
        { 
            return ConvertRtfToPdf("", rtfFileContent, deleteTmpFile);
        }

        public static string ConvertRtfToPdf(string fileName, byte[] rtfFileContent, bool deleteTmpFile = true)
        {
            string tempPath = System.IO.Path.GetTempPath();
            string pdfBase64Content = "";

            string randomFileName = System.IO.Path.GetRandomFileName();

            string datetimestring = DateTime.Now.ToFileTime() + "";
            string fullRtfFileName = tempPath + fileName + datetimestring + randomFileName + ".rtf";
            string fullPdfFileName = tempPath + fileName + datetimestring + randomFileName + ".pdf";


            // Create work folder if it does not exist already
            bool isTempFolderCreated = CreateTempFolderIfNotExists(tempPath);

            if (isTempFolderCreated)
            {
                // perform conversion steps
                SaveRtfToFile(fullRtfFileName, rtfFileContent);
                ExportToPdf(fullRtfFileName, fullPdfFileName);

                // read the pdf document bytes
                pdfBase64Content = ReadPdfContent(fullPdfFileName);

                //cleanup temporary files.
                if (deleteTmpFile)
                {
                    DeleteTempFile(fullRtfFileName);
                    DeleteTempFile(fullPdfFileName);
                }
            }

            return pdfBase64Content;
        }


        private static bool CreateTempFolderIfNotExists(string tempFolder)
        {
            try
            {
                if (!System.IO.Directory.Exists(tempFolder))
                {
                    System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(tempFolder);
                }
            }
            catch (Exception e)
            {
                string msg = e.Message;
                return false;
            }

            return true;
        }

        private static string SaveRtfToFile(string fullFileName, byte[] rtfFileContent) 
        {
            try
            {
                System.IO.File.WriteAllBytes(fullFileName, rtfFileContent);
            }
            catch(Exception e) 
            {
                string msg = e.Message;
            }

            return fullFileName;
        }


        private static void ExportToPdf(string fullRtfFileName, string fullPdfFileName)
        {
            Microsoft.Office.Interop.Word.Application wordApplication = new Application();
            Document wordDocument = null;

            object paramSourceDocPath = fullRtfFileName;
            object paramMissing = Type.Missing;

            string paramExportFilePath = fullPdfFileName;
            WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor =
                WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks =
                WdExportCreateBookmarks.wdExportCreateNoBookmarks; //wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            try
            {

                /*
                 
                 docs.Open(
                             COleVariant("C:\\Test.doc",VT_BSTR),
                             covFalse,    // Confirm Conversion.
                             covFalse,    // ReadOnly.
                             covFalse,    // AddToRecentFiles.
                             covOptional, // PasswordDocument.
                             covOptional, // PasswordTemplate.
                             covFalse,    // Revert.
                             covOptional, // WritePasswordDocument.
                             covOptional, // WritePasswordTemplate.
                             covOptional) // Format. // Last argument for Word 97
                                covOptional, // Encoding // New for Word 2000/2002
                                covTrue,     // Visible
                                covOptional, // OpenConflictDocument
                                covOptional, // OpenAndRepair
                                (long)0,     // DocumentDirection wdDocumentDirection LeftToRight
                                covOptional  // NoEncodingDialog
                                )  // Close Open parameters
                 */
                // Open the source document.
                wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

                // Export it in the specified format.
                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                        paramExportFormat, paramOpenAfterExport,
                        paramExportOptimizeFor, paramExportRange, paramStartPage,
                        paramEndPage, paramExportItem, paramIncludeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1,
                        ref paramMissing);
            }
            catch (Exception ex)
            {
                // Respond to the error
                string msg = ex.Message;
            }
            finally
            {
                // Close and release the Document object.
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordDocument = null;
                }

                // Quit Word and release the ApplicationClass object.
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }


        private static string ReadPdfContent(string fullPdfFileName)
        {
            string base64PdfContent = "";
            try
            {
                byte [] fileBytes = System.IO.File.ReadAllBytes(fullPdfFileName);
                base64PdfContent = Convert.ToBase64String(fileBytes);
            }
            catch (Exception e)
            {
                string msg = e.Message;
            }

            return base64PdfContent;
        }


        private static bool DeleteTempFile(string fullFileName)
        {
            try
            {
                if(System.IO.File.Exists(fullFileName))
                    System.IO.File.Delete(fullFileName);
            }
            catch (Exception e)
            {
                string msg = e.Message;
                return false;
            }

            return true;
        }
    }
}
