using AntiVirus;
using log4net;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;

namespace FileConverter
{
    class FileConverter
    {
        private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private static String InputFolder, OutputFolder, ErrorFolder, CSVSeparator;
        private static Regex MSWordFileExtensions, MSExcelFileExtensions, MSPowerPointFileExtensions, OOWriterFileExtensions,
            OOCalcFileExtensions, OOImpressFileExtensions, TextFileExtensions;

        static void Main()
        {
            Init();
            foreach (string filePath in Directory.GetFiles(InputFolder))
            {

                TrovaVirus trovaVirus = new TrovaVirus();
                ScanResult scanResult = trovaVirus.ControllaVirus(filePath);

                if (scanResult == ScanResult.FileNotExist)
                    continue;
                else if (scanResult == ScanResult.VirusFound)
                {
                    log.Warn("Ho eliminato il file infetto " + filePath);
                    continue;
                }

                string fileName = Path.GetFileName(filePath);
                if (MSWordFileExtensions.Match(fileName).Success)
                    ConvertWordToPdf(filePath);
                else if (MSExcelFileExtensions.Match(fileName).Success)
                    ConvertExcelToPdf(filePath);
                else if (MSPowerPointFileExtensions.Match(fileName).Success)
                    ConvertPowerPointToPdf(filePath);
                else if (TextFileExtensions.Match(fileName).Success)
                    ConvertCSVToPDF(filePath);
                else if (OOWriterFileExtensions.Match(fileName).Success || OOCalcFileExtensions.Match(fileName).Success || 
                    OOImpressFileExtensions.Match(fileName).Success)
                    ConvertToPdf(filePath);
                else
                {
                    log.Error("Tipo di file non supportato: " + fileName);
                    File.Move(filePath, ErrorFolder + fileName);
                }
            }
        }

        private static void Init()
        {
            log4net.Config.XmlConfigurator.Configure();
            NameValueCollection appSettings = ConfigurationManager.AppSettings;
            if (appSettings["InputFolder"] == null)
            {
                log.Error("Non è stato specificato il valore di InputFolder.");
                Environment.Exit(-1);
            }
            InputFolder = appSettings["InputFolder"];
            if (!InputFolder.EndsWith(Path.DirectorySeparatorChar.ToString()))
                InputFolder += Path.DirectorySeparatorChar;
            log.Debug("InputFolder: " + InputFolder);
            if (appSettings["OutputFolder"] == null)
            {
                OutputFolder = InputFolder + "Outuput" + Path.DirectorySeparatorChar;
                log.Warn("Non è stato specificato il valore di OutputFolder. Utilizzo la directory di default sotto InputFolder: " + OutputFolder);
            }
            else
            {
                OutputFolder = appSettings["OutputFolder"];
                if (!OutputFolder.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    OutputFolder += Path.DirectorySeparatorChar;
                log.Debug("OutputFolder: " + OutputFolder);
            }
            if (!Directory.Exists(OutputFolder))
                Directory.CreateDirectory(OutputFolder);
            if (appSettings["ErrorFolder"] == null)
            {
                ErrorFolder = OutputFolder + "Errors" + Path.DirectorySeparatorChar;
                log.Warn("Non è stato specificato il valore di ErrorFolder. Utilizzo la directory di default sotto InputFolder: " + ErrorFolder);
            }
            else
            {
                ErrorFolder = appSettings["ErrorFolder"];
                if (!ErrorFolder.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    ErrorFolder += Path.DirectorySeparatorChar;
                log.Debug("ErrorFolder: " + ErrorFolder);
            }
            if (!Directory.Exists(ErrorFolder))
                Directory.CreateDirectory(ErrorFolder);
            if (appSettings["MSWordFileExtensions"] == null)
            {
                MSWordFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di MSWordFileExtensions. I file Word non verranno considerati.");
            }
            else
            {
                MSWordFileExtensions = new Regex(appSettings["MSWordFileExtensions"]);
                log.Debug("MSWordFileExtensions: " + MSWordFileExtensions);
            }
            if (appSettings["MSExcelFileExtensions"] == null)
            {
                MSExcelFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di MSExcelFileExtensions. I file Excel non verranno considerati.");
            }
            else
            {
                MSExcelFileExtensions = new Regex(appSettings["MSExcelFileExtensions"]);
                log.Debug("MSExcelFileExtensions: " + MSExcelFileExtensions);
            }
            if (appSettings["MSPowerPointFileExtensions"] == null)
            {
                MSPowerPointFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di MSPowerPointFileExtensions. I file PowerPoint non verranno considerati.");
            }
            else
            {
                MSPowerPointFileExtensions = new Regex(appSettings["MSPowerPointFileExtensions"]);
                log.Debug("MSPowerPointFileExtensions: " + MSPowerPointFileExtensions);
            }
            if (appSettings["OOWriterFileExtensions"] == null)
            {
                OOWriterFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di OOWriterFileExtensions. I file Word non verranno considerati.");
            }
            else
            {
                OOWriterFileExtensions = new Regex(appSettings["OOWriterFileExtensions"]);
                log.Debug("OOWriterFileExtensions: " + OOWriterFileExtensions);
            }
            if (appSettings["OOCalcFileExtensions"] == null)
            {
                OOCalcFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di OOCalcFileExtensions. I file Excel non verranno considerati.");
            }
            else
            {
                OOCalcFileExtensions = new Regex(appSettings["OOCalcFileExtensions"]);
                log.Debug("OOCalcFileExtensions: " + OOCalcFileExtensions);
            }
            if (appSettings["OOImpressFileExtensions"] == null)
            {
                OOImpressFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di OOImpressFileExtensions. I file PowerPoint non verranno considerati.");
            }
            else
            {
                OOImpressFileExtensions = new Regex(appSettings["OOImpressFileExtensions"]);
                log.Debug("OOImpressFileExtensions: " + OOImpressFileExtensions);
            }
            if (appSettings["TextFileExtensions"] == null)
            {
                TextFileExtensions = new Regex("^$");
                log.Warn("Non è stato specificato il valore di TextFileExtensions. I file di testo non verranno considerati.");
            }
            else
            {
                TextFileExtensions = new Regex(appSettings["TextFileExtensions"]);
                log.Debug("TextFileExtensions: " + TextFileExtensions);
            }
            if (appSettings["CSVSeparator"] == null)
            {
                CSVSeparator = "";
                log.Warn("Non è stato specificato il valore di CSVSeparator. I file di testo varrano convertiti senza che il contenuto venga separato.");
            }
            else
            {
                CSVSeparator = appSettings["CSVSeparator"];
                log.Debug("CSVSeparator: " + CSVSeparator);
            }
        }

        private static void ConvertWordToPdf(string filePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string outputFilePath = OutputFolder + fileNameWithoutExtension + ".pdf";
            Microsoft.Office.Interop.Word.ApplicationClass applicationClass = new Microsoft.Office.Interop.Word.ApplicationClass();
            Document document = null;
            try
            {
                document = applicationClass.Documents.Open(filePath);
                document.ExportAsFixedFormat(outputFilePath, WdExportFormat.wdExportFormatPDF);
            }
            catch (System.Exception e)
            {
                log.Error("Eccezione durante la conversione del file Word " + filePath);
                log.Error(e.StackTrace);
                File.Move(filePath, ErrorFolder + Path.GetFileName(filePath));
            }
            finally
            {
                if (document != null)
                    document.Close(WdSaveOptions.wdDoNotSaveChanges);
                if (applicationClass != null)
                {
                    applicationClass.Quit(WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.FinalReleaseComObject(applicationClass);
                }
                CloseProcess("winword");
                File.Delete(filePath);
            }
        }

        private static void ConvertExcelToPdf(string filePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string outputFilePath = OutputFolder + fileNameWithoutExtension + ".pdf";
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false
            };
            Workbook workbook = null;
            Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                workbook = workbooks.Open(filePath);
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFilePath, XlFixedFormatQuality.xlQualityStandard, 
                    true, true, Type.Missing, Type.Missing, false, Type.Missing);
            }
            catch (System.Exception e)
            {
                log.Error("Eccezione durante la conversione del file Excel " + filePath);
                log.Error(e.StackTrace);
                File.Move(filePath, ErrorFolder + Path.GetFileName(filePath));
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(XlSaveAction.xlDoNotSaveChanges);
                    while (Marshal.FinalReleaseComObject(workbook) != 0);
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    while (Marshal.FinalReleaseComObject(workbooks) != 0);
                }
                if (application != null)
                {
                    application.Quit();
                    application.Application.Quit();
                    while (Marshal.FinalReleaseComObject(application) != 0) ;
                }
                CloseProcess("EXCEL");
                File.Delete(filePath);
            }
        }

        private static void ConvertPowerPointToPdf(string filePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string outputFilePath = OutputFolder + fileNameWithoutExtension + ".pdf";
            Microsoft.Office.Interop.PowerPoint.ApplicationClass applicationClass = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            Presentation presentation = null;
            try
            {
                Presentations presentations = applicationClass.Presentations;
                presentation = presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                presentation.ExportAsFixedFormat(outputFilePath, PpFixedFormatType.ppFixedFormatTypePDF,
                    PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState.msoFalse,
                    PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides,
                    MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintAll, string.Empty, false, true, true, true, false, Type.Missing);
            }
            catch (System.Exception e)
            {
                log.Error("Eccezione durante la conversione del file PowerPoint " + filePath);
                log.Error(e.StackTrace);
                File.Move(filePath, ErrorFolder + Path.GetFileName(filePath));
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }
                if (applicationClass != null)
                {
                    applicationClass.Quit();
                    Marshal.ReleaseComObject(applicationClass);
                }
                CloseProcess("powerpnt");
                File.Delete(filePath);
            }
        }

        private static void ConvertCSVToPDF(string filePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string outputFilePath = OutputFolder + fileNameWithoutExtension + ".pdf";
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            application.Visible = false;
            Workbook workbook = null;
            Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                workbooks.OpenText(Filename: filePath, DataType: XlTextParsingType.xlDelimited, Semicolon: true);
                workbook = workbooks[1];
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputFilePath, XlFixedFormatQuality.xlQualityStandard,
                    true, true, Type.Missing, Type.Missing, false, Type.Missing);
            }
            catch (System.Exception e)
            {
                log.Error("Eccezione durante la conversione del file Excel " + filePath);
                log.Error(e.StackTrace);
                File.Move(filePath, ErrorFolder + Path.GetFileName(filePath));
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(XlSaveAction.xlDoNotSaveChanges);
                    while (Marshal.FinalReleaseComObject(workbook) != 0) ;
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    while (Marshal.FinalReleaseComObject(workbooks) != 0) ;
                }
                if (application != null)
                {
                    application.Quit();
                    application.Application.Quit();
                    while (Marshal.FinalReleaseComObject(application) != 0) ;
                }
                CloseProcess("EXCEL");
                File.Delete(filePath);
            }
        }

        private static void ConvertToPdf(string filePath)
        {
            Process[] ps = Process.GetProcessesByName("soffice.exe");
            if (ps.Length != 0)
                throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");
            if (ps.Length > 0)
                return;
            Process p = new Process
            {
                StartInfo =
                        {
                            Arguments = "-headless -nofirststartwizard",
                            FileName = "soffice.exe",
                            CreateNoWindow = true
                        }
            };
            bool result = p.Start();
            if (result == false)
                throw new InvalidProgramException("OpenOffice failed to start.");
            XComponentContext xLocalContext = Bootstrap.bootstrap();
            XMultiServiceFactory xRemoteFactory = (XMultiServiceFactory)xLocalContext.getServiceManager();
            XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
            XComponent xComponent = null;
            try
            {
                PropertyValue[] openProps = new PropertyValue[1];
                openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };
                xComponent = aLoader.loadComponentFromURL(PathConverter(filePath), "_blank", 0, openProps);
                while (xComponent == null)
                    Thread.Sleep(1000);
                PropertyValue[] propertyValues = new PropertyValue[2];
                propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };
                propertyValues[0] = new PropertyValue { Name = "FilterName", Value = new Any(ConvertExtensionToFilterType(Path.GetExtension(filePath))) };
                ((XStorable) xComponent).storeToURL(PathConverter(OutputFolder + Path.GetFileNameWithoutExtension(filePath) + ".pdf"), propertyValues);
            }
            catch (System.Exception e)
            {
                log.Error("Eccezione la conversione del file OpenOffice " + filePath);
                log.Error(e.StackTrace);
                File.Move(filePath, ErrorFolder + Path.GetFileName(filePath));
            }
            finally
            {
                if (xComponent != null) xComponent.dispose();
                CloseProcess("soffice");
                File.Delete(filePath);
            }
        }

        private static string PathConverter(string file)
        {
            if (string.IsNullOrEmpty(file))
                throw new NullReferenceException("Null or empty path passed to OpenOffice");
            return String.Format("file:///{0}", file.Replace(@"\", "/"));
        }

        private static string ConvertExtensionToFilterType(string extension)
        {
            switch (extension)
            {
                case ".html":
                case ".HTML":
                case ".htm":
                case ".HTM":
                case ".xml":
                case ".XML":
                case ".odt":
                case ".ODT":
                case ".wps":
                case ".WPS":
                case ".wpd":
                case ".WPD":
                    return "writer_pdf_Export";
                case ".xlsb":
                case ".XLSB":
                case ".ods":
                case ".ODS":
                    return "calc_pdf_Export";
                case ".odp":
                case ".ODP":
                    return "impress_pdf_Export";
                default:
                    return "";
            }
        }

        private static void CloseProcess(string processName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(processName);
                if (processes != null)
                    foreach (Process process in processes)
                    {
                        process.Kill();
                        process.WaitForExit();
                    }
            }
            catch (System.Exception e)
            {
                log.Fatal("Eccezione durante la chiusura dei processi " + processName);
                log.Fatal(e.StackTrace);
                Environment.Exit(-1);
            }
        }
    }
}