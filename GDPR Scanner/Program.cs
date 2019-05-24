using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Xceed.Words.NET;
using IronOcr;
using System.Drawing;

namespace GDPR_Scanner
{
    class FileHit
    {
        public string FilePath { get; set; }
        public int FileID { get; set; }
    }
    class Program
    {
        public static List<string> CheckSums = new List<string>();
        public static List<string> AlreadyDangerous = new List<string>();
        public static string StringCheck = "";
        static void Main(string[] args)
        {
            if (1 == 1)
            {
                MailHandler MH = new MailHandler();
                DataManager DM = new DataManager();
                if (1 == 1)
                {
                    using (SqlConnection SC = DM.getStatsConnection())
                    {
                        Console.WriteLine("Approve validated files...");
                        SC.Open();
                        SqlDataReader DFReader = null;
                        SqlCommand DFCommand = new SqlCommand("select FilePath from FileScannerDangerousFiles where Validated is null", SC);
                        DFReader = DFCommand.ExecuteReader();
                        while (DFReader.Read())
                        {
                            AlreadyDangerous.Add(DFReader[0].ToString());
                        }


                        SqlDataReader approveFileReader = null;
                        SqlCommand approveFileCommand = new SqlCommand("select FilePath from FileScannerDangerousFiles where Validated = 1 and filePath not like 'http%'", SC);
                        approveFileReader = approveFileCommand.ExecuteReader();
                        while (approveFileReader.Read())
                        {
                            try
                            {
                                FileInfo fi = new System.IO.FileInfo(@approveFileReader[0].ToString());

                                SqlCommand cmd = new SqlCommand("insert into FileScanner (FilePath,UnixLastEdit) VALUES ('" + approveFileReader[0].ToString().Replace("'", "''") + "','" + getUnixTime(fi.LastWriteTime) + "')", SC);
                                cmd.ExecuteNonQuery();

                                SqlCommand cmd1 = new SqlCommand("delete from FileScannerDangerousFiles where FilePath ='" + approveFileReader[0].ToString().Replace("'", "''") + "'", SC);
                                cmd1.ExecuteNonQuery();
                            }
                            catch { }
                        }



                        SqlDataReader approveWebReader = null;
                        SqlCommand approveWebCommand = new SqlCommand("select FilePath from FileScannerDangerousFiles where Validated = 1 and filePath like 'http%'", SC);
                        approveWebReader = approveWebCommand.ExecuteReader();
                        while (approveWebReader.Read())
                        {
                            try
                            {
                                SqlCommand cmd = new SqlCommand("insert into FileScanner (FilePath,UnixLastEdit) VALUES ('" + approveWebReader[0].ToString().Replace("'", "''") + "','0')", SC);
                                cmd.ExecuteNonQuery();

                                SqlCommand cmd1 = new SqlCommand("delete from FileScannerDangerousFiles where FilePath ='" + approveWebReader[0].ToString().Replace("'", "''") + "'", SC);
                                cmd1.ExecuteNonQuery();
                            }
                            catch { }
                        }







                        SqlDataReader itemReader = null;
                        SqlCommand itemCommand = new SqlCommand("select BasePath, OwnerEmail from FileScannerPaths", SC);
                        itemReader = itemCommand.ExecuteReader();
                        while (itemReader.Read())
                        {
                            string basePath = itemReader[0].ToString();
                            Console.WriteLine("Load Checksums");
                            CheckSums = GetChecks(basePath);

                            List<FileHit> temp = new List<FileHit>();
                            List<FileHit> Documents = LoadDirectories(basePath, temp);
                            Console.WriteLine("Documents found: " + Documents.Count);
                        }
                    }

                }
                Console.WriteLine("All Done!");
            }


        }

        


        public static List<FileHit> LoadDirectories(string DirectoryPath, List<FileHit> T)
        {
            List<FileHit> r = T;

            DataManager DM = new DataManager();
            try
            {
                foreach (string fileName in Directory.GetFiles(@DirectoryPath))
                {
                    try
                    {
                        
                        if (!CheckSums.Contains(fileName) && !AlreadyDangerous.Contains(fileName))
                        {
                            FileInfo fi = null;
                        fi = new System.IO.FileInfo(@fileName);
                            //Do check
                            Console.WriteLine("TEST FILE: " + fileName);
                            string FileContent = "";
                            switch (fi.Extension.ToLower())
                            {
                                case ".pdf":
                                    Console.WriteLine("PDF CHECK");
                                    FileContent = ReadPDFFile(fileName);
                                    break;
                                case ".xls":
                                    Console.WriteLine("EXCEL CHECK");
                                    FileContent = ReadExcelFile(fileName);
                                    break;
                                case ".xlsx":
                                    Console.WriteLine("EXCEL CHECK");
                                    FileContent = ReadExcelFile(fileName);
                                    break;
                                case ".html":
                                    Console.WriteLine("HTML CHECK");
                                    FileContent = ReadTextFile(fileName);
                                    break;
                                case ".htm":
                                    Console.WriteLine("HTML CHECK");
                                    FileContent = ReadTextFile(fileName);
                                    break;
                                case ".txt":
                                    Console.WriteLine("TXT CHECK");
                                    FileContent = ReadTextFile(fileName);
                                    break;
                                case ".xml":
                                    Console.WriteLine("XML CHECK");
                                    FileContent = ReadTextFile(fileName);
                                    break;
                                case ".xsl":
                                    Console.WriteLine("XSL CHECK");
                                    FileContent = ReadTextFile(fileName);
                                    break;
                                case ".doc":
                                    Console.WriteLine("DOC CHECK");
                                    FileContent = ReadDocFile(fileName);
                                    break;
                                case ".docx":
                                    Console.WriteLine("DOCX CHECK");
                                    FileContent = ReadDocFile(fileName);
                                    break;

                            }

                            string CPRF = "";
                            bool HasCPR = false;
                            int CPRCount = 0;
                            if (FileContent.Length > 1)
                            {
                                foreach (string Word in FileContent.Split(new Char[] { '.', ',', ' ', '\n' }))
                                {
                                    string FixWord = Word.Trim();
                                    if (CPRValid(FixWord))
                                    {
                                        CPRF = FixWord;
                                        if (FixWord.Contains('-'))
                                        {
                                            HasCPR = true;
                                            Console.WriteLine("CPR Found: " + FixWord);
                                            FileHit FH = new FileHit();
                                            FH.FilePath = fileName;
                                            FH.FileID = AddDangerousFile(fileName);
                                            r.Add(FH);
                                        }
                                        else
                                        {
                                            HasCPR = true;
                                            CPRCount++;
                                        }
                                        break;
                                    }


                                }
                            }
                            else
                            {
                                if (fi.Extension.ToLower().Equals(".pdf"))
                                {
                                    try
                                    {
                                        Console.WriteLine("OCR Read document");
                                        FileContent = OCRPDF(fileName);
                                        if (FileContent.Length > 1)
                                        {
                                            foreach (string Word in FileContent.Split(new Char[] { '.', ',', ' ', '\n' }))
                                            {
                                                string FixWord = Word.Trim();
                                                if (CPRValid(FixWord))
                                                {
                                                    CPRF = FixWord;
                                                    if (FixWord.Contains('-'))
                                                    {
                                                        HasCPR = true;
                                                        Console.WriteLine("CPR Found: " + FixWord);
                                                        FileHit FH1 = new FileHit();
                                                        FH1.FilePath = fileName;
                                                        FH1.FileID = AddDangerousFile(fileName);
                                                        r.Add(FH1);
                                                    }
                                                    else
                                                    {
                                                        HasCPR = true;
                                                        CPRCount++;
                                                    }
                                                    break;
                                                }
                                            }
                                        }

                                    }
                                    catch
                                    {
                                        FileContent = "ERROR - DISREGARD";
                                        HasCPR = false;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Nothing in file...");
                                }
                            }

                            if (CPRCount > 0 && CPRCount < 5000)
                            {
                                Console.WriteLine("CPR Found: " + CPRF);
                                FileHit FH = new FileHit();
                                FH.FilePath = fileName;
                                FH.FileID = AddDangerousFile(fileName);
                                r.Add(FH);
                                HasCPR = true;
                            }





                            if (!HasCPR)
                            {
                                using (SqlConnection SC = DM.getStatsConnection())
                                {
                                    SC.Open();
                                    SqlCommand cmd = new SqlCommand("insert into FileScanner (FilePath,UnixLastEdit) VALUES ('" + fileName.Replace("'", "''") + "','" + getUnixTime(fi.LastWriteTime) + "')", SC);
                                    cmd.ExecuteNonQuery();
                                    Console.WriteLine("File Cleared...");
                                }
                            }
                            GC.Collect();

                        }
                        else
                        {
                            Console.WriteLine("DO NOT TEST FILE: " + fileName);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("ERROR IN FILE: " + fileName);
                    }
                }
            }
            catch { }
            try
            {
                foreach (string SubDirectories in Directory.GetDirectories(DirectoryPath))
                {
                    foreach (FileHit FH in LoadDirectories(SubDirectories, T))
                    {
                        if (!r.Contains(FH))
                            r.Add(FH);
                    }
                }
            }
            catch { }

            return r;
        }


        public static string OCRPDF(string filePath)
        { 
            try
            {
                var OCR = new AdvancedOcr()
                {
                    Language = IronOcr.Languages.Danish.OcrLanguagePack,
                    ColorSpace = AdvancedOcr.OcrColorSpace.GrayScale,
                    EnhanceResolution = false,
                    EnhanceContrast = true,
                    CleanBackgroundNoise = true,
                    ColorDepth = 4,
                    RotateAndStraighten = false,
                    DetectWhiteTextOnDarkBackgrounds = false,
                    ReadBarCodes = false,
                    Strategy = AdvancedOcr.OcrStrategy.Fast,
                    InputImageType = AdvancedOcr.InputTypes.Document
                };
                var result = OCR.ReadPdf(filePath);
                return result.Text;
            }
            catch { return ""; }
        }

        public static string ReadDocFile(string FilePath)
        {
            try
            {
                using (DocX doc = DocX.Load(File.OpenRead(@FilePath)))
                {
                    return doc.Text;
                }
            }
            catch
            {
                return "";
            }
        }

        public static int AddDangerousFile(string FilePath)
        {
            int i = 0;
            //FileScannerDangerousFiles
            DataManager DM = new DataManager();
            using (SqlConnection SC = DM.getStatsConnection())
            {

                SC.Open();
                SqlDataReader itemReader = null;
                SqlCommand itemCommand = new SqlCommand("select ID from FileScannerDangerousFiles where FilePath = '" + FilePath.Replace("'","''") + "'", SC);
                itemReader = itemCommand.ExecuteReader();
                while (itemReader.Read())
                {
                    i = Convert.ToInt32(itemReader[0].ToString());
                }
                if (i == 0)
                {
                    //Add + ReCall
                    SqlCommand cmd = new SqlCommand("insert into FileScannerDangerousFiles (FilePath) VALUES ('" + FilePath.Replace("'", "''") + "')", SC);
                    cmd.ExecuteNonQuery();
                    i = AddDangerousFile(FilePath);
                }
            }

            return i;

        }

        private static int getUnixTime(DateTime DT)
        {
            return Convert.ToInt32((DT - new DateTime(1970, 1, 1, 0, 0, 0, 0)).TotalSeconds);
        }

        private static bool CPRValid(string cprNummer)
        {
            string cpr = cprNummer.Replace("-", "").Trim();
            if (cpr.Length != 10) return false;
            int sum = 0;
            int t;
            if (Int32.TryParse(cpr, out t))
            {
                try
                {
                    //3112990000
                    int day = Convert.ToInt32(cpr.Substring(0, 2));
                    int month = Convert.ToInt32(cpr.Substring(2, 2));
                    if (day < 32 && month < 13 && day > 0 && month > 0)
                    {
                        int[] check = new int[] { 4, 3, 2, 7, 6, 5, 4, 3, 2, 1 };
                        try
                        {
                            for (int i = 0; i < check.Length; i++)
                            {
                                sum += int.Parse(cpr.Substring(i, 1)) * check[i];
                            }
                        }
                        catch
                        {
                            sum = 1;
                        }
                    }
                    else
                    { sum = 1; }
                }
                catch
                {
                    sum = 1;
                }
                return sum % 11 == 0;
            }
            else
            { return false; }
        }

        public static string ReadTextFile(string FilePath)
        {
            return System.IO.File.ReadAllText(@FilePath);
        }

        public static List<string> GetChecks(string basePath)
        {
            List<string> LS = new List<string>();
            DataManager DM = new DataManager();
            using (SqlConnection SC = DM.getStatsConnection())
            {

                SC.Open();
                SqlDataReader itemReader = null;
                SqlCommand itemCommand = new SqlCommand("select FilePath, UnixLastEdit from FileScanner where FilePath like '"+ basePath +"%'", SC);
                itemReader = itemCommand.ExecuteReader();
                while (itemReader.Read())
                {
                    LS.Add(itemReader[0].ToString());
                    //StringCheck = StringCheck + "[" + itemReader[0].ToString() + "|" + itemReader[1].ToString() + "]";
                }
            }
            return LS;
        }

        public static string ReadExcelFile(string FilePath)
        {
            StringBuilder sb = new System.Text.StringBuilder();
            try
            {
                FileStream stream = File.Open(@FilePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;
                string FileType = System.IO.Path.GetExtension(@FilePath).ToUpper();
                //1. Reading Excel file
                if (System.IO.Path.GetExtension(@FilePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }


                foreach (DataTable table in excelReader.AsDataSet().Tables)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        foreach (object item in row.ItemArray)
                        {
                            sb.Append(item.ToString() + " ");
                        }
                    }
                }
            }
            catch { }
            return sb.ToString();
        }

        public static string ReadPDFFile(string FilePath)
        {
            try
            {
                PdfReader reader = new PdfReader(@FilePath);
                string text = string.Empty;
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    text += PdfTextExtractor.GetTextFromPage(reader, page);
                }
                reader.Close();
                return text;
            }
            catch
            {
                return "";
            }
        }
    }
}
