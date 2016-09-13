using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Reflection;
using Ionic.Zip;
using System.Text.RegularExpressions;

namespace ProblemSetGenerator
{
    
    public class WordDocumentMerger
    {
        private Application objApp = null;
        private Document objDocLast = null;
        private Document objDocBeforeLast = null;
        public WordDocumentMerger()
        {
            objApp = new Application();
            objApp.Visible = false;
            objApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        }
        #region 打开文件
        private void Open(string tempDoc)
        {
            object objTempDoc = tempDoc;
            object objMissing = System.Reflection.Missing.Value;

            objDocLast = objApp.Documents.Open(
                 ref objTempDoc,    //FileName
                 ref objMissing,   //ConfirmVersions
                 ref objMissing,   //ReadOnly
                 ref objMissing,   //AddToRecentFiles
                 ref objMissing,   //PasswordDocument
                 ref objMissing,   //PasswordTemplate
                 ref objMissing,   //Revert
                 ref objMissing,   //WritePasswordDocument
                 ref objMissing,   //WritePasswordTemplate
                 ref objMissing,   //Format
                 ref objMissing,   //Enconding
                 ref objMissing,   //Visible
                 ref objMissing,   //OpenAndRepair
                 ref objMissing,   //DocumentDirection
                 ref objMissing,   //NoEncodingDialog
                 ref objMissing    //XMLTransform
                 );

            objDocLast.Activate();
        }
        #endregion

        #region 保存文件到输出模板
        private void SaveAs(string outDoc)
        {
            object objMissing = System.Reflection.Missing.Value;
            object objOutDoc = outDoc;
            objDocLast.SaveAs(
              ref objOutDoc,      //FileName
              ref objMissing,     //FileFormat
              ref objMissing,     //LockComments
              ref objMissing,     //PassWord     
              ref objMissing,     //AddToRecentFiles
              ref objMissing,     //WritePassword
              ref objMissing,     //ReadOnlyRecommended
              ref objMissing,     //EmbedTrueTypeFonts
              ref objMissing,     //SaveNativePictureFormat
              ref objMissing,     //SaveFormsData
              ref objMissing,     //SaveAsAOCELetter,
              ref objMissing,     //Encoding
              ref objMissing,     //InsertLineBreaks
              ref objMissing,     //AllowSubstitutions
              ref objMissing,     //LineEnding
              ref objMissing      //AddBiDiMarks
              );
        }
        #endregion

        #region 循环合并多个文件（复制合并重复的文件）
        /// <summary>
        /// 循环合并多个文件（复制合并重复的文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void CopyMerge(string tempDoc, string[] arrCopies, string outDoc)
        {
            object objMissing = Missing.Value;
            object objFalse = false;
            object objTarget = WdMergeTarget.wdMergeTargetSelected;
            object objUseFormatFrom = WdUseFormattingFrom.wdFormattingFromSelected;
            try
            {
                //打开模板文件
                Open(tempDoc);
                foreach (string strCopy in arrCopies)
                {
                    objDocLast.Merge(
                      strCopy,                //FileName    
                      ref objTarget,          //MergeTarget
                      ref objMissing,         //DetectFormatChanges
                      ref objUseFormatFrom,   //UseFormattingFrom
                      ref objMissing          //AddToRecentFiles
                      );
                    objDocBeforeLast = objDocLast;
                    objDocLast = objApp.ActiveDocument;
                    if (objDocBeforeLast != null)
                    {
                        objDocBeforeLast.Close(
                          ref objFalse,     //SaveChanges
                          ref objMissing,   //OriginalFormat
                          ref objMissing    //RouteDocument
                          );
                    }
                }
                //保存到输出文件
                SaveAs(outDoc);
                foreach (Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }
            }
            finally
            {
                objApp.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                objApp = null;
            }
        }
        /// <summary>
        /// 循环合并多个文件（复制合并重复的文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void CopyMerge(string tempDoc, string strCopyFolder, string outDoc)
        {
            string[] arrFiles = Directory.GetFiles(strCopyFolder);
            CopyMerge(tempDoc, arrFiles, outDoc);
        }
        #endregion

        #region 循环合并多个文件（插入合并文件）
        /// <summary>
        /// 循环合并多个文件（插入合并文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void InsertMerge(string tempDoc, string[] arrCopies, string outDoc)
        {
            object objMissing = Missing.Value;
            object objFalse = false;
            object confirmConversion = false;
            object link = false;
            object attachment = false;
            int cnt = 0;
            try
            {
                //打开模板文件
                Open(tempDoc);
                int len = arrCopies.Length;
                for (int i = 0; i < len; ++i)
                {
                    var strCopy = arrCopies[i];
                    objApp.Selection.InsertFile(
                        strCopy,
                        ref objMissing,
                        ref confirmConversion,
                        ref link,
                        ref attachment
                        );
                    if (i != len - 1)
                    {
                        objApp.Selection.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                        objApp.Selection.EndKey(WdUnits.wdStory);
                    }
                }
                //保存到输出文件
                SaveAs(outDoc);
                foreach (Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }
            }
            finally
            {
                objApp.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                objApp = null;
            }
        }
        /// <summary>
        /// 循环合并多个文件（插入合并文件）
        /// </summary>
        /// <param name="tempDoc">模板文件</param>
        /// <param name="arrCopies">需要合并的文件</param>
        /// <param name="outDoc">合并后的输出文件</param>
        public void InsertMerge(string tempDoc, string strCopyFolder, string outDoc)
        {
            string[] arrFiles = Directory.GetFiles(strCopyFolder);
            InsertMerge(tempDoc, arrFiles, outDoc);
        }
        #endregion
    }
    class Helper
    {
        public static string LogFile = "log" + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".txt";
        public static void AppendLog(string msg) {
            File.AppendAllText(LogFile, DateTime.Now.ToString() + ":\t" + msg + "\r\n");

        }
        public static void CloseWordExcel() {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
            process = System.Diagnostics.Process.GetProcessesByName("WinWord");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }


        public static void GenerateProblemSet(string outputDirectory, string id)
        {
            var excelMerger = new ExcelMerger();
            Directory.CreateDirectory(outputDirectory);
            DirectoryInfo workDir = new DirectoryInfo(outputDirectory);
            workDir.Create();
            Helper.Clear(workDir);

            string root = MainWindow.workDirectory;

            DirectoryInfo d1 = new DirectoryInfo("ch1");
            DirectoryInfo d2 = new DirectoryInfo("ch2");
            DirectoryInfo d3 = new DirectoryInfo("ch3");
            DirectoryInfo d4 = new DirectoryInfo("ch4");

            Helper.ConvertWordToPdf(root + "\\000.docx", root + "\\000.pdf");
            /*
            Helper.Word2PdfInDir(d1);
            Helper.Word2PdfInDir(d2);
            Helper.Word2PdfInDir(d3);
            Helper.Word2PdfInDir(d4);
            */
            if (id == null || id.Length == 0)
            {
                foreach (var f1 in d1.GetFiles())
                    foreach (var f2 in d2.GetFiles())
                        foreach (var f3 in d3.GetFiles())
                            foreach (var f4 in d4.GetFiles())
                            {
                                if (!f1.FullName.EndsWith(".docx") || f1.Name.StartsWith("~$")) continue;
                                if (!f2.FullName.EndsWith(".docx") || f2.Name.StartsWith("~$")) continue;
                                if (!f3.FullName.EndsWith(".docx") || f3.Name.StartsWith("~$")) continue;
                                if (!f4.FullName.EndsWith(".docx") || f4.Name.StartsWith("~$")) continue;

                                var id1 = Path.GetFileNameWithoutExtension(f1.Name);
                                var id2 = Path.GetFileNameWithoutExtension(f2.Name);
                                var id3 = Path.GetFileNameWithoutExtension(f3.Name);
                                var id4 = Path.GetFileNameWithoutExtension(f4.Name);
                                string foldName = id1 + id2 + id3 + id4;
                                //log.AppendText("Generating problem set " + foldName + "...\n");
                                DirectoryInfo current = new DirectoryInfo(workDir.FullName + "\\" + foldName);
                                current.Create();
                                Helper.MergeProblems(current, f1, f2, f3, f4);

                            }
            }
            else
            {
                string s1 = id.Substring(0, 3);
                string s2 = id.Substring(3, 3);
                string s3 = id.Substring(6, 3);
                string s4 = id.Substring(9, 3);

                var f1 = d1.GetFiles().Single(x => x.Name.StartsWith(s1));
                var f2 = d2.GetFiles().Single(x => x.Name.StartsWith(s2));
                var f3 = d3.GetFiles().Single(x => x.Name.StartsWith(s3));
                var f4 = d4.GetFiles().Single(x => x.Name.StartsWith(s4));

                var id1 = Path.GetFileNameWithoutExtension(f1.Name);
                var id2 = Path.GetFileNameWithoutExtension(f2.Name);
                var id3 = Path.GetFileNameWithoutExtension(f3.Name);
                var id4 = Path.GetFileNameWithoutExtension(f4.Name);

                string foldName = id1 + id2 + id3 + id4;
                //log.AppendText("Generating problem set " + foldName + "...\n");
                DirectoryInfo current = new DirectoryInfo(workDir.FullName + "\\" + foldName);
                current.Create();
                Helper.MergeProblems(current, f1, f2, f3, f4);
            }

            excelMerger.Close();
        }

        public static void Clear(DirectoryInfo dir)
        {
            foreach (FileInfo file in dir.GetFiles())
            {
                file.Delete();
            }

            foreach (DirectoryInfo dirs in dir.GetDirectories())
            {
                dirs.Delete(true);
            }
        }

        public static void ConvertWordToPdf(string input, string output)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            word.Visible = false;
            word.ScreenUpdating = false;

            // Cast as Object for word Open method
            Object filename = (Object)input;

            // Use the dummy value as a placeholder for optional arguments
            Document doc = word.Documents.Open(ref filename, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            

            object outputFileName = output;
            object fileFormat = WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;

            // word has to be cast to type _Application so that it will find
            // the correct Quit method.
            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
        }
        public static void ConvertFeedbackToPdf(string input, string output) {
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            app.DisplayAlerts = false;
            app.Workbooks.Add(input);
            //while (wkb.Worksheets.Count > 1)
            //    wkb.Worksheets[1].Delete();
            //wkb.Worksheets[1].Cells[2, 2] = "'";
            //wkb.Worksheets[1].Range[wkb.Worksheets[1].Cells[2, 2], wkb.Worksheets[1].Cells[2, 16]].UnMerge();

            Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Add("");
            app.Workbooks[1].Worksheets[9].Move(app.Workbooks[2].Worksheets[1]);

            wkb.Worksheets[1].PageSetup.Zoom = false;
            wkb.Worksheets[1].PageSetup.FitToPagesWide = 1;
            wkb.Worksheets[1].PageSetup.FitToPagesTall = false;
            wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, output);
            wkb.Close();
            app.Workbooks.Close();
            app.Quit();

            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }

            app = null;
        }
        public static void Word2PdfInDir(DirectoryInfo dir)
        {
            foreach (var f in dir.GetFiles())
            {
                if (f.FullName.EndsWith(".docx") && !f.Name.StartsWith("~$"))
                {
                    var k = f.FullName.Split(new[] { '.' });
                    k[k.Length - 1] = "pdf";
                    var newName = string.Join(".", k);

                    ConvertWordToPdf(f.FullName, newName);
                }
            }
        }
        public static string ReplaceFirst(string text, string search, string replace)
        {
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
        public static void CopyDirectory(string SourcePath, string DestinationPath)
        {
            Directory.CreateDirectory(DestinationPath);
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(SourcePath, "*",
                SearchOption.AllDirectories))
                Directory.CreateDirectory(dirPath.Replace(SourcePath, DestinationPath));

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(SourcePath, "*.*",
                SearchOption.AllDirectories))
                File.Copy(newPath, ReplaceFirst(newPath, SourcePath, DestinationPath), true);
        }

        static Object sendlock = new Object();
        public static void SendEmail(string from, string to, string account, string password, string attachment, string subject, string body, string smtpAddr, int smtpPort)
        {
            lock (sendlock)
            {
                using (var mail = new System.Net.Mail.MailMessage(from, to))
                {
                    using (var client = new System.Net.Mail.SmtpClient())
                    {
                        client.Port = smtpPort;
                        client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                        client.UseDefaultCredentials = false;
                        client.Credentials = new System.Net.NetworkCredential(account, password);

                        //client.Host = "smtp." + from.Split(new[] { '@' })[1];
                        client.Host = smtpAddr;
                        mail.Subject = subject;
                        mail.Body = body;
                        mail.IsBodyHtml = true;
                        if (attachment != null)
                            mail.Attachments.Add(new System.Net.Mail.Attachment(attachment));

                        client.Send(mail);
                    }
                }
            }
            /*
            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();  
            msg.To.Add(to);
            //msg.To.Add("FDBK007@163.com");
            msg.From = new System.Net.Mail.MailAddress("admin@no-reply.com", "hyw");
            msg.Subject = subject;//邮件标题   
            msg.Body = body;//邮件内容  
            msg.IsBodyHtml = true;//是否是HTML邮件   

            if (attachment != null)
                msg.Attachments.Add(new System.Net.Mail.Attachment(attachment));
  
            var client = new System.Net.Mail.SmtpClient();
            //client.UseDefaultCredentials = true;
            client.Host = "localhost";
            //client.Port = 25;
            client.Send(msg);   
             * */
            
        }

        public static void Zip(string input, string output)
        {
            try {
                new FileInfo(output).Delete();
            }
            catch { }
            using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
            {
                zip.UseUnicodeAsNecessary = true;  // utf-8
                zip.AddDirectory(input);
                zip.Comment = "This zip was created at " + System.DateTime.Now.ToString("G");
                zip.Save(output);
            }
        }

        public static void MergeProblems(DirectoryInfo workDir, FileInfo f1, FileInfo f2, FileInfo f3, FileInfo f4) {
            if (!f1.FullName.EndsWith(".docx") || f1.Name.StartsWith("~$")) return;
            if (!f2.FullName.EndsWith(".docx") || f2.Name.StartsWith("~$")) return;
            if (!f3.FullName.EndsWith(".docx") || f3.Name.StartsWith("~$")) return;
            if (!f4.FullName.EndsWith(".docx") || f4.Name.StartsWith("~$")) return;

            
            
            string root = Directory.GetCurrentDirectory();
            var id1 = Path.GetFileNameWithoutExtension(f1.Name);
            var id2 = Path.GetFileNameWithoutExtension(f2.Name);
            var id3 = Path.GetFileNameWithoutExtension(f3.Name);
            var id4 = Path.GetFileNameWithoutExtension(f4.Name);
            string outputName = id1 + id2 + id3 + id4;
            //Console.WriteLine("Generating problem set " + workDir.Name + "...");

            try
            {
                var excelMerger = new ExcelMerger();
                excelMerger.Merge(new[] { id1, id2, id3, id4 }, workDir.FullName + "\\" + outputName + ".xlsx");
                excelMerger.Close();
                Helper.CloseWordExcel();
            }
            catch { }

            
            WordDocumentMerger m = new WordDocumentMerger();
            m.InsertMerge(root + "\\template.docx", new string[] { root + "\\000.docx", f1.FullName, f2.FullName, f3.FullName, f4.FullName }, workDir.FullName + "\\" + outputName + ".docx");
            //Helper.CloseWordExcel();
            

            Helper.ConvertWordToPdf(workDir.FullName + "\\" + outputName + ".docx", workDir.FullName + "\\" + outputName + ".pdf");
            

            try
            {
                Helper.CopyDirectory(root + "\\" + "ref for homework", workDir.FullName + "\\" + "ref for homework");
            }
            catch { }

            
        }

        public static string ReplaceExt(string path, string ext) {
            var k = path.Split(new[] { '.' });
            k[k.Length - 1] = ext;
            return string.Join(".", k);
        }
        public static string ReplaceFileName(string path, string name)
        {
            var k = path.Split(new[] { '\\' });
            k[k.Length - 1] = name;
            return string.Join("\\", k);
        }

        public static bool IsEnglish(string inputstring)
        {
            Regex regex = new Regex(@"[A-Za-z0-9 .,-=+(){}\[\]\\]");
            MatchCollection matches = regex.Matches(inputstring);

            if (matches.Count.Equals(inputstring.Length))
                return true;
            else
                return false;
        }
    }
}
