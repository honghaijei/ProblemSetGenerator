using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Linq;
using OpenPop.Pop3;
using OpenPop.Mime;

namespace ProblemSetGenerator
{
    public class ExcelMerger
    {
        private Application app;
        public ExcelMerger() {
            app = new Application();
            
            string root = Directory.GetCurrentDirectory();
            app.DisplayAlerts = false;
            
            app.Visible = false;
            app.Workbooks.Add("");
            app.Workbooks.Add(root + "\\ch1.xlsx");
            app.Workbooks.Add(root + "\\ch2.xlsx");
            app.Workbooks.Add(root + "\\ch3.xlsx");
            app.Workbooks.Add(root + "\\ch4.xlsx");
        }
        public void Merge(string[] names, string output)
        {
            int cnt = 0;
            List<string> t = new List<string>();
            bool first = true;
            for (int i = 2; i <= app.Workbooks.Count; i++)
            {
                app.Workbooks[i].Sheets.Add(After: app.Workbooks[i].Sheets[app.Workbooks[i].Sheets.Count]); 
                for (int j = 1; j < app.Workbooks[i].Worksheets.Count;)
                {
                    if (app.Workbooks[i].Worksheets[j].Name.StartsWith(names[i - 2]))
                    {
                        Worksheet ws = app.Workbooks[i].Worksheets[j];
                        string orgName = ws.Name;
                        if (first && ws.Name.EndsWith("feedback"))
                        {
                            ws.Name = names[0] + names[1] + names[2] + names[3];
                            ws.Copy(app.Workbooks[1].Worksheets[cnt + 1]);
                            first = false;
                        }
                        string newName = orgName[0] + new string(orgName.Where(x => char.IsLetter(x)).ToArray());
                        ws.Name = newName;

                        ws.Move(app.Workbooks[1].Worksheets[++cnt]);
                        
                        //Copy problem number to summary sheet
                        /*
                        if (app.Workbooks[1].Worksheets[cnt].Name.EndsWith("feedback"))
                        {
                            for (int k = 6; ; ++k)
                            {
                                var num = app.Workbooks[i].Worksheets[j].Cells[k, 2].Value;
                                if (num == null) break;
                                string number = num.ToString();
                                number = number.Trim();
                                if (number == "") break;
                                t.Add(num);
                            }
                        }
                         * */
                        //ws.Name = orgName;
                    }
                    else ++j;
                }
            }
            var last = app.Workbooks[1].Worksheets[app.Workbooks[1].Worksheets.Count - 1];
            for (int k = 6; k < 6 + t.Count; ++k) {
                last.Cells[k, 2] = t[k - 6];
            }

            while (true)
            {
                try
                {
                    app.Workbooks[1].Worksheets[app.Workbooks[1].Worksheets.Count].Delete();
                    break;
                }
                catch {

                }
            }


            app.Workbooks[1].SaveCopyAs(output);
        }
        public void Close() {
            app.DisplayAlerts = false;
            int len = app.Workbooks.Count;
            for (int i = len; i >= 1; --i)
                app.Workbooks[i].Close();
            app.Quit();
            app = null;
        }
    }
    public class Replicator {
        public static void Replicate(string filename, int number) {
            Application app = new Application();

            string root = Directory.GetCurrentDirectory();
            app.DisplayAlerts = false;

            app.Visible = false;
            app.Workbooks.Add(filename);
            int len = app.Workbooks[1].Worksheets.Count;
            Worksheet ws1 = app.Workbooks[1].Worksheets[len - 1];
            Worksheet ws2 = app.Workbooks[1].Worksheets[len];
            string h1 = new string(ws1.Name.Where(char.IsLetter).ToArray());
            string h2 = new string(ws2.Name.Where(char.IsLetter).ToArray());
            int id = int.Parse(new string(ws1.Name.Where(char.IsDigit).ToArray()));

            Func<string, string> bq = (string a) => { 
                if (a.Length == 0) return "000" + a; 
                else if (a.Length == 1) return "00" + a;
                else if (a.Length == 2) return "0" + a;
                else return a;
            };
            for (int i = 0; i < number; ++i) {
                ws1.Copy(After: app.Workbooks[1].Worksheets[len + i * 2 + 0]);
                app.Workbooks[1].Worksheets[len + i * 2 + 1].Name = h1 + bq((id + i + 1).ToString());
                ws2.Copy(After: app.Workbooks[1].Worksheets[len + i * 2 + 1]);
                app.Workbooks[1].Worksheets[len + i * 2 + 2].Name = h2 + bq((id + i + 1).ToString());
            }

            app.Workbooks[1].SaveCopyAs(filename);
        }
    }
    public class HomeworkGenerator {
        private static Application app;
        
        public static List<List<FileInfo>> GetCandidates() {
            string root = Directory.GetCurrentDirectory();

            DirectoryInfo d1 = new DirectoryInfo("ch1");
            DirectoryInfo d2 = new DirectoryInfo("ch2");
            DirectoryInfo d3 = new DirectoryInfo("ch3");
            DirectoryInfo d4 = new DirectoryInfo("ch4");

            var ans = new List<List<FileInfo>>();
            foreach (var f1 in d1.GetFiles())
                foreach (var f2 in d2.GetFiles())
                    foreach (var f3 in d3.GetFiles())
                        foreach (var f4 in d4.GetFiles())
                        {
                            if (!f1.FullName.EndsWith(".docx") || f1.Name.StartsWith("~$")) continue;
                            if (!f2.FullName.EndsWith(".docx") || f2.Name.StartsWith("~$")) continue;
                            if (!f3.FullName.EndsWith(".docx") || f3.Name.StartsWith("~$")) continue;
                            if (!f4.FullName.EndsWith(".docx") || f4.Name.StartsWith("~$")) continue;
                            var cur = new List<FileInfo>() { f1, f2, f3, f4 };
                            ans.Add(cur);
                        }
            return ans;
        }
        public static void GenerateID() {
            app = new Microsoft.Office.Interop.Excel.Application();

            string root = Directory.GetCurrentDirectory();

            var Problems = GetCandidates();

            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            int pos = 0;
            for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
            {
                if (app.Workbooks[1].Worksheets[i].Name == "number")
                {
                    pos = i;
                }
            }
            var rd = new Random();
            var Numbers = new Dictionary<string, List<FileInfo>>();
            for (int i = 7; ; ++i)
            {
                var num = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                if (num == null) break;
                string number = num.ToString();
                number = number.Trim();
                if (number == "") break;
                var value = Problems[rd.Next(Problems.Count)];
                string id = string.Join("", value.Select(x => Path.GetFileNameWithoutExtension(x.Name)));
                app.Workbooks[1].Worksheets[pos].Cells[i, 3] = "'" + id.Split(new char[]{'\\'}).Last();
                Numbers[number] = value;
            }
            app.DisplayAlerts = false;
            app.Workbooks[1].SaveAs(root + "\\feedback.xlsx");
            app.Workbooks[1].Close();

        }

        public static void GenerateByID(int type)
        {
            string root = Directory.GetCurrentDirectory();
            //Helper.ConvertWordToPdf(root + "\\000.docx", root + "\\000.pdf");
            //Helper.CloseWordExcel();

            app = new Microsoft.Office.Interop.Excel.Application();

            

            //var Problems = GetCandidates();

            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            int pos = 0;
            for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
            {
                if (app.Workbooks[1].Worksheets[i].Name == "number")
                {
                    pos = i;
                }
            }
            var Numbers = new Dictionary<string, List<FileInfo>>();
            for (int i = 7; ; ++i)
            {
                var num = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                if (num == null) break;
                string number = num.ToString();
                number = number.Trim();
                if (number == "") break;
                //var value = Problems[rd.Next(Problems.Count)];
                //string id = string.Join("", value.Select(x => Path.GetFileNameWithoutExtension(x.Name)));

                string vs = new string(((string)app.Workbooks[1].Worksheets[pos].Cells[i, 3].Value).Where(char.IsDigit).ToArray());
                var id1 = new FileInfo("ch1\\" + vs.Substring(0, 3) + ".docx");
                var id2 = new FileInfo("ch2\\" + vs.Substring(3, 3) + ".docx");
                var id3 = new FileInfo("ch3\\" + vs.Substring(6, 3) + ".docx");
                var id4 = new FileInfo("ch4\\" + vs.Substring(9, 3) + ".docx");
                var value = new List<FileInfo>() { id1, id2, id3, id4 };

                Numbers[number] = value;
            }
            app.DisplayAlerts = false;
            app.Workbooks[1].SaveAs(root + "\\feedback.xlsx");
            app.Workbooks[1].Close();
            Helper.CloseWordExcel();
            DirectoryInfo gradeDir = null, homeworkDir = null;
            if (type == 2)
            {
                gradeDir = new DirectoryInfo(root + "\\for grade");
                gradeDir.Create();
                Helper.Clear(gradeDir);
            }
            else
            {
                homeworkDir = new DirectoryInfo(root + "\\for homework");
                homeworkDir.Create();
                Helper.Clear(homeworkDir);
            }


            foreach (var number in Numbers.Keys)
            {
                if (type == 1)
                {
                    Console.WriteLine("Generating " + number + "'s homework...");
                    //Helper.CopyDirectory(Numbers[number], homeworkDir.FullName + "\\" + number);
                    var xx = new DirectoryInfo(homeworkDir.FullName + "\\" + number);
                    xx.Create();
                    Helper.MergeProblems(xx, Numbers[number][0], Numbers[number][1], Numbers[number][2], Numbers[number][3]);
                    foreach (var f in xx.GetFiles().Where(x => x.Name.EndsWith(".xlsx") || x.Name.EndsWith(".docx")))
                        f.Delete();
                    foreach (var f in xx.GetFiles().Where(x => x.Name.EndsWith(".pdf")).Take(1))
                    {
                        File.Move(f.FullName, Path.Combine(f.DirectoryName, number + ".pdf"));
                    }
                    Helper.CloseWordExcel();
                }
            }

            foreach (var number in Numbers.Keys)
            {
                if (type == 2)
                {
                    Console.WriteLine("Generating " + number + "'s grade...");
                    DirectoryInfo sd = new DirectoryInfo(gradeDir.FullName + "\\" + number + "homework");
                    sd.Create();
                    Helper.MergeProblems(sd, Numbers[number][0], Numbers[number][1], Numbers[number][2], Numbers[number][3]);
                    var feedback = sd.GetFiles().FirstOrDefault(x => x.Name.EndsWith(".xlsx"));
                    app = new Microsoft.Office.Interop.Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;
                    if (feedback != null)
                    {
                        app.Workbooks.Add(feedback.FullName);
                        var wb = app.Workbooks[app.Workbooks.Count];
                        wb.Worksheets[wb.Worksheets.Count].Name = number + "feedback";
                        wb.SaveAs(feedback.FullName);
                        wb.Close();
                        File.Move(feedback.FullName, feedback.DirectoryName + "\\" + number + "feedback.xlsx");
                    }
                    Helper.CloseWordExcel();
                    //Helper.CopyDirectory(Numbers[number], gradeDir.FullName + "\\" + number + "homework");
                }
            }
            


        }
        public static void SendHomework(string from, string account, string password, string subject, string body, string smtpAddr, int smtpPort) {
            app = new Application();

            string root = Directory.GetCurrentDirectory();

            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            int pos = 0;
            for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
            {
                if (app.Workbooks[1].Worksheets[i].Name == "number")
                {
                    pos = i;
                }
            }
            var Numbers = new List<Tuple<string, string>>();
            for (int i = 7; ; ++i)
            {
                var num = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                if (num == null) break;
                string number = num.ToString();
                number = number.Trim();
                if (number == "") break;
                Numbers.Add(new Tuple<string, string>(number, app.Workbooks[1].Worksheets[pos].Cells[i, 10].Value));
            }

            string probDir = root + "\\" + "for homework";

            foreach (var pr in Numbers) {
                string number = pr.Item1;
                string tname = probDir + "\\" + number;
                string to = number + "@tongji.edu.cn";
                if (pr.Item2 != null) to = pr.Item2.ToString();
                Helper.Zip(tname, tname + ".zip");
                Console.WriteLine("Sending Email to " + to + "...");
                int tryNum = 5;

                while (tryNum-- > 0)
                {
                    try
                    {
                        Helper.SendEmail(from, to, account, password, tname + ".zip", subject, body, smtpAddr, smtpPort);
                        break;
                    }
                    catch (Exception ex) {
                        Helper.AppendLog(string.Format("An error happened when sending homework to {0}.", number));
                        Helper.AppendLog(ex.Message + "\r\n");
                        System.Threading.Thread.Sleep(1000 * 60 * 15);
                    }
                }
                if (tryNum >= 0)
                    Helper.AppendLog(string.Format("Send homework to {0} successfully.\r\n", number));
                //new FileInfo(tname + ".zip").Delete();
            }
        }
        public static void CollectHomework(string email, string password, string pattern)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                var t = email.Split(new[] { '@' });
                t[0] = "pop";
                string server = string.Join(".", t);
                client.Connect(server, 110, false);

                // Authenticate ourselves towards the server
                client.Authenticate(email, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);
                Console.WriteLine("Number of email : " + messageCount);
                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = 1; i <= messageCount; i++)
                {
                    Console.WriteLine("Scanning the " + i + "th email.");
                    if (client.GetMessageHeaders(i).Subject.ToLower().StartsWith(pattern.ToLower()))
                    {
                        allMessages.Add(client.GetMessage(i));
                    }
                }
                allMessages = allMessages.OrderBy(x => x.Headers.DateSent).ToList();

                var homeworkDic = new Dictionary<string, List<Message>>();
                foreach (var item in allMessages) {
                    var k = item.Headers.Subject.Substring(pattern.Length).Trim();
                    if (!homeworkDic.ContainsKey(k))
                    {
                        homeworkDic[k] = new List<Message>();
                    }
                    homeworkDic[k].Add(item);
                }

                string root = Directory.GetCurrentDirectory();
                var gradeDir = new DirectoryInfo(root + "\\for grade");
                var receivedSet = new HashSet<string>();
                foreach (var st in gradeDir.GetDirectories()) {
                    var number = new string(st.Name.Where(char.IsDigit).ToArray());
                    if (homeworkDic.ContainsKey(number)) {
                        receivedSet.Add(number);
                        var mails = homeworkDic[number];
                        Dictionary<string, MessagePart> dic = new Dictionary<string, MessagePart>();
                        foreach (var p in mails.SelectMany(x => x.FindAllAttachments())) {
                            dic[p.FileName] = p;
                        }
                        foreach (var item in dic.Values)
                        {
                            var file = new FileInfo(st.FullName + "\\" + item.FileName);
                            
                            //if (file.Exists) continue;
                            item.Save(file.Create());
                            
                        }
                    }
                }
                app = new Microsoft.Office.Interop.Excel.Application();

                app.Visible = false;
                app.DisplayAlerts = false;
                app.Workbooks.Add(root + "\\feedback.xlsx");

                int pos = 0;
                for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
                {
                    if (app.Workbooks[1].Worksheets[i].Name == "number")
                    {
                        pos = i;
                    }
                }
                var Numbers = new Dictionary<string, List<FileInfo>>();
                for (int i = 7; ; ++i)
                {
                    var val = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                    if (val == null) break;
                    var num = val.ToString();
                    if (receivedSet.Contains(num))
                        app.Workbooks[1].Worksheets[pos].Cells[i, 11].Value = "收到邮件";
                }
                app.Workbooks[1].SaveAs(root + "\\feedback.xlsx");
                app.Workbooks[1].Close();
                Helper.CloseWordExcel();
            }
        }
        public static void ClassifyHomework() 
        {
            app = new Microsoft.Office.Interop.Excel.Application();

            string root = Directory.GetCurrentDirectory();

            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            int pos = 0;
            for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
            {
                if (app.Workbooks[1].Worksheets[i].Name == "number")
                {
                    pos = i;
                }
            }
            var Numbers = new Dictionary<string, List<string>>();
            var usingPaper = new HashSet<string>();
            for (int i = 7; ; ++i)
            {
                var num = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                if (num == null) break;
                string number = num.ToString();
                number = number.Trim();
                if (number == "") break;

                string vs = new string(((string)app.Workbooks[1].Worksheets[pos].Cells[i, 3].Value).Where(char.IsDigit).ToArray());
                var id1 = vs.Substring(0, 3);
                var id2 = vs.Substring(3, 3);
                var id3 = vs.Substring(6, 3);
                var id4 = vs.Substring(9, 3);
                Numbers[number] =  new List<string>() { id1, id2, id3, id4 };

                var byPaper = app.Workbooks[1].Worksheets[pos].Cells[i, 14].Value;
                if (byPaper != null && byPaper.ToString() == "1")
                    usingPaper.Add(number);
            }
            Helper.CloseWordExcel();

            DirectoryInfo category = new DirectoryInfo(root + "\\classified");
            if (!category.Exists) category.Create();
            Helper.Clear(category);
            for (int i = 1; i <= 4; ++i)
            {
                category.CreateSubdirectory("ch" + i);
            }

            
            foreach (var num in Numbers.Keys) {
                var hwDir = new DirectoryInfo(root + "\\for grade\\" + num + "homework\\");
                for (int i = 0; i < 4; ++i)
                {
                    var ch = category.FullName + "\\ch" + (i + 1);
                    var t = new DirectoryInfo(ch + "\\" + (usingPaper.Contains(num) ? "print//" : "elect//") + Numbers[num][i]);
                    if (!t.Exists) t.Create();
                    var target = t.FullName + "\\" + num + "\\";
                    
                    Helper.CopyDirectory(hwDir.FullName, target);
                    try
                    {
                        new FileInfo(target + string.Join("", Numbers[num]) + ".docx").Delete();
                        new FileInfo(target + string.Join("", Numbers[num]) + ".pdf").Delete();
                    }catch{}
                    try
                    {
                        app = new Microsoft.Office.Interop.Excel.Application();
                        app.DisplayAlerts = false;
                        var fdbk = new DirectoryInfo(target).GetFiles().First(x => x.FullName.ToLower().EndsWith("feedback.xlsx"));
                        app.Workbooks.Add(fdbk.FullName);
                        while (app.Workbooks[1].Worksheets.Count > 1 && !app.Workbooks[1].Worksheets[1].Name.StartsWith((i + 1).ToString()))
                            app.Workbooks[1].Worksheets[1].Delete();
                        while (app.Workbooks[1].Worksheets.Count > 2)
                            app.Workbooks[1].Worksheets[3].Delete();
                        app.Workbooks[1].SaveAs(fdbk.FullName);
                        app.Workbooks[1].Close();
                        Helper.CloseWordExcel();
                        File.Move(fdbk.FullName, fdbk.DirectoryName + "\\ch" + (i + 1) + ".xlsx");
                        var probDocx = target + Numbers[num][i] + ".docx";
                        File.Copy(root + "\\ch" + (i + 1) + "\\" + Numbers[num][i] + ".docx", probDocx);
                        Helper.ConvertWordToPdf(probDocx, target + Numbers[num][i] + ".pdf");
                        new FileInfo(probDocx).Delete();
                    }
                    catch {
                        Helper.CloseWordExcel();
                    }
                }
                    
            }
            
        }
        public static void SendFeedback(string from, string account, string password, string subject, string body, string smtpAddr, int smtpPort)
        {
            string root = Directory.GetCurrentDirectory();


            app = new Application();
            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            int pos = 0;
            for (int i = 1; i <= app.Workbooks[1].Worksheets.Count; ++i)
            {
                if (app.Workbooks[1].Worksheets[i].Name == "number")
                {
                    pos = i;
                }
            }
            var Numbers = new Dictionary<string, string>();
            for (int i = 7; ; ++i)
            {
                var num = app.Workbooks[1].Worksheets[pos].Cells[i, 4].Value;
                if (num == null) break;
                string number = num.ToString();
                number = number.Trim();
                if (number == "") break;
                if (app.Workbooks[1].Worksheets[pos].Cells[i, 15].Value != null) {
                    string otherMailAddr = app.Workbooks[1].Worksheets[pos].Cells[i, 15].Value;
                    Numbers[number] = otherMailAddr;
                }
            }

            var gradeDir = new DirectoryInfo(root + "\\for grade");
            foreach (var st in gradeDir.GetDirectories()) {
                var feedbackXlsx = st.GetFiles().Single(x => x.Name.EndsWith(".xlsx"));
                string snum = new string(st.Name.Where(char.IsDigit).ToArray());
                var target = Helper.ReplaceFileName(feedbackXlsx.FullName, snum + "feedback.pdf");

                var to = snum + "@tongji.edu.cn";
                if (Numbers.ContainsKey(snum))
                    to = Numbers[snum];


                Helper.ConvertFeedbackToPdf(feedbackXlsx.FullName, target);
                
                Console.WriteLine("Sending feedback to " + to + "...");

                int tryNum = 5;

                while (tryNum-- > 0)
                {
                    try
                    {
                        Helper.SendEmail(from, to, account, password, target, subject, body, smtpAddr, smtpPort);
                        break;
                    }
                    catch (Exception ex)
                    {
                        Helper.AppendLog(string.Format("An error happened when sending homework to {0}.", snum));
                        Helper.AppendLog(ex.Message + "\r\n");
                        System.Threading.Thread.Sleep(1000 * 60 * 15);
                    }
                }
                if (tryNum >= 0)
                    Helper.AppendLog(string.Format("Send feedback to {0} successfully.\r\n", snum));
            }
        }

        public static void MergeFeedback() {
            string root = Directory.GetCurrentDirectory();
            var dic = new Dictionary<string, List<FileInfo>>();
            for (int t = 1; t <= 4; ++t) {
                DirectoryInfo dir = new DirectoryInfo(root + "\\classified\\ch" + t);
                var fd = dir.GetDirectories().SelectMany(x => x.GetDirectories()).SelectMany(x => x.GetDirectories());
                foreach (var d in fd) {
                    if (!dic.ContainsKey(d.Name)) {
                        dic[d.Name] = new List<FileInfo>();
                    }
                    var xlsx = d.GetFiles().FirstOrDefault(x => x.Name.EndsWith(".xlsx"));
                    dic[d.Name].Add(xlsx);
                }
            }
            Application app;
            var summary = new Dictionary<string, string[]>();
            var msg = new List<string>();
            foreach (var num in dic.Keys.OrderBy(x => x)) {
                try
                {
                    if (dic[num].Count != 4) throw new Exception();
                    app = new Application();
                    app.DisplayAlerts = false;
                    
                    app.Visible = false;
                    app.Workbooks.Add("");
                    app.Workbooks.Open(dic[num][0].FullName);
                    app.Workbooks.Open(dic[num][1].FullName);
                    app.Workbooks.Open(dic[num][2].FullName);
                    app.Workbooks.Open(dic[num][3].FullName);
                    app.Workbooks.Open(root + "\\feedback.xlsx");
                    
                    app.Workbooks[6].Worksheets["feedback"].Copy(app.Workbooks[1].Worksheets[1]);

                    var from = app.Workbooks[2].Worksheets[2].Range("B3:H16");
                    var to = app.Workbooks[1].Worksheets[1].Range("B6:H19");
                    from.Copy(to);

                    from = app.Workbooks[3].Worksheets[2].Range("B3:H17");
                    from.Copy(app.Workbooks[1].Worksheets[1].Range("B20:H34"));

                    from = app.Workbooks[4].Worksheets[2].Range("B3:H20");
                    from.Copy(app.Workbooks[1].Worksheets[1].Range("B35:H52"));

                    from = app.Workbooks[5].Worksheets[2].Range("B3:H14");
                    from.Copy(app.Workbooks[1].Worksheets[1].Range("B53:H64"));

                    app.Workbooks[2].Worksheets.Add(After: app.Workbooks[2].Worksheets[app.Workbooks[2].Worksheets.Count]);
                    app.Workbooks[3].Worksheets.Add(After: app.Workbooks[3].Worksheets[app.Workbooks[3].Worksheets.Count]);
                    app.Workbooks[4].Worksheets.Add(After: app.Workbooks[4].Worksheets[app.Workbooks[4].Worksheets.Count]);
                    app.Workbooks[5].Worksheets.Add(After: app.Workbooks[5].Worksheets[app.Workbooks[5].Worksheets.Count]);

                    app.Workbooks[5].Worksheets[1].Move(app.Workbooks[1].Worksheets[1]);
                    app.Workbooks[5].Worksheets[1].Move(app.Workbooks[1].Worksheets[2]);
                    app.Workbooks[4].Worksheets[1].Move(app.Workbooks[1].Worksheets[1]);
                    app.Workbooks[4].Worksheets[1].Move(app.Workbooks[1].Worksheets[2]);
                    app.Workbooks[3].Worksheets[1].Move(app.Workbooks[1].Worksheets[1]);
                    app.Workbooks[3].Worksheets[1].Move(app.Workbooks[1].Worksheets[2]);
                    app.Workbooks[2].Worksheets[1].Move(app.Workbooks[1].Worksheets[1]);
                    app.Workbooks[2].Worksheets[1].Move(app.Workbooks[1].Worksheets[2]);

                    while (true)
                    {
                        try
                        {
                            app.Workbooks[1].Worksheets[10].Delete();
                            break;
                        }
                        catch
                        {

                        }
                    }
                    app.Workbooks[1].Worksheets[9].Cells[4, 6] = num;

                    var total = (app.Workbooks[1].Worksheets[9].Cells[4, 8].Value ?? "").ToString();
                    var sc1 = (app.Workbooks[1].Worksheets[9].Cells[7, 8].Value ?? "").ToString();
                    var sc2 = (app.Workbooks[1].Worksheets[9].Cells[21, 8].Value ?? "").ToString();
                    var sc3 = (app.Workbooks[1].Worksheets[9].Cells[36, 8].Value ?? "").ToString();
                    var sc4 = (app.Workbooks[1].Worksheets[9].Cells[54, 8].Value ?? "").ToString();

                    string cheat1 = (app.Workbooks[1].Worksheets[9].Cells[18, 8].Value ?? "").ToString();
                    string cheat2 = (app.Workbooks[1].Worksheets[9].Cells[33, 8].Value ?? "").ToString();
                    string cheat3 = (app.Workbooks[1].Worksheets[9].Cells[51, 8].Value ?? "").ToString();
                    string cheat4 = (app.Workbooks[1].Worksheets[9].Cells[62, 8].Value ?? "").ToString();
                    string cheat4Candidate = (app.Workbooks[1].Worksheets[9].Cells[63, 8].Value ?? "").ToString();
                    if (!Helper.IsEnglish(cheat4Candidate))
                        cheat4 = cheat4Candidate;

                    summary[num] = new string[] { total, sc1, sc2, sc3, sc4, cheat1, cheat2, cheat3, cheat4 };

                    app.Workbooks[1].SaveCopyAs(root + "\\for grade\\" + num + "homework\\" + num + "feedback.xlsx");
                    app.Workbooks[1].Close();
                }
                catch (Exception ex) {
                    msg.Add(num);
                }
                Helper.CloseWordExcel();
            }

            app = new Application();
            app.DisplayAlerts = false;
            app.Visible = false;
            app.Workbooks.Add(root + "\\feedback.xlsx");

            for (int i = 7; ; ++i) {
                if (app.Workbooks[1].Worksheets["number"].Cells[i, 4].Value == null) break;
                string num = app.Workbooks[1].Worksheets["number"].Cells[i, 4].Value.ToString();
                if (!summary.ContainsKey(num)) continue;
                for (int j = 31; j < 31 + 9; ++j) {
                    app.Workbooks[1].Worksheets["number"].Cells[i, j] = summary[num][j - 31];
                }
            }
            app.Workbooks[1].SaveAs(root + "\\feedback.xlsx");

            if (msg.Count > 0) {
                var result = System.Windows.MessageBox.Show(string.Join(",", msg), "有错误", System.Windows.MessageBoxButton.OK);
            }
            
        }
    }
}
