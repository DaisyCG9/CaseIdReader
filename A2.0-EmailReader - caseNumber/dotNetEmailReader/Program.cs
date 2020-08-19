using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http;
using static dotNetEmailReader.writeExcel;
using NPOI.HSSF.UserModel;
using static NPOI.XSSF.UserModel.Helpers.ColumnHelper;

namespace dotNetEmailReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("-----caseNumber-----");
            //CaseNumberReader();
           // Console.WriteLine(string.Join("\n",CaseNumberReader()));
            
            ExcelTest();
   
            Console.ReadKey();
        }

        

        public static List<con> CaseNumberReader()
        {
            List<con> writeCon = new List<con>();
            //List<string> id = new List<string>();
           // List<DateTime> time = new List<DateTime>();
            var mails = OutlookEmails.ReadMailItems();

            foreach (var mail in mails)
            {
                String pattern1 = @"[1]\d{14}";
                String CaseReader = mail.EmailSubject;
                Match match1 = Regex.Match(CaseReader, pattern1, RegexOptions.IgnoreCase);
                if (match1.Success)
                {
                    con data = new con() { caseId = match1.Value, time = mail.EmailDate, alias = mail.EmailTo };
                    
                    String pattern2 = @"Task";
                    Match match2 = Regex.Match(CaseReader, pattern2, RegexOptions.IgnoreCase);
                    if(match2.Success)
                    {
                        data.isTask = "Collaboration Task";
                    }
                    else
                    {
                        data.isTask = "Case";
                    }
                    StringBuilder ser = new StringBuilder();
                    ser.Append(CaseReader);
                    ser.Append(mail.EmailBody);
                    string pattern3 = @"\s[A|B|C]\s";
                    Match match3 = Regex.Match(ser.ToString(), pattern3, RegexOptions.IgnoreCase);
                    if (match3.Success)
                    {
                        data.severity = match3.Value;
                    }
                    else
                    {
                        data.severity = "";
                    }
                    writeCon.Add(data);
                    //id.Add(match1.Value);
                    //time.Add(mail.EmailDate);
                   
                }    
            }
            //Sort the numbers from the oldest to the lastest.
            // id.Sort();
            //remove the dulplicate numbers
            // List<string> id1=id.Distinct().ToList();
            // time.ToString();
            // Console.WriteLine(string.Join("\n", time)); 
            //  return id1;
           List<con> nonDuplicateList = new List<con>();
            foreach (con mem in writeCon)
            {
                if (nonDuplicateList.Exists(x=>x.caseId==mem.caseId)==false)
                {
                    nonDuplicateList.Add(mem);
                }
            }
           var sortedData =
             (from s in nonDuplicateList
              select new
              {
                  s.caseId,
                  s.time,
                  s.alias,
                  s.severity
              }).Distinct().OrderBy(x => x.caseId).ToList();
            foreach (var i in sortedData)
            {
                Console.WriteLine("caseId:   " + i.caseId + "          " + "SentTime:   " + i.time + "          " + "Alias:   " + i.alias);

            }
            return nonDuplicateList;
        }
        public static void ExcelTest()
        {
            //导出：将数据库中的数据，存储到一个excel中
            List<con> id = CaseNumberReader();
            var sd =
            (from s in id
             select new
             {
                 s.caseId,
                 s.time,
                 s.alias,
                 s.isTask,
                 s.severity
             }).Distinct().OrderBy(x => x.caseId).ToList();

            //1、查询数据库数据  
            //2、  生成excel
            //2_1、生成workbook
            //2_2、生成sheet
            //2_3、遍历集合，生成行
            //2_4、根据对象生成单元格
            HSSFWorkbook workbook = new HSSFWorkbook();
            //创建工作表
            var sheet = workbook.CreateSheet("信息表");
            //创建标题行（重点） 从0行开始写入
            var row = sheet.CreateRow(0);
            //创建单元格
            var cellid = row.CreateCell(0);
            cellid.SetCellValue("nums");
            var cellname = row.CreateCell(1);
            cellname.SetCellValue("CaseNumber");
            var cellpwd = row.CreateCell(2);
            cellpwd.SetCellValue("Alias");
            var date = row.CreateCell(3);
            date.SetCellValue("Date");
            var isTask = row.CreateCell(4);
            isTask.SetCellValue("Item Type");
            var severity = row.CreateCell(5);
            severity.SetCellValue("Severity");

            //遍历集合，生成行
            int index = 1; //从1行开始写入
            for (int i = 0; i < sd.Count; i++)
            {
                int x = index + i;
                var rowi = sheet.CreateRow(x);
                var seq = rowi.CreateCell(0);
                seq.SetCellValue(i+1);
                var ids = rowi.CreateCell(1);
                ids.SetCellValue(sd[i].caseId);
                var name = rowi.CreateCell(2);
                name.SetCellValue(sd[i].alias);
                var d = rowi.CreateCell(3);
                d.SetCellValue(sd[i].time.ToString());
                var t = rowi.CreateCell(4);
                t.SetCellValue(sd[i].isTask);
                var s = rowi.CreateCell(5);
                s.SetCellValue(sd[i].severity);
            }
            for (int k = 0; k<7; k++)
            {
                sheet.AutoSizeColumn(k);
            }
            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\Daisy\Desktop\inf.xls");
           /* String rootFolder = @"C:\Users\Daisy\Desktop";
            string file = "inf1.xls";
            try
            {
                if (File.Exists(Path.Combine(rootFolder, file)))
                {
                    File.Delete(Path.Combine(rootFolder, file));
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("This process failed:{0}", e.Message);
            }*/
            string w = @"C:\Users\Daisy\Desktop\"+ "CaseReader_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";
            FileStream file1 = new FileStream(w, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file1);
            file1.Dispose();
            Console.WriteLine("File has been finished!");
        }


    }
}
