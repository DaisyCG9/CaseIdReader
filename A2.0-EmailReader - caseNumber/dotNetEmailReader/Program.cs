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
            // CaseNumberReader();
            Console.WriteLine("-----caseNumber-----");
            //CaseNumberReader();
            Console.WriteLine(string.Join("\n",CaseNumberReader()));
             ExcelTest();
   
            Console.ReadKey();
        }

        

        public static List<String> CaseNumberReader()
        {
            List<string> id = new List<string>();
            var mails = OutlookEmails.ReadMailItems();

            foreach (var mail in mails)
            {
                String pattern1 = @"[1]\d{14}";
                String CaseReader = mail.EmailSubject;
                Match match1 = Regex.Match(CaseReader, pattern1, RegexOptions.IgnoreCase);
                if (match1.Success)
                {
                    id.Add(match1.Value);
                   // Console.WriteLine( mail.EmailTo);
                }    
            }
            //Sort the numbers from the oldest to the lastest.
            id.Sort();
            //remove the dulplicate numbers
            List<string> id1=id.Distinct().ToList();
            //Console.WriteLine(string.Join("\n", id1));
           
            return id1;
        }
        public static void ExcelTest()
        {
            //导出：将数据库中的数据，存储到一个excel中
            List<String> id = CaseNumberReader();
            List<Nums> list = new List<Nums>();
            //1、查询数据库数据  
            for (int i = 0; i < CaseNumberReader().Count; i++)
            {
                Nums num = new Nums() { CaseNumber = i + 1, Number = id[i] };
                list.Add(num);
            }
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
            cellpwd.SetCellValue("Number");
            //遍历集合，生成行
            int index = 1; //从1行开始写入
            for (int i = 0; i < list.Count; i++)
            {
                int x = index + i;
                var rowi = sheet.CreateRow(x);
                var ids = rowi.CreateCell(0);
                ids.SetCellValue(list[i].CaseNumber);
                var name = rowi.CreateCell(1);
                name.SetCellValue(list[i].Number);
            }
            for (int k = 0; k < 6; k++)
            {
                sheet.AutoSizeColumn(k);
            }
            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\Daisy\Desktop\inf.xls");
            String rootFolder = @"C:\Users\Daisy\Desktop";
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
            }
            FileStream file1 = new FileStream(@"C:\Users\Daisy\Desktop\inf1.xls", FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file1);
            file1.Dispose();
            Console.WriteLine("File has been finished!");
        }

        }
}
