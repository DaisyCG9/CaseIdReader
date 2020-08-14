using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace dotNetEmailReader
{
    class Program
    {
        public void caseNumberReader()
        {
            var mails = OutlookEmails.ReadMailItems();
            int i = 1;
            foreach (var mail in mails)
            {
                String pattern = @"[1]\d{14}";
                String reader = mail.EmailSubject;
                Match match = Regex.Match(reader, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    Console.WriteLine(match.Value);
                }
            }
            Console.ReadKey();
        }
        static void Main(string[] args)
        {
            caseNumberReader();
        }
    }
}
