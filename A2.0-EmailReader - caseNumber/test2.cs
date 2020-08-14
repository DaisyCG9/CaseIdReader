var mails = OutlookEmails.ReadMailItems();
int i = 1;
                foreach (var mail in mails)
                {
                    /*  // C:\project\dotNetEmailReader
                      Console.WriteLine("Mail No: " + i);
                      Console.WriteLine("Mail received from: " + mail.EmailFrom);
                     // Console.WriteLine("Mail received from: " + mail.EmailFrom);
                      Console.WriteLine("Mail subject: " + mail.EmailSubject);
                     // Console.WriteLine("Mail body: " + mail.EmailBody);
                      Console.WriteLine("");*/
                    String pattern = @"[1]\d{14}";
String reader = mail.EmailSubject;
Match match = Regex.Match(reader, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        //string time = Convert.ToDateTime(mail.EmailDate).ToString("hh:mm:ss");
                        Console.WriteLine(match.Value);
                        //Console.WriteLine("Mail received Time: " + time);
                        //Console.WriteLine(match);
                    }
                    //i += 1;
                    /*StreamWriter sw = new StreamWriter(@"C:\project\dotNetEmailReader");
                     Console.SetOut(sw); 
                     Console.WriteLine("Here is the result:");
                     Console.WriteLine("Processing......");
                     Console.WriteLine("OK!"); 
                     sw.Flush(); 
                     sw.Close();
                 */
                }
                Console.ReadKey();