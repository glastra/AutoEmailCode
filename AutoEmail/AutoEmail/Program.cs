using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;

namespace AutoEmail
{
    class Program
    {
        public static string FileName;
        public static List<string> Emails;
        public static List<Object[,]> ValueArrayList;
        public static bool Continue;

        static void Main(string[] args)
        {

            try
            {
                if (args.Any())
                 FileName = args[0];
                else
                  FileName = "c:\\GilCats\\AutoEmail\\Data\\autoemails.xlsx";
                                
                if (File.Exists(FileName))
                {
                    Continue = false;
                    // Here we read the emails from the Excel 
                    ValueArrayList = GetEmailsFromFile(FileName);
                    if (ValueArrayList.Any())
                      Emails = GetEmails(ValueArrayList);
                    else
                    {
                        Console.WriteLine("No records found");
                       // return;
                    }

                    if (Emails.Any())
                      SendEmails();
                    else
                    {
                        Console.WriteLine("No emails found");
                       // return;
                    }

                }
                else
                {
                    Console.WriteLine("File does not exist");
                    //return;
                }

            }
            catch (System.NullReferenceException ex)
            {
                throw new NullReferenceException("Error: Nothing to process.");
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();

        }



        public static void SendEmails()
        {

            DisplayEmails();
            Console.WriteLine("Proceed sending emails: (Y/N)");
            string r = Console.ReadLine();
            if (r.Contains("Y")||r.Contains("y"))
            {
                foreach (var e in Emails)
                {
                    OutlookInterop outlook = new OutlookInterop();
                    outlook.CreateItemFromTemplate(e);
                    Thread.Sleep(500);
                }
            }

          }

        public static List<string> GetEmails(List<Object[,]> objectArray)
            {
            List<string> result = new List<string>();
            foreach (var a in objectArray)
               {
                try
                {
                    object[,] value = a;
                    foreach (var v in value)
                    {
                        string s = v.ToString();
                        if (s.Contains("@"))
                            result.Add(s);
                    }
                }
                catch (System.NullReferenceException ex)
                {
                    continue;
                }
               }
            return result;
            }


        public static List<object[,]> GetEmailsFromFile(string filename)
        {
            ExcelInterop excel = new ExcelInterop();
            excel.OpenSpreadsheets(filename);
            List<object[,]> ValueArrayList = excel.ValueArraysList;

            return ValueArrayList;
        }

        public static void DisplayEmails()
        {
            Console.WriteLine("You are about to send emails to:");
            foreach (var e in Emails)
            {
                Console.WriteLine(e);

            }


        }


    }
}
