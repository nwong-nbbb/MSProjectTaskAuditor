using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.MSProject;

namespace MS_Project_Task_Auditor
{
    internal class Program
    {

        static void Main(string[] args)
        {
            bool runBackup = false;
            if (runBackup) 
            { MSProjectExtensions.DailyBackup(); }
            
            TaskUIDReport.RunAll();

            Console.WriteLine();
            //ConfirmPrompt();
        }

        public static void ConfirmPrompt()
        {
            Console.WriteLine("Press Any Key To Continue");
            Console.ReadLine();
        }
    }
}
