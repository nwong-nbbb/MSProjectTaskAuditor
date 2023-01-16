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
    // All assignments converted enum based references.
    // After refactoring for more general use, Excel code was moved to Excel Extensions.
    internal class TaskUIDReport
    {
        public static void RunCurrent()
        {
            Microsoft.Office.Interop.MSProject.Application activeProjectSession = new Microsoft.Office.Interop.MSProject.Application();
            Project currentProject = null;

            try { currentProject = activeProjectSession.ActiveProject; } catch { }

            RunTaskReport(currentProject);
        }

        public static void RunAll()
        {
            RunTaskReport();
        }

        // All project UID and Cost Code report
        protected static void RunTaskReport(Project initialProject = null)
        {
            string dateStamp = string.Format("{0}-{1}-{2}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            string primaryTag = "Master";
            string projectTag = "";
            string outputFile = String.Format("{0}\\Project\\TaskLists\\{2}\\{1} - TaskList - {2}.xlsx", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), primaryTag, dateStamp);

            if (!Directory.Exists(Path.GetDirectoryName(outputFile))) { Directory.CreateDirectory(Path.GetDirectoryName(outputFile)); }

            Microsoft.Office.Interop.MSProject.Application activeProjectSession = new Microsoft.Office.Interop.MSProject.Application();
            activeProjectSession.OptionsSave(AutomaticSave: false, AutomaticSavePrompt: false);
            Project currentProject = initialProject;

            List<ProjectReportValues> reportValueList = ((ProjectReportValues[])Enum.GetValues(typeof(ProjectReportValues))).ToList();

            Microsoft.Office.Interop.Excel.Application excelSession = new Microsoft.Office.Interop.Excel.Application();
            excelSession.Visible = true;
            Workbook primaryWorkbook = excelSession.Workbooks.Add();
            Worksheet primarySheet = primaryWorkbook.ActiveSheet;
            Workbook activeProjectWorkbook;
            Worksheet activeProjectWorksheet;

            int currentLine = 1;
            int projectLine = 1;
            bool removeHeader = false;

            if (currentProject != null)
            { currentLine = currentProject.WriteToExcel(primarySheet, currentLine, reportValueList); }
            else
            {
                List<PWAProjectList> projectList = ((PWAProjectList[])Enum.GetValues(typeof(PWAProjectList))).ToList();
                foreach (PWAProjectList nextProject in projectList)
                {
                    Console.WriteLine("Beginning {0} export", nextProject.ToString());

                    projectLine = 1;
                    activeProjectWorkbook = excelSession.Workbooks.Add();
                    activeProjectWorksheet = activeProjectWorkbook.ActiveSheet;

                    currentProject = activeProjectSession.OpenProject(nextProject, noMessages: true);
                    projectTag = currentProject.Name.Replace("<>", "").Replace(".mpp", "");                    
                    projectLine = currentProject.WriteToExcel(activeProjectWorksheet, projectLine, reportValueList);
                    currentLine += activeProjectWorksheet.CopyToSheet(primarySheet, currentLine, refColumn: "B", removeHeader: removeHeader);
                    activeProjectSession.FileCloseEx(Save: PjSaveType.pjDoNotSave, NoAuto: true, CheckIn: true);
                    //activeProjectSession.DocClose();

                    removeHeader = true;
                    activeProjectWorksheet.SaveAs(outputFile.Replace(primaryTag, projectTag));
                    activeProjectWorkbook.Close();
                }
                Console.WriteLine("Project iteration complete");
            }

            if (File.Exists(outputFile)) { File.Delete(outputFile); }
            primarySheet.SaveAs(outputFile);
            excelSession.Workbooks.Close();
            excelSession.Quit();
        }
    }
}
