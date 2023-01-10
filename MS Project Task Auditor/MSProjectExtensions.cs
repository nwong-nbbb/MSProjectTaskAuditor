using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.MSProject;

namespace MS_Project_Task_Auditor
{

    public static class MSProjectExtensions
    {

        // Generate daily backup path
        public static string GetBackupPath(string fileName = "")
        {
            string dateStamp = string.Format("{0}-{1}-{2}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            if (fileName.Length > 0) { fileName = "\\" + fileName.Replace(".", "-" + dateStamp + "."); }
            string backupPath = string.Format("{0}\\Project\\Backup\\{1}\\{2}", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), dateStamp, fileName);

            return backupPath;
        }

        // Check if daily backup exists
        public static bool BackupExists()
        {
            string backupPath = GetBackupPath();
            return (Directory.Exists(backupPath) && Directory.GetFiles(backupPath).Count() > 0);
        }

        // Create a local backup for all PWAProjects
        public static void DailyBackup()
        {
            if (!BackupExists())
            {
                List<PWAProjectList> projectList = ((PWAProjectList[])Enum.GetValues(typeof(PWAProjectList))).ToList();
                DailyBackup(projectList);
            }
        }

        // Create a local backup for give project
        public static void DailyBackup(PWAProjectList targetProject)
        {
            List<PWAProjectList> projectList = new List<PWAProjectList>() { targetProject };
            DailyBackup(projectList);
        }

        // Create a local backup for listed projects
        public static void DailyBackup(List<PWAProjectList> projectList)
        {
            Microsoft.Office.Interop.MSProject.Application activeProjectSession = new Microsoft.Office.Interop.MSProject.Application();
            activeProjectSession.Visible = true;
            Project currentProject = null;

            foreach (PWAProjectList nextProject in projectList)
            {
                currentProject = activeProjectSession.OpenProject(nextProject);

                string projectName = currentProject.Name.Replace("<>", "").Replace(".mpp", "");
                Console.WriteLine("Creating backup for {0}", currentProject.FullName.Replace("<>", ""));
                string backupPath = GetBackupPath(projectName + "-Backup.mpp");
                Console.WriteLine("Creating backup {0}", backupPath);
                if (!Directory.Exists(backupPath)) { Directory.CreateDirectory(Path.GetDirectoryName(backupPath)); }
                if (!File.Exists(backupPath)) { currentProject.SaveAs(backupPath); }
                activeProjectSession.FileCloseEx(Save: PjSaveType.pjDoNotSave, NoAuto: true, CheckIn: true);
            }
        }

        // Method to simplify OpenProject access
        public static Project OpenProject(this Application activeProjectSession, Enum pwaProjectEntry, bool noMessages = false)
        {
            string projectRefPath = EnumHelper.StringValueOf(pwaProjectEntry);

            if (noMessages) { activeProjectSession.DisplayAlerts = false; }
            bool openSuccess = activeProjectSession.FileOpenEx(projectRefPath, ReadOnly: PjYesNoAutomatic.pjYes);
            activeProjectSession.DisplayAlerts = true;

            if (openSuccess)
            { return activeProjectSession.ActiveProject; }
            else
            { return null; }
        }

        // Task Value mapping
        // Most values should be picked up by the default case if enum descriptions accurately match field names
        // Specific values or formatting will need to be accounted for within the other switch statements
        // Customize relative to specific environment requirements
        public static string GetValue(this Microsoft.Office.Interop.MSProject.Task thisTask, ProjectReportValues reportFieldID)
        {
            string fieldValue = "";

            switch (reportFieldID)
            {
                case ProjectReportValues.projectName:
                    fieldValue = thisTask.GetField(PjField.pjTaskProject); break;
                case ProjectReportValues.detailCode:
                    fieldValue = string.Format("{0}_{1}_", thisTask.GetValue(ProjectReportValues.projectID), thisTask.GetValue(ProjectReportValues.taskUID)); break;
                case ProjectReportValues.resourceCount:
                    fieldValue = thisTask.Assignments.Count.ToString(); break;
                case ProjectReportValues.workHours:
                    fieldValue = thisTask.Work.ToString(); break;
                case ProjectReportValues.taskDescription:
                    fieldValue = thisTask.Name; break;
                case ProjectReportValues.isParent:
                    fieldValue = thisTask.OutlineChildren.Count > 0 ? "Yes" : "No"; break;
                default:
                    string fieldName = EnumHelper.StringValueOf(reportFieldID);
                    fieldValue = thisTask.GetField(thisTask.Application.FieldNameToFieldConstant(fieldName));
                    break;
            }

            return fieldValue;
        }

        public static List<string> GetValue(this Microsoft.Office.Interop.MSProject.Task thisTask, List<ProjectReportValues> fieldList)
        {
            List<string> valueTable = new List<string>();

            foreach(ProjectReportValues nextFieldID in fieldList)
            { valueTable.Add(thisTask.GetValue(nextFieldID)); }

            return valueTable;
        }

    }

    // Simplifies accessing enum Description Values
    public class EnumHelper
    {
        public static string StringValueOf(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());
            DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (attributes.Length > 0)
            {
                return attributes.First().Description;
            }
            else
            {
                return value.ToString();
            }
        }
    }
}
