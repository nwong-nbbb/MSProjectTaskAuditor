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
    public enum EGeneralTaskChecks
    {
        [Description(@"Auto Schedule Status")]
        AutoScheduleStatus,
        [Description(@"Predecessor Status")]
        PredecessorStatus,
        [Description(@"Task Constraint Type")]
        TaskConstraintType,
        //TODO add implimentation with event based reports / actions
        //[Description(@"Release to Production Status")]
        //ReleaseToProductionStatus,
        [Description(@"Work Status")]
        WorkStatus,
    }

    public enum EParentTaskChecks
    {
        [Description(@"Work assignments for children")]
        ChildAssignments,
        [Description(@"Parent Release to Production Status")]
        ParentReleaseToProductionStatus,
        [Description(@"Parent Contains Actuals")]
        ParentContainsActuals,
    }

    public enum EResourceChecks
    {
        [Description(@"Assigned Resources")]
        AssignedResources,
        [Description(@"Single Resource Assignment")]
        SingleResourceAssignment,
        [Description(@"Valid Cost Codes")]
        ValidCostCodes
    }

    public enum EWorkChecks
    {
        [Description(@"Contains Actuals")]
        ContainsActuals,
        [Description(@"Contains Work")]
        ContainsWork,
        [Description(@"Actuals in the future")]
        ActualsFuture,
        [Description(@"Scheduled time in the past")]
        SchedulingPast,
        [Description(@"Work Type")]
        WorkType
    }

    public enum EProjectChecks
    {
        [Description(@"Duplicate Tasks")]
        DuplicateTasks,
    }

    public interface IProjectMonitor
    {
        Microsoft.Office.Interop.MSProject.Task AssignedTask { get; set; }
    }

    public class ProjectMonitorBase : IProjectMonitor
    {
        protected Microsoft.Office.Interop.MSProject.Task thisTask;
        protected string projectName;
        protected string monitorName;
        public string Name { get; }
        public Microsoft.Office.Interop.MSProject.Task AssignedTask
        {
            get { return thisTask; }
            set { thisTask = value; VerifyTask(); }
        }
        protected bool writeLog;
        protected bool fixErrors;
        protected bool skipParents;

        protected string statusMessage;
        public string StatusMessage { get => statusMessage; }

        protected string reportLog;
        public string LogPath { get => reportLog; }



        public ProjectMonitorBase(Project thisProject, bool skipParents = false, string logDir = "")
        {
            this.skipParents = skipParents;
            this.projectName = thisProject.FullName;

            string projectName = thisProject.Name.Replace("<>", "").Replace(".mpp", "");
            Console.WriteLine("Generating report for {0}", this.projectName.Replace("<>", ""));
            string logPath = logDir;

            string dateStamp = string.Format("{0}-{1}-{2}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

            if (logPath.Length < 1 || !Directory.Exists(logPath)) { logPath = String.Format("{0}\\Project", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)); }
            logPath += string.Format("\\Reports\\{0}\\{0}-{1}-log.txt", projectName, dateStamp);

            reportLog = logPath;
            writeLog = logPath.Length > 0;

            statusMessage = "";
        }

        protected void InitLog()
        {
            reportLog = reportLog.Replace("log.txt", monitorName + "log.txt");
            if (!Directory.Exists(reportLog)) { Directory.CreateDirectory(Path.GetDirectoryName(reportLog)); }
            if (File.Exists(reportLog)) { File.Delete(reportLog); }
            using (StreamWriter logStream = File.CreateText(reportLog))
            {
                logStream.WriteLine("Task Report for {0} generated on {1}", this.projectName, DateTime.Now.ToString());
                logStream.Close();
            }
        }

        // Must override during implimentation
        //public virtual IProjectMonitor NewMonitor(Project currentProject, Boolean skipParents, string logFilePath) { return this; }

        protected virtual void VerifyTask() { }
        protected virtual void CorrectTask() { } 

        protected bool WriteMessage()
        {
            if (statusMessage.Length > 0)
            {
                statusMessage = string.Format("[{0}]::", thisTask.Name) + statusMessage;
                if (writeLog) { WriteToLog(reportLog, statusMessage).Wait(); Console.Write("."); }
                else { Console.WriteLine(statusMessage); }
                return false;
            }
            else
                return true;
        }

        protected static async System.Threading.Tasks.Task WriteToLog(string logPath, string logMessage)
        {
            StreamWriter logFileStream = new StreamWriter(logPath, append: true);
            await logFileStream.WriteLineAsync(logMessage);
            logFileStream.Close();
        }
    }

}
