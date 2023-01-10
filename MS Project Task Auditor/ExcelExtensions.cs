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
    public static class ExcelExtensions
    {
        // Copies all contents of an excel sheet to another excel sheet based on insert point argument
        // Optionally provide table to use as Row length reference
        // Optionally remove header row
        public static int CopyToSheet(this Worksheet inputSheet, Worksheet outputSheet, int insertLineIndex, string refColumn = "A", bool removeHeader = false)
        {
            int firstRow = 1;
            if (removeHeader) { firstRow++; }
            int rowRange = inputSheet.Range[refColumn + inputSheet.Rows.Count.ToString()].End[XlDirection.xlUp].Row;
            int rowCount = rowRange + 1 - firstRow;
            int columnRange = inputSheet.Range["A1"].End[XlDirection.xlToRight].Column;

            string lastColumnID = Convert.ToChar(Convert.ToInt32('A') - 1 + columnRange).ToString();

            string inputRange = string.Format("{0}{1}:{2}{3}", "A", firstRow, lastColumnID, rowRange);
            string outputRange = string.Format("{0}{1}:{2}{3}", "A", insertLineIndex, lastColumnID, insertLineIndex + rowCount - 1);

            inputSheet.Range[inputRange].Copy(outputSheet.Range[outputRange]);

            return rowCount;
        }

        // Project UID and Cost Code report
        public static int WriteToExcel(this Project currentProject, Worksheet activeSheet, int currentLine, List<ProjectReportValues> reportValueList)
        {
            int rowIndex = currentLine;

            // Set Headers
            if (rowIndex == 1)
            { activeSheet.WriteHeaders(reportValueList); rowIndex++; }

            List<string> taskValueList = new List<string>();

            // Write Task Details
            foreach (Microsoft.Office.Interop.MSProject.Task nextTask in currentProject.Tasks)
            {
                taskValueList = nextTask.GetValue(reportValueList);
                activeSheet.WriteToRow(rowIndex, taskValueList);
                rowIndex++;
            }

            return rowIndex;
        }

        // Writes header row based on Values
        public static void WriteHeaders(this Worksheet activeSheet, List<ProjectReportValues> reportValueList)
        {
            List<string> stringList = new List<string>();

            foreach (ProjectReportValues reportValue in reportValueList)
            { stringList.Add(Enum.GetName(typeof(ProjectReportValues), reportValue)); }

            activeSheet.WriteToRow(1, stringList);
        }

        // Writes header row based on Values
        public static void WriteHeaders(this Worksheet activeSheet, List<string> reportValueList)
        {
            activeSheet.WriteToRow(1, reportValueList);
        }

        // Writes row based on Values
        public static void WriteToRow(this Worksheet activeSheet, int rowIndex, List<string> rowValueList)
        {
            int columnIndex = 1;
            foreach(string nextValue in rowValueList)
            { activeSheet.Cells[rowIndex, columnIndex] = nextValue; columnIndex++; }
        }

    }
}
