using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.MSProject;


namespace MS_Project_Task_Auditor
{
    public class ResourceTable
    {
        int sourceCostIndex = 2;
        int sourceResourceNameIndex = 4;

        private string defaultResourceTableDocument = @"C:\Users\nwong\Documents\Nichols\Cost Code Resource List.xlsx";
        List<(string ResourceName, string ResourceID)> resourceTable = new List<(string ResourceName, string ResourceID)>();
        List<(string CostCode, string ResourceName, string ResourceID)> costCodeTable = new List<(string CostCode, string ResourceName, string ResourceID)>();

        private bool tableLoaded;
        public bool ValidTable { get { return tableLoaded; } }

        public ResourceTable(Resources resourceList, string ResourceFile = "")
        {
            if (ResourceFile != "") { defaultResourceTableDocument = ResourceFile; }
            BuildResources(resourceList);
            BuildTable();
        }

        private void BuildResources(Resources resourceList)
        {
            foreach (Resource nextResource in resourceList)
            { resourceTable.Add((nextResource.Name, nextResource.ID.ToString())); }
        }

        private void BuildTable()
        {
            if (!File.Exists(defaultResourceTableDocument))
            { Console.WriteLine("Resource list not found, resource corrections will not be made"); tableLoaded = false; return; }
            else { tableLoaded = true; }

            Microsoft.Office.Interop.Excel.Application excelSession = new Microsoft.Office.Interop.Excel.Application();
            excelSession.Workbooks.Open(defaultResourceTableDocument);
            Worksheet activeSheet = excelSession.ActiveSheet;

            int i = 2;
            string nextCostCode = activeSheet.Cells[i, sourceCostIndex].Value;
            string nextResourceName = activeSheet.Cells[i,sourceResourceNameIndex].Value;
            string nextResourceID = LookupID(nextResourceName);
            if (nextResourceID == "0" ) { nextResourceID = LookupID(activeSheet.Cells[i, sourceResourceNameIndex + 1].Value); }

            while (nextCostCode != null)
            {
                costCodeTable.Add((nextCostCode, nextResourceName, nextResourceID));
                i++;
                nextCostCode = activeSheet.Cells[i, sourceCostIndex].Value;
                nextResourceName = activeSheet.Cells[i, sourceResourceNameIndex].Value;
                nextResourceID = LookupID(nextResourceName);
                if (nextResourceID == "0") { nextResourceID = LookupID(activeSheet.Cells[i, sourceResourceNameIndex + 1].Value); }
                if (nextCostCode != null) { if (nextCostCode.Length == 0) { nextCostCode = null; } }
            }
            excelSession.Workbooks.Close();
            excelSession.Quit();
        }

        public string LookupID(string resourceName)
        {
            string nextID = "0";
            List<(string ResourceName, string ResourceID)> resultList = resourceTable.Where(nextResource => nextResource.ResourceName == resourceName).ToList();
            if (resultList.Count > 0) { nextID = resultList.First().ResourceID; }
            return nextID;
        }

        public string GetResourceID(string CostCode)
        {
            if (CostCode != null)
            {
                if (CostCode.Length > 0)
                    return costCodeTable.Where(resource => resource.CostCode == CostCode).First().ResourceID;

            }
            return null;
        }

        public string GetResourceName(string CostCode)
        {
            if (CostCode != null)
            {
                if (CostCode.Length > 0)
                    return costCodeTable.Where(resource => resource.CostCode == CostCode).First().ResourceName;
            }
            return null;
        }

    }
}
