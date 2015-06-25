using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace xl2vsoAddIn
{
    class Converter
    {
        string[] columns = { "junk"/* just to make column index to match with excel column*/,
                               "Result",
                               "Id", 
                               "Title",
                               "Description",
                               "AssignedTo",
                               "Priority",
                               "WorkItemType",
                               "OriginalEstimate",
                               "StartDate",
                               "TargetDate",
                               "AreaPath",
                               "IterationPath",
                               "ParentId",
                               "ProjectName"
                           };
        private Excel.Worksheet activeWorksheet = null;
        public Converter(Excel.Worksheet _activeWorksheet)
        {
            activeWorksheet = _activeWorksheet;
        }
        public  bool CheckColumnSchema()
        {
            char curRow = '1';
            char curCol = 'A';
            for(int i = 1 ; i < columns.Length; ++i) {
                string range = curCol + curRow.ToString();
                Excel.Range row = activeWorksheet.get_Range(range);
                string cellValue = row.Value2;
                if(string.Compare(cellValue, columns[i], true) != 0) {
                    return false;
                } //if
                curCol++;
            } // for
            return true;
        } // CheckColumnSchema ends
        public xl2vsoModel ConvertToVSOModel(int rowIndex)
        {
            string rowStr = rowIndex.ToString();
            xl2vsoModel model = new xl2vsoModel();
            // B =id
            string range = 'B' + rowStr;
            Excel.Range row = activeWorksheet.get_Range(range);
            if (row.Value2 != null)
            {
                model.Id = row.Value2.ToString(); ;
            }

            // C =title
            model.Title = activeWorksheet.get_Range('C' + rowStr).Value2; ;

            // D =description
            var description = activeWorksheet.get_Range('D' + rowStr);
            if (description != null)
            {
                model.Description = description.Value2;
            }

            // E =assignedTo
            var assignedTo = activeWorksheet.get_Range('E' + rowStr);
            if (assignedTo != null)
            {
                model.AssignedTo = assignedTo.Value2;
            }

            // F = priority
            model.Priority = (int)activeWorksheet.get_Range('F' + rowStr).Value2;

            // G = workItemType
            model.WorkItemType = activeWorksheet.get_Range('G' + rowStr).Value2;

            // H = originalEstimate
            var estimate = activeWorksheet.get_Range('H' + rowStr);
            if (estimate != null)
            {
                model.OriginalEstimate = (int)estimate.Value2;
            }

            // I = StartDate
            var date = activeWorksheet.get_Range('I' + rowStr);
            if (date != null)
            {
                model.StartDate = date.Value2;
            }

            // K = TargetDate 
            var targetDate = activeWorksheet.get_Range('J' + rowStr);
            if (date != null)
            {
                model.TargetDate = targetDate.Value2;
            }
            // K = areaPath
            model.AreaPath = activeWorksheet.get_Range('K' + rowStr).Value2;

            // L = iterationPath
            model.IterationPath = activeWorksheet.get_Range('L' + rowStr).Value2;

            // M = parentId
            var parentId = activeWorksheet.get_Range('M' + rowStr);
            if (parentId != null)
            {
                model.ParentID = (int)parentId.Value2;
            }

            // N = projectName
            model.ProjectName = activeWorksheet.get_Range('N' + rowStr).Value2;

            return model;
        }
    }
}
