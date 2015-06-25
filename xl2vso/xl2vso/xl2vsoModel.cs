using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xl2vso
{
    class xl2vsoModel
    {
        public string Title { get; set; } //System.Title

        public string AssignedTo { get; set; } //"System.AssignedTo":"Arvind Saxena"

        public string AreaPath { get; set; } //System.AreaPath:OExt\\Developer Experience and Analytics\\Dev Exp\\0-60 On Boarding

        public string IterationPath { get; set; } // "System.IterationPath":"OExt\\QR10\\04-Apr"

        public int Priority { get; set; } //"Microsoft.VSTS.Common.Priority":1,

        public int ParentID { get; set; } // use relations

        public string CreatedBy { get; set; } //"System.CreatedBy":"Keyur Patel"

        public string Description { get; set; } //"System.Description":"<ul><li>Add Office Dev center title to the App Launcher in O365</li><li>On the new Office 365....

        public string ProjectName { get; set; } // "System.TeamProject":"OExt"

        public int OriginalEstimate { get; set; } //     "Microsoft.VSTS.Scheduling.OriginalEstimate": 8,

        public string WorkItemType { get; set; } // Bug, Feature, Task

        public string TargetDate { get; set; } //2015-07-07

    }
}
