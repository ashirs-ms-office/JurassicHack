using System;
using VisualStudioOnline.Api.Rest.V1.Client;
using VisualStudioOnline.Api.Rest.V1.Model;

namespace xl2vso
{
    class VSOWorker
    {
        public enum WorkItemType
        {
            Bug,
            Feature
        }
        private VsoClient _vsoClient = null;
        private IVsoWit _workItemClient = null;
        public VSOWorker()
        {
            _vsoClient = new VsoClient(VSOConfig.accountName, new System.Net.NetworkCredential(VSOConfig.userName, VSOConfig.password));
            _workItemClient = _vsoClient.GetService<IVsoWit>();
        }

        public void QuickCreate(WorkItemType type)
        {
            var item = new WorkItem();
            item.Fields["System.Title"] = "Build Feature 2";
            item.Fields["System.AssignedTo"] = "Ashirvad Sahu";
            item.Rev = 1;
            item = _workItemClient.CreateWorkItem(VSOConfig.projectName, type.ToString(), item).Result;
        }

        public void GetWorkItem()
        {
            var item = _workItemClient.GetWorkItem(7);
        }

        private void _office_CreateWorkItem()
        {
            var bug = new WorkItem();
            bug.Fields["System.Title"] = "Build Feature 2";
            bug.Fields["System.AssignedTo"] = "Ashirvad Sahu";
            bug.Fields["System.AreaPath"] = @"OExt\Developer Experience and Analytics";
            bug.Fields["System.Iteration"] = @"OExt\Current";
            bug.Rev = 1;
            bug = _workItemClient.CreateWorkItem(VSOConfig.projectName, "Bug", bug).Result;
        }
    }
}
