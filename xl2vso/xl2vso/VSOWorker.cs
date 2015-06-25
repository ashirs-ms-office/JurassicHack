using System;
using System.IO;
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
            item.Fields["System.Title"] = "Bug with snapshot";
            item.Fields["System.AssignedTo"] = "Ashirvad Sahu";
            item.Rev = 1;

            //create attachment
            byte[] content = GetBytesFromFile(@"C:\Users\Ashirvad\Desktop\thX4SG17TP.jpg");
            var fileRef = _workItemClient.UploadAttachment(null, null, "test.jpg", content).Result;

            // add attachment to the bug
            item.Relations.Add(new WorkItemRelation()
            {
                Url = fileRef.Url,
                Rel= "AttachedFile",
                Attributes = new RelationAttributes() { Comment = "this is a weird bug"}
            });
            item = _workItemClient.CreateWorkItem(VSOConfig.projectName, type.ToString(), item).Result;
        }

        public void GetWorkItem()
        {
            var item = _workItemClient.GetWorkItem(7);
        }

        public void _office_CreateWorkItem()
        {
            xl2vsoModel dataModel = new xl2vsoModel();
            dataModel.Title = "New work item";
            dataModel.ProjectName = "OExt";
            dataModel.AssignedTo = "Ashirvad Sahu";
            dataModel.AreaPath = @"OExt\Developer Experience and Analytics\Dev Exp\0-60 On Boarding";
            dataModel.IterationPath = @"OExt\Current";
            dataModel.Priority = 4;
            dataModel.CreatedBy = "Keyur Patel";
            dataModel.Description = "here is the description";
            dataModel.Estimate = 7;
            dataModel.WorkItemType = "Task";

            CreateWorkItem(dataModel);
        }

        private void ParkedCode()
        {
            var bug = new WorkItem();
            bug.Fields["System.Title"] = "Build Feature 2";
            bug.Fields["System.AssignedTo"] = "Ashirvad Sahu";
            bug.Fields["System.AreaPath"] = @"OExt\Developer Experience and Analytics";
            bug.Fields["System.Iteration"] = @"OExt\Current";
            bug.Rev = 1;
            bug = _workItemClient.CreateWorkItem(VSOConfig.projectName, "Bug", bug).Result;
        }
        private void CreateWorkItem(xl2vsoModel dataModel)
        {
            // No need to check param here, this will be checked by another function before calling this API
            var workItem = new WorkItem();
            //working
            workItem.Fields["System.Title"] = dataModel.Title;
            workItem.Fields["Microsoft.VSTS.Common.Priority"] = dataModel.Priority;
            workItem.Fields["System.AreaPath"] = dataModel.AreaPath;
            workItem.Fields["System.TeamProject"] = dataModel.Title;
            workItem.Fields["System.AssignedTo"] = dataModel.AssignedTo;
            workItem.Fields["System.Description"] = dataModel.Description;

            // not shown in actual work item
            workItem.Fields["System.CreatedBy"] = dataModel.CreatedBy;
            workItem.Fields["Microsoft.VSTS.Scheduling.Effort"] = dataModel.Estimate;

            // run time error
            //workItem.Fields["System.Iteration"] = dataModel.IterationPath;

            //workItem.Rev = 1;
            workItem = _workItemClient.CreateWorkItem(VSOConfig.projectName, dataModel.WorkItemType, workItem).Result;
        }
        private byte[] GetBytesFromFile(string fullFilePath)
        {
            // this method is limited to 2^32 byte files (4.2 GB)

            FileStream fs = null;
            try
            {
                fs = File.OpenRead(fullFilePath);
                byte[] bytes = new byte[fs.Length];
                fs.Read(bytes, 0, Convert.ToInt32(fs.Length));
                return bytes;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs.Dispose();
                }
            }

        }
    }
}
