using Aspose.Tasks;
using Aspose.Tasks.Util;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using ProjectService.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ConstraintType = Microsoft.ProjectServer.Client.ConstraintType;

namespace ProjectService.Controllers
{
    public class ProjectController : ApiController
    {
        int defaultTimeoutSeconds = int.Parse(ConfigurationManager.AppSettings["DefaultTimeoutSeconds"]);

        [AcceptVerbs("POST")]
        [Route("api/Project/CreateProject")]
        public Result CreateProject([FromBody]JObject value)
        {
            string CurrentUser = RequestContext.Principal.Identity.Name;
            dynamic JsonValue = value;

            Guid id = JsonValue == null || JsonValue.id == null ? Guid.NewGuid() : Guid.Parse(Convert.ToString(JsonValue.id));

            if (JsonValue != null)
            {
                if (JsonValue.name != null && JsonValue.date != null)
                {
                    try
                    {
                        int defaultTimeoutSeconds = int.Parse(ConfigurationManager.AppSettings["DefaultTimeoutSeconds"]);
                        string url = ConfigurationManager.AppSettings["Site"];
                        var context = new ProjectContext(url)
                        {
                            //for accessing to server
                            Credentials = new NetworkCredential(ConfigurationManager.AppSettings["domain_user"], ConfigurationManager.AppSettings["domain_password"])
                        };

                        PublishedProject project = context.Projects.Add(new ProjectCreationInformation()
                        {
                            Id = id,
                            Name = JsonValue.name,
                            Start = Convert.ToDateTime(JsonValue.date),
                            Description = JsonValue.description == null ? string.Empty : JsonValue.description,
                        });

                        JobState jobState = context.WaitForQueue(context.Projects.Update(), defaultTimeoutSeconds);
                        IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                        context.ExecuteQuery();

                        var ppr = projs.FirstOrDefault();

                        DraftProject draft = project.CheckOut();

                        var users = context.LoadQuery(context.Web.SiteUsers);
                        context.Load(context.CustomFields);
                        context.ExecuteQuery();


                        //sample field handler
                        // var sampleField = context.CustomFields.FirstOrDefault(cf => cf.Name == "SampleName");
                        // draft[sampleField.InternalName] = JsonValue.realizationIndicator == null ? string.Empty : Convert.ToString(JsonValue.sampleName);


                        if (JsonValue.owner != null)
                        {
                            string ownerName = Convert.ToString(JsonValue.owner).ToLower();
                            var user = users.FirstOrDefault(x => x.LoginName.Split('\\').Count() > 1 && x.LoginName.Split('\\')[1].ToLower().Equals(ownerName));
                            draft.Owner = user;
                        }

                        draft.Update();

                        JobState jobSta5te = context.WaitForQueue(draft.Publish(false), defaultTimeoutSeconds);
                        context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);
                    }
                    catch (Exception ex)
                    {
                        return new Result() { succeeded = false, code = 400, data = "", message = ex.Message };
                    }

                }
                else
                {
                    return new Result() { succeeded = false, code = 400, data = "", message = "Wrong Parameters" };
                }
            }

            return new Result() { succeeded = true, code = 200, data = id, message = "Registered Successfully" };
        }

        [AcceptVerbs("POST")]
        [Route("api/Project/UpdateProjectTasks")]
        public Result UpdateProjectTasks([FromBody]JObject value)
        {
            string CurrentUser = RequestContext.Principal.Identity.Name;
            dynamic JsonValue = value;
            if (JsonValue != null)
            {
                if (JsonValue.id != null && (JsonValue.file != null || JsonValue.address != null))
                {
                    Guid id = Guid.Parse(Convert.ToString(JsonValue.id));

                    string url = ConfigurationManager.AppSettings["Site"];
                    var context = new ProjectContext(url)
                    {
                        Credentials = new NetworkCredential(ConfigurationManager.AppSettings["domain_user"], ConfigurationManager.AppSettings["domain_password"])
                    };

                    context.RequestTimeout = 10 * 60 * 1000;

                    IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                    context.ExecuteQuery();

                    PublishedProject project = projs.FirstOrDefault();

                    DraftProject draft = project.CheckOut();
                    var prj = JsonValue.file == null ? new Aspose.Tasks.Project((string)JsonValue.address) : new Aspose.Tasks.Project((Stream)JsonValue.file);
                    var collector = new ChildTasksCollector();
                    TaskUtils.Apply(prj.RootTask, collector, 0);

                    var counter = 1;

                    var main = Guid.Parse(collector.Tasks.First().Get(Tsk.Guid));
                    var list = new List<Guid>();

                    CreateItems(counter, main, true, collector.Tasks.First(), draft, list);

                    draft.Update();

                    JobState jobState = context.WaitForQueue(draft.Publish(true), defaultTimeoutSeconds);
                    context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);

                    System.Threading.Thread.Sleep(60 * 30);

                    projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                    context.ExecuteQuery();

                    draft = project.CheckOut();

                    foreach (var item in prj.TaskLinks)
                    {
                        var related = new TaskLinkCreationInformation
                        {
                            StartId = Guid.Parse(item.PredTask.Get(Tsk.Guid)),
                            EndId = Guid.Parse(item.SuccTask.Get(Tsk.Guid)),
                            DependencyType = (DependencyType)((int)item.LinkType),
                        };

                        draft.TaskLinks.Add(related);
                    }

                    draft.Update();

                    jobState = context.WaitForQueue(draft.Publish(true), defaultTimeoutSeconds);
                    context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);

                    System.Threading.Thread.Sleep(60 * 30);

                    projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                    context.ExecuteQuery();

                    draft = project.CheckOut();

                    context.Load(draft.Tasks);
                    context.ExecuteQuery();

                    foreach (var item in draft.Tasks)
                    {
                        item.RefreshLoad();

                        if (item.ConstraintType != ConstraintType.AsSoonAsPossible && item.ConstraintType != ConstraintType.AsLateAsPossible)
                        {
                            if (item.ConstraintType == ConstraintType.FinishNoEarlierThan || item.ConstraintType == ConstraintType.FinishNoLaterThan || item.ConstraintType ==ConstraintType.MustFinishOn)
                            {
                                item.ConstraintStartEnd = item.Finish;
                            }
                            else
                            {
                                item.ConstraintStartEnd = item.Start;
                            }
                        }
                    }

                    draft.Update();

                    jobState = context.WaitForQueue(draft.Publish(true), defaultTimeoutSeconds);
                    context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);
                }
                else
                {
                    return new Result() { succeeded = false, code = 400, data = "", message = "Wrong Parameters" };
                }
            }

            return new Result() { succeeded = true, code = 200, data = "", message = "Registered Successfully" };
        }

        [AcceptVerbs("POST")]
        [Route("api/Project/UpdateProject")]
        public Result UpdateProject([FromBody]JObject value)
        {
            string CurrentUser = RequestContext.Principal.Identity.Name;
            dynamic JsonValue = value;
            if (JsonValue != null)
            {
                if (JsonValue.id != null)
                {
                    Guid id = Guid.Parse(Convert.ToString(JsonValue.id));

                    int defaultTimeoutSeconds = int.Parse(ConfigurationManager.AppSettings["DefaultTimeoutSeconds"]);
                    string url = ConfigurationManager.AppSettings["Site"];
                    var context = new ProjectContext(url)
                    {
                        Credentials = new NetworkCredential(ConfigurationManager.AppSettings["domain_user"], ConfigurationManager.AppSettings["domain_password"])
                    };

                    try
                    {
                        IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                        context.ExecuteQuery();

                        var ppr = projs.FirstOrDefault();
                        DraftProject draft = ppr.CheckOut();

                        context.Load(context.CustomFields);
                        context.ExecuteQuery();

                        // for get custum fields you can use this code
                        var sampleField = context.CustomFields.FirstOrDefault(cf => cf.Name == "SampleName").InternalName;

                        context.Load(draft, 
                            d => d.Name, 
                            d => d.StartDate, 
                            d => d.Description,
                        d => d[sampleField]);

                        var enterprises = context.LoadQuery(context.EnterpriseResources);
                        var users = context.LoadQuery(context.Web.SiteUsers);
                        context.Load(draft.Tasks);

                        context.ExecuteQuery();

                        draft[sampleField] = JsonValue.realizationIndicator == null ? draft[sampleField] : Convert.ToString(JsonValue.sampleField);

                        if (JsonValue.owner != null)
                        {
                            string ownerName = Convert.ToString(JsonValue.owner).ToLower();
                            var user = users.FirstOrDefault(x => x.LoginName.Split('\\').Count() > 1 && x.LoginName.Split('\\')[1].ToLower().Equals(ownerName));
                            draft.Owner = user;
                        }

                        if (JsonValue.resource != null)
                        {
                            string resourceName = Convert.ToString(JsonValue.resource);
                            var resourceId = Guid.NewGuid();
                            var user = enterprises.FirstOrDefault(x => x.Email != null && x.Email.StartsWith(resourceName));

                            draft.ProjectResources.AddEnterpriseResource(user);

                            foreach (var item in draft.Tasks)
                            {
                                if (!item.IsSummary)
                                {
                                    draft.Assignments.Add(new AssignmentCreationInformation { TaskId = item.Id, Finish = item.Finish, Start = item.Start, ResourceId = user.Id });
                                }
                            }
                        }

                        draft.Update();

                        JobState jobSta5te = context.WaitForQueue(draft.Publish(false), defaultTimeoutSeconds);
                        context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);
                    }
                    catch (Exception ex)
                    {
                        return new Result() { succeeded = false, code = 400, data = "", message = ex.Message };
                    }

                }
                else
                {
                    return new Result() { succeeded = false, code = 400, data = "", message = "Wrong Parameters" };
                }
            }

            return new Result() { succeeded = true, code = 200, data = "", message = "Registered Successfully" };
        }

        [AcceptVerbs("POST")]
        [Route("api/Project/UpdateProjectResources")]
        public Result UpdateProjectResources([FromBody]JObject value)
        {
            string CurrentUser = RequestContext.Principal.Identity.Name;
            dynamic JsonValue = value;
            if (JsonValue != null)
            {
                if (JsonValue.id != null)
                {
                    Guid id = Guid.Parse(Convert.ToString(JsonValue.id));
                    string url = ConfigurationManager.AppSettings["Site"];
                    var context = new ProjectContext(url)
                    {
                        Credentials = new NetworkCredential(ConfigurationManager.AppSettings["domain_user"], ConfigurationManager.AppSettings["domain_password"])
                    };

                    IEnumerable<PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Id == id));
                    context.ExecuteQuery();

                    PublishedProject project = projs.FirstOrDefault();
                    DraftProject draft = project.CheckOut();

                    context.Load(draft.Tasks);
                    context.Load(draft.Assignments);
                    context.ExecuteQuery();

                    var enterprises = context.LoadQuery(context.EnterpriseResources);
                    var users = context.LoadQuery(context.Web.SiteUsers);

                    context.ExecuteQuery();

                    if (JsonValue.owner != null)
                    {
                        string ownerName = Convert.ToString(JsonValue.owner).ToLower();
                        var user = users.FirstOrDefault(x => x.LoginName.Split('\\').Count() > 1 && x.LoginName.Split('\\')[1].ToLower().Equals(ownerName));
                        draft.Owner = user;
                    }

                    if (JsonValue.resource != null)
                    {
                        string resourceName = Convert.ToString(JsonValue.resource);
                        var resourceId = Guid.NewGuid();
                        var user_item = users.FirstOrDefault(x => x.LoginName.Split('\\').Count() > 1 && x.LoginName.Split('\\')[1].ToLower().Equals(resourceName));
                        var user = enterprises.FirstOrDefault(x => x.Email != null && x.Email.Equals(user_item.Email));


                        draft.ProjectResources.AddEnterpriseResource(user);

                        foreach (var item in draft.Tasks)
                        {
                            if (!item.IsSummary)
                            {
                                draft.Assignments.Add(new AssignmentCreationInformation { TaskId = item.Id, Finish = item.Finish, Start = item.Start, ResourceId = user.Id });
                            }
                        }
                    }

                    draft.Update();

                    JobState jobState = context.WaitForQueue(draft.Publish(false), defaultTimeoutSeconds);
                    context.WaitForQueue(draft.CheckIn(true), defaultTimeoutSeconds);
                }
                else
                {
                    return new Result() { succeeded = false, code = 400, data = "", message = "Wrong Parameters" };
                }
            }

            return new Result() { succeeded = true, code = 200, data = "", message = "Registered Successfully" };
        }

        public Aspose.Tasks.Task CreateItems(int count, Guid main, bool isFirst, Aspose.Tasks.Task task, DraftProject draft, List<Guid> list)
        {
            var task_item = new TaskCreationInformation
            {
                Id = Guid.Parse(task.Get(Tsk.Guid)),
                Name = task.Get(Tsk.Name),
                IsManual = false,
                Start = task.Get(Tsk.Start),
                Duration = task.Get(Tsk.Duration).ToString(),
            };

            if ((task.ParentTask != null && !isFirst && !Guid.Parse(task.ParentTask.Get(Tsk.Guid)).Equals(main)) || (task.ParentTask != null && isFirst))
            {
                task_item.ParentId = Guid.Parse(task.ParentTask.Get(Tsk.Guid));

            }

            if (list.Count > 0)
            {
                task_item.AddAfterId = list.Last();
            }

            if (!isFirst)
            {
                list.Add(Guid.Parse(task.Get(Tsk.Guid)));

                var draft_item = draft.Tasks.Add(task_item);


                var constraintType = (ConstraintType)((int)task.Get(Tsk.ConstraintType) + 1);

                if (constraintType != ConstraintType.AsSoonAsPossible && constraintType != ConstraintType.AsLateAsPossible)
                {
                    draft_item.ConstraintStartEnd = task.Get(Tsk.ConstraintDate);
                }

                draft_item.OutlineLevel = task.Get(Tsk.OutlineLevel);
                draft_item.Priority = task.Get(Tsk.Priority);
                draft_item["WBS"] = task.Get(Tsk.OutlineNumber);
                draft_item.IsManual = task.Get(Tsk.IsManual);
                draft_item.ConstraintType = constraintType;
            }

            foreach (var item in task.Children)
            {
                CreateItems(count, main, false, item, draft, list);
            }

            return task;
        }
    }
}
