using Rally.RestApi;
using Rally.RestApi.Response;
using System;
using System.Collections.Generic;

namespace RallyUtilization
{
    public class RallyClass
    {
        private readonly string password;

        private readonly RallyRestApi restApi;

        private readonly string serverUrl;

        private readonly string username;

        private readonly string workSpaceRef;

        private readonly string startDate;

        private readonly string endDate;

        private readonly string ownerName;

        private Owner owner;

        public RallyClass(string start, string end, string member)
        {
            this.workSpaceRef = "/workspace/";//workspaceID
            this.username = "";//username
            this.password = "";//password
            this.serverUrl = "https://rally1.rallydev.com";
            this.restApi = new RallyRestApi(null, "v2.0", 3, TraceFieldEnum.Data | TraceFieldEnum.Headers | TraceFieldEnum.Cookies);
            this.restApi.Authenticate(this.username, this.password, this.serverUrl, null, false);
            this.startDate = start;
            this.endDate = end;
            this.ownerName = member;
            this.owner = new Owner();
        }

        public Owner LoadWorkSpaceDetails()
        {
            Owner owner;
            try
            {
                Request tasksRequest = new Request("Task")
                {
                    Workspace = this.workSpaceRef,
                    Query = (new Query("Owner.Name", Query.Operator.Equals, this.ownerName))
                        .And((new Query(string.Concat("(Iteration.StartDate = ", this.startDate, ")")))
                            .And(new Query(string.Concat("(Iteration.EndDate = ", this.endDate, ")"))))
                };
                QueryResult queryTaskResult = this.restApi.Query(tasksRequest);
                this.owner.taskCount = queryTaskResult.TotalResultCount;
                foreach (dynamic t in queryTaskResult.Results)
                {
                    this.owner.name = (string)t["Owner"]._refObjectName;
                    this.owner.iteration = (string)t["Iteration"]._refObjectName;
                    owner = this.owner;
                    owner.estimate = (decimal)(owner.estimate + (t["Estimate"] == (dynamic)null ? 0 : t["Estimate"]));
                    //owner = this.owner;
                    owner.toDo = (int)(owner.toDo + (t["ToDo"] == (dynamic)null ? 0 : t["ToDo"]));
                    //owner = this.owner;
                    owner.actuals = (decimal)(owner.actuals + (t["Actuals"] == (dynamic)null ? 0 : t["Actuals"]));
                    owner.timespent = (decimal)(owner.timespent + (t["TimeSpent"] == (dynamic)null ? 0 : t["TimeSpent"]));
                }
                Request request = new Request("UserIterationCapacity")
                {
                    Fetch = new List<string>()
                    {
                        "Capacity",
                        "User",
                        "Role",
                        "EmailAddress",
                        "DisplayName",
                        "UserName"
                    },
                    Query = (new Query("Project.Name", Query.Operator.Equals, "Team1")).
                        Or(new Query("Project.Name", Query.Operator.Equals, "Team2")).
                        Or(new Query("Project.Name", Query.Operator.Equals, "Team3")).
                        Or(new Query("Project.Name", Query.Operator.Equals, "Team4")).
                        And(new Query("User", Query.Operator.Equals, this.ownerName)).
                        And((new Query(string.Concat("(Iteration.StartDate = ", this.startDate, ")"))).
                            And(new Query(string.Concat("(Iteration.EndDate = ", this.endDate, ")"))))
                };
                foreach (dynamic s in this.restApi.Query(request).Results)
                {
                    if (s["Capacity"] == (dynamic)null ||  (decimal)s["Capacity"] == 0 )
                    {
                        continue;
                    }
                    this.owner.capacity = (decimal)s["Capacity"];
                }
                this.owner.timespentPercentage = (int)Math.Round((new decimal(100) * this.owner.timespent) / this.owner.capacity);
                this.owner.estimatePercentage = (int)Math.Round((new decimal(100) * this.owner.estimate) / this.owner.capacity);

                owner = LoadIterationInfo();
            }
            catch (Exception exception)
            {
                owner = null;
            }
            return owner;
        }

        internal Owner LoadIterationInfo()
        {

            int storyCount = 0;
            int defectCount = 0;
            int testSetCount = 0;
            decimal temp = 0;
            decimal totalUSPoints = 0;
            decimal totalDEPoints = 0;
            decimal totalTSPoints = 0;
            decimal totalPoints = 0;

            Request request = new Request("HierarchicalRequirement")
            {
                Workspace = this.workSpaceRef,
                Query = (new Query("Owner.Name", Query.Operator.Equals, this.ownerName))
                    .And((new Query(string.Concat("(Iteration.StartDate = ", this.startDate, ")")))
                        .And(new Query(string.Concat("(Iteration.EndDate = ", this.endDate, ")"))))
            };

            request.PageSize = 500;
            QueryResult queryResults = restApi.Query(request);

            foreach (var s in queryResults.Results)
            {
                Console.WriteLine(" Name: " + s["Name"] + " ;Project: " + s["Project"]._refObjectName + ";FormattedID: " + s["FormattedID"] + s["PlanEstimate"]);
                storyCount++;
                if (s["PlanEstimate"] != null && decimal.TryParse(s["PlanEstimate"].ToString(), out temp))
                {
                    totalUSPoints += temp;
                }

            }

            request = new Request("defect")
            {
                Workspace = this.workSpaceRef,
                Query = (new Query("Owner.Name", Query.Operator.Equals, this.ownerName))
                    .And((new Query(string.Concat("(Iteration.StartDate = ", this.startDate, ")")))
                        .And(new Query(string.Concat("(Iteration.EndDate = ", this.endDate, ")"))))
            };

            request.PageSize = 500;

            queryResults = restApi.Query(request);
            foreach (var result in queryResults.Results)
            {
                Console.WriteLine("Defect name: " + result["Name"] + " ;Project: " + result["Project"]._refObjectName + ";FormattedID: " + result["FormattedID"]);
                defectCount++;
                if (result["PlanEstimate"] != null && decimal.TryParse(result["PlanEstimate"].ToString(), out temp))
                {
                    totalDEPoints += temp;
                }
            }
            request = new Request("testset")
            {
                Workspace = this.workSpaceRef,
                Query = (new Query("Owner.Name", Query.Operator.Equals, this.ownerName))
                    .And((new Query(string.Concat("(Iteration.StartDate = ", this.startDate, ")")))
                        .And(new Query(string.Concat("(Iteration.EndDate = ", this.endDate, ")"))))
            };

            request.PageSize = 500;

            queryResults = restApi.Query(request);
            foreach (var result in queryResults.Results)
            {
                Console.WriteLine("Testset name: " + result["Name"] + " ;Project: " + result["Project"]._refObjectName + ";FormattedID: " + result["FormattedID"]);
                testSetCount++;
                if (result["PlanEstimate"] != null && decimal.TryParse(result["PlanEstimate"].ToString(), out temp))
                {
                    totalTSPoints += temp;
                }
            }
            totalPoints = totalUSPoints + totalDEPoints + totalTSPoints;
            owner.totalPoints = totalPoints;
            return owner;
        }
    }
}