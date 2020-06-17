using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Brugeradministration
{
    public class GetUserRoles : IPlugin
    {

  /*      [Output("Roller")]
        public OutArgument<string> Roller { get; set; }

        [Output("Licens")]
        public OutArgument<string> Licens { get; set; }*/

        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
               serviceProvider.GetService(typeof(IPluginExecutionContext)); ;

            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory =
               (IOrganizationServiceFactory)serviceProvider.GetService
               (typeof(IOrganizationServiceFactory));
            IOrganizationService service =
               serviceFactory.CreateOrganizationService(context.UserId);

            

            var SystemUser = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));

            // make query for roles
            var query = new QueryExpression();
            query.EntityName = "role";
            query.ColumnSet = new ColumnSet(true);

            // link entity - fra rolle til rolle relation
            LinkEntity role = new LinkEntity();
            role.LinkFromEntityName = "role";
            role.LinkFromAttributeName = "roleid";
            role.LinkToEntityName = "systemuserroles";
            role.LinkToAttributeName = "roleid";

            // fra rolle relation til bruger
            LinkEntity userRoles = new LinkEntity();
            userRoles.LinkFromEntityName = "systemuserroles";
            userRoles.LinkFromAttributeName = "systemuserid";
            userRoles.LinkToEntityName = "systemuser";
            userRoles.LinkToAttributeName = "systemuserid";

            // lav condition til query
            ConditionExpression conditionExpression = new ConditionExpression();
            conditionExpression.AttributeName = "systemuserid";
            conditionExpression.Operator = ConditionOperator.Equal;
            conditionExpression.Values.Add(SystemUser.Id);

            userRoles.LinkCriteria = new FilterExpression();
            userRoles.LinkCriteria.Conditions.Add(conditionExpression);

            // link entiterne sammen
            role.LinkEntities.Add(userRoles);
            query.LinkEntities.Add(role);

            // definer regler!
            var sales = new List<string> { "Lead", "Opportunity", "Goal", "Contract", "Quote", "Order", "Invoice", "Competitor", "Campaign", "List", "cdi_" }; // CLICKDIMENSION


            // få alle roller fra user
            var UserRoles = service.RetrieveMultiple(query).Entities;

            List<string> allRolesList = new List<string>();

            foreach (var Role in UserRoles)
            {
                allRolesList.AddRange(getMasterRole(Role.GetAttributeValue<string>("name"), service, true));

            }

            // få alle roller fra team
            var TeamRoles = GetRolesFromTeam(SystemUser, service);

            foreach (var teamRole in TeamRoles)
            {
                allRolesList.AddRange(getMasterRole(teamRole, service, false));
            }

            // fjern dupletter
            var allRolesListDistinct = allRolesList.Distinct();
            var allRolesListDistinctList = allRolesListDistinct.ToList();

            // returner en streng.
            var returnString = new StringBuilder();

            var salesString = new StringBuilder();
            var customEntitiesString = new StringBuilder();

            var salesCounter = 0;
            var customEntityCounter = 0;


            var customEntities = new List<string>();
            var salesEntities = new List<string>();

            string[] seperator = { "_" };

            // tjek om der er roller der matcher med sales -> stor licens!
            foreach (var privilege in allRolesListDistinctList)
            {

                for (int i = 0; i < sales.Count; i++)
                {
                    var salesSplit = privilege.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
                    var salesLastMember = salesSplit.Length - 1;

                    var salesName = salesSplit.GetValue(salesLastMember).ToString();

                    // "SalesProcess" bliver ekskluderet - det er en adgang til BPF
                    if (privilege.Contains(sales.ElementAt(i)) && !privilege.Contains("SalesProcess") && !salesEntities.Contains(salesName))
                    {
                        salesEntities.Add(salesName);
                        salesCounter = salesCounter + 1;
                        salesString.AppendLine(privilege);
                    }
                }


                if ((privilege.Contains("sdu_") || privilege.Contains("mp_") || privilege.Contains("nn_") || privilege.Contains("alumne_")) && !privilege.Contains("bpf_"))
                {
                    var entitySplit = privilege.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
                    var entityLastMember = entitySplit.Length - 1;

                    var entityName = entitySplit.GetValue(entityLastMember).ToString();



                    if (!customEntities.Contains(entityName))
                    {
                        customEntities.Add(entityName);
                        customEntityCounter = 1 + customEntityCounter;
                        customEntitiesString.AppendLine(privilege);
                    }
                    else
                    {
                        continue;
                    }

                }
            }

            Entity systemUser = new Entity("systemuser");
            systemUser.Id = SystemUser.Id;
            systemUser["sdu_aktivtjek"] = CheckLastLoginDate(service, SystemUser.Id);
            

            if (salesCounter > 0 || customEntityCounter > 15)
            {
                systemUser["sdu_rolletype"] = "Sales";
            }
            else
            {
                systemUser["sdu_rolletype"] = "Team Member";
            }

            returnString.AppendLine("- SALES: " + salesCounter.ToString());
            returnString.AppendLine("- Custom Entities: " + customEntityCounter.ToString());
            returnString.AppendLine("");
            returnString.AppendLine("Sales: ");
            returnString.AppendLine(salesString.ToString());
            returnString.AppendLine("");
            returnString.AppendLine("Custom Entities: ");
            returnString.AppendLine(customEntitiesString.ToString());

            systemUser["sdu_rettigheder"] = returnString.ToString();

            service.Update(systemUser);
        }

        


        public List<string> getMasterRole(string roleName, IOrganizationService service, bool user)
        {
            var query = new QueryExpression("role");
            query.ColumnSet.AddColumns("name", "roleid");

            query.Criteria.AddCondition("name", ConditionOperator.Equal, roleName);
            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, "ecd7eb59-0441-e111-8cfb-0050568a0051"); // id corresponds to sdupro

            var results = service.RetrieveMultiple(query).Entities;

            var masterPrivileges = new List<string>();

            foreach (var record in results)
            {
                masterPrivileges = GetPrivileges(record.GetAttributeValue<Guid>("roleid").ToString(), roleName, service, user);
            }

            return masterPrivileges;

        }

        public List<string> GetPrivileges(string roleID, string roleName, IOrganizationService service, bool user)
        {
            var privQuery = new QueryExpression();
            privQuery.EntityName = "privilege";
            privQuery.ColumnSet = new ColumnSet(true);

            // link entity - fra rolle til rolle relation
            LinkEntity privToRolePriv = new LinkEntity();
            privToRolePriv.LinkFromEntityName = "privilege";
            privToRolePriv.LinkFromAttributeName = "privilegeid";
            privToRolePriv.LinkToEntityName = "roleprivileges";
            privToRolePriv.LinkToAttributeName = "privilegeid";

            // link entity - fra rolle til rolle relation
            LinkEntity rolePrivToRole = new LinkEntity();
            rolePrivToRole.LinkFromEntityName = "roleprivileges";
            rolePrivToRole.LinkFromAttributeName = "roleid";
            rolePrivToRole.LinkToEntityName = "role";
            rolePrivToRole.LinkToAttributeName = "roleid";

            // lav condition til query
            ConditionExpression privConditionExpression = new ConditionExpression();
            privConditionExpression.AttributeName = "roleid";
            privConditionExpression.Operator = ConditionOperator.Equal;
            privConditionExpression.Values.Add(roleID); // ROLLEID fra query

            rolePrivToRole.LinkCriteria = new FilterExpression();
            rolePrivToRole.LinkCriteria.Conditions.Add(privConditionExpression);

            // link entiterne sammen
            privToRolePriv.LinkEntities.Add(rolePrivToRole);
            privQuery.LinkEntities.Add(privToRolePriv);

            var rolePrivileges = service.RetrieveMultiple(privQuery).Entities;


            var allPrivilegesList = new List<string>();

            //var allPrivilegesListKeyValue = new List<KeyValuePair<string, string>>();


            foreach (var priv in rolePrivileges)
            {
                var nameOfPriv = priv.GetAttributeValue<string>("name");
                if (nameOfPriv.Contains("Write") || nameOfPriv.Contains("Create"))
                {
                    if (user)
                    {
                        allPrivilegesList.Add("(U): " + roleName + " - " + priv.GetAttributeValue<string>("name"));
                    }
                    else
                    {
                        allPrivilegesList.Add("(T): " + roleName + " - " + priv.GetAttributeValue<string>("name"));

                    }
                }
            }

            return allPrivilegesList;
        }

        public List<string> GetRolesFromTeam(Entity systemUser, IOrganizationService service)
        {
            var query = new QueryExpression();
            query.EntityName = "role";
            query.ColumnSet = new ColumnSet(true);

            LinkEntity roleToTeamRole = new LinkEntity();
            roleToTeamRole.LinkFromEntityName = "role";
            roleToTeamRole.LinkFromAttributeName = "roleid";
            roleToTeamRole.LinkToEntityName = "teamroles";
            roleToTeamRole.LinkToAttributeName = "roleid";

            LinkEntity TeamRolestoTeam = new LinkEntity();
            TeamRolestoTeam.LinkFromEntityName = "teamroles";
            TeamRolestoTeam.LinkFromAttributeName = "teamid";
            TeamRolestoTeam.LinkToEntityName = "team";
            TeamRolestoTeam.LinkToAttributeName = "teamid";

            LinkEntity TeamToTeamMembership = new LinkEntity();
            TeamToTeamMembership.LinkFromEntityName = "team";
            TeamToTeamMembership.LinkFromAttributeName = "teamid";
            TeamToTeamMembership.LinkToEntityName = "teammembership";
            TeamToTeamMembership.LinkToAttributeName = "teamid";

            LinkEntity teamMembershipToUser = new LinkEntity();
            teamMembershipToUser.LinkFromEntityName = "teammembership";
            teamMembershipToUser.LinkFromAttributeName = "systemuserid";
            teamMembershipToUser.LinkToEntityName = "systemuser";
            teamMembershipToUser.LinkToAttributeName = "systemuserid";

            // lav condition til query
            ConditionExpression conditionExpression = new ConditionExpression();
            conditionExpression.AttributeName = "systemuserid";
            conditionExpression.Operator = ConditionOperator.Equal;
            conditionExpression.Values.Add(systemUser.Id);

            teamMembershipToUser.LinkCriteria = new FilterExpression();
            teamMembershipToUser.LinkCriteria.Conditions.Add(conditionExpression);

            // LINK
            TeamToTeamMembership.LinkEntities.Add(teamMembershipToUser);

            TeamRolestoTeam.LinkEntities.Add(TeamToTeamMembership);

            roleToTeamRole.LinkEntities.Add(TeamRolestoTeam);

            query.LinkEntities.Add(roleToTeamRole);

            var teamRolesQuery = service.RetrieveMultiple(query).Entities;

            List<string> teamRoles = new List<string>();

            //Console.WriteLine(teamRolesQuery.Count);
            foreach (var teamRole in teamRolesQuery)
            {
                teamRoles.Add(teamRole.GetAttributeValue<string>("name"));
            }

            return teamRoles;


        }


        public DateTime CheckLastLoginDate(IOrganizationService service, Guid systemUserId)
        {
            try
            {
                // Retrieve the audit history for the account and display it.
                RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                changeRequest.Target = new EntityReference("systemuser", systemUserId);

                RetrieveRecordChangeHistoryResponse changeResponse = (RetrieveRecordChangeHistoryResponse)service.Execute(changeRequest);

                AuditDetailCollection details = changeResponse.AuditDetailCollection;

                var auditRecords = details.AuditDetails;

                for (int i = 0; i < auditRecords.Count; i++)
                {
                    // READ: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/audit?view=dynamics-ce-odata-9#see-also
                    // 64 = "User Access via Web"
                    // 65 = "User Access via Web Services"

                    var auditActionType = auditRecords[i].AuditRecord.GetAttributeValue<OptionSetValue>("action").Value;
                    var auditCreatedOn = auditRecords[i].AuditRecord.GetAttributeValue<DateTime>("createdon").ToLocalTime();

                    if ((auditActionType == 64 || auditActionType == 65) && auditCreatedOn > DateTime.Today.AddMonths(-6))
                    {
                        return auditCreatedOn;
                    }

                }
                return new DateTime(1900, 01, 01);
                // EVT: https://docs.microsoft.com/en-us/dynamics365/customerengagement/on-premises/developer/sample-disable-user
            }
            catch (Exception)
            {

                throw;
            }


        }


    }


}
