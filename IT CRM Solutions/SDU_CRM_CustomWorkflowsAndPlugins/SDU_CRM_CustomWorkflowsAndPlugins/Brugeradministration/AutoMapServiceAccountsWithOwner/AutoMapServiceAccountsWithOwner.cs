using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Brugeradministration
{
    public class AutoMapServiceAccountsWithOwner : CodeActivity
    {
        [Input("AD Manager")]
        public InArgument<string> AdManager { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Get the context service.
                IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

                // Use the context service to create an instance of IOrganizationService.             
                IOrganizationService service = serviceFactory.CreateOrganizationService(Icontext.InitiatingUserId);


                var adManagerPresent = String.IsNullOrEmpty(AdManager.Get(context));
                var fullnameOfOwner = adManagerPresent == true ? "" : GetFullnameOfOwner(AdManager.Get(context));

                if (fullnameOfOwner != "")
                {
                    // Define Condition Values
                    var QEsdu_brugeradministration_statecode = 0;
                    var QEsdu_brugeradministration_contact_fullname = "%" + fullnameOfOwner + "%";

                    // Instantiate QueryExpression QEsdu_brugeradministration
                    var QEsdu_brugeradministration = new QueryExpression("sdu_brugeradministration");

                    // Add columns to QEsdu_brugeradministration.ColumnSet
                    QEsdu_brugeradministration.ColumnSet.AddColumns("sdu_brugeradministrationid", "sdu_name", "createdon");
                    QEsdu_brugeradministration.AddOrder("sdu_name", OrderType.Ascending);

                    // Define filter QEsdu_brugeradministration.Criteria
                    QEsdu_brugeradministration.Criteria.AddCondition("statecode", ConditionOperator.Equal, QEsdu_brugeradministration_statecode);

                    // Add link-entity QEsdu_brugeradministration_contact
                    var QEsdu_brugeradministration_contact = QEsdu_brugeradministration.AddLink("contact", "sdu_person", "contactid");
                    QEsdu_brugeradministration_contact.EntityAlias = "ab";

                    // Define filter QEsdu_brugeradministration_contact.LinkCriteria
                    QEsdu_brugeradministration_contact.LinkCriteria.AddCondition("fullname", ConditionOperator.Like, QEsdu_brugeradministration_contact_fullname);
                    QEsdu_brugeradministration_contact.LinkCriteria.AddCondition("sdu_medarbejdernummer", ConditionOperator.NotNull);


                    var results = service.RetrieveMultiple(QEsdu_brugeradministration);

                    if (results.Entities.Count == 1)
                    {
                        UpdateCurrentRecord(Icontext, service, results[0]);
                    }

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static string GetFullnameOfOwner(string Dn)
        {
                var firstComma = Dn.IndexOf(',');
                var firstEqual = Dn.IndexOf('=') + 1;

                if (firstComma != -1 && firstEqual != -1)
                {
                    var fullNameOfOwner = Dn.Substring(firstEqual, firstComma - firstEqual);

                    if (fullNameOfOwner == "" || fullNameOfOwner == null)
                    {
                        return "";
                    }
                    else
                    {
                        return fullNameOfOwner;
                    }
                } else
            {
                return "";
            }
            }

        public static void UpdateCurrentRecord(IWorkflowContext Icontext, IOrganizationService Service, Entity brugeradministrationOwner)
        {
            var entity = new Entity(Icontext.PrimaryEntityName)
            {
                Id = Icontext.PrimaryEntityId
            };

            entity["sdu_ansvarligid"] = new EntityReference(brugeradministrationOwner.LogicalName, brugeradministrationOwner.Id);

            Service.Update(entity);
        }
    }
}


