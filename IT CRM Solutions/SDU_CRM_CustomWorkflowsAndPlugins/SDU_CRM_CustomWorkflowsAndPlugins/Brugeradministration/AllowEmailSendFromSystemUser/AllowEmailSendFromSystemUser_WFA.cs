using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;

namespace Brugeradministration
{
    public class Brugeradministration_AllowEmailSendFromSystemUser_WFA : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {
            // Obtain the execution context from the service provider.
            // Get the context service.
            IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.             
            IOrganizationService service = serviceFactory.CreateOrganizationService(Icontext.InitiatingUserId);

            // get current record
            var systemUser = service.Retrieve(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, new ColumnSet("defaultmailbox", "sdu_brugeradministration", "domainname", "businessunitid"));

            var username = systemUser.GetAttributeValue<string>("domainname").Split('\\')[1];

            // search for brugeradministration
            var query = new QueryExpression("contact")
            {
                ColumnSet = new ColumnSet("sdu_brugernavn", "sdu_adgange")
            };

            var condition = new ConditionExpression("sdu_brugernavn", ConditionOperator.Equal, username);

            query.Criteria.AddCondition(condition);

            var result = service.RetrieveMultiple(query);

            // only if one result
            if (result.Entities.Count == 1)
            {
                var contact = result.Entities[0];
                var brugeradmRef = contact.GetAttributeValue<EntityReference>("sdu_adgange");

                var SystemUser = new Entity(Icontext.PrimaryEntityName)
                {
                    Id = Icontext.PrimaryEntityId
                };
                SystemUser["sdu_brugeradministration"] = brugeradmRef;

                service.Update(SystemUser);
            }


            // get the business unit
            var businessUnitRef = systemUser.GetAttributeValue<EntityReference>("businessunitid");
            var businessUnitEntity = service.Retrieve(businessUnitRef.LogicalName, businessUnitRef.Id, new ColumnSet("name"));
            var nameOfBU = businessUnitEntity.GetAttributeValue<string>("name");

            // update issendasallow attribute on user settings
            var userSettings = new Entity("usersettings")
            {
                Id = Icontext.PrimaryEntityId
            };
            userSettings["issendasallowed"] = true;

            if (nameOfBU == "ASS_930 Bestilling af specialkonti")
            {
                userSettings["homepagearea"] = "Settings";
                userSettings["homepagesubarea"] = "sdu_kontobestilling";
            }
            else if (nameOfBU.StartsWith("ASS_"))
            {
                userSettings["homepagearea"] = "Workplace";
                userSettings["homepagesubarea"] = "Dashboards";
            }

            service.Update(userSettings);


            // update delivery method of default mailbox
            var defaultMailBox = new Entity("mailbox")
            {
                Id = systemUser.GetAttributeValue<EntityReference>("defaultmailbox").Id
            };
            defaultMailBox["outgoingemaildeliverymethod"] = new OptionSetValue(2);

            service.Update(defaultMailBox);
        }

    }
}
