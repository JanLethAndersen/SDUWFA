using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Brugeradministration
{
    public class Brugeradministration_AllowEmailSendFromSystemUser : IPlugin
    {

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

            // get current record
            var systemUser = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("defaultmailbox", "sdu_brugeradministration", "domainname", "businessunitid"));

            var username = systemUser.GetAttributeValue<string>("domainname").Split('\\')[1];

            // search for brugeradministration
            var query = new QueryExpression("contact");
            query.ColumnSet = new ColumnSet("sdu_brugernavn", "sdu_adgange");

            var condition = new ConditionExpression("sdu_brugernavn", ConditionOperator.Equal, username);

            query.Criteria.AddCondition(condition);

            var result = service.RetrieveMultiple(query);

            // only if one result
            if (result.Entities.Count == 1) {
                var contact = result.Entities[0];
                var brugeradmRef = contact.GetAttributeValue<EntityReference>("sdu_adgange");

                var SystemUser = new Entity(context.PrimaryEntityName)
                {
                    Id = context.PrimaryEntityId
                };
                SystemUser["sdu_brugeradministration"] = brugeradmRef;

                service.Update(SystemUser);
            }

            // get the business unit
            var businessUnitRef = systemUser.GetAttributeValue<EntityReference>("businessunitid");
            var businessUnitEntity = service.Retrieve(businessUnitRef.LogicalName, businessUnitRef.Id, new ColumnSet("name"));
            var nameOfBU = businessUnitEntity.GetAttributeValue<string>("name");

            
            var userSettings = new Entity("usersettings")
            {
                Id = context.PrimaryEntityId
            };
            // update issendasallow attribute on user settings
            userSettings["issendasallowed"] = true;

            if (nameOfBU == "ASS_930 Bestilling af specialkonti")
            {
                // update the startpage for people only ordering user accounts
                userSettings["homepagearea"] = "Settings";
                userSettings["homepagesubarea"] = "sdu_kontobestilling";
            } else if (nameOfBU.StartsWith("ASS_"))
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
