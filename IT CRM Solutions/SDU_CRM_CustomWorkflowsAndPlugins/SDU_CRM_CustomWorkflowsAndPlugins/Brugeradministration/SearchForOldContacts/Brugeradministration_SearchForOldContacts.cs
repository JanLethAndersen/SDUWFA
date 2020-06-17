using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;

namespace Brugeradministration
{
    public class SearchForOldContacts : CodeActivity
    {
        [Input("username")]
        public InArgument<string> Username { get; set; }

        [Input("birthdate")]
        public InArgument<DateTime> Birthdate { get; set; }

        [Input("firstname")]
        public InArgument<string> Firstname { get; set; }

        [Input("lastname")]
        public InArgument<string> Lastname { get; set; }

        [Input("Status")]
        public InArgument<bool> RecordStatus { get; set; }

        [Input("CRM Udløbet er maks (måneder) gammelt")]
        public InArgument<int> ExpirationDateLimit { get; set; }

        [Output("contact ID")]
        public OutArgument<string> ContactID { get; set; }

        [Output("fullname")]
        public OutArgument<string> FullnameOutput { get; set; }

        [Output("omkstedID")]
        public OutArgument<string> OmkStedId { get; set; }

        [Output("domæne")]
        public OutArgument<string> Domaene { get; set; }

        [Output("email domæne")]
        public OutArgument<string> EmailDomaene { get; set; }

        [Output("lokation")]
        public OutArgument<string> Lokation { get; set; }

        [Output("stillingsbetegnelse")]
        public OutArgument<string> StillingsBetegnelse { get; set; }

        [Output("birthdate output")]
        public OutArgument<DateTime> BirthDateOutput { get; set; }

        [Output("username output")]
        public OutArgument<string> UsernameOutput { get; set; }



        public QueryExpression query { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                // Get the context service.
                IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

                // Use the context service to create an instance of IOrganizationService.             
                IOrganizationService service = serviceFactory.CreateOrganizationService(Icontext.InitiatingUserId);

                //get input fields
                var birthDate = Birthdate.Get(context) == new DateTime(0001, 01, 01) ? new DateTime(1753, 01, 01) : Birthdate.Get(context);
                var firstName = Firstname.Get(context);
                var lastName = Lastname.Get(context);
                var userName = Username.Get(context) == null ? "" : Username.Get(context);

                // made in fetchXML builder 
                // Instantiate QueryExpression QEcontact
                var QEcontact = new QueryExpression("contact");

                // Add columns to QEcontact.ColumnSet
                QEcontact.ColumnSet.AddColumns("fullname", "firstname", "lastname", "birthdate", "sdu_brugernavn", "sdu_domne", "emailaddress1", "address1_city", "jobtitle", "sdu_crmomkostningssted", "sdu_brugernavn", "address1_line1");

                // Define filter QEcontact.Criteria // all must match birthdate + sdu_crmudlb + opdateret fra fim
                var QEcontact_Criteria = new FilterExpression();
                QEcontact.Criteria.AddFilter(QEcontact_Criteria);

                // Define filter QEcontact_Criteria // either match on username, or firstname + lastname
                QEcontact_Criteria.FilterOperator = LogicalOperator.Or;
                QEcontact_Criteria.AddCondition("sdu_brugernavn", ConditionOperator.Equal, userName);

                var QEcontact_Criteria_name = new FilterExpression();
                QEcontact_Criteria.AddFilter(QEcontact_Criteria_name);

                // Define filter QEcontact_Criteria_name_birthdate
                QEcontact_Criteria_name.AddCondition("firstname", ConditionOperator.Like, "%" + firstName + "%");
                QEcontact_Criteria_name.AddCondition("lastname", ConditionOperator.Like, "%" + lastName + "%");
                QEcontact_Criteria_name.AddCondition("birthdate", ConditionOperator.On, birthDate);

                var QEcontact_Criteria_state = new FilterExpression();
                QEcontact.Criteria.AddFilter(QEcontact_Criteria_state);

                // define filter QEcontact_Criteria_dates
                QEcontact_Criteria_state.AddCondition("statecode", ConditionOperator.Equal, RecordStatus.Get(context) == false ? 1 : 0); // kontoen er inaktiv
                QEcontact_Criteria_state.AddCondition("sdu_crmudlb", ConditionOperator.OnOrAfter, DateTime.Today.AddMonths(ExpirationDateLimit.Get(context) * -1)); // CRM udløb er maksimalt 13 måneder gammelt
                QEcontact_Criteria_state.AddCondition("sdu_crmudlb", ConditionOperator.OnOrBefore, DateTime.Today.AddDays(-1)); // CRM er igår eller før
                QEcontact_Criteria_state.AddCondition("jobtitle", ConditionOperator.NotLike, "%Studentermedhjælp%");

                //find records
                var queryResult = service.RetrieveMultiple(QEcontact);

                if (queryResult.Entities.Count == 1)
                {

                    foreach (var record in queryResult.Entities)
                    {
                        // fullname
                        FullnameOutput.Set(context, record.GetAttributeValue<string>("fullname"));

                        // contactid
                        ContactID.Set(context, record.GetAttributeValue<Guid>("contactid").ToString());

                        // omkostningssted
                        var omkStedIdLocal = SearchForRecord(service, "sdu_brugeradmomksted",
                             new KeyValuePair<string, string>("sdu_omksted", record.GetAttributeValue<EntityReference>("sdu_crmomkostningssted").Id.ToString()),
                             new KeyValuePair<string, string>("", ""),
                             "sdu_brugeradmomkstedid");

                        OmkStedId.Set(context, omkStedIdLocal);

                        // domæne 
                        String[] seperator_domaene = { "_" };

                        if (omkStedIdLocal != "")
                        {
                            Domaene.Set(context, SearchForRecord(service,
                                  "sdu_domner",
                                  new KeyValuePair<string, string>("sdu_brugeradmomksted", omkStedIdLocal.Split(seperator_domaene, StringSplitOptions.RemoveEmptyEntries).GetValue(0).ToString()),
                                  new KeyValuePair<string, string>("sdu_name", record.GetAttributeValue<string>("sdu_domne")),
                                  "sdu_domnerid"));

                            // email domæne
                            String[] seperator_emailDomaene = { "@" };
                            var emailDomainFromContact = "@" + record.GetAttributeValue<string>("emailaddress1").Split(seperator_emailDomaene, StringSplitOptions.RemoveEmptyEntries).GetValue(1).ToString();

                            EmailDomaene.Set(context, SearchForRecord(service,
                                "sdu_emaildomne",
                                new KeyValuePair<string, string>("sdu_brugeradmomksted", omkStedIdLocal.Split(seperator_domaene, StringSplitOptions.RemoveEmptyEntries).GetValue(0).ToString()),
                                new KeyValuePair<string, string>("sdu_name", emailDomainFromContact.Replace(" ", "")), // remove whitespaces
                                "sdu_emaildomneid"));
                        }
                        else
                        {
                            // set output parameters to empty strings, if no omk sted
                            Domaene.Set(context, "");
                            EmailDomaene.Set(context, "");
                        }

                        // lokation + arbejdsadresse
                        var LokationOptionSetValue = "";

                        switch (record.GetAttributeValue<string>("address1_city"))
                        {
                            case "Odense":
                                LokationOptionSetValue = "100000000";
                                break;
                            case "Odense M":
                                LokationOptionSetValue = "100000000";
                                break;
                            case "Odense C":
                                LokationOptionSetValue = "100000000";
                                break;
                            case "Sønderborg":
                                LokationOptionSetValue = "100000001";
                                break;
                            case "Esbjerg":
                                LokationOptionSetValue = "100000002";
                                break;
                            case "Slagelse":
                                LokationOptionSetValue = "100000003";
                                break;
                            case "Kolding":
                                LokationOptionSetValue = "100000004";
                                break;
                            default:
                                break;
                        }

                        var workAddress = record.GetAttributeValue<string>("address1_line1") == null ? "" : record.GetAttributeValue<string>("address1_line1");

                        if (workAddress.Contains("Campusvej"))
                        {
                            LokationOptionSetValue = LokationOptionSetValue + "_" + "100000000";
                        }
                        else if (workAddress.Contains("J.B. Winsløws Vej"))
                        {
                            LokationOptionSetValue = LokationOptionSetValue + "_" + "100000001";
                        }

                        Lokation.Set(context, LokationOptionSetValue);


                        // stillingsbetegnelse
                        StillingsBetegnelse.Set(context, SearchForRecord(service,
                         "sdu_stikogruppe",
                         new KeyValuePair<string, string>("sdu_name", record.GetAttributeValue<string>("jobtitle") == null ? "" : record.GetAttributeValue<string>("jobtitle")),
                         new KeyValuePair<string, string>("", ""),
                         "sdu_stikogruppeid"));


                        // fødselsdato
                        BirthDateOutput.Set(context, record.GetAttributeValue<DateTime>("birthdate") < new DateTime(1900, 01, 01) ? DateTime.Now : record.GetAttributeValue<DateTime>("birthdate"));

                        // brugernavn
                        UsernameOutput.Set(context, record.GetAttributeValue<string>("sdu_brugernavn"));

                    }
                }
                else
                {
                    FullnameOutput.Set(context, "");
                    ContactID.Set(context, "");
                    OmkStedId.Set(context, "");
                    Domaene.Set(context, "");
                    EmailDomaene.Set(context, "");
                    Lokation.Set(context, "");
                    StillingsBetegnelse.Set(context, "");
                    BirthDateOutput.Set(context, DateTime.Now);
                    UsernameOutput.Set(context, "");
                }

            }
            catch (Exception ex)
            {
                throw (ex);
            }


        }

        public string SearchForRecord(IOrganizationService service, string entity, KeyValuePair<string, string> firstKVP, KeyValuePair<string, string> secondKVP, string fieldToExtract)
        {
            try
            {
                var brugerAdmOmkId = "";


                var query = new QueryExpression(entity);
                query.ColumnSet = new ColumnSet(true);

                if (firstKVP.Key == null || firstKVP.Key != "")
                {
                    ConditionExpression firstCondition = new ConditionExpression
                    {
                        AttributeName = firstKVP.Key,
                        Operator = ConditionOperator.Equal,
                        Values = { firstKVP.Value }
                    };

                    ConditionExpression secondCondition = new ConditionExpression
                    {
                        AttributeName = secondKVP.Key,
                        Operator = ConditionOperator.Equal,
                        Values = { secondKVP.Value }
                    };

                    FilterExpression filter = new FilterExpression();
                    filter.AddCondition(firstCondition);

                    if (secondKVP.Key != "")
                    {
                        filter.AddCondition(secondCondition);
                    }

                    query.Criteria.AddFilter(filter);

                    DataCollection<Entity> queryResult = service.RetrieveMultiple(query).Entities;

                    foreach (var result in queryResult)
                    {
                        brugerAdmOmkId = result.GetAttributeValue<Guid>(fieldToExtract).ToString() + "_" + result.GetAttributeValue<string>("sdu_name").ToString();
                    }

                    return brugerAdmOmkId;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception ex)
            {

                throw (ex);
            }

        }
    }
}
