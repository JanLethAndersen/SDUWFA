using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Brugeradministration
{
    public class CreateEmailTemplate : CodeActivity
    {
        [ArgumentRequired]
        [Input("Record Dynamic URL")]
        public InArgument<string> RecordDynamicUrl { get; set; }

        [ArgumentRequired]
        [Input("Rights Area")]
        public InArgument<string> RightsArea { get; set; }

        [Input("Parameter reference")]
        [ReferenceTarget("sdu_parameter")]
        public InArgument<EntityReference> ParameterReference { get; set; }

        [Input("Does the user have existing  rights?")]
        public InArgument<bool> ExistingRights { get; set; }

        [Input("Should the rights be removed completely?")]
        public InArgument<bool> RemoveRightsCompletely { get; set; }

        [Output("Email Body - Rettigheder")]
        public OutArgument<string> EmailRettigheder { get; set; }

        [Output("Email subject")]
        public OutArgument<string> EmailSubject { get; set; }

        public Guid id { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Get the context service.
            IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.             
            IOrganizationService service = serviceFactory.CreateOrganizationService(new Guid("cfd6fed3-7115-e911-8129-0050568a45d1")); // user = SDU\SQLCRMDynWorkF

            var Url = RecordDynamicUrl.Get(context);

            string[] parameters = Url.TrimStart('?').Split('&');

            foreach (string param in parameters)
            {
                var keyValue = param.Split('=');

                // 0 fordi jeg kigge på label og ikke værdi
                switch (keyValue[0])
                {
                    case "id":
                        id = new Guid(keyValue[1]);
                        break;
                    default:
                        break;
                }
            }

            // get all fields in the entity
            var entity = service.Retrieve(Icontext.PrimaryEntityName, id, new ColumnSet(true));

            // get name of current record
            var nameOfRecord = entity.GetAttributeValue<string>("sdu_name");

            // get person - dvs. contact (som det drejer sig om)
            var personReference = entity.GetAttributeValue<EntityReference>("sdu_person");
            var person = new Person(service, personReference);

            // get user
            var user = new User(service, Icontext.InitiatingUserId);

            // get brugeradministration
            var brugeradministrationReference = entity.GetAttributeValue<EntityReference>("sdu_brugeradministration");
            var brugeradministrationEntity = service.Retrieve(brugeradministrationReference.LogicalName, brugeradministrationReference.Id, new ColumnSet(true));
            var brugeradministration = new Brugeradministration(service, brugeradministrationReference);

            // which area?
            var AreaFromInput = RightsArea.Get(context).ToLower();
            var areaFromInput = AreaFromInput.ToUpper();

            // is it new or is it changes to exsisting?
            bool existingRights = ExistingRights.Get(context);
            bool removeRightsCompletely = RemoveRightsCompletely.Get(context);

            // instantiate class
            var brugeradministrationRettighedsFlow = new BrugeradministrationRettighedsFlow(user, person, service, nameOfRecord, existingRights, removeRightsCompletely);

            // Parameters
            var Parameters = ParameterReference.Get(context);
            var ParametersRecord = service.Retrieve(Parameters.LogicalName, Parameters.Id, new ColumnSet(true));

            var oessKey = ParametersRecord.GetAttributeValue<String>("sdu_sskey");
            var hrKey = ParametersRecord.GetAttributeValue<String>("sdu_hrkey");
            var qlikviewKey = ParametersRecord.GetAttributeValue<String>("sdu_qlikviewkey");
            var dansKey = ParametersRecord.GetAttributeValue<String>("sdu_danskey");
            var stadsKey = ParametersRecord.GetAttributeValue<String>("sdu_stadskey");
            var acadreKey = ParametersRecord.GetAttributeValue<String>("sdu_acadrekey");
            var kmeibrugKey = ParametersRecord.GetAttributeValue<String>("sdu_kemibrugkey");
            var emailAfsender = ParametersRecord.GetAttributeValue<EntityReference>("sdu_baemailafsender");


            // CRM Mailbox
            var crmMailbox = "mscrm@sdu.dk";


            switch (AreaFromInput)
            {
                case "øss":
                    var oessEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    // ØSS
                    var oessEmailRecipient = "oesshrbv@sdu.dk";
                    object[] kreditor = { "Kreditor forespørgsel", entity.GetAttributeValue<bool>("sdu_sskreditorforesprgsel"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_kreditorforesprgsel") };
                    object[] debitor = { "Debitor forespørgsel", entity.GetAttributeValue<bool>("sdu_ssdebitorforesprgsel"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_debitorforesprgsel") };
                    object[] eforms = { "Eforms Modtager", entity.GetAttributeValue<bool>("sdu_sseformsmodtager"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_eformsmodtager") };
                    object[] rekvirent = { "EF Rekvirent", entity.GetAttributeValue<bool>("sdu_ssefrekvirent"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_ssef") };
                    //object[] omkontering = { "Omkontering", entity.GetAttributeValue<bool>("sdu_ssomkontering") };
                    object[] internHandel = { "Intern handel og ompostering", entity.GetAttributeValue<bool>("sdu_ssinternhandelogompostering"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_ssinternhandelogompostering") };
                    object[] bevillingsIndtastning = { "Bevillingsindtastning", entity.GetAttributeValue<bool>("sdu_ssbevillingsindtastning"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_ssbevillingsindtastning") };
                    //object[] indkob = { "Indkøb", entity.GetAttributeValue<bool>("sdu_ssindkb"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_indkb") };
                    //object[] standardIndkober = { "Standardgodkender", entity.GetAttributeValue<EntityReference>("sdu_ssstandardgodkenderindkb_contact")?.Name, brugeradministrationEntity.GetAttributeValue<EntityReference>("sdu_ssstandardgodkenderindkb_contact") == null ? "" : brugeradministrationEntity.GetAttributeValue<EntityReference>("sdu_ssstandardgodkenderindkb_contact").Name };
                    object[] oessNoter = { "Noter", entity.GetAttributeValue<string>("sdu_ssnoter"), "" };
                    //object[] omkSted = { "<br>Omkostningssted", person.OmkSted, "" };

                    object[] forventetStartDatoOess = { "<br>Forventet startdato", brugeradministration.forventetStart != null ? brugeradministration.forventetStart.ToString("d", new CultureInfo("da-DK")) : null, new DateTime() };


                    object[] oess = { kreditor, debitor, eforms, rekvirent, internHandel, bevillingsIndtastning, oessNoter, forventetStartDatoOess/*omkSted*/ };

                    DataCollection<Entity> oessOmkontering = null;

                    if (!existingRights)
                    {
                        oessOmkontering = QueryForRelatedRecords(true, id, service, "sdu_oessadgang");
                        oessEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, oessOmkontering, oessKey, oess);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "ss");
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        oessOmkontering = QueryForRelatedRecords(false, id, service, "sdu_oessadgang");
                        oessEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, oessOmkontering, oessKey, oess);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "ss");
                    }
                    else if (removeRightsCompletely)
                    {
                        oessEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, oessKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "ss");
                    }

                    // set the generated email
                    string oessEmailBodyString = oessEmailBodyAndSubject.Item1.ToString();

                    EmailRettigheder.Set(context, oessEmailBodyString);
                    EmailSubject.Set(context, oessEmailBodyAndSubject.Item2);


                    break;
                case "hr":
                    // HR 
                    var hrEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    var hrEmailRecipient = "oesshrbv@sdu.dk";

                    //object[] fravaersAdm = { "HR - Fraværsadm.", entity.GetAttributeValue<bool>("sdu_hrfravrsadm") };
                    //object[] foresporgsel = { "HR - Forespørgsel", entity.GetAttributeValue<bool>("sdu_hrforesprgsel") };
                    object[] hrNoter = { "HR - Noter", entity.GetAttributeValue<string>("sdu_hrnoter"), "" };
                    object[] forventetStartDatoHr = { "<br>Forventet startdato", brugeradministration.forventetStart != null ? brugeradministration.forventetStart.ToString("d", new CultureInfo("da-DK")) : null, new DateTime() };


                    // object[] omkStedHr = { "<br>Omkostningssted", person.OmkSted, "" };


                    object[] hr = { hrNoter, forventetStartDatoHr/*, omkStedHr*/ };

                    DataCollection<Entity> hrOmkontering = null;


                    if (!existingRights)
                    {
                        hrOmkontering = QueryForRelatedRecords(true, id, service, "sdu_hradgang");
                        hrEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, hrOmkontering, hrKey, hr);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "hr");

                        // Send Email
                        SendEmail(service, hrOmkontering, emailAfsender.Id, entity.Id, hrEmailRecipient, hrEmailBodyAndSubject.Item1.ToString(), hrEmailBodyAndSubject.Item2);

                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        hrOmkontering = QueryForRelatedRecords(false, id, service, "sdu_hradgang");
                        hrEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, hrOmkontering, hrKey, hr);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "hr");

                        // Send Email
                        SendEmail(service, hrOmkontering, emailAfsender.Id, entity.Id, hrEmailRecipient, hrEmailBodyAndSubject.Item1.ToString(), hrEmailBodyAndSubject.Item2);
                    }
                    else if (removeRightsCompletely)
                    {
                        hrEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, hrKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "hr");
                    }

                    string hrEmailBodyString = hrEmailBodyAndSubject.Item1.ToString();

                    EmailRettigheder.Set(context, hrEmailBodyString);
                    EmailSubject.Set(context, hrEmailBodyAndSubject.Item2);


                    break;
                case "qlikview":
                    //qlikview
                    var qvEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    var qvEmailRecipient = "oesshrbv@sdu.dk";

                    //object[] finansForesporgsel = { "Finans forespørgsel", entity.GetAttributeValue<bool>("sdu_qlikviewfinansforesprgsel") };
                    //object[] lonForeporgsel = { "Løn forespørgsel", entity.GetAttributeValue<bool>("sdu_qlikviewlnforesprgsel") };
                    //object[] lonPrognose = { "Lønprognose", entity.GetAttributeValue<bool>("sdu_qlikviewlnprognose") };
                    object[] qlikviewNoter = { "Noter", entity.GetAttributeValue<string>("sdu_qlikviewnoter"), "" };

                    object[] forventetStartDatoQv = { "<br>Forventet startdato", brugeradministration.forventetStart != null ? brugeradministration.forventetStart.ToString("d", new CultureInfo("da-DK")) : null, new DateTime() };

                    //object[] omkStedQv = { "<br>Omkostningssted", person.OmkSted, "" };


                    object[] qlikView = { qlikviewNoter, forventetStartDatoQv/*omkStedQv*/ };

                    DataCollection<Entity> qlikViewOmkontering = null;


                    if (!existingRights)
                    {
                        qlikViewOmkontering = QueryForRelatedRecords(true, id, service, "sdu_qlikviewadgang");
                        qvEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, qlikViewOmkontering, qlikviewKey, qlikView);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "qlikview");

                        // SendEmail
                        SendEmail(service, qlikViewOmkontering, emailAfsender.Id, entity.Id, qvEmailRecipient, qvEmailBodyAndSubject.Item1.ToString(), qvEmailBodyAndSubject.Item2);
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        qlikViewOmkontering = QueryForRelatedRecords(false, id, service, "sdu_qlikviewadgang");
                        qvEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, qlikViewOmkontering, qlikviewKey, qlikView);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "qlikview");

                        // Send Email
                        SendEmail(service, qlikViewOmkontering, emailAfsender.Id, entity.Id, qvEmailRecipient, qvEmailBodyAndSubject.Item1.ToString(), qvEmailBodyAndSubject.Item2);

                    }
                    else if (removeRightsCompletely)
                    {
                        qvEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, qlikviewKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "qlikview");
                    }

                    string qvEmailBodyString = qvEmailBodyAndSubject.Item1.ToString();

                    EmailRettigheder.Set(context, qvEmailBodyString);
                    EmailSubject.Set(context, qvEmailBodyAndSubject.Item2);

                    break;
                case "stads":
                    //stads
                    var stadsEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    var stadsEmailRecipient = "stads@sdu.dk";

                    object[] delestaaModul = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_delestmodul_stads", service).DisplayName.UserLocalizedLabel.Label,
                                               entity.GetAttributeValue<bool>("sdu_delestmodul_stads"),
                                               brugeradministrationEntity.GetAttributeValue<bool>("sdu_delestmodul_stads")
                                                };



                    object[] eksamensAdministration = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_eksamensadministration_stads", service).DisplayName.UserLocalizedLabel.Label,
                                                        entity.GetAttributeValue<bool>("sdu_eksamensadministration_stads"),
                                                        brugeradministrationEntity.GetAttributeValue<bool>("sdu_eksamensadministration_stads")};

                    object[] opslagsMenu = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_opslagsmenu_stads", service).DisplayName.UserLocalizedLabel.Label,
                                             entity.GetAttributeValue<bool>("sdu_opslagsmenu_stads"),
                                             brugeradministrationEntity.GetAttributeValue<bool>("sdu_opslagsmenu_stads")   };

                    object[] optagelsesMenu = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_optagelsesmenu_stads", service).DisplayName.UserLocalizedLabel.Label,
                                                entity.GetAttributeValue<bool>("sdu_optagelsesmenu_stads"),
                                                brugeradministrationEntity.GetAttributeValue<bool>("sdu_optagelsesmenu_stads")    };

                    object[] sekretrMenu = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_sekretrmenu_stads", service).DisplayName.UserLocalizedLabel.Label,
                                             entity.GetAttributeValue<bool>("sdu_sekretrmenu_stads"),
                                             brugeradministrationEntity.GetAttributeValue<bool>("sdu_sekretrmenu_stads")};

                    object[] su = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_statensuddannelsesstttesu_stads", service).DisplayName.UserLocalizedLabel.Label,
                                    entity.GetAttributeValue<bool>("sdu_statensuddannelsesstttesu_stads"),
                                    brugeradministrationEntity.GetAttributeValue<bool>("sdu_statensuddannelsesstttesu_stads")};
            

                    object[] stadsNoter = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_stadsnoter", service).DisplayName.UserLocalizedLabel.Label,
                                            entity.GetAttributeValue<string>("sdu_stadsnoter"),
                                            ""};

                    object[] stads = { delestaaModul, eksamensAdministration, opslagsMenu, optagelsesMenu, sekretrMenu, su, stadsNoter };


                    if (!existingRights)
                    {
                        stadsEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, null, stadsKey, stads);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "stads");
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        stadsEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, null, stadsKey, stads);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "stads");
                    }
                    else if (removeRightsCompletely)
                    {
                        stadsEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, stadsKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "stads");
                    }


                    string stadsEmailBodyString = stadsEmailBodyAndSubject.Item1.ToString();

                    EmailRettigheder.Set(context, stadsEmailBodyString);
                    EmailSubject.Set(context, stadsEmailBodyAndSubject.Item2);

                    break;
                case "dans":
                    // dans
                    var dansEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    var dansEmailRecipient = "dans@sdu.dk";

                    object[] optagelseSagsbehandler = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_optagelsesagsbehandler_dans", service).DisplayName.UserLocalizedLabel.Label,
                                                        entity.GetAttributeValue<bool>("sdu_optagelsesagsbehandler_dans"),
                                                        brugeradministrationEntity.GetAttributeValue<bool>("sdu_optagelsesagsbehandler_dans")};

                    object[] optagelseSuperbruger = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_optagelsesuperbruger_dans", service).DisplayName.UserLocalizedLabel.Label,
                                                      entity.GetAttributeValue<bool>("sdu_optagelsesuperbruger_dans"),
                                                      brugeradministrationEntity.GetAttributeValue<bool>("sdu_optagelsesuperbruger_dans")};

                    object[] kiggeAdgang = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_kiggeadgang_dans", service).DisplayName.UserLocalizedLabel.Label,
                                             entity.GetAttributeValue<bool>("sdu_kiggeadgang_dans"),
                                             brugeradministrationEntity.GetAttributeValue<bool>("sdu_kiggeadgang_dans")};

                    object[] dansNoter = { "Noter", entity.GetAttributeValue<string>("sdu_dansnoter"), "" };

                    object[] dans = { optagelseSagsbehandler, optagelseSuperbruger, kiggeAdgang, dansNoter };


                    if (!existingRights)
                    {
                        dansEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, null, dansKey, dans);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "dans");
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        dansEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, null, dansKey, dans);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "dans");
                    }
                    else if (removeRightsCompletely)
                    {
                        dansEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, dansKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "dans");
                    }

                    string dansEmailBodyString = dansEmailBodyAndSubject.Item1.ToString();

                    EmailRettigheder.Set(context, dansEmailBodyString);
                    EmailSubject.Set(context, dansEmailBodyAndSubject.Item2);
                    break;
                case "acadre":
                    var acadreEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    var acadreEmailRecipient = "acadre@sdu.dk";

                    object[] primaryAcadreAfdeling = { "Primær Acadre Afdeling", entity.GetAttributeValue<EntityReference>("sdu_acadreafdeling")?.Name, brugeradministrationEntity.GetAttributeValue<EntityReference>("sdu_acadreafdeling")?.Name };
                    object[] studenterSager = { "Studentersager", entity.GetAttributeValue<bool>("sdu_acadrestudentersager"), brugeradministrationEntity.GetAttributeValue<bool>("sdu_acadrestudentersager") };
                    object[] typeAfAdgang = { RetrieveAttributeMetadata(Icontext.PrimaryEntityName, "sdu_acadretypeafadgang", service).DisplayName.UserLocalizedLabel.Label,
                                              GetOptionsetText(service, "sdu_acadretypeafadgang", entity.GetAttributeValue<OptionSetValue>("sdu_acadretypeafadgang") != null ? entity.GetAttributeValue<OptionSetValue>("sdu_acadretypeafadgang").Value : 0),
                                               GetOptionsetText(service, "sdu_acadretypeafadgang", brugeradministrationEntity.GetAttributeValue<OptionSetValue>("sdu_acadretypeafadgang") != null ? brugeradministrationEntity.GetAttributeValue<OptionSetValue>("sdu_acadretypeafadgang").Value : 0)};

                    object[] forventetStartDatoAcadre = { "<br>Forventet startdato", brugeradministration.forventetStart != null ? brugeradministration.forventetStart.ToString("d", new CultureInfo("da-DK")) : null, new DateTime() };

                    // added 03-12-2019
                    object[] jobTitle = { "<br>Stillingsbetegnelse", person.JobTitle, "" };


                    object[] acadreNoter = { "Noter", entity.GetAttributeValue<string>("sdu_acadrenoter"), "" };

                    object[] acadre = { primaryAcadreAfdeling, studenterSager, typeAfAdgang, acadreNoter, forventetStartDatoAcadre, jobTitle };

                    // query for multiple records in relation
                    /* var acadreAfdelingerQuery = new QueryExpression
                     {
                         EntityName = "sdu_acadreafdelingerspecifikkeadgange",
                         ColumnSet = new ColumnSet(true),
                         Criteria = new FilterExpression
                         {
                             FilterOperator = LogicalOperator.And,
                             Conditions =
                             {
                                 new ConditionExpression {AttributeName = "sdu_brugeradministrationsndring",
                                                          Operator = ConditionOperator.Equal,
                                                          Values = { id } }
                             }
                         }
                     };*/

                    DataCollection<Entity> acadreAfdelinger = null;

                    //acadreEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(acadre, areaFromInput, acadreEmailRecipient,acadreAfdelinger);


                    if (!existingRights)
                    {
                        acadreAfdelinger = QueryForRelatedRecords(true, id, service, "sdu_acadreafdelingerspecifikkeadgange");
                        acadreEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, acadreAfdelinger, acadreKey, acadre);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "acadre");
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        acadreAfdelinger = QueryForRelatedRecords(false, id, service, "sdu_acadreafdelingerspecifikkeadgange");
                        acadreEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, acadreAfdelinger, acadreKey, acadre);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "acadre");
                    }
                    else if (removeRightsCompletely)
                    {
                        acadreEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, acadreKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "acadre");
                    }

                    EmailRettigheder.Set(context, acadreEmailBodyAndSubject.Item1.ToString());
                    EmailSubject.Set(context, acadreEmailBodyAndSubject.Item2);


                    break;


                case "kemibrug":
                    var kemibrugEmailBodyAndSubject = new Tuple<StringBuilder, String>(new StringBuilder(), "");

                    object[] kemibrug = { };

                    DataCollection<Entity> kemibrugRettigheder = null;

                    if (!existingRights)
                    {
                        kemibrugRettigheder = QueryForRelatedRecords(true, id, service, "sdu_kemibrugadgange");
                        kemibrugEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_NoExistingRights(areaFromInput, crmMailbox, kemibrugRettigheder, kmeibrugKey);
                        brugeradministrationRettighedsFlow.updateRecord(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, "kemibrug");
                    }
                    else if (existingRights && !removeRightsCompletely)
                    {
                        kemibrugRettigheder = QueryForRelatedRecords(false, id, service, "sdu_kemibrugadgange");
                        kemibrugEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_ChangesInRights(areaFromInput, crmMailbox, kemibrugRettigheder, kmeibrugKey);
                    }
                    else if (removeRightsCompletely)
                    {
                        kemibrugEmailBodyAndSubject = brugeradministrationRettighedsFlow.buildEmailBody_RemoveAllRights(areaFromInput, crmMailbox, kmeibrugKey);
                    }

                    EmailRettigheder.Set(context, kemibrugEmailBodyAndSubject.Item1.ToString());
                    EmailSubject.Set(context, kemibrugEmailBodyAndSubject.Item2);


                    break;

                default:
                    break;
            }

        }

        class BrugeradministrationRettighedsFlow
        {
            public User user { get; set; }
            public Person person { get; set; }
            public IOrganizationService service { get; set; }
            public string nameOfRecord { get; set; }
            public bool existingRights { get; set; }
            public bool removeRightsCompletely { get; set; }


            public BrugeradministrationRettighedsFlow(User user, Person person, IOrganizationService service, string nameOfRecord, bool existingRights, bool removeRightsCompletely)
            {
                this.user = user;
                this.person = person;
                this.service = service;
                this.nameOfRecord = nameOfRecord;
                this.existingRights = existingRights;
                this.removeRightsCompletely = removeRightsCompletely;

            }

            public string acceptRightsKey = "703fa7b2-dced-4cda-9172-611927168bf5";
            public string declineRightsKey = "cfead9a7-3e56-4c65-8dbf-8a1cb1056992";
            public string changesRightsKey = "cb85ba1c-f33c-4a19-9657-189496890ea3";

            public Tuple<StringBuilder, String> buildEmailBody_NoExistingRights(string rightsAreaName, string emailRecipient, DataCollection<Entity> relatedRecords = null, string rightsKey = "", Object[] fieldsToAppend = null)
            {
                var emailSubject = "Ansøgning til " + rightsAreaName + ": " + nameOfRecord;


                var stringBuilder = new StringBuilder();

                stringBuilder.Append("<b>Til lederen: Du skal forholde dig passivt til nedenstående, da tildelingen sker fra dem der teknisk giver adgang til systemet. Din godkendelse er passiv. Hvis du finder nedenstående uhensigtsmæssigt, kontakt da den bestillende sekretær.</b><hr><br>");
                stringBuilder.AppendFormat("Hej {0}", rightsAreaName);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("Der er blevet søgt om adgang til {0} for {1} ({2})", rightsAreaName, person.Name, person.Email);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("Bestillingen er foretaget af: {0} ({1})", user.Name, user.Email);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.Append("De ønskede rettigheder er:");
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");



                if (fieldsToAppend != null)
                {
                    foreach (var field in fieldsToAppend)
                    {
                        var specificField = (object[])field;

                        if (specificField[1] == null)
                        {
                            continue;
                        }
                        if (specificField[1].GetType() == typeof(bool))
                        {
                            if ((bool)specificField[1] == false)
                            {
                                continue;
                            }
                            else
                            {
                                stringBuilder.Append("<b>" + specificField[0] + "</b>");
                                stringBuilder.Append(": Ja<br>");
                            }
                        }
                        else
                        {
                            stringBuilder.Append("<b>" + specificField[0] + "</b>");
                            stringBuilder.Append(": ");
                            stringBuilder.Append(specificField[1]);
                            stringBuilder.Append("<br>");

                        }


                    }
                }
                ////// related - null?
                if (relatedRecords != null)
                {
                    var RelatedRecords = getRelatedRecords(relatedRecords, rightsAreaName, true);
                    stringBuilder.Append(RelatedRecords);

                }

                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet tildelt - {2} | {3}'><b>Rettighederne er tildelt</b></a>", emailRecipient, emailSubject, acceptRightsKey, rightsKey);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet afvist - {2} | {3}'><b>Rettighederne er blevet afvist</b></a>", emailRecipient, emailSubject, declineRightsKey, rightsKey);
                stringBuilder.Append("<br>");

                return Tuple.Create<StringBuilder, String>(stringBuilder, emailSubject);
            }


            public Tuple<StringBuilder, String> buildEmailBody_ChangesInRights(string rightsAreaName, string emailRecipient, DataCollection<Entity> relatedRecords = null, string rightsKey = "", Object[] fieldsToAppend = null)
            {

                var emailSubject = "Ændringer for " + rightsAreaName + ": " + nameOfRecord;


                var stringBuilder = new StringBuilder();

                stringBuilder.Append("<b>Til lederen: Du skal forholde dig passivt til nedenstående, da tildelingen sker fra dem der teknisk giver adgang til systemet. Din godkendelse er passiv. Hvis du finder nedenstående uhensigtsmæssigt, kontakt da den bestillende sekretær.</b><hr><br>");
                stringBuilder.AppendFormat("Hej {0}", rightsAreaName);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("Der er sket ændringer i adgang til {0} for {1} ({2})", rightsAreaName, person.Name, person.Email);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("Ændringerne er foretaget af: {0} ({1})", user.Name, user.Email);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendLine("Der er foretaget ændringer for følgende rettigheder: <br>");

                if (fieldsToAppend != null)
                {
                    foreach (var field in fieldsToAppend)
                    {

                        var specificField = (object[])field;

                        if (specificField[1] == null)
                        {
                            continue;
                        }

                        if (specificField[1].ToString() != specificField[2].ToString())
                        {
                            if (specificField[1].GetType() == typeof(bool))
                            {
                                if ((bool)specificField[1] == false)
                                {
                                    stringBuilder.Append("<b>" + specificField[0] + "</b>");
                                    stringBuilder.Append(": Nej<br>");
                                }
                                else
                                {
                                    stringBuilder.Append("<b>" + specificField[0] + "</b>");
                                    stringBuilder.Append(": Ja<br>");
                                }
                            }
                            else
                            {
                                stringBuilder.Append("<b>" + specificField[0] + "</b>");
                                stringBuilder.Append(": ");
                                stringBuilder.Append(specificField[1]);
                                stringBuilder.Append("<br>");

                            }
                        }


                    }
                }

                ////// related
                if (relatedRecords != null)
                {
                    var RelatedRecords = getRelatedRecords(relatedRecords, rightsAreaName, false);
                    stringBuilder.Append(RelatedRecords);
                }


                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet tildelt - {2} | {3}'><b>Rettighederne er tildelt</b></a>", emailRecipient, emailSubject, changesRightsKey, rightsKey);
                stringBuilder.Append("<br>");
                stringBuilder.Append("<br>");
                stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet afvist - {2} | {3}'><b>Rettighederne er blevet afvist</b></a>", emailRecipient, emailSubject, declineRightsKey, rightsKey);
                stringBuilder.Append("<br>");

                return Tuple.Create<StringBuilder, String>(stringBuilder, emailSubject);

            }


            public Tuple<StringBuilder, String> buildEmailBody_RemoveAllRights(string rightsAreaName, string emailRecipient, string rightsKey = "")
            {
                // values for emails 

                var emailSubject = "Fjernelse af rettigheder til " + rightsAreaName + ": " + nameOfRecord;


                var stringBuilder = new StringBuilder();

                stringBuilder.Append("<b>Til lederen: Du skal forholde dig passivt til nedenstående, da tildelingen sker fra dem der teknisk giver adgang til systemet. Din godkendelse er passiv. Hvis du finder nedenstående uhensigtsmæssigt, kontakt da den bestillende sekretær.</b><hr><br>");
                stringBuilder.AppendFormat("Hej {0}", rightsAreaName);
                stringBuilder.AppendLine("");
                stringBuilder.AppendLine("");
                stringBuilder.AppendFormat("{0} ({1}) har fået frataget sine rettigheder til {2}.", person.Name, person.Email, rightsAreaName);
                stringBuilder.AppendLine("");
                stringBuilder.AppendLine("");
                stringBuilder.AppendFormat("Ændringerne er foretaget af: {0} ({1})", user.Name, user.Email);
                stringBuilder.AppendLine("");


                /*foreach (var field in fieldsToAppend)
                {
                    var specificField = (object[])field;

                    if (specificField[1] == null)
                    {
                        continue;
                    }
                    if (specificField[1].GetType() == typeof(bool))
                    {
                        if ((bool)specificField[1] == false)
                        {
                            stringBuilder.Append("<b>" + specificField[0] + "</b>");
                            stringBuilder.AppendLine(": Nej");
                        }
                        else
                        {
                            stringBuilder.Append("<b>" + specificField[0] + "</b>");
                            stringBuilder.AppendLine(": Ja");
                        }
                    }
                    else
                    {
                        stringBuilder.Append("<b>" + specificField[0] + "</b>");
                        stringBuilder.Append(": ");
                        stringBuilder.Append(specificField[1]);
                        stringBuilder.AppendLine("");
                    }


            }*/


                stringBuilder.AppendLine("");
                stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet fjernet - {2} | {3}'><b>Rettighederne er fjernet</b></a>", emailRecipient, emailSubject, acceptRightsKey, rightsKey);
                stringBuilder.AppendLine("");

                return Tuple.Create<StringBuilder, String>(stringBuilder, emailSubject);

            }



            public void updateRecord(string entityName, Guid entityId, string area)
            {
                // update the record
                Entity currentRecord = new Entity(entityName);
                // optionset value == godkendelse afsendt
                currentRecord["sdu_" + area + "status"] = new OptionSetValue(100000000);
                currentRecord.Id = entityId;
                service.Update(currentRecord);
            }


            public StringBuilder getRelatedRecords(DataCollection<Entity> relatedRecords, string rightsArea, bool NoExsistingRights)
            {
                var stringBuilder = new StringBuilder();

                // IF RELATED RECORDS! (aka records in a subgrid) 
                // hardcoded atm for acadre - not optimal
                if (relatedRecords.Count > 0)
                {
                    // generic generation
                    //stringBuilder.Append("<br>");

                    foreach (var record in relatedRecords)
                    {
                        // MOVED 25-02-2020 
                        /*stringBuilder.Append("<br><b>");

                        if (rightsArea.ToLower() == "acadre")
                        {
                            stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_acadreafdeling") != null ? record.GetAttributeValue<EntityReference>("sdu_acadreafdeling").Name : null);
                        }
                        else
                        {
                            stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_omkostningssted") != null ? record.GetAttributeValue<EntityReference>("sdu_omkostningssted").Name : null);
                        }
                        stringBuilder.Append("<br></b>");*/

                        // ---------------------------------------------- ?!?!

                        var masterRecordReference = record.GetAttributeValue<EntityReference>("sdu_masterrecord") != null ? record.GetAttributeValue<EntityReference>("sdu_masterrecord") : null;

                        var masterRecord = new Entity();
                        if (masterRecordReference != null)
                        {
                            masterRecord = service.Retrieve(masterRecordReference.LogicalName, masterRecordReference.Id, new ColumnSet(true));
                        }

                        bool OmkPlacedInString = false;

                        foreach (var attribute in record.FormattedValues)
                        {
                            var attriubteMetaData = RetrieveAttributeMetadata(record.LogicalName, attribute.Key, service);

                            if (attribute.Key.StartsWith("sdu_") && attriubteMetaData.AttributeType.ToString() != "Lookup" && !attribute.Value.ToLower().Contains("bestilling"))
                            {
                                // IF NO RIGHTS
                                if (NoExsistingRights && attribute.Value == "Yes")
                                {

                                    // PLACE OMK/ACADRE AFDELING BEFORE THE RIGHTS - BUT ONLY ONCE!
                                    if (OmkPlacedInString == false)
                                    {
                                        stringBuilder.Append("<br><b>");
                                        if (rightsArea.ToLower() == "acadre")
                                        {
                                            OmkPlacedInString = true;
                                            stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_acadreafdeling") != null ? record.GetAttributeValue<EntityReference>("sdu_acadreafdeling").Name : null);
                                        }
                                        else
                                        {
                                            OmkPlacedInString = true;
                                            stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_omkostningssted") != null ? record.GetAttributeValue<EntityReference>("sdu_omkostningssted").Name : null);
                                        }
                                        stringBuilder.Append("<br></b>");

                                    }
                                       // place the value
                                    stringBuilder.Append("- " + attriubteMetaData.DisplayName.UserLocalizedLabel.Label + ": " + attribute.Value + "<br>");
                                }
                                // IF EXSISTING RIGHTS
                                else if (!NoExsistingRights)
                                {
                                    if (masterRecord.Id != new Guid())
                                    {
                                        if (CompareWithMasterRecord(masterRecord, attribute))
                                        {

                                            // PLACE OMK/ACADRE AFDELING BEFORE THE RIGHTS - BUT ONLY ONCE!
                                            if (OmkPlacedInString == false)
                                            {
                                                stringBuilder.Append("<br><b>");
                                                if (rightsArea.ToLower() == "acadre")
                                                {
                                                    OmkPlacedInString = true;
                                                    stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_acadreafdeling") != null ? record.GetAttributeValue<EntityReference>("sdu_acadreafdeling").Name : null);
                                                }
                                                else
                                                {
                                                    OmkPlacedInString = true;
                                                    stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_omkostningssted") != null ? record.GetAttributeValue<EntityReference>("sdu_omkostningssted").Name : null);
                                                }
                                                stringBuilder.Append("<br></b>");
                                                    
                                            }

                                            stringBuilder.Append("- " + attriubteMetaData.DisplayName.UserLocalizedLabel.Label + ": " + attribute.Value + "<br>");
                                        }
                                    }
                                    else if (attribute.Value == "Yes")
                                    {
                                        // PLACE OMK/ACADRE AFDELING BEFORE THE RIGHTS - BUT ONLY ONCE!
                                        if (OmkPlacedInString == false)
                                        {
                                            stringBuilder.Append("<br><b>");
                                            if (rightsArea.ToLower() == "acadre")
                                            {
                                                OmkPlacedInString = true;
                                                stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_acadreafdeling") != null ? record.GetAttributeValue<EntityReference>("sdu_acadreafdeling").Name : null);
                                            }
                                            else
                                            {
                                                OmkPlacedInString = true;
                                                stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_omkostningssted") != null ? record.GetAttributeValue<EntityReference>("sdu_omkostningssted").Name : null);
                                            }
                                            stringBuilder.Append("<br></b>");
                                        }


                                        stringBuilder.Append("- " + attriubteMetaData.DisplayName.UserLocalizedLabel.Label + ": " + attribute.Value + "<br>");
                                    }

                                }
                            }
                        }
                    }



                    /*

                    if (rightsArea.ToLower() == "acadre")
                    {
                        stringBuilder.AppendLine("");
                        stringBuilder.AppendLine("Der ønskes rettigheder til disse Acadre afdelinger: ");
                        stringBuilder.AppendLine("");

                        foreach (var record in relatedRecords)
                        {
                            stringBuilder.AppendLine("<b>");
                            stringBuilder.AppendLine(record.GetAttributeValue<EntityReference>("sdu_acadreafdeling") != null ? record.GetAttributeValue<EntityReference>("sdu_acadreafdeling").Name : null);
                            stringBuilder.AppendLine("</b>");


                            foreach (var attribute in record.FormattedValues)
                            {
                                if (attribute.Key.StartsWith("sdu"))
                                {
                                    if (attribute.Value.ToLower() == "yes")
                                    {
                                        stringBuilder.AppendLine("- " + RetrieveAttributeMetadata("sdu_acadreafdelingerspecifikkeadgange", attribute.Key, service) + ": Ja");
                                    }
                                }

                            }

                            stringBuilder.AppendLine(record.GetAttributeValue<string>("sdu_noter") != null ? "- Noter: " + record.GetAttributeValue<string>("sdu_noter") : "");

                        }
                    }
                    else if (rightsArea.ToLower() == "kemibrug") // ...... something something
                    {
                        foreach (var record in relatedRecords)
                        {
                            var newRight = record.GetAttributeValue<EntityReference>("sdu_kemibrug") != null ? record.GetAttributeValue<EntityReference>("sdu_kemibrug").Name : null;
                            var exsistingRight = record.GetAttributeValue<EntityReference>("sdu_eksisterendekemibrugadgang") != null ? record.GetAttributeValue<EntityReference>("sdu_eksisterendekemibrugadgang").Name : null;

                            if (newRight != null)
                            {
                                stringBuilder.AppendLine("Tilføj rettighed til: " + record.GetAttributeValue<EntityReference>("sdu_kemibrug").Name);
                            }
                            else if (exsistingRight != null)
                            {
                                String[] seperator = { "_" };
                                var kemiBrugSplit = record.GetAttributeValue<EntityReference>("sdu_eksisterendekemibrugadgang").Name.Split(seperator, StringSplitOptions.RemoveEmptyEntries);

                                stringBuilder.AppendLine("Fjern rettighed til: " + kemiBrugSplit.GetValue(0));
                            }
                        }
                    }*/

                }



                return stringBuilder;
            }


        }

        public class Person
        {
            public string Name { get; set; }
            public string OmkSted { get; set; }
            public string Email { get; set; }

            public string JobTitle { get; set; }

            public Person(IOrganizationService service, EntityReference person)
            {
                var personEntity = service.Retrieve(person.LogicalName, person.Id, new ColumnSet(new[] { "fullname", "sdu_crmomkostningssted", "emailaddress1", "jobtitle" }));

                Name = personEntity.GetAttributeValue<string>("fullname");
                OmkSted = personEntity.GetAttributeValue<EntityReference>("sdu_crmomkostningssted") == null ? null : personEntity.GetAttributeValue<EntityReference>("sdu_crmomkostningssted").Name;
                Email = personEntity.GetAttributeValue<string>("emailaddress1");
                JobTitle = personEntity.GetAttributeValue<string>("jobtitle");
            }
        }

        public class Brugeradministration
        {
            public DateTime forventetStart { get; set; }

            public Brugeradministration(IOrganizationService service, EntityReference brugeradministration)
            {
                var brugeradministrationEntity = service.Retrieve(brugeradministration.LogicalName, brugeradministration.Id, new ColumnSet(new[] { "sdu_forventetstartdato" }));

                forventetStart = brugeradministrationEntity.GetAttributeValue<DateTime>("sdu_forventetstartdato");
            }
        }

        public class User
        {
            public string Name { get; set; }
            public string Domain { get; set; }
            public string Email { get; set; }

            public User(IOrganizationService service, Guid userID)
            {
                // get initiating user 
                var initiatingUser = service.Retrieve("systemuser", userID, new ColumnSet(new[] { "domainname", "fullname", "internalemailaddress" }));

                Name = initiatingUser.GetAttributeValue<string>("fullname");
                Domain = initiatingUser.GetAttributeValue<string>("domain");
                Email = initiatingUser.GetAttributeValue<string>("internalemailaddress");
            }
        }

        public static DataCollection<Entity> QueryForRelatedRecords(bool isThisNew, Guid recordID, IOrganizationService service, string relatedEntityName)
        {
            var query = new QueryExpression();

            if (isThisNew)
            {
                query = new QueryExpression
                {
                    EntityName = relatedEntityName,
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        //FilterOperator = LogicalOperator.And,
                        Conditions =
                            {
                                new ConditionExpression("sdu_brugeradministrationsndring", ConditionOperator.Equal,  recordID),
                                new ConditionExpression("statuscode", ConditionOperator.Equal,  1)

                            }
                    }
                };
            }
            else // is this even necessary?
            {
                query = new QueryExpression
                {
                    EntityName = relatedEntityName,
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        FilterOperator = LogicalOperator.And,
                        Conditions =
                            {
                                new ConditionExpression("sdu_brugeradministrationsndring", ConditionOperator.Equal,  recordID)
                                //new ConditionExpression("sdu_rettighedsstatus", ConditionOperator.Equal, 100000000)
                            }
                    }
                };
            }

            return service.RetrieveMultiple(query).Entities;

        }

        public static string GetOptionsetText(IOrganizationService service, string optionsetName, int optionsetValue)
        {
            string optionsetSelectedText = string.Empty;
            try
            {
                RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = optionsetName
                    };

                // Execute the request.
                RetrieveOptionSetResponse retrieveOptionSetResponse =
                    (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    if (optionMetadata.Value == optionsetValue)
                    {
                        optionsetSelectedText = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                        break;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return optionsetSelectedText;
        }

        public void SendEmail(IOrganizationService service, DataCollection<Entity> relatedRecords, Guid EmailSender, Guid RegardingID, string EmailTo, string EmailBody, string EmailSubject, string AdditionalCC = "")
        {
            var ccToIntitutlederList = new List<Entity>();
            var ccToDekanList = new List<Entity>();
            var lonPrognose = new Entity("activityparty");
            bool lonPrognosePresent = false;
            bool useDekan = false;


            if (relatedRecords != null)
            {

                foreach (var record in relatedRecords)
                {
                    var BrugerAdmOmkostningsstedReference = record.GetAttributeValue<EntityReference>("sdu_omkostningssted");

                    var BrugerAdmOmkostningssted = service.Retrieve(BrugerAdmOmkostningsstedReference.LogicalName, BrugerAdmOmkostningsstedReference.Id, new ColumnSet(true));

                    var OmkostningsstedReference = BrugerAdmOmkostningssted.GetAttributeValue<EntityReference>("sdu_omksted");

                    var Omkostningssted = service.Retrieve(OmkostningsstedReference.LogicalName, OmkostningsstedReference.Id, new ColumnSet(true));

                    var institutLeder = Omkostningssted.GetAttributeValue<EntityReference>("sdu_institutleder");


                    // check for xx -> dekan skal CC istedet for institutleder
                    if (BrugerAdmOmkostningssted.GetAttributeValue<string>("sdu_name").Contains("xx"))
                    {
                        var dekan = Omkostningssted.GetAttributeValue<EntityReference>("sdu_dekan");

                        Entity CCToDekan = new Entity("activityparty");
                        CCToDekan["partyid"] = new EntityReference(dekan.LogicalName, dekan.Id);

                        ccToDekanList.Add(CCToDekan);

                        useDekan = true;
                    }

                    Entity CCToLeder = new Entity("activityparty");
                    CCToLeder["partyid"] = new EntityReference(institutLeder.LogicalName, institutLeder.Id);

                    ccToIntitutlederList.Add(CCToLeder);

                    // QLIKVIEW
                    if (record.LogicalName == "sdu_qlikviewadgang")
                    {
                        if (record.GetAttributeValue<bool>("sdu_lnprognoseeksternvirksomhed") || record.GetAttributeValue<bool>("sdu_lnprognoseordinrvirksomhed"))
                        {
                            lonPrognosePresent = true;
                            lonPrognose["addressused"] = "qlikview@sdu.dk";
                        }

                    }

                }
            }

            if (lonPrognosePresent)
            {
                ccToIntitutlederList.Add(lonPrognose);
                ccToDekanList.Add(lonPrognose);
            }

            // forudsætter at man ikke vælger andet end eget fakultet!
            var ccTo = useDekan == true ? ccToDekanList.ToArray() : ccToIntitutlederList.ToArray();

            Entity emailFrom = new Entity("activityparty");
            Entity emailTo = new Entity("activityparty");
            //Entity CCTo = new Entity("activityparty");


            emailFrom["partyid"] = new EntityReference("systemuser", EmailSender);
            emailTo["addressused"] = EmailTo;

            //CCTo["partyid"] = new EntityReference("contact", new Guid("956c7143-d0d5-e811-8123-0050568a4348"));

            Entity email = new Entity("email");

            email["from"] = new Entity[] { emailFrom };
            email["to"] = new Entity[] { emailTo };

            email["cc"] = ccTo;


            email["regardingobjectid"] = new EntityReference("sdu_brugeradministrationsndringer", RegardingID);
            email["subject"] = EmailSubject;
            email["description"] = EmailBody;
            email["directioncode"] = true;
            Guid emailId = service.Create(email);

            // Use the SendEmail message to send an e-mail message.
            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);

        }

        public static bool CompareWithMasterRecord(Entity masterRecord, KeyValuePair<string, string> attribute)
        {
           /* if (attribute.Value.ToLower() != "true" && attribute.Value.ToLower() != "false" && attribute.Value != "No" && attribute.Value != "Yes")
            {
                return true;
            }*/

            if ((masterRecord.GetAttributeValue<bool>(attribute.Key) == true && attribute.Value == "No") || (masterRecord.GetAttributeValue<bool>(attribute.Key) == false && attribute.Value == "Yes")) /*|| (masterRecord.GetAttributeValue<bool>(attribute.Key) == true && attribute.Value == "False") || (masterRecord.GetAttributeValue<bool>(attribute.Key) == false && attribute.Value == "True"))/*|| (masterRecord.GetAttributeValue<bool>(attribute.Key) == true && attribute.Value == "Yes")*/
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        // get the display name for the attribute
        public static AttributeMetadata RetrieveAttributeMetadata(string EntitySchemaName, string AttributeSchemaName, IOrganizationService service)
        {

            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = EntitySchemaName,
                LogicalName = AttributeSchemaName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            return (AttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
        }
    }
}
