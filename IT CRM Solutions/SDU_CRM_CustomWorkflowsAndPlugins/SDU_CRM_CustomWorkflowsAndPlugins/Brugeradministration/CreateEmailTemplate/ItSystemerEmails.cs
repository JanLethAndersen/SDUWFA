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
    public class ItSystemerEmails : CodeActivity
    {
        [Input("Contact")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> Contact { get; set; }

        [Input("Parameter reference")]
        [ReferenceTarget("sdu_parameter")]
        public InArgument<EntityReference> ParameterReference { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            // Get the context service.
            IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.             
            IOrganizationService service = serviceFactory.CreateOrganizationService(Icontext.UserId); // user = SDU\SQLCRMDynWorkF

            var tracingService = context.GetExtension<ITracingService>();

            tracingService.Trace("EXECUTE.....");

            try
            {

                var relatedRecords = CreateEmailTemplate.QueryForRelatedRecords(true, Icontext.PrimaryEntityId, service, "sdu_adgangtilitsystem");
                var contact = service.Retrieve(Contact.Get(context).LogicalName, Contact.Get(context).Id, new ColumnSet(true));
                tracingService.Trace("CONTACTID " + contact.Id.ToString());
                var parameters = service.Retrieve(ParameterReference.Get(context).LogicalName, ParameterReference.Get(context).Id, new ColumnSet(true));
                tracingService.Trace("PARAMETERS " + parameters.Id.ToString());
                var emailSender = parameters.GetAttributeValue<EntityReference>("sdu_baemailafsender");
                tracingService.Trace("EMAILSENDER " + emailSender.Id.ToString());



                if (relatedRecords.Count > 0)
                {
                    GroupRightsTogether(service, emailSender.Id, Icontext, relatedRecords, new Person() { FullName = contact.GetAttributeValue<string>("fullname"), Email = contact.GetAttributeValue<string>("emailaddress1") }, parameters, tracingService);
                    //var emailBody = ConstructEmailBody(groupedRigths);
                    //SendEmail(service, Icontext.UserId, Icontext.PrimaryEntityId, "jlan@sdu.dk", emailBody, ".....");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(ex.ToString());
                throw;
            }


        }

        public void GroupRightsTogether(IOrganizationService service, Guid emailSender, IWorkflowContext Icontext, DataCollection<Entity> records, Person person, Entity parameters, ITracingService tracingService)
        {

            try
            {
                tracingService.Trace("starting...");
                var rightsGroup = new List<RightsObject>();

                // determine the rigth groups
                foreach (var record in records)
                {

                    var rettighedIitSystemRef = record.GetAttributeValue<EntityReference>("sdu_itsystem");
                    var rettighedIitSystemEnt = service.Retrieve(rettighedIitSystemRef.LogicalName, rettighedIitSystemRef.Id, new ColumnSet(true));

                    // get name of the system + id
                    var itSystemRef = rettighedIitSystemEnt.GetAttributeValue<EntityReference>("sdu_itsystem");
                    var itSystemEnt = service.Retrieve(itSystemRef.LogicalName, itSystemRef.Id, new ColumnSet(true));
                    //var systemName = itSystemEnt["sdu_name"].ToString();
                    //var systemId = record.GetAttributeValue<string>("sdu_itsystemid");
                    //var systemMailbox = itSystemEnt.GetAttributeValue<string>("sdu_systemmailboks");

                    var ItSystem = new ItSystem(itSystemEnt["sdu_name"].ToString(), itSystemEnt["sdu_godkendelsesid"].ToString(), itSystemEnt.GetAttributeValue<bool>("sdu_krverinstitutledergodkendelse"), itSystemEnt.GetAttributeValue<bool>("sdu_cctilinstitutleder"), itSystemEnt["sdu_systemmailboks"].ToString());

                    var godkendtInstitutleder = record.GetAttributeValue<bool>("sdu_godkendtinstitutleder");
                    var institutLederGodkendelse = itSystemEnt.GetAttributeValue<bool>("sdu_krverinstitutledergodkendelse");

                    var masterRecord = record.GetAttributeValue<EntityReference>("sdu_masterrecord");
                    var rigthRemoved = record.GetAttributeValue<bool>("sdu_adgangskalfjernes");
                    var afdeling = record.GetAttributeValue<EntityReference>("sdu_afdeling") == null ? null : record.GetAttributeValue<EntityReference>("sdu_afdeling").Name;


                    var rigth = rettighedIitSystemEnt.GetAttributeValue<string>("sdu_name");

                    var afsendtInstitutleder = record.GetAttributeValue<DateTime>("sdu_afsendttilinstitutleder");
                    var afsendtSystemejer = record.GetAttributeValue<DateTime>("sdu_afsendttilsystemejer");


                    // only add if there is no master record = new right, or the right is going to removed
                    // also check if the right has already been sent
                    if ((masterRecord == null || rigthRemoved == true) && (afsendtSystemejer.Year < 2000 || (institutLederGodkendelse == true && afsendtInstitutleder.Year > 2000)))
                    {
                        rightsGroup.Add(new RightsObject() { Rigth = rigth, ToBeDeleted = rigthRemoved, GodkendtAfInstitutLeder = godkendtInstitutleder, Record = new Entity(record.LogicalName) { Id = record.Id }, ItSystem = ItSystem, Afdeling = afdeling });
                    }
                }

                tracingService.Trace("Righsgroup checking...");

                if (rightsGroup.Count > 0)
                {
                    tracingService.Trace("Rightsgroups = TRUE");
                    var grouped = rightsGroup.GroupBy(group => group.ItSystem.Name);

                    tracingService.Trace("Group count: " + grouped.Count().ToString());

                    foreach (var group in grouped)
                    {
                        // get all the element which is going to be deleted
                        var toBeGrantedGroup = group.Where(item => item.ToBeDeleted == false);

                        var subject = "Angående rettigheder til: " + group.First().ItSystem.Name + " for " + person.FullName + " (" + person.Email + ")";

                        // do not cc institutleder if all is going to be deleted
                        if (group.First().ItSystem.InstitutlederGodkendelse && toBeGrantedGroup.Count() >= 1)
                        {
                            if (!group.First().GodkendtAfInstitutLeder)
                            {
                                tracingService.Trace("Institutleder tjek " + group.First().GodkendtAfInstitutLeder.ToString());
                                SendEmail(service, emailSender, Icontext.PrimaryEntityId, GetInstitutLeder(service, Icontext).Email, ConstructEmailBody(group, group.First().ItSystem.InstitutlederGodkendelse, subject, parameters["sdu_emailgodkendtinstitutleder"].ToString(), parameters["sdu_emailafvistinstitutleder"].ToString(), group.First().ItSystem.Id, person), subject, new Person());
                                tracingService.Trace("Send email");
                                UpdateRecordDatesAfterEmailSend(group, service, "sdu_afsendttilinstitutleder", DateTime.Now);
                            }
                            else
                            {
                                SendEmail(service, emailSender, Icontext.PrimaryEntityId, group.First().ItSystem.SystemMailbox, ConstructEmailBody(group, false, subject, parameters["sdu_emailgodkendelser"].ToString(), parameters["sdu_emailafvisninger"].ToString(), group.First().ItSystem.Id, person), subject, new Person());
                                UpdateRecordDatesAfterEmailSend(group, service, "sdu_afsendttilsystemejer", DateTime.Now);
                            }
                        }
                        else
                        {
                            SendEmail(service, emailSender, Icontext.PrimaryEntityId, group.First().ItSystem.SystemMailbox, ConstructEmailBody(group, false, subject, parameters["sdu_emailgodkendelser"].ToString(), parameters["sdu_emailafvisninger"].ToString(), group.First().ItSystem.Id, person), subject, group.First().ItSystem.CCTilInstitutleder == true ? GetInstitutLeder(service, Icontext) : new Person());
                            UpdateRecordDatesAfterEmailSend(group, service, "sdu_afsendttilsystemejer", DateTime.Now);
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(ex.ToString());
                throw ex;
            }
            //var rigthsGroups = new List<string>();




        }


        public void UpdateRecordDatesAfterEmailSend(IGrouping<string, RightsObject> group, IOrganizationService service, string fieldToUpdate, DateTime value)
        {
            foreach (var item in group)
            {
                var entity = item.Record;
                entity[fieldToUpdate] = value;
                service.Update(entity);
            }

        }

        public class RightsObject
        {
            public string Rigth { get; set; }
            public bool ToBeDeleted { get; set; }
            public bool GodkendtAfInstitutLeder { get; set; }

            public Entity Record { get; set; }
            public ItSystem ItSystem { get; set; }
            public string Afdeling { get; set; }

        }

        public class ItSystem
        {
            public string Name { get; set; }
            public string Id { get; set; }
            public bool InstitutlederGodkendelse { get; set; }
            public bool CCTilInstitutleder { get; set; }
            public string SystemMailbox { get; set; }

            public ItSystem(string name, string id, bool institutLederGodkendelse, bool ccTilInstitutleder, string systemMailBox)
            {
                Name = name;
                Id = id;
                InstitutlederGodkendelse = institutLederGodkendelse;
                CCTilInstitutleder = ccTilInstitutleder;
                SystemMailbox = systemMailBox;
            }

        }

        public class Person
        {
            public string FullName { get; set; }
            public string Email { get; set; }

        }


        public string ConstructEmailBody(IGrouping<string, RightsObject> rigths, bool InstitutlederGodkRequired, string emailSubject, string acceptRightsKey, string declineKey, string rightsKey, Person person)
        {
            var stringBuilder = new StringBuilder();

            stringBuilder.AppendLine("Hej<br><br>");

            if (InstitutlederGodkRequired)
            {
                stringBuilder.AppendLine("<b>Følgende rettigheder kræver din godkendelse, førend de kan tildeles personen. Du bedes derfor tage stilling, ved at trykke på ét af de links i bunden af emailen, og sende emailen retur.</b><br><br>");
            }

            stringBuilder.AppendLine("Rettigheder for: <b>" + rigths.First().ItSystem.Name + "</b> til " + person.FullName + "<br><br>");

            var newRigths = rigths.Where(newRigth => newRigth.ToBeDeleted == false);
            var removeRigths = rigths.Where(removeRigth => removeRigth.ToBeDeleted == true);


            if (newRigths.Any())
            {
                stringBuilder.AppendLine("Følgende rettigheder ønskes tildelt: <br>");

                var RightsAfdelinger = newRigths.GroupBy(rights => rights.Afdeling);

                if (RightsAfdelinger.Any())
                {
                    foreach (var group in RightsAfdelinger)
                    {
                        if (group.First().Afdeling == "" || group.First().Afdeling == null)
                        {
                            stringBuilder.AppendLine("<br><b>Diverse:</b> <br>");
                        }
                        else
                        {
                            stringBuilder.AppendLine("<br><b>" + group.First().Afdeling + "</b> <br>");
                        }
                        foreach (var right in group)
                        {
                            stringBuilder.AppendLine("- " + right.Rigth + "<br>");
                        }
                    }
                }
                else
                {
                    foreach (var rigth in newRigths)
                    {
                        stringBuilder.AppendLine("- " + rigth.Rigth + "<br>");
                    }
                }

            }

            if (removeRigths.Any())
            {
                stringBuilder.AppendLine("<br>Følgende rettigheder ønskes frataget: <br>");
                foreach (var rigth in removeRigths)
                {

                    stringBuilder.AppendLine("- " + rigth.Rigth + "<br>");

                }
            }

            //var reasonEncoded = "%0D%0A%0D%0ADu kan indenfor de to hashtags (#) herunder angive årsagen til afvisningen: %0D%0A%0D%0A# %0D%0A... %0D%0A#";

            stringBuilder.Append("<br>");
            stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet tildelt - {2} | {3} |'><b>Rettighederne er godkendt</b></a>", "mscrm@sdu.dk", emailSubject, acceptRightsKey, rightsKey);
            stringBuilder.Append("<br>");
            stringBuilder.Append("<br>");
            stringBuilder.AppendFormat("<a href='mailto:{0}?subject={1}&body=Rettighederne er blevet afvist - {2} | {3} |'><b>Rettighederne er afvist</b></a>", "mscrm@sdu.dk", emailSubject, declineKey, rightsKey /*reasonEncoded*/);
            stringBuilder.Append("<br>");

            return stringBuilder.ToString();

        }

        public void SendEmail(IOrganizationService service, Guid EmailSender, Guid RegardingID, string EmailTo, string EmailBody, string EmailSubject, Person CCTo)
        {

            Entity emailFrom = new Entity("activityparty");
            Entity emailTo = new Entity("activityparty");
            Entity[] emailToArray = new Entity[] { };
            Entity emailCC = new Entity("activityparty");
            //Entity CCTo = new Entity("activityparty");

            Entity email = new Entity("email");

            var multipleEmails = EmailTo.Contains(";");

            if (multipleEmails)
            {
                var emailToSplitted = EmailTo.Split(';');

                for (int i = 0; i < emailToSplitted.Count(); i++)
                {
                    var emailAtPos = new Entity("activityparty");
                    emailAtPos["addressused"] = emailToSplitted[i];
                    Array.Resize(ref emailToArray, emailToArray.Length + 1);
                    emailToArray.SetValue(emailAtPos, i);
                }

                email["to"] = emailToArray;
            }
            else
            {
                emailTo["addressused"] = EmailTo;
                email["to"] = new Entity[] { emailTo };
            }




            emailFrom["partyid"] = new EntityReference("systemuser", EmailSender);
            //emailTo["addressused"] = EmailTo;




            email["from"] = new Entity[] { emailFrom };


            if (!String.IsNullOrEmpty(CCTo.Email))
            {
                emailCC["addressused"] = CCTo.Email;
                email["cc"] = new Entity[] { emailCC };
            }

            //CCTo["partyid"] = new EntityReference("contact", new Guid("956c7143-d0d5-e811-8123-0050568a4348"));

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


        public Person GetInstitutLeder(IOrganizationService service, IWorkflowContext brugeradmChange)
        {

            var BrugerAdmEntity = service.Retrieve(brugeradmChange.PrimaryEntityName, brugeradmChange.PrimaryEntityId, new ColumnSet("sdu_person"));
            var ContactReference = BrugerAdmEntity.GetAttributeValue<EntityReference>("sdu_person");
            var ContactEntity = service.Retrieve(ContactReference.LogicalName, ContactReference.Id, new ColumnSet("sdu_crmomkostningssted"));
            var OmkostningsstedReference = ContactEntity.GetAttributeValue<EntityReference>("sdu_crmomkostningssted");
            var OmkostningsstedsEntity = service.Retrieve(OmkostningsstedReference.LogicalName, OmkostningsstedReference.Id, new ColumnSet("sdu_institutleder"));
            var InstitutLederRef = OmkostningsstedsEntity.GetAttributeValue<EntityReference>("sdu_institutleder");
            var InstitutLederEntity = service.Retrieve(InstitutLederRef.LogicalName, InstitutLederRef.Id, new ColumnSet("emailaddress1", "fullname"));

            return new Person() { FullName = InstitutLederEntity.GetAttributeValue<string>("fullname"), Email = InstitutLederEntity.GetAttributeValue<string>("emailaddress1") };
        }
    }
}
