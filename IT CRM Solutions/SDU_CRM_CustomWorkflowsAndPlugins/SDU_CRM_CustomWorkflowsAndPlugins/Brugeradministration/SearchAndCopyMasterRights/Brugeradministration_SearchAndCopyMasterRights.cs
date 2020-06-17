using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text;

namespace Brugeradministration
{
    public class Brugeradministration_SearchAndCopyMasterRights : IPlugin
    {
        public const int BestillingsStatusGodkendt = 100000001;

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
            var callingRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("sdu_brugeradministration"));
            var brugeradministrationId = callingRecord.GetAttributeValue<EntityReference>("sdu_brugeradministration").Id;

            // current record entity refrence
            var callingRecordReference = new EntityReference("sdu_brugeradministrationsndringer", callingRecord.Id);

            // CALLING PARAMETERS?!
            //var test = WorkflowInvoker.Invoke(new GetParameters());


            // DEFINE PER AREA ///////////////////////////

            //// ØSS //////
            const string OessEntityName = "sdu_oessadgang";

            // search for the master-records aka the accepted ones
            var approvedrightsOess = QueryForApprovedRights(service, brugeradministrationId, OessEntityName, "sdu_bestillingsstatus", BestillingsStatusGodkendt);

            // Which fields should be copied?
            var oessFieldsToCopy = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("Bool", "sdu_omkontering"),
                new Tuple<string, string>("EntityReference", "sdu_omkostningssted")
            };

            // call the method to copy the fields
            if (approvedrightsOess.Entities.Count > 0)
            {
                CreateCopiesOfApprovedRights(service, approvedrightsOess, oessFieldsToCopy, callingRecordReference, OessEntityName, new string[] { "sdu_omkostningssted" });
                UpdateRightsOrderWithNumberOfRecordsCreated(service, callingRecordReference, approvedrightsOess.Entities.Count, "sdu_oessnumberofrightscreated");
            }

            //// QlikView //////
            const string QlikViewEntityName = "sdu_qlikviewadgang";

            var approvedrightsQlikView = QueryForApprovedRights(service, brugeradministrationId, QlikViewEntityName, "sdu_bestillingsstatus", BestillingsStatusGodkendt);

            // Which fields should be copied?
            var qlikViewFieldsToCopy = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("Bool", "sdu_finansforesprgsel"),
                new Tuple<string, string>("Bool", "sdu_lnprognoseeksternvirksomhed"),
                new Tuple<string, string>("Bool", "sdu_lnprognoseordinrvirksomhed"),
                new Tuple<string, string>("Bool", "sdu_lnforesprgsel"),
                new Tuple<string, string>("EntityReference", "sdu_omkostningssted")
            };

            // call the method to copy the fields
            if (approvedrightsQlikView.Entities.Count > 0)
            {
                CreateCopiesOfApprovedRights(service, approvedrightsQlikView, qlikViewFieldsToCopy, callingRecordReference, QlikViewEntityName, new string[] { "sdu_omkostningssted" });
                UpdateRightsOrderWithNumberOfRecordsCreated(service, callingRecordReference, approvedrightsQlikView.Entities.Count, "sdu_qlikviewnumberofrightscreated");
            }

            //// HR //////
            const string HREntityName = "sdu_hradgang";

            var approvedrightsHR = QueryForApprovedRights(service, brugeradministrationId, HREntityName, "sdu_bestillingsstatus", BestillingsStatusGodkendt);

            // Which fields should be copied?
            var hrFieldsToCopy = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("Bool", "sdu_foresprgsel"),
                new Tuple<string, string>("Bool", "sdu_fravrsadministration"),
                new Tuple<string, string>("EntityReference", "sdu_omkostningssted")
            };

            // call the method to copy the fields
            if (approvedrightsHR.Entities.Count > 0)
            {
                CreateCopiesOfApprovedRights(service, approvedrightsHR, hrFieldsToCopy, callingRecordReference, HREntityName, new string[] { "sdu_omkostningssted" });
                UpdateRightsOrderWithNumberOfRecordsCreated(service, callingRecordReference, approvedrightsHR.Entities.Count, "sdu_hrnumberofrightscreated");

            }


            //// Acadre //////
            const string AcadreEntityName = "sdu_acadreafdelingerspecifikkeadgange";

            var approvedrightsAcadre = QueryForApprovedRights(service, brugeradministrationId, AcadreEntityName, "sdu_rettighedsstatus", BestillingsStatusGodkendt);

            // Which fields should be copied?
            var AcadreFieldsToCopy = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("Bool", "sdu_bevillingssagerxb"),
                new Tuple<string, string>("Bool", "sdu_arbejdsmilj"),
                new Tuple<string, string>("Bool", "sdu_chefch"),
                new Tuple<string, string>("Bool", "sdu_studentersagst"),
                new Tuple<string, string>("Bool", "sdu_susagsu"),
                new Tuple<string, string>("Bool", "sdu_sikkerhedsbrudab"),
                new Tuple<string, string>("Bool", "sdu_samarbejdsaftalersa"),
                new Tuple<string, string>("Bool", "sdu_inkassoin"),
                new Tuple<string, string>("Bool", "sdu_phdstudiumph"),
                new Tuple<string, string>("Bool", "sdu_personalesagpe"),
                new Tuple<string, string>("Bool", "sdu_fortroligfo"),
                new Tuple<string, string>("Bool", "sdu_rekruttering"),
                new Tuple<string, string>("Bool", "sdu_patentpa"),
                new Tuple<string, string>("Bool", "sdu_ansgningxa"),
                new Tuple<string, string>("Bool", "sdu_byggesagerbs"),
                new Tuple<string, string>("Bool", "sdu_forskningfk"),
                new Tuple<string, string>("EntityReference", "sdu_acadreafdeling")
            };

            // call the method to copy the fields
            if (approvedrightsAcadre.Entities.Count > 0)
            {
                CreateCopiesOfApprovedRights(service, approvedrightsAcadre, AcadreFieldsToCopy, callingRecordReference, AcadreEntityName, new string[] { "sdu_acadreafdeling" });
                UpdateRightsOrderWithNumberOfRecordsCreated(service, callingRecordReference, approvedrightsAcadre.Entities.Count, "sdu_acadrenumberofrightscreated");

            }


            //// ITsystemer //////
            const string itSystemer = "sdu_adgangtilitsystem";

            var approvedrightsitSystemer = QueryForApprovedRights(service, brugeradministrationId, itSystemer, "sdu_adgangsstatus", BestillingsStatusGodkendt);

            // Which fields should be copied?
            var itSystemerFieldsToCopy = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("EntityReference", "sdu_itsystem"),
                new Tuple<string, string>("EntityReference", "sdu_itsystem_master"),
                new Tuple<string,string>("EntityReference", "sdu_afdeling")
            };

            // call the method to copy the fields
            if (approvedrightsitSystemer.Entities.Count > 0)
            {
                CreateCopiesOfApprovedRights(service, approvedrightsitSystemer, itSystemerFieldsToCopy, callingRecordReference, itSystemer, new string[] { "sdu_itsystem", "sdu_afdeling" });
                //UpdateRightsOrderWithNumberOfRecordsCreated(service, callingRecordReference, approvedrightsAcadre.Entities.Count, "sdu_acadrenumberofrightscreated");
            }

        }

        public EntityCollection QueryForApprovedRights(IOrganizationService service, Guid CallingRecordId, string Entity, string QueryKey, int QueryValue)
        {

            // search for matching record
            var query = new QueryExpression(Entity)
            {
                ColumnSet = new ColumnSet(true)
            };

            var condition_id = new ConditionExpression()
            {
                AttributeName = "sdu_brugeradministration",
                Operator = ConditionOperator.Equal,
                Values = { CallingRecordId }
            };

            var condition_status = new ConditionExpression()
            {
                AttributeName = QueryKey,
                Operator = ConditionOperator.Equal,
                Values = { QueryValue } // == godkendt
            };

            var condition_active = new ConditionExpression()
            {
                AttributeName = "statuscode",
                Operator = ConditionOperator.Equal,
                Values = { 1 } // == ACTIVE
            };

            query.Criteria.AddCondition(condition_id);
            query.Criteria.AddCondition(condition_status);
            query.Criteria.AddCondition(condition_active);


            return service.RetrieveMultiple(query);
        }

        public void CreateCopiesOfApprovedRights(IOrganizationService service, EntityCollection QueryResults, List<Tuple<string, string>> FieldsToCopy, EntityReference CallingRecordReference, string TargetEntityName, string[] altKeyGen)
        {
            foreach (var record in QueryResults.Entities)
            {
                // initialize request - to make sure it inherits the correct mappings from CRM.
                var request = new InitializeFromRequest()
                {
                    EntityMoniker = CallingRecordReference,
                    TargetEntityName = TargetEntityName
                };

                var response = (InitializeFromResponse)service.Execute(request);

                // get entity from the response
                Entity newRecord = response.Entity;


                // copy values from master record
                // t1 = dataType, t2 = logicalname
                foreach (var field in FieldsToCopy)
                {
                    switch (field.Item1)
                    {
                        case "Bool":
                            newRecord[field.Item2] = record.GetAttributeValue<bool>(field.Item2);
                            break;
                        case "EntityReference":
                            var checkValue = record.GetAttributeValue<EntityReference>(field.Item2) == null ? false : true;
                            if (checkValue)
                            {
                                newRecord[field.Item2] = new EntityReference(record.GetAttributeValue<EntityReference>(field.Item2).LogicalName, record.GetAttributeValue<EntityReference>(field.Item2).Id);

                            }
                            break;
                        default:
                            break;
                    }
                }

                // take the id from the master and insert into lookupfield on child
                newRecord["sdu_masterrecord"] = new EntityReference(TargetEntityName, record.Id);
                newRecord["sdu_alternatekey"] = GenerateAlternateKey(service, newRecord, "sdu_brugeradministrationsndring", altKeyGen);


                //newRecord["sdu_omkostningssted"] = new EntityReference("sdu_bruger", record.GetAttributeValue<EntityReference>("sdu_omkostningssted").Id);


                var createdEntity = service.Create(newRecord);

            }
        }

        public void UpdateRightsOrderWithNumberOfRecordsCreated(IOrganizationService service, EntityReference caliingEntityReference, int count, string fieldToUpdate)
        {
            var UpdatedEntity = new Entity(caliingEntityReference.LogicalName)
            {
                Id = caliingEntityReference.Id
            };
            UpdatedEntity[fieldToUpdate] = count;

            service.Update(UpdatedEntity);
        }

        public string GenerateAlternateKey(IOrganizationService service, Entity entityToUpdate, string keyOne, string[] extraKeys)
        {
            var extraKeyIds = new StringBuilder();
            var keyOneId = entityToUpdate.GetAttributeValue<EntityReference>(keyOne) != null ? entityToUpdate.GetAttributeValue<EntityReference>(keyOne).Id.ToString() : "";

            if (extraKeys.Length > 0)
            {
                foreach (var key in extraKeys)
                {
                    extraKeyIds.Append(entityToUpdate.GetAttributeValue<EntityReference>(key) != null ? entityToUpdate.GetAttributeValue<EntityReference>(key).Id.ToString() : "");
                }
            }


            //var keyTwoId = entityToUpdate.GetAttributeValue<EntityReference>(extraKeys) != null ? entityToUpdate.GetAttributeValue<EntityReference>(extraKeys).Id.ToString() : "";

            var completeKey = keyOneId + extraKeyIds.ToString();

            return completeKey;
        }
    }
}
