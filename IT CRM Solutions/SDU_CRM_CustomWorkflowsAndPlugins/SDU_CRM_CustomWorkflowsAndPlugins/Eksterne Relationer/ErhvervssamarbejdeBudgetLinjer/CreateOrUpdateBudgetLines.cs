using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;

namespace EksterneRelationer
{
    public class ER_CreateOrUpdateBudgetLines : CodeActivity
    {

        protected override void Execute(CodeActivityContext context)
        {
            // Get the context service.
            IWorkflowContext Icontext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.             
            IOrganizationService service = serviceFactory.CreateOrganizationService(Icontext.UserId); // user = SDU\SQLCRMDynWorkF


            // get current record
            var fakProjekt = service.Retrieve(Icontext.PrimaryEntityName, Icontext.PrimaryEntityId, new ColumnSet(true));

            var projectStartDate = fakProjekt.GetAttributeValue<DateTime>("sdu_projektstart");
            var projectEndDate = fakProjekt.GetAttributeValue<DateTime>("sdu_projektslut");

            // put the difference in years into list
            List<int> years = new List<int>();

            if (projectStartDate != null && projectEndDate != null)
            {
                for (int i = projectStartDate.ToLocalTime().Year; i <= projectEndDate.ToLocalTime().Year; i++)
                {
                    years.Add(i);
                }
            }

            // get bevilling + medf SDU
            var bevilling = fakProjekt.GetAttributeValue<Money>("sdu_bevillingtilsdu")?.Value;
            var medf = fakProjekt.GetAttributeValue<Money>("sdupro_medfinansieringbevilling")?.Value;

            decimal amountBev;
            decimal amountMedf;

            // amount to distribute
            if (years.Count > 0)
            {
                if (bevilling != null)
                {
                    amountBev = (decimal)bevilling / years.Count;
                }
                else
                {
                    amountBev = 0;
                }

                if (medf != null)
                {
                    amountMedf = (decimal)medf / years.Count;
                }
                else
                {
                    amountMedf = 0;
                }
            }
            else
            {
                amountMedf = 0;
                amountBev = 0;
            }

            // definition
            var fieldIndexDefinition = new List<Tuple<int, string>>
            {
                new Tuple<int, string>(0, "sdu_er_budgetlinje_year1"),
                new Tuple<int, string>(1, "sdu_er_budgetlinje_year2"),
                new Tuple<int, string>(2, "sdu_er_budgetlinje_year3"),
                new Tuple<int, string>(3, "sdu_er_budgetlinje_year4"),
                new Tuple<int, string>(4, "sdu_er_budgetlinje_year5"),
                new Tuple<int, string>(5, "sdu_er_budgetlinje_year6"),
                new Tuple<int, string>(6, "sdu_er_budgetlinje_year7"),
                new Tuple<int, string>(7, "sdu_er_budgetlinje_year8"),
                new Tuple<int, string>(8, "sdu_er_budgetlinje_year9"),
                new Tuple<int, string>(9, "sdu_er_budgetlinje_year10")
            };

            // loop through 10 years MAX and field the corresponding logical name in the definition.
            for (int i = 0; i < years.Count; i++)
            {
                // get the field name corresponding to the item1 in definition
                var crmFieldName = fieldIndexDefinition.Find(tuple => tuple.Item1 == i)?.Item2;
                if (!String.IsNullOrEmpty(crmFieldName) && (amountBev != 0 || amountMedf != 0))
                {
                    var crmField = fakProjekt.GetAttributeValue<EntityReference>(crmFieldName);

                    CreateOrUpdateBudgetLines(service, years[i].ToString(), amountBev, amountMedf, fakProjekt, crmField, crmFieldName);

                }
            }

            CheckForExtraYears(service, fakProjekt, years.Count, fieldIndexDefinition);

        }

        public void CreateOrUpdateBudgetLines(IOrganizationService service, string Year, decimal AmountBev, decimal AmountMedf, Entity FakProjekt, EntityReference BudgetLine, string crmFieldName)
        {
            // establish entity to either create or update
            var budgetLine = new Entity("sdu_erhvervssamarbejdebudgetlinjer");
            budgetLine["sdu_name"] = Year;
            budgetLine["sdu_belb"] = new Money(AmountBev);
            budgetLine["sdu_medfinansieringsbelb"] = new Money(AmountMedf);
            budgetLine["sdu_fakultetsprojekt"] = new EntityReference(FakProjekt.LogicalName, FakProjekt.Id);
            budgetLine["sdu_internt"] = true;

            // Update
            if (BudgetLine != null)
            {
                budgetLine.Id = BudgetLine.Id;
                service.Update(budgetLine);
            }
            else // Create
            {
                var budgetLineCreated = service.Create(budgetLine);

                var updateFakProj = new Entity(FakProjekt.LogicalName) { Id = FakProjekt.Id };
                updateFakProj[crmFieldName] = new EntityReference(budgetLine.LogicalName, budgetLineCreated);

                service.Update(updateFakProj);
            }



        }

        public void CheckForExtraYears(IOrganizationService service, Entity fakProjekt, int years, List<Tuple<int, string>> fieldDefinition)
        {
            var exsistingYears = new List<EntityReference>();

            // check for number of exsisting budget lines
            foreach (var field in fieldDefinition)
            {
                // use the logical name from the fielDef list to fetch each field
                var yearLine = fakProjekt.GetAttributeValue<EntityReference>(field.Item2);

                if (yearLine != null)
                {
                    exsistingYears.Add(yearLine);
                }
            }

            // compare the results to the amount of years calculated
            if (years != exsistingYears.Count)
            {
                // delete the remaining years
                for (int i = years; i < exsistingYears.Count; i++)
                {
                    service.Delete(exsistingYears[i].LogicalName, exsistingYears[i].Id);
                }
            }


        }




    }
}
