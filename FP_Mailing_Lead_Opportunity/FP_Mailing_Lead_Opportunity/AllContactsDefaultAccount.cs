using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace FPMailingLeadOpportunity
{
    public class AllContactsDefaultAccount
    {
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, EntityReference sDefaultAccount)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, ContactId, LeadEnable, Last Opportunity Assignment
            from
                Contact a
                    inner join
                Account b
                    on a.parentcustomerid = b.id
            where 
            (
	            a.pearl_assignlead = "True" and b.AccountID = "incoming string"
            )
            */

            /*QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = Contact.EntityLogicalName,
                ColumnSet = new ColumnSet("contactid", "pearl_assignlead", "pearl_lastopportunityassignment"),
                LinkEntities = 
        {
            new LinkEntity 
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "parentcustomerid",
                LinkFromEntityName = Contact.EntityLogicalName,
                //LinkToAttributeName = "contactid",
                LinkToAttributeName = "contact_customer_accounts",
                LinkToEntityName = Account.EntityLogicalName
            }
          
        },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions = 
                    {
                        new ConditionExpression("a.pearl_assignlead", ConditionOperator.Equal, "True"),
                        new ConditionExpression("b.accountid", ConditionOperator.Equal, sDefaultAccount)
                    },
                }
            }
                }
            };
            */
            
            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = Contact.EntityLogicalName,
                ColumnSet = new ColumnSet(true),
                LinkEntities =
                {
                },
                Criteria =
                {
                    Filters = 
                {
                    new FilterExpression
                    {
                        Conditions = 
                        {
                            new ConditionExpression("pearl_assignlead", ConditionOperator.Equal, true),
                            new ConditionExpression("parentcustomerid".ToString(), ConditionOperator.Equal, sDefaultAccount.Id.ToString())
                        },
                    }
                }
                }
            };
            entityCollection = service.RetrieveMultiple(query).Entities;
            return entityCollection;
        }
    }
}