﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;
using System.ServiceModel;

namespace FPMailingLeadOpportunity
{
    public class AllContactsLeadAssigned
    {
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sAccountID)
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
	            a.pearl_assignlead = "True" and b.id = "incoming string"
            )
            */

            // int e = 1;
            // string eText = "";

            try
            {
                QueryExpression query = new QueryExpression()
                {
                    Distinct = false,
                    EntityName = Contact.EntityLogicalName,
                    ColumnSet = new ColumnSet(true),
                    LinkEntities = 
                {
                    new LinkEntity 
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "parentcustomerid",
                        LinkFromEntityName = Contact.EntityLogicalName,
                        LinkToAttributeName = "id",
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
                                new ConditionExpression("a.pearl_assignlead", ConditionOperator.Equal, true),
                                new ConditionExpression("b.ID", ConditionOperator.Equal, sAccountID)
                            },
                        }
                    }
                    }
                };

                entityCollection = service.RetrieveMultiple(query).Entities;

                return entityCollection;
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("*** An error occurred in the AllContactsLeadAssigned query.", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("** An error occurred in the AllContactsLeadAssigned query: {0}", ex.ToString());
                throw;
            }
        }
    }
}
