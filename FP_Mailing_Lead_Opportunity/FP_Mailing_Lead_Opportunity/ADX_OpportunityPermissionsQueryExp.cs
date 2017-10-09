using System;
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
    public class ADX_OpportunityPermissionsQueryExp
    {
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, Guid gContact, Guid gAccount, int dSelfScope, int dAccountScope)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select *
            from
                adx_opportunitypermissions a
            where 
            (
	            a.Contact = gContact and a.Account = gAccount and (a.Scope = dSelfScope or a.Scope = dAccountScope)
            )
            */

            int e = 1;
            try
            {
                QueryExpression query = new QueryExpression()
                {
                    Distinct = false,
                    EntityName = adx_opportunitypermissions.EntityLogicalName,
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
                                FilterOperator = LogicalOperator.And,
                                Conditions = 
                                {
                                    new ConditionExpression("adx_contactid", ConditionOperator.Equal, gContact),
                                    new ConditionExpression("adx_accountid", ConditionOperator.Equal, gAccount)
                                },
                            },
                            new FilterExpression
                            {
                                FilterOperator = LogicalOperator.Or,
                                Conditions = 
                                {
                                    new ConditionExpression("adx_scope", ConditionOperator.Equal, dSelfScope),
                                    new ConditionExpression("adx_scope", ConditionOperator.Equal, dAccountScope)
                                },
                            }
                        }
                    }
                };

                e = 2;
                
                entityCollection = service.RetrieveMultiple(query).Entities;

                return entityCollection;
            }
             catch (FaultException<OrganizationServiceFault> ex)
             {
                 throw new InvalidPluginExecutionException("*****Error in position " + e + "***** ** Error in ADX_OpportunityPermissionsQueryExp query : ", ex);
             }
             catch (Exception ex)
             {
                 tracingService.Trace("*****Error in position " + e + "***** ** Error in ADX_OpportunityPermissionsQueryExp query : {0}", ex.ToString());
                 throw;
             }
        }
    }
}