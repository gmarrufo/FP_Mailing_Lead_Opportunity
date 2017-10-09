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
    public class AllLeadSourceEqualOpportunity
    {
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, Guid sLeadSource, Guid sLeadRoutingType)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, Lead_Source, Start_Date, End_Date, Special_Processing_Rules, Lead_Routing_Type
            from
                Lead_Distribution a
                    inner join
                Lead_Routing_Type b
                    on a.Lead_Routing_Type = b.ID
                    inner join 
                Opportunity c
                    on b.ID = c.Routing_Type
            where 
            (
	            c.Lead_Source = "incoming string"
            )
            */

        //    QueryExpression query = new QueryExpression()
        //    {
        //        Distinct = false,
        //        EntityName = pearl_leaddistribution.EntityLogicalName,
        //        ColumnSet = new ColumnSet("ID", "Lead_Source", "Start_Date", "End_Date", "Special_Processing_Rules", "Lead_Routing_Type"),
        //        LinkEntities = 
        //{
        //    new LinkEntity 
        //    {
        //        JoinOperator = JoinOperator.Inner,
        //        LinkFromAttributeName = "Lead_Routing_Type",
        //        LinkFromEntityName = pearl_leaddistribution.EntityLogicalName,
        //        LinkToAttributeName = "ID",
        //        LinkToEntityName = pearl_leadroutingtype.EntityLogicalName
        //    },
        //    new LinkEntity
        //    {
        //        JoinOperator = JoinOperator.Inner,
        //        LinkFromAttributeName = "Routing_Type",
        //        LinkFromEntityName = pearl_leadroutingtype.EntityLogicalName,
        //        LinkToAttributeName = "ID",
        //        LinkToEntityName = Opportunity.EntityLogicalName,
        //    }
        //},
        //        Criteria =
        //        {
        //            Filters = 
        //    {
        //        new FilterExpression
        //        {
        //            Conditions = 
        //            {
        //                new ConditionExpression("Lead_Source", ConditionOperator.Equal, sLeadSource)
        //            },
        //        }
        //    }
        //        }
        //    };

            try
            {
                QueryExpression query = new QueryExpression()
                {
                    Distinct = false,
                    EntityName = pearl_leaddistribution.EntityLogicalName,
                    ColumnSet = new ColumnSet(true),
                    LinkEntities = 
                {
                    new LinkEntity 
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "pearl_leadsource",
                        LinkFromEntityName = pearl_leaddistribution.EntityLogicalName,
                        LinkToAttributeName = "pearl_leadsourceid",
                        LinkToEntityName = pearl_leadsource.EntityLogicalName,// pearl_leadroutingtype.EntityLogicalName
                    },
                    new LinkEntity
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "pearl_leadtype",
                        LinkFromEntityName = pearl_leaddistribution.EntityLogicalName,
                        LinkToAttributeName = "pearl_leadroutingtypeid",
                        LinkToEntityName = pearl_leadroutingtype.EntityLogicalName, //Opportunity.EntityLogicalName,
                    }
                },
                    Criteria =
                    {
                        Filters = 
                    {
                        new FilterExpression
                        {
                            Conditions = 
                            {
                                new ConditionExpression("pearl_leadsource", ConditionOperator.Equal, sLeadSource),
                                new ConditionExpression("pearl_leadtype",ConditionOperator.Equal,sLeadRoutingType)
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
                throw new InvalidPluginExecutionException("** Error in ALSEO ** An error occurred in the AllLeadSourceEqualOpportunity query.", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("** Error in ALSEO ** AllLeadSourceEqualOpportunity query: {0}", ex.ToString());
                throw;
            }
        }
    }
}
