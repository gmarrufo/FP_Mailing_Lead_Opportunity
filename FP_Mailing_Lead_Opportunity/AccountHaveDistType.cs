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
    public class AccountHaveDistType
    {
        DataCollection<Entity> entityCollection = null;

        public string Execute(IServiceProvider serviceProvider, Guid gLeadSource, Guid gLeadRoutingType, Guid gAccount, DateTime dToday)
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
            int e = 1;
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
                                new ConditionExpression("pearl_leadsource", ConditionOperator.Equal, gLeadSource),
                                new ConditionExpression("pearl_leadtype",ConditionOperator.Equal,gLeadRoutingType),
                                new ConditionExpression("pearl_allowedleaddistributionid",ConditionOperator.Equal,gAccount)
                            },
                        }
                    }
                    }
                };
                e = 2;
                entityCollection = service.RetrieveMultiple(query).Entities;
                string authorized = "X";
                pearl_leaddistribution ld = null;
                e = 3;
                foreach (pearl_leaddistribution plD in entityCollection)
                {
                    e = 4;
                    if (ld == null)
                    {
                        e = 5;
                        if (plD.pearl_StartDate == null && plD.pearl_EndDate == null)
                        {
                            e = 6;
                            authorized = "Y";
                            ld = plD;
                        }
                        else if (plD.pearl_StartDate == null && plD.pearl_EndDate != null)
                        {
                            e = 7;
                            // (DATE NOW BEFORE END == Good / DATE NOW AFTER ED == Bad)
                            if (dToday < plD.pearl_EndDate)
                            {
                                e = 8;
                                authorized = "Y";
                                ld = plD;
                            }
                        }
                        else if (plD.pearl_StartDate != null && plD.pearl_EndDate == null)
                        {
                            e = 9;
                            // (DATE NOW AFTER SD == GOOD / DATE NOW BEFORE SD == BAD)
                            if (dToday > plD.pearl_StartDate)
                            {
                                e = 10;
                                authorized = "Y";
                                ld = plD;
                            }
                        }
                        else
                        {
                            e = 11;
                            // DEPENDS ON DATE NOW
                            if ((plD.pearl_StartDate < dToday) && (dToday < plD.pearl_EndDate))
                            {
                                e = 12;
                                authorized = "Y";
                                ld = plD;
                            }
                        }
                    }
                    else
                        authorized = "N";
                    e = 13;
                }
                e = 14;
                if (ld != null)
                    if ((bool)ld.pearl_SpecialProcessingRules)
                        authorized = "S";

                e = 15;
                return authorized;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("*****Error in position " + e + "***** ** Error in AccountHaveDistType: ", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("*****Error in position " + e + "***** ** Error in AccountHaveDistType : {0}", ex.ToString());
                throw;
            }
        }
    }
}
