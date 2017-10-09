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
    public class RentalLineByAccount
    {
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sAccountName)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, new_dealcode, new_GenProdPostingGroup, new_DocumentNo, new_LineNo
            from
                Rental_Line a
                    inner join
                Rental_Header b
                    on a.new_rentalhead_to_rentallineid = b.new_rentalheaderid
                    inner join 
                Account c
                    on b.new_SelltoCustomerNo = c.new_stccustomerno
            where 
            (
	            c.new_AccountName = "incoming string" and
	            a.new_DocumentNo is not null and
	            a.new_LineNo <> 0 and
	            a.new_GenProdPostingGroup = "meter" and
	            (a.new_dealcode not like 'D%' or a.new_dealcode not like 'd%' or a.new_dealcode not like 'C%' or a.new_dealcode not like 'c%')
            )
            */
            int e = 150;
            string eText = "RLBA";
            try
            {
                QueryExpression query = new QueryExpression()
                {
                    Distinct = false,
                    EntityName = new_rentalline.EntityLogicalName,
                    ColumnSet = new ColumnSet(true),
                    LinkEntities = 
                    {
                        new LinkEntity
                        {
                            JoinOperator = JoinOperator.Inner,
                            LinkFromAttributeName = "new_rentalhead_to_rentallineid",
                            LinkFromEntityName = new_rentalline.EntityLogicalName,
                            LinkToAttributeName = "new_rentalheaderid",
                            LinkToEntityName = new_rentalheader.EntityLogicalName
                        },
                        new LinkEntity
                        {
                            JoinOperator = JoinOperator.Inner,
                            LinkFromAttributeName = "new_selltocustno",
                            LinkFromEntityName = new_rentalheader.EntityLogicalName,
                            LinkToAttributeName = "name",
                            LinkToEntityName = Account.EntityLogicalName,
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
                                    new ConditionExpression("c.name", ConditionOperator.Equal, sAccountName),
                                    new ConditionExpression("a.new_documentno", ConditionOperator.NotNull),
                                    new ConditionExpression("a.new_lineno", ConditionOperator.NotEqual, 0),
                                    new ConditionExpression("a.new_genprodpostinggroup", ConditionOperator.Equal, "meter")
                                },
                            },
                            new FilterExpression
                            {
                                FilterOperator = LogicalOperator.Or,
                                Conditions = 
                                {
                                    new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "D%"),
                                    new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "d%"),
                                    new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "C%"),
                                    new ConditionExpression("a.new_dealcode", ConditionOperator.NotLike, "c%")
                                },
                            }
                        }
                    }
                };
                e = 151;
                try
                {
                    entityCollection = service.RetrieveMultiple(query).Entities;
                }
                catch { }
                e = 152;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("** Error in position: " + e + " ** ***Etext: " + eText + "*** An error occurred in the RentalLineByAccount query.", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("** Error in position: " + e + " ** ***Etext: " + eText + "*** RentalLineByAccount query: {0}", ex.ToString());
                throw;
            }
            return entityCollection;
        }
    }
}
