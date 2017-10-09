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
    public class AllAccountsInTerritory
    {
        DataCollection<Entity> entityCollection = null;
        DataCollection<Entity> entityCollection2 = null;//{ get; set; }// = null;
       // DataCollection<Entity> entityCollection3 { get; set; }

        public List<Account> Execute(IServiceProvider serviceProvider, string sZip)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, Contacts, Name, Counties, Last Opportunity Assignment, Distribution Rules
            from
                Account a
                    inner join
                Counties b
                    on a.pearl_account_pearl_county = b.pearl_county
                    inner join 
                ZipCodes c
                    on b.pearl_county_pearl_zipcode = c.pearl_zipcode
            where 
            (
	            a.Name = "incoming string"
            )
                */
            /*
            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Contacts", "Name", "Counties", "Last Opportunity Assignment", "Distribution Rules"),
                LinkEntities = 
                {
                    new LinkEntity 
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "a.pearl_account_pearl_county",
                        LinkFromEntityName = Account.EntityLogicalName,
                        LinkToAttributeName = "b.pearl_county",
                        LinkToEntityName = pearl_county.EntityLogicalName
                    },
                    new LinkEntity
                    {
                        JoinOperator = JoinOperator.Inner,
                        LinkFromAttributeName = "b.pearl_county_pearl_zipcode",
                        LinkFromEntityName = pearl_county.EntityLogicalName,
                        LinkToAttributeName = "c.pearl_zipcode",
                        LinkToEntityName = pearl_zipcode.EntityLogicalName,
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
                                new ConditionExpression("name", ConditionOperator.Equal, sName)
                            },
                        }
                    }
                }
            };
            */

            //bool firsttime = true;
            string eText = "";
            int e = 1;
            
            /*IList<Entity> list = new IList<Entity>();
            DataCollection<Entity> entityCollection3 = new DataCollection<Entity>(Entity);*/
            try
            {
                //DataCollection<Entity> entityCollection3 = (DataCollection<Entity>);

                e++; string entity1 = pearl_zipcode.EntityLogicalName;
                e++; string entity2 = pearl_county.EntityLogicalName;
                e++; string relationshipname = pearl_county_pearl_zipcode.EntityLogicalName;
                e++; QueryExpression q = new QueryExpression(entity2);
                e++; q.ColumnSet = new ColumnSet(true);
                e++; LinkEntity linkEntity1 = new LinkEntity(entity2, relationshipname, "pearl_countyid", "pearl_countyid", JoinOperator.Inner);
                e++; LinkEntity linkEntity2 = new LinkEntity(relationshipname, entity1, "pearl_zipcodeid", "pearl_zipcodeid", JoinOperator.Inner);
                e++; linkEntity1.LinkEntities.Add(linkEntity2);
                e++; q.LinkEntities.Add(linkEntity1);
                e++; linkEntity2.LinkCriteria = new FilterExpression();
                e++; linkEntity2.LinkCriteria.AddCondition(new ConditionExpression("pearl_zipcode", ConditionOperator.Equal, sZip));

                e++; entityCollection = service.RetrieveMultiple(q).Entities;
                List<Account> alist = new List<Account>();
                eText += "first query";
                if (entityCollection.Count > 0)
                    foreach (pearl_county co in entityCollection)
                    {

                        eText += " co: " + co.pearl_name;
                        e = 50;
                        e++; entity1 = pearl_county.EntityLogicalName;
                        e++; entity2 = Account.EntityLogicalName;
                        e++; relationshipname = pearl_account_pearl_county.EntityLogicalName;
                        e++; QueryExpression q2 = new QueryExpression(entity2);
                        e++; q2.ColumnSet = new ColumnSet(true);
                        e++; linkEntity1 = new LinkEntity(entity2, relationshipname, "accountid", "accountid", JoinOperator.Inner);
                        e++; linkEntity2 = new LinkEntity(relationshipname, entity1, "pearl_countyid", "pearl_countyid", JoinOperator.Inner);
                        e++; linkEntity1.LinkEntities.Add(linkEntity2);
                        e++; q2.LinkEntities.Add(linkEntity1);
                        e++; linkEntity2.LinkCriteria = new FilterExpression();
                        e++; linkEntity2.LinkCriteria.AddCondition(new ConditionExpression("pearl_countyid", ConditionOperator.Equal, co.pearl_countyId));

                        entityCollection2 = null;
                        //62
                        e++; entityCollection2 = service.RetrieveMultiple(q2).Entities;
                        //63
                        e++;
                        eText += " actCount: " + entityCollection2.Count;
                        if (entityCollection2.Count > 0)
                            foreach (Account act in entityCollection2)
                            {
                                if (act != null)
                                {
                                    e = 70;
                                    e++; eText += " act: " + act.Name;
                                    //if (firsttime)
                                    //entityCollection3.Add(act);
                                    //else 
                                    //if (!entityCollection3.Contains(act))
                                    //{
                                    e++;
                                    bool found = false;
                                    if (alist != null)
                                    {
                                        e = 80;
                                        foreach (Account actCk in alist)
                                        {
                                            found = (actCk.Name == act.Name);
                                            //if (found) break;
                                        }
                                    }
                                    e = 90;
                                    if (!found)
                                        alist.Add(act);
                                    
                                }
                            }
                    }
//Entity retDistribution = service.Retrieve(pearl_leaddistribution.EntityLogicalName, new Guid(), new ColumnSet(true));
                e = 198;
                //entityCollection3.Add(aList);
                e++; return alist;


                /*
                QueryExpression query = new QueryExpression()
                {
                    Distinct = false,
                    EntityName = entity2.EntityLogicalName,
                    ColumnSet = new ColumnSet("pearl_countyid"),
                    LinkEntities =
                    {
                       new LinkEntity
                       {
                           JoinOperator= JoinOperator.Inner,
                           LinkFromAttributeName = "pearl_county_pearl_zipcode",
                           LinkFromEntityName = entity2.EntityLogicalName,
                           LinkToAttributeName = "pearl_zipcode",
                           LinkToEntityName = pearl_zipcode.EntityLogicalName
                       }
                    },
                    Criteria=
                    {
                        Filters= 
                        {
                            new FilterExpression
                            {
                                Conditions=
                                {
                                    new ConditionExpression("pearl_zipcode", ConditionOperator.Equal,sZip)
                                },
                            }
                        }
                    }
                };
                entityCollection = service.RetrieveMultiple(query).Entities;
            
                foreach (pearl_county county in entityCollection)
                {
                    QueryExpression query2 = new QueryExpression()
                    {
                        Distinct = false,
                        EntityName = Account.EntityLogicalName,
                        ColumnSet = new ColumnSet(true),
                        LinkEntities =
                    {
                       new LinkEntity
                       {
                           JoinOperator= JoinOperator.Inner,
                           LinkFromAttributeName = "pearl_account_pearl_county",
                           LinkFromEntityName = Account.EntityLogicalName,
                           LinkToAttributeName = "pearl_county",
                           LinkToEntityName = entity2.EntityLogicalName
                       }
                    },
                        Criteria =
                        {
                            Filters = 
                        {
                            new FilterExpression
                            {
                                Conditions=
                                {
                                    new ConditionExpression("pearl_countyid", ConditionOperator.Equal,county.pearl_countyId)
                                },
                            }
                        }
                        }
                    };
                    entityCollection2 = service.RetrieveMultiple(query2).Entities;
                    foreach (Entity en in entityCollection2)
                    {
                        entityCollection3.Add(en);
                    }
                }

                return entityCollection3;*/
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("** LOCATION: " + e + " etext " + eText + " ** Error in AllAccountsInTerritory query.", ex);
            }
            catch (Exception ex)
            {
                tracingService.Trace("** LOCATION: " + e + " etext " + eText + " ** Error in AllAccountsInTerritory query: {0}", ex.ToString());
                throw;
            }
        }
    }
}
