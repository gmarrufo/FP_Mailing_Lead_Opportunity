// <copyright file="InsertLeadProcess.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>04/12/16 10:18:00 AM</date>
// <summary>Implements the UpdateQualifyLead Process to determine qualification, non qualification among others to the raw lead prior to assignment.</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace FPMailingLeadOpportunity
{
    public class UpdateQualifyLead : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                // Lead Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                // Instantiate Organization Service Interfaces
                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);
                int e = 0;
                string eText = "";
                try
                {
                    #region STEP -1 Qualified & Not Qualified
                    //entity.Attributes["pearl_campaignid"] = "update";

                    var curQuery = new QueryExpression("organization");
                    curQuery.ColumnSet = new ColumnSet("basecurrencyid");
                    var curResult = service.RetrieveMultiple(curQuery);
                    var currencyId = (EntityReference)curResult.Entities[0]["basecurrencyid"];
                    Lead updLead = (service.Retrieve("lead", entity.Id, new ColumnSet(true))).ToEntity<Lead>();
                    e = 1;
                    bool qualify = false;
                    //updLead.pearl_CampaignID = "preQualify";

                    // GMC - 04/25/16 - To check for updLead Object exists
                    if (updLead != null)
                    {
                        e = 5;
                        if ((updLead.pearl_RoutingType != null &&
                            updLead.pearl_LeadType != null &&
                            (bool)updLead.pearl_AutoQualified))
                            qualify = true;
                        e = 10;
                        if ((updLead.pearl_RoutingType != null &&
                            updLead.pearl_LeadType != null &&
                            (bool)updLead.new_Qualified &&
                            updLead.new_NotQualifiedReason.Value == 0) &&
                            updLead.new_ContractEndDate != null)
                            qualify = true;
                    }

                    if (qualify)
                    //if ((bool)updLead.pearl_AutoQualified)
                    {
                        e = 3101; /************  Code Position #3101  ************/

                        var qlreq = new QualifyLeadRequest
                        {
                            CreateOpportunity = true,
                            OpportunityCurrencyId = currencyId,

                            // GMC - 04/15/16 - Change based on Alex Patrickus email
                            // CreateAccount = false,
                            CreateAccount = true,
                            CreateContact = true,
                            
                            //OpportunityCustomerId = null,
                            Status = new OptionSetValue(3),//.Qualified),                                
                            //Status = new OptionSetValue(1), 
                           
                            
                            // GMC - 04/25/16 - Change to avoid duplicate record issue
                            // LeadId = new EntityReference(updLead.LogicalName, updLead.Id)//entity.Id)//entity.ToEntityReference(),// ((EntityReference)entity.Id),
                            LeadId = updLead.ToEntityReference()
                        };
                        e = 3102; /************  Code Position #3104  ************/
                        //var qlres = (QualifyLeadResponse)
                        service.Execute(qlreq);
                        //service.Execute(qlreq);

                        e++; /************  Code Position #3103  ************/

                        //entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                        //entity.Attributes["statuscode"] = new OptionSetValue(3); //"Qualified because pearl_AutoQualified is true";
                    }

                    if (((bool)updLead.pearl_AutoQualified || (bool)updLead.new_Qualified) && !qualify)
                    {
                        e = 30;
                        updLead.pearl_LeadStatus = "No Lead Type or Routing Type";
                        //service.Update(updLead);
                        ExecuteMultipleRequest multipleRequest = new ExecuteMultipleRequest()
                        {
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError = true,
                                ReturnResponses = true
                            },
                            Requests = new OrganizationRequestCollection()
                        };
                        e++; UpdateRequest updateRequest = new UpdateRequest { Target = updLead };
                        e++; multipleRequest.Requests.Add(updateRequest);
                        e++; service.Execute(multipleRequest);
                    }

                    /*
                    if ((bool)updLead.new_Qualified && updLead.new_NotQualifiedReason.Value == null)
                    {
                        e = 40;
                        var qlreq = new QualifyLeadRequest
                        {
                            CreateOpportunity = false,
                            OpportunityCurrencyId = currencyId,
                            CreateAccount = false,
                            CreateContact = false,
                            //OpportunityCustomerId = null,
                            Status = new OptionSetValue(7),//.Qualified),                                
                            //Status = new OptionSetValue(1),    
                            //LeadId = new EntityReference(updLead.LogicalName, updLead.Id)//entity.Id)//entity.ToEntityReference(),// ((EntityReference)entity.Id),                                

                        };
                        e = 3102; 
                        //var qlres = (QualifyLeadResponse)
                        service.Execute(qlreq);
                    }*/
                    // /
                    //service.Update(updLead);

                    // Determine Qualified - depending on the pearl_autoqualified value

                    /* QUALIFIED */

                    // Obtain Raw Lead Object

                    //if (entity.Attributes["pearl_autoqualified"].ToString().Equals("Yes"))//context.InputParameters.Contains("pearl_autoqualified"))
                    //{
                    //if (entity.Contains("pearl_autoqualified"))
                    //{
                    //    if (Convert.ToBoolean(entity.Attributes["pearl_autoqualified"]))
                    //    {


                    //entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                    //entity.Attributes["statuscode"] = new OptionSetValue(3); //"Qualified";
                    ////entity.Attributes["pearl_campaignid"] = "plugin worked";
                    /*QualifyLeadResponse qualifyLeadResp;
                    QualifyLeadRequest qualifyLead;

                    qualifyLead = new QualifyLeadRequest
                    {
                        CreateOpportunity = true,
                        CreateAccount = false,
                        CreateContact = false,
                        Status = new OptionSetValue(1),
                        LeadId = new EntityReference("lead", entity.Id)
                    };
                    qualifyLeadResp = (QualifyLeadResponse)service.Execute(qualifyLead);*/
                    //    }
                    //}
                    //}
                    /*
                    if (context.InputParameters.Contains("pearl_autoqualified"))
                    {
                        //  Check for the pearl_autoqualified value 
                        if (entity.Attributes["pearl_autoqualified"].ToString().Equals("Yes"))
                        {

                            // Qualify the Incoming Lead 
                            entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                            entity.Attributes["statuscode"] = new OptionSetValue(3); //"Qualified";
                        }
                        else
                        {
                            entity.Attributes["new_numberoflocations"] = 66;
                        }
                    }*/

                    /*
                    if (context.Depth > 1)
                    {
                        tracingService.Trace("Plugin has called itself. Exiting.");
                        return;
                    }
                    */

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("** Error in position: " + e + " ** ***Etext: " + eText + "*** An error occurred in the UpdateQualifyLead plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("** Error in position: " + e + " ** ***Etext: " + eText + "*** UpdateQualifyLead : {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
