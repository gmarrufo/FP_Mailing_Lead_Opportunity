// <copyright file="PostLeadCreate.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>11/1/2015 12:23:59 PM</date>
// <summary>Implements the PostLeadCreate Process to fulfill extra requirements after assignment but prior to opportunity state.</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;

namespace FPMailingLeadOpportunity
{
    public class PostLeadCreate : IPlugin
    {
        double dDateDiff = 0.0;
        double dDateDiffCompare = 0.0;
        TimeSpan t;
        readonly DateTime dToday = DateTime.Now;
        private ITracingService tracingService;

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                // Opportunity Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);
                int e = 500;
                string eText = "";
                try
                {
                    #region STEP 6 - Business Rules & Assign Sales Channel
                    e = 601;/*************   Code Position 601   *****************/
                    /* Lead – Sales Channel Assignment, process based on following criteria:

                    Identifies which product type (Seg A, Seg B, F/I and OEM) the lead is interested in and which Sales Channel it gets assigned to.

                    MAP? Yes – Go to Step 7. This includes all Servicing Dealers that are MAP accounts. 
                    Retention? Yes – Go to Step 7. If no, it will go to existing customer check.
                    Cancellation? Check if existing customer? Yes – It will go to the Servicing Dealer. If Servicing Dealer is House (Between 4000 and 4999), it will go to Cancellations Team. If not, it will go to the Dealer.  

                    */

                    // Check if Object Lead is MAP == TRUE, if equal then ASSIGN LEAD, otherwise
                    // Check if Object Lead is Retention == TRUE, if equal then ASSIGN LEAD, otherwise
                    // Check if Object Lead is Cancellation == TRUE, if equal then LEAD ASSIGNED TO CANCELLATIONS TEAM, otherwise
                    // Check if Object Lead is Existing Customer == TRUE, if equal then Check if Servicing Dealer Criteria is met == TRUE, if equal then LEAD ASSIGNED TO CANCELLATIONS TEAM, otherwise LEAD ASSIGNED TO DEALER.

                    // Obtain the Lead Entity Object based on Opportunity Entity Originating Lead Id
                    /**                     
                     * map custom fields from lead to opportunity
                     */
                    e = 660;
                    e++; Entity leadQuery = service.Retrieve("lead", ((EntityReference)entity.Attributes["originatingleadid"]).Id, new ColumnSet(true));
                    e++; Lead foundLead = leadQuery.ToEntity<Lead>();
                    e++; eText = foundLead.Id.ToString();
                    

                    e++; PearlAnyTwoKeyMapping pA2KM = new PearlAnyTwoKeyMapping();
                    e++; DataCollection<Entity> L2OMapping = pA2KM.Execute(serviceProvider, 4, "lead", "opportunity");

                    e++; eText = "Number of mappings: "+L2OMapping.Count().ToString() + "\r\n";
                    foreach (pearl_anytwokey a2k in L2OMapping)
                    {
                        eText += a2k.pearl_Text1 + " " + a2k.pearl_Text2 + " " + a2k.pearl_Text3 + " " + a2k.pearl_Text4 + "\r\n";
                        if (leadQuery.Attributes.Contains(a2k.pearl_Text2))
                            if (leadQuery.Attributes[a2k.pearl_Text2] != null)
                                entity.Attributes[a2k.pearl_Text4]= leadQuery.Attributes[a2k.pearl_Text2];
                    }

                    e = 652;

                    //QueryByAttribute queryLeadEntity = new QueryByAttribute
                    //{
                    //    EntityName = "lead",
                    //    ColumnSet = new ColumnSet(true)
                    //};
                    //e = 602;/*************   Code Position 602   *****************/
                    //queryLeadEntity.AddAttributeValue("lead", entity.Attributes["originatingleadid"]);
                    //EntityCollection entResultLeadEntity = service.RetrieveMultiple(queryLeadEntity);
                    //e = 603;/*************   Code Position 603   *****************/
                    //if (checkLeadPos(entResultLeadEntity))
                    //{
                    //    e = 6101;/*************   Code Position 6101   *****************/
                    //     /*  Check for MAP */
                    //    if (leadQuery.Attributes["pearl_saleschannel"].Equals("MAP"))
                    //    {
                    //        entity.Attributes["pearl_leadroutingtype"] = "MAP";
                    //    }
                    //    /*  Check for Retention */
                    //    else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Retention"))
                    //    {
                    //        entity.Attributes["pearl_leadroutingtype"] = "Retention";
                    //    }
                    //    /*  Check for Cancellations */
                    //    else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Cancellations"))
                    //    {
                    //        entity.Attributes["pearl_leadroutingtype"] = "Cancellations";
                    //    }
                    //    else
                    //    {
                    //         /*  Check for EXISTING CUSTOMER */
                    //        if((bool)entResultLeadEntity[0].Attributes["pearl_existingcustomer"])
                    //        {
                    //            /*  Check for SERVICING DEALER CRITERIA */
                    //            if(entResultLeadEntity[0].Attributes["new_servicingsalesperson"].Equals("")) // TO DO - Check with Alex if this is still valid ????
                    //            {
                    //                entity.Attributes["pearl_leadroutingtype"] = "Cancellations";
                    //            }
                    //            else
                    //            {
                    //                entity.Attributes["pearl_leadroutingtype"] = "Dealer";
                    //            }
                    //        }
                    //    }
                    //    e = 6102;/*************   Code Position 6102   *****************/
                    //    /* New Sales Product Type Distribution, process based on following criteria:

                    //    To determine product type
                    //    •	If meter, determine if A or B segment
                    //    •	If NOT meter, determine if F/I or OEM
                    //    •	If NO product type is given, A segment is default

                    //    */

                    //    // Check if Object Lead is meter == TRUE, if equal then check if Object Lead is segment == A || == B, if equal then assign product type = meter, otherwise check if Object Lead is product type == F/I || == OEM,
                    //    // if equal then assign product type == to F/I or OEM, otherwise check if Object Lead is product type == “”, if equal then assign product type = meter, segment = A.

                    //    /*  Check for NEW CURRENT METER */
                    //    if(entResultLeadEntity[0].Attributes["new_currentmeter"].Equals(""))
                    //    {
                    //        /*  Check for SEGMENT NOT A or NOT B */
                    //        if(!entResultLeadEntity[0].Attributes["pearl_producttype"].Equals("Seg A") || !entResultLeadEntity[0].Attributes["pearl_producttype"].Equals("Seg B"))
                    //        {
                    //            entity.Attributes["pearl_leadtype"] = "Seg A";
                    //        }
                    //    }
                    //}
                    e = 604;/*************   Code Position 604   *****************/
                    eText = "";
                    #endregion

                    #region STEP 7 - Assign Lead
                    
                    /* Process based on following criteria: */

                    // NEW PROCESS

                    /*
                    Opportunity to Lead Routing Type to Lead Distribution to get All Lead Distributions that have same Lead Source as the Opportunity
                    Opportunity to Lead Source to get All Lead Source and compare against view above with a Start Date and End Date.
                    From Zip Code to Counties to Account to get All Account by Zip Code and County and compare against view result from above.
                    All accounts available for Distribution.
                    Round Robin in Account and Round Robin in the Contact to Assign the LEAD
                    IF NOTHING IS FOUND:
                    Opportunity to Lead Routing Type to Account to get Default Account Information
                    Round Robin for the Contact associated with Default Contact Above.
                    IF NO CONTACTS AVAILABLE THEN ASSIGN TO DEFAULT CONTACT in the Lead Routing Type.
                    */
                    e = 721;/*************   Code Position 701   *****************/
                    // Obtain Lead Source Name
                    bool sendToDefaultAct = false;
                    Account distAccount = null;
                    Account servicingDealer = null;
                    Guid _actRouting =  new Guid();
                    Guid _conRouting = new Guid();
                    Contact distContact = null;
                    pearl_leadroutingtype leadRoutingType = null;

                    bool specialProcess = false;

                    
                    if (entity.Attributes.Contains("pearl_leadsource"))//(context.InputParameters.Contains("pearl_leadsource"))
                    {
                        e = 7222;
                        Entity retLeadSource = service.Retrieve(pearl_leadsource.EntityLogicalName, ((EntityReference)entity.Attributes["pearl_leadsource"]).Id, new ColumnSet(true));
                    }
                    if (entity.Attributes.Contains("pearl_routingtype"))//(context.InputParameters.Contains("pearl_routingtype"))
                    {
                        e = 7333;
                        Entity retLeadRouting = service.Retrieve(pearl_leadroutingtype.EntityLogicalName, ((EntityReference)entity.Attributes["pearl_routingtype"]).Id, new ColumnSet(true));
                        leadRoutingType = retLeadRouting.ToEntity<pearl_leadroutingtype>();
                        e = 7334;
                        _actRouting = leadRoutingType.pearl_DefaultAccount.Id;
                        _conRouting = leadRoutingType.pearl_DefaultContact.Id;
                        e = 7335;
                        eText = "act routing " + _actRouting + " conRouting " + _conRouting + " ";
                    }
                    
                    if (foundLead.Address1_PostalCode != null)
                    {
                        e = 722;
                        string postCode = foundLead.Address1_PostalCode.Substring(0, 5);
                        eText += " postcode " + postCode + " ";
                        AllAccountsInTerritory aait = new AllAccountsInTerritory();
                        AccountHaveDistType ahdt = new AccountHaveDistType();
                        e = 723;
                        List<Account> listAait = aait.Execute(serviceProvider, postCode);
                        
                        e = 724;
                        eText += " list count: " + listAait.Count;
                        eText += "allowed for dist: \r\n";
                        if (listAait.Count > 0)
                        {
                            foreach (Account act in listAait)
                            {
                                bool first = true;
                                e = 725;
                                //check to see if they hace proper distribution
                                //if they do, add to list
                                string result = ahdt.Execute(serviceProvider,
                                    ((EntityReference)entity.Attributes["pearl_leadsource"]).Id,
                                    ((EntityReference)entity.Attributes["pearl_routingtype"]).Id,
                                    act.Id,
                                    dToday);
                                if (result == "Y" || result == "S")
                                {
                                    e = 726;
                                    t = dToday - (DateTime)act.pearl_LastOpportunityAssignment;
                                    dDateDiffCompare = t.TotalMilliseconds;
                                    e = 727;

                                    if (dDateDiffCompare > dDateDiff || first)
                                    {
                                        e = 728;
                                        if (result == "S")
                                            specialProcess = true;
                                        else
                                            specialProcess = false;
                                        dDateDiff = dDateDiffCompare;
                                        distAccount = null;
                                        distAccount = act;
                                        eText += " account set: " + act.Name;
                                        first = true;
                                    }
                                }
                            }
                        }
                        if (distAccount == null)
                            sendToDefaultAct = true;
                    } //end postcode != null
                    else                    
                        sendToDefaultAct = true;

                    e = 730;
                    if (sendToDefaultAct)
                    {
                        Entity retDistAccount = service.Retrieve(Account.EntityLogicalName, _actRouting, new ColumnSet(true));
                        distAccount = retDistAccount.ToEntity<Account>();
                        servicingDealer = distAccount;
                    }

                    if (specialProcess)
                    {
                        servicingDealer = distAccount;
                        Entity retDistAccount = service.Retrieve(Account.EntityLogicalName, _actRouting, new ColumnSet(true));
                        distAccount = retDistAccount.ToEntity<Account>();
                    }
                    else
                        servicingDealer = distAccount;
                    e = 733;
                    AllContactsDefaultAccount aCLA = new AllContactsDefaultAccount();
                    DataCollection<Entity> entityCollection6 = aCLA.Execute(serviceProvider, distAccount.ToEntityReference());
                    e = 1705;/*************   Code Position 1705   *****************/
                    // Round Robin in Contact
                    dDateDiff = 0.0;
                    dDateDiffCompare = 0.0;
                    e = 734;
                    foreach (Contact cT in entityCollection6)
                    {

                        t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                        dDateDiffCompare = t.TotalDays;

                        if (dDateDiffCompare > dDateDiff)
                        {
                            dDateDiff = dDateDiffCompare;
                            distContact = cT;
                        }
                    }
                    e = 735;
                    if (distContact == null)
                    {
                        Entity retDistContact = service.Retrieve(Contact.EntityLogicalName, _conRouting, new ColumnSet(true));
                        distContact = retDistContact.ToEntity<Contact>();  
                    }
                    e = 736;
                    /************THIS CAUSES AND ERROR FOR TESTING*************/
                    //Entity retDistribution = service.Retrieve(pearl_leaddistribution.EntityLogicalName, new Guid(), new ColumnSet(true));


                    e = 737;
                    if (distContact != null)
                    {
                        e = 738;

                        if (distContact.Attributes.Contains("adx_systemuserid"))// != null)
                        {
                            if (distContact.Attributes["adx_systemuserid"] != null)
                                entity.Attributes["ownerid"] = distContact.Attributes["adx_systemuserid"];
                        }

                        //e++; distContact.Attributes["pearl_assignlead"] = true;
                        e++; distContact.pearl_LastOpportunityAssignment = dToday;// distContact.Attributes["pearl_lastopportunityassignment"] = dToday;
                        /*40*/e++; distAccount.pearl_LastOpportunityAssignment = dToday;//distAccount.Attributes["pearl_lastopportunityassignment"] = dToday;
                        e++; servicingDealer.pearl_LastOpportunityAssignment = dToday;
                        e++; entity.Attributes["msa_partnerid"] = (EntityReference)distContact.Attributes["parentcustomerid"];
                        e++; entity.Attributes["msa_partneroppid"] = distContact.ToEntityReference();
                        //todo add parter settings for opportunity


                        /*44*/e++; entity.Attributes["pearl_sellingdealer"] = distAccount.ToEntityReference();
                        /*45*/e++; entity.Attributes["pearl_servicingdealer"] = servicingDealer.ToEntityReference();


                        //todo add permissions to view opportunites on the portal
                        adx_opportunitypermissions oppPerm = new adx_opportunitypermissions();

                        ExecuteMultipleRequest multipleRequest = new ExecuteMultipleRequest()
                        {
                            Settings = new ExecuteMultipleSettings()
                            {
                                ContinueOnError=true,
                                ReturnResponses=true
                            },
                            Requests = new OrganizationRequestCollection()
                        };
                        e++; UpdateRequest updateRequest = new UpdateRequest { Target = distContact };
                        e++; multipleRequest.Requests.Add(updateRequest);
                        e++; updateRequest = new UpdateRequest { Target = distAccount };
                        e++; multipleRequest.Requests.Add(updateRequest);
                        e++; updateRequest = new UpdateRequest { Target = servicingDealer };
                        e++; multipleRequest.Requests.Add(updateRequest);
                        e++; service.Execute(multipleRequest);
                    }
                    e = 750; 
                    /*****************************************************************************************************************************************
                     * ****************************************************************************************************************************************
                     * ****************************************************************************************************************************************/


                    /*  
                    Afterwards, the lead will go to the primary contact at the enabled Dealer with a link to the 
                    opportunity information. The Dealer will have to login to view info in Portal. From Portal,
                    they will be able to view full copy of the opportunity and acknowledge they received the lead,
                    edit content, qualify (criteria) and complete order.

                    RSM’s get a full copy of the opportunity via email with a link the opportunity. The RSM has
                    the ability to view info in CRM, edit content and qualify (criteria). 

                                    RSM                         DEALER
                    EMAIL	        Full opportunity Info	    View opportunity link only
                    ABILITIES	    View info in CRM            Edit Content
                                    Qualify (criteria)          View full opportunity info in Portal
                                    Edit Content                Qualify (criteria)
                                                                Complete Order

                    Sales Channel Default Distribution

                    If no email address is on file for sales channel, send Opportunity to corresponding manager below.
                    MAP – emails will go to MAP Manager
                    Retention - emails will go to Cancellations Manager
                    Cancellation - emails will go to Cancellations Manager
                    Inside Sales -- emails will go to Inside Sales Manager
                    Dealers – emails will go to Marketing (marketing@fp-usa.com) 
                    */

                    // TO DO - Email Variables used
                    string sFirstName = null;
                    string sLastName = null;
                    string sEmailAddress = null;
                    string sSubject = null;
                    string sDescription = null;
                    string sOpportunityInformation = null;
                    e = 999;
                    // TO DO - If no email address is on file for sales channel, send Opportunity to corresponding manager below.
                    // TO DO - Which is the object that contains the email address for sales channel????
                    int i = 0;
                    if(i > 1)
                    {
                        sOpportunityInformation = "URL Link TBD Later"; // TO DO - Obtain information
                        SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                    }
                    else
                    {
                         /*  Check for MAP */
                        /*if (leadQuery.Attributes["pearl_saleschannel"].Equals("MAP"))
                        {
                            sFirstName = "MAP Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Fitzpatrick";
                            sEmailAddress = "dfitzpatrick@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Retention"))
                        {
                            sFirstName = "Retention Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Hannon";
                            sEmailAddress = "mhannon@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Cancellations"))
                        {
                            sFirstName = "Cancellation Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Hannon";
                            sEmailAddress = "mhannon@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Inside Sales"))
                        {
                            sFirstName = "Inside Sales Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Charatin";
                            sEmailAddress = "dcharatin@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (leadQuery.Attributes["pearl_saleschannel"].Equals("Dealer"))
                        {
                            sFirstName = "Dealers Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Thompson";
                            sEmailAddress = "marketing@fp-usa.com;kthompson@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }*/
                    }

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("** Error in code position: "+ e + " ** || *** eText: " + eText + " *** An error occurred in the InsertLeadProcess plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("** Error in code position: " + e + " ** || *** eText: " + eText + " *** InsertLeadProcess: {0}", ex.ToString());
                    throw;
                }
            }
        }

        public Boolean checkLead(EntityCollection ent)
        {
            Boolean bolResult = false;
            if (ent.Entities.Count <= 0)
            {
                bolResult = true;
            }
            return bolResult;
        }

        public Boolean checkLeadPos(EntityCollection ent)
        {
            Boolean bolResult = false;
            if (ent.Entities.Count >= 0)
            {
                bolResult = true;
            }
            return bolResult;
        }

        public void SendEmail(ITracingService tracingService, string sFirstName, string sLastName, string sEmailAddress, string sSubject, string sDescription, string sOpportunityInformation)
        {
            try
            {
                // Obtain the target organization's Web address and client logon 
                // credentials from the user.
                ServerConnection serverConnect = new ServerConnection();
                ServerConnection.Configuration config = serverConnect.GetServerConfiguration();

                SendEmail app = new SendEmail();
                app.Run(config, tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);

            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace("Timestamp: {0}", ex.Detail.Timestamp);
                tracingService.Trace("Code: {0}", ex.Detail.ErrorCode);
                tracingService.Trace("Message: {0}", ex.Detail.Message);
                tracingService.Trace("Plugin Trace: {0}", ex.Detail.TraceText);
                tracingService.Trace("Inner Fault: {0}",
                    null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
            }
            catch (System.TimeoutException ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace("Message: {0}", ex.Message);
                tracingService.Trace("Stack Trace: {0}", ex.StackTrace);
                tracingService.Trace("Inner Fault: {0}",
                    null == ex.InnerException.Message ? "No Inner Fault" : ex.InnerException.Message);
            }
            catch (System.Exception ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace(ex.Message);

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    tracingService.Trace(ex.InnerException.Message);

                    FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe =
                        ex.InnerException
                        as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                    if (fe != null)
                    {
                        tracingService.Trace("Timestamp: {0}", fe.Detail.Timestamp);
                        tracingService.Trace("Code: {0}", fe.Detail.ErrorCode);
                        tracingService.Trace("Message: {0}", fe.Detail.Message);
                        tracingService.Trace("Plugin Trace: {0}", fe.Detail.TraceText);
                        tracingService.Trace("Inner Fault: {0}",
                            null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
                    }
                }
            }
            // Additonal exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
            // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.

            finally
            {
                // Any other code to finalize process if needed.
            }
        }
    }
}