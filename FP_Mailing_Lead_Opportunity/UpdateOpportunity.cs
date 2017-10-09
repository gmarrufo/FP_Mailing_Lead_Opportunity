// <copyright file="PostLeadCreate.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>04/12/16 10:24:00 AM</date>
// <summary>Implements the CreateOpportunity Process to fulfill extra requirements after assignment but prior to opportunity state.</summary>

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
    public class UpdateOpportunity : IPlugin
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

                    #region STEP 7 - Assign Lead
                    if (entity.Attributes.Contains("pearl_redistribute"))
                        if ((bool)entity.Attributes["pearl_redistribute"])
                        {

                            Opportunity op = service.Retrieve(Opportunity.EntityLogicalName, entity.Id, new ColumnSet(true)).ToEntity<Opportunity>();
                            if (op.OriginatingLeadId == null)
                                return;
                            e++; Entity leadQuery = service.Retrieve("lead", op.OriginatingLeadId.Id, new ColumnSet(true));
                            e++; Lead foundLead = leadQuery.ToEntity<Lead>();
                            e++; eText = foundLead.Id.ToString();
                            
                            entity.Attributes["pearl_redistribute"] = false;
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
                            Guid _actRouting = new Guid();
                            Guid _conRouting = new Guid();
                            Contact distContact = null;
                            pearl_leadroutingtype leadRoutingType = null;

                            bool specialProcess = false;

                            if (op.pearl_LeadSource != null)
                            {
                                e = 7222;
                                Entity retLeadSource = service.Retrieve(pearl_leadsource.EntityLogicalName, op.pearl_LeadSource.Id, new ColumnSet(true));
                            }
                            if (op.pearl_RoutingType != null)
                            {
                                e = 7333;
                                Entity retLeadRouting = service.Retrieve(pearl_leadroutingtype.EntityLogicalName, op.pearl_RoutingType.Id, new ColumnSet(true));
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
                                            op.pearl_LeadSource.Id,
                                            op.pearl_RoutingType.Id,
                                            act.Id,
                                            dToday);
                                        if (result == "Y" || result == "S")
                                        {
                                            e = 726;
                                            eText += " account: " + act.Name + " ";
                                            if (act.pearl_LastOpportunityAssignment == null)
                                                act.pearl_LastOpportunityAssignment = dToday;
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
                            bool firstContact = true;
                            foreach (Contact cT in entityCollection6)
                            {

                                t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                dDateDiffCompare = t.TotalMilliseconds;

                                if (dDateDiffCompare > dDateDiff || firstContact)
                                {
                                    firstContact = false;
                                    dDateDiff = dDateDiffCompare;
                                    distContact = cT;
                                }
                            }

                            if (op.pearl_ExistingCustomer != null)
                            {
                                if (!(bool)op.pearl_ExistingCustomer && (bool)leadRoutingType.pearl_ManualRouting)
                                {
                                    distContact = service.Retrieve(Contact.EntityLogicalName, leadRoutingType.pearl_DefaultContact.Id, new ColumnSet(true)).ToEntity<Contact>();
                                    distAccount = service.Retrieve(Account.EntityLogicalName, leadRoutingType.pearl_DefaultAccount.Id, new ColumnSet(true)).ToEntity<Account>();
                                }
                                else if ((bool)op.pearl_ExistingCustomer)
                                {
                                    distAccount = service.Retrieve(Account.EntityLogicalName, op.ParentAccountId.Id, new ColumnSet(true)).ToEntity<Account>();
                                    DataCollection<Entity> entityCollection7 = aCLA.Execute(serviceProvider, distAccount.ToEntityReference());
                                    e = 17051;/*************   Code Position 1705   *****************/
                                    // Round Robin in Contact
                                    dDateDiff = 0.0;
                                    dDateDiffCompare = 0.0;
                                    e = 17052;
                                    firstContact = true;
                                    foreach (Contact cT in entityCollection7)
                                    {

                                        t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                        dDateDiffCompare = t.TotalMilliseconds;

                                        if (dDateDiffCompare > dDateDiff || firstContact)
                                        {
                                            firstContact = false;
                                            dDateDiff = dDateDiffCompare;
                                            distContact = cT;
                                        }
                                    }
                                }
                            }
                            e = 735;
                            if (distContact == null)
                            {
                                Account defaultAccount = service.Retrieve(Account.EntityLogicalName, _actRouting, new ColumnSet(true)).ToEntity<Account>();
                                DataCollection<Entity> entityCollection8 = aCLA.Execute(serviceProvider, defaultAccount.ToEntityReference());
                                e = 17051;/*************   Code Position 1705   *****************/
                                // Round Robin in Contact
                                dDateDiff = 0.0;
                                dDateDiffCompare = 0.0;
                                e = 17052;
                                firstContact = true;
                                foreach (Contact cT in entityCollection8)
                                {

                                    t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                    dDateDiffCompare = t.TotalMilliseconds;

                                    if (dDateDiffCompare > dDateDiff || firstContact)
                                    {
                                        firstContact = false;
                                        dDateDiff = dDateDiffCompare;
                                        distContact = cT;
                                    }
                                }
                                if (distContact == null)
                                {
                                    Entity retDistContact = service.Retrieve(Contact.EntityLogicalName, _conRouting, new ColumnSet(true));
                                    distContact = retDistContact.ToEntity<Contact>();
                                }
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

                                e++; distContact.pearl_LastOpportunityAssignment = dToday;
                                /*40*/
                                e++; distAccount.pearl_LastOpportunityAssignment = dToday;
                                e++; servicingDealer.pearl_LastOpportunityAssignment = dToday;
                                e++; entity.Attributes["msa_partnerid"] = (EntityReference)distContact.Attributes["parentcustomerid"];
                                e++; entity.Attributes["msa_partneroppid"] = distContact.ToEntityReference();
                                entity.Attributes["adx_delivereddate"] = DateTime.Now;
                                /*44*/
                                e++; entity.Attributes["pearl_sellingdealer"] = distAccount.ToEntityReference();
                                /*45*/
                                e++; entity.Attributes["pearl_servicingdealer"] = servicingDealer.ToEntityReference();
                                e++; entity.Attributes["statuscode"] = new OptionSetValue(10000001);
                                // todo add permissions to view opportunites on the portal

                                /*                        
                                Entity
                                adx_opportunitypermissions
                                Number of records = 2
                                Record 1
                                Name = distContact.fullname
                                Contact = distContact
                                Account = distContact.parentcustomerid
                                Scope = self  -- new OptionSet(100000000)
                                Read = true
                                Write = true
                                Create = true
                                Delete = false
                                Accept/decline = true
                                Assign = true
                         
                                Record 2
                                Name = distContact.fullname
                                Contact = distContact
                                Account = distContact.parentcustomerid
                                Scope = account -- new OptionSet(100000001)
                                Read = true
                                Write = true
                                Create = true
                                Delete = false
                                Accept/decline = true
                                Assign = true
                                */

                                ExecuteMultipleRequest multipleRequest = new ExecuteMultipleRequest()
                                {
                                    Settings = new ExecuteMultipleSettings()
                                    {
                                        ContinueOnError = true,
                                        ReturnResponses = true
                                    },
                                    Requests = new OrganizationRequestCollection()
                                };
                                CreateRequest createRequest = new CreateRequest();

                                ADX_OpportunityPermissionsQueryExp oppPerQ = new ADX_OpportunityPermissionsQueryExp();
                                DataCollection<Entity> entityCollection = oppPerQ.Execute(serviceProvider, distContact.Id, distAccount.Id, 100000000, 100000001);

                                adx_opportunitypermissions oppPerm1 = new adx_opportunitypermissions();
                                adx_opportunitypermissions oppPerm2 = new adx_opportunitypermissions();

                                if (entityCollection.Count <= 0)
                                {
                                    oppPerm1.adx_ContactId = distContact.ToEntityReference();
                                    oppPerm1.adx_AccountId = distAccount.ToEntityReference();
                                    oppPerm1.adx_name = distContact.FullName;
                                    oppPerm1.adx_Read = true;
                                    oppPerm1.adx_Write = true;
                                    oppPerm1.adx_Create = true;
                                    oppPerm1.adx_Delete = false;
                                    oppPerm1.adx_AcceptDecline = true;
                                    oppPerm1.adx_Assign = true;
                                    oppPerm1.adx_Scope = new OptionSetValue(100000000);
                                    createRequest = new CreateRequest { Target = oppPerm1 };
                                    multipleRequest.Requests.Add(createRequest);

                                    //oppPerm2.adx_ContactId = distContact.ToEntityReference();
                                    //oppPerm2.adx_AccountId = distAccount.ToEntityReference();
                                    //oppPerm2.adx_name = distContact.FullName;
                                    //oppPerm2.adx_Read = true;
                                    //oppPerm2.adx_Write = true;
                                    //oppPerm2.adx_Create = true;
                                    //oppPerm2.adx_Delete = false;
                                    //oppPerm2.adx_AcceptDecline = true;
                                    //oppPerm2.adx_Assign = true;
                                    //oppPerm2.adx_Scope = new OptionSetValue(100000001);
                                    //createRequest = new CreateRequest { Target = oppPerm2 };
                                    //multipleRequest.Requests.Add(createRequest);
                                }

                                /*
                                Entity
                                adx_accountaccess
                                Number of records = 1
                                Name=distContact.fullname
                                Account = distContact.parentcustomerid
                                Contact = distContact
                                Read = true
                                Write = true
                                Manage permissions = true
                                */

                                ADX_AccountAccessQueryExp accAccQ = new ADX_AccountAccessQueryExp();
                                entityCollection = accAccQ.Execute(serviceProvider, distContact.Id, distAccount.Id);

                                Adx_accountaccess accACC = new Adx_accountaccess();

                                if (entityCollection.Count <= 0)
                                {
                                    accACC.adx_contactid = distContact.ToEntityReference();
                                    accACC.adx_accountid = distAccount.ToEntityReference();
                                    accACC.Adx_name = distContact.FullName;
                                    accACC.Adx_Read = true;
                                    accACC.Adx_Write = true;
                                    accACC.Adx_ManagePermissions = true;
                                    createRequest = new CreateRequest { Target = accACC };
                                    multipleRequest.Requests.Add(createRequest);
                                }

                                /*
                                Entity
                                adx_contactaccess
                                Number of records = 1
                                Name = distContact.fullname
                                Account = distContact.parentcustomerid
                                Contact = distContact
                                Scope = account -- new OptionSet(2)
                                Read = true
                                Write = true
                                Create = true
                                Delete = false
                                */

                                ADX_ContactAccessQueryExp accCA = new ADX_ContactAccessQueryExp();
                                entityCollection = accCA.Execute(serviceProvider, distContact.Id, distAccount.Id, 2);

                                Adx_contactaccess conACC = new Adx_contactaccess();

                                if (entityCollection.Count <= 0)
                                {
                                    conACC.adx_contactid = distContact.ToEntityReference();
                                    conACC.adx_accountid = distAccount.ToEntityReference();
                                    conACC.Adx_name = distContact.FullName;
                                    conACC.Adx_Read = true;
                                    conACC.Adx_Write = true;
                                    conACC.Adx_Create = true;
                                    conACC.Adx_Delete = true;
                                    conACC.Adx_Scope = new OptionSetValue(2);
                                    createRequest = new CreateRequest { Target = conACC };
                                    multipleRequest.Requests.Add(createRequest);
                                }

                                /*
                                Entity
                                adx_channelpermissions
                                Number of records = 1
                                Name = distContact.fullname
                                Account = distContact.parentcustomerid
                                Contact = distContact
                                Read = true
                                Write = true
                                Create = true
                                Delete = false
                                */

                                ADX_ChannelPermissionsQueryExp chaPer = new ADX_ChannelPermissionsQueryExp();
                                entityCollection = chaPer.Execute(serviceProvider, distContact.Id, distAccount.Id);
                                adx_channelpermissions chPr = new adx_channelpermissions();

                                if (entityCollection.Count <= 0)
                                {
                                    chPr.adx_ContactId = distContact.ToEntityReference();
                                    chPr.adx_AccountId = distAccount.ToEntityReference();
                                    chPr.adx_name = distContact.FullName;
                                    chPr.adx_Read = true;
                                    chPr.adx_Write = true;
                                    chPr.adx_Create = true;
                                    chPr.adx_Delete = true;
                                    createRequest = new CreateRequest { Target = chPr };
                                    multipleRequest.Requests.Add(createRequest);
                                }


                                e++; UpdateRequest updateRequest = new UpdateRequest { Target = distContact };
                                e++; multipleRequest.Requests.Add(updateRequest);
                                e++; updateRequest = new UpdateRequest { Target = distAccount };
                                e++; multipleRequest.Requests.Add(updateRequest);
                                e++; updateRequest = new UpdateRequest { Target = servicingDealer };
                                e++; multipleRequest.Requests.Add(updateRequest);
                                e++; service.Execute(multipleRequest);
                            }
                        }
                    e = 750;

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("** Error in code position: " + e + " ** || *** eText: " + eText + " *** An error occurred in the UpdateOpportunity plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("** Error in code position: " + e + " ** || *** eText: " + eText + " *** UpdateOpportunity: {0}", ex.ToString());
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
