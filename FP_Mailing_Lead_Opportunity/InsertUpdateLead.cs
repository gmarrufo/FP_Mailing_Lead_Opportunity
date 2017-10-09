// <copyright file="InsertUpdateLead.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>04/12/16 10:10:00 AM</date>
// <summary>Implements the InsertUpdateLead Process to check for multiple conditions (customer check, duplicate check, retention check, etc.) prior to determine lead assignment.</summary>

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

namespace FPMailingLeadOpportunity
{
    public class InsertUpdateLead : IPlugin
    {
        // Create EntityCollection variables to hold results from querying
        EntityCollection resultsDupCust = new EntityCollection();
        EntityCollection accResultsDupCust = new EntityCollection();
        EntityCollection resultsDupLead = new EntityCollection();
        EntityCollection accResultsDupLead = new EntityCollection();
        DataCollection<Entity> entityCollection = null;
        DataCollection<Entity> entityCollection1 = null;
        DataCollection<Entity> entityCollection2 = null;
        DataCollection<Entity> entityCollection3 = null;

        // Create a Boolean variable to hold result of duplicate customer check
        bool bolDupCustCheck = false;

        // Create a Boolean variable to hold result of duplicate lead check
        bool bolDupLeadCheck = false;

        // Create following variables for assigning and querying process: FirstName, LastName, CompanyName, EmailAddress, Address Number, Zip Code
        string strFirstName = null;
        string strLastName = null;
        string strCompanyName = null;
        string strEmailAddress = null;
        string strAddressLine1 = null;
        string strAddressNumber = null;
        string strZipCode = null;

        // Useful variables
        string strChk4DupCust = null;
        double dDateDiff = 0.0;
        double dDateDiffCompare = 0.0;
        TimeSpan t;
        DateTime dToday = DateTime.Now;
        DateTime dDaysToExpire;

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
                
                string eText = "";
                int e = 101;

                try
                {
                    if (entity.Attributes.Contains("statecode"))
                        if (entity.Attributes["statecode"].Equals(2) || entity.Attributes["statecode"].Equals(1))
                            return;

                    var curQuery = new QueryExpression("organization");
                    curQuery.ColumnSet = new ColumnSet("basecurrencyid");
                    var curResult = service.RetrieveMultiple(curQuery);
                    var currencyId = (EntityReference)curResult.Entities[0]["basecurrencyid"];

                    #region STEP -2 Duplicate Customer Check

                    /* Determine duplication based on following criteria:

                    Customer Check by Search Fields: 
                    o	First Name
                    o	Last Name
                    o	Company Name (Strip abbreviations: CO, INC, THE, OF and CORP)
                    o	Email Address
                    o	Address Number (123 Main St. Use Only 123)
                    o	Zip Code

                    In the rules below, active customer match is defined below:

                    •	Email Address is an exact match then there is a customer match OR
                    •	First Name, Last name, Address Number and Zip Code is a customer match OR
                    •	First Name, Last Name, Zip Code and Company Name is a customer match OR
                    •	Company Name, Address Number and Zip Code is a customer match OR 

                     If a match is found for a Cancelled Customer (An account with no active meter line for more than 6 months), this is NOT a customer match. If an existing customer is a match, check box as “existing customer”.  
                    
                    */

                    // Assign Raw Lead Object elements to created variables
                    strFirstName = getFieldData(entity, "firstname") ?? "";
                    strLastName = getFieldData(entity, "lastname") ?? "";
                    strCompanyName = getFieldData(entity, "companyname") ?? "";
                    strEmailAddress = getFieldData(entity, "emailaddress1") ?? "";
                    strAddressLine1 = getFieldData(entity, "address1_line1") ?? "";
                    strZipCode = getFieldData(entity, "address1_postalcode") ?? "";

                    e++; /************  Code Position #102  ************/
                    int pc = -1;

                    if (entity.Attributes.Contains("new_productcategory"))
                    {
                        pc = ((Microsoft.Xrm.Sdk.OptionSetValue)(entity.Attributes["new_productcategory"])).Value;
                    }

                    // Create an object that will contain a list of company abbreviations. 
                    // Use the pearl_anytwokey table in action = 3

                    e++; /************  Code Position #103  ************/
                    if (!string.IsNullOrEmpty(strCompanyName))
                    {
                        PearlAnyTwoKeyValue pATK3 = new PearlAnyTwoKeyValue();
                        entityCollection = pATK3.Execute(serviceProvider, 3);
                        List<string> listCompanyAbbreviations = new List<string>();

                        foreach (pearl_anytwokey plD in entityCollection)
                        {
                            listCompanyAbbreviations.Add(plD.Attributes["pearl_text1"].ToString());
                        }
                        // Strip Company Name abbreviations

                        // Loop through List with foreach.
                        foreach (String strAbbreviation in listCompanyAbbreviations)
                        {
                            strCompanyName = strCompanyName.SafeReplace(strAbbreviation, "", true);
                        }
                    }
                    e++; /************  Code Position #104  ************/
                    if (!string.IsNullOrEmpty(strAddressNumber))
                    {
                        // Strip Address Number from Addresses
                        strAddressNumber = StringExtensions.StripAddressNumber(strAddressLine1);
                    }
                    if (!string.IsNullOrEmpty(strZipCode))
                    {
                        // Strip last 4 numbers from Zip Code
                        strZipCode = StringExtensions.StripZipCode(strZipCode);
                    }
                    e++; /************  Code Position #105  ************/

                    // Check existing Organization Service alongside the Raw Lead Object for:

                    // Email form Organization Service == Email from Raw Lead Object

                    // ZipCode, Company Name from Organization Service == ZipCode, Company Name from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    bolDupCustCheck = false;
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strCompanyName))
                        if (DuplicateCustomerCheck(service, "account", new string[] { "new_stczipcode", "new_accountname" }, new string[] { strZipCode, strCompanyName }).Entities.Count > 0 && bolDupCustCheck == false)
                        {
                            e = 1003; /************  Code Position #1003  ************/
                            bolDupCustCheck = true;
                            accResultsDupCust = resultsDupCust;
                            strChk4DupCust = "account";
                        }
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strAddressNumber))
                        if (DuplicateCustomerCheck(service, "account", new string[] { "new_stcaddress", "new_stczipcode" }, new string[] { strAddressNumber, strZipCode }).Entities.Count > 0 && bolDupCustCheck == false)
                        {
                            e = 1004; /************  Code Position #1004  ************/
                            bolDupCustCheck = true;
                            accResultsDupCust = resultsDupCust;
                            strChk4DupCust = "account";
                        }
                    // ZipCode, Company Name from Organization Service == ZipCode, Company Name from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strCompanyName))
                        if (DuplicateCustomerCheck(service, "account", new string[] { "new_billtopostcode", "new_accountname" }, new string[] { strZipCode, strCompanyName }).Entities.Count > 0 && bolDupCustCheck == false)
                        {
                            e = 1005; /************  Code Position #1005  ************/
                            bolDupCustCheck = true;
                            accResultsDupCust = resultsDupCust;
                            strChk4DupCust = "account";
                        }
                    //Address Number, Zip Code from Organization Service == Address Number, Zip Code from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strAddressNumber))
                        if (DuplicateCustomerCheck(service, "account", new string[] { "new_billtostreet1", "new_billtopostcode" }, new string[] { strAddressNumber, strZipCode }).Entities.Count > 0 && bolDupCustCheck == false)
                        {
                            e = 1002; /************  Code Position #1002  ************/
                            bolDupCustCheck = true;
                            accResultsDupCust = resultsDupCust;
                            strChk4DupCust = "account";
                        }
                    // (if equal then set Boolean variable to true, otherwise no) 
                    if (!string.IsNullOrEmpty(strEmailAddress))
                        if (DuplicateCustomerCheck(service, "account", new string[] { "emailaddress1" }, new string[] { strEmailAddress }).Entities.Count > 0 && bolDupCustCheck == false)
                        {
                            e = 1001; /************  Code Position #1001  ************/
                            bolDupCustCheck = true;
                            accResultsDupCust = resultsDupCust;
                            strChk4DupCust = "account";
                        }
                    //TODO fix customer check adding addresses
                    // Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    //else if (DuplicateCustomerCheck(service, "account", new string[] { "new_accountname", "new_billtostreet1", "new_billtopostcode" }, new string[] { strCompanyName, strAddressNumber, strZipCode }).Entities.Count > 0 && bolDupCustCheck == false)
                    //{
                    //    e=1006; /************  Code Position #1006  ************/
                    //    bolDupCustCheck = true;
                    //    accResultsDupCust = resultsDupCust;
                    //    strChk4DupCust = "account";
                    //}

                    // On a positive match process:
                    if (bolDupCustCheck && strChk4DupCust == "account")
                    {
                        // Active Rental Header -- Rental Line -- Deal code that doesn't start with D or C and Has GEN PROD POS = Meter, Document and Line No.
                        // More that one Active Real Header for the same account - BEWARE!!!!!!
                        RentalLineByAccount rLBA = new RentalLineByAccount();
                        entityCollection3 = rLBA.Execute(serviceProvider, accResultsDupCust[0].Attributes["name"].ToString());
                        e = 107; /************  Code Position #107  ************/
                        if (entityCollection3 != null)
                        {
                            if (entityCollection3.Count > 0)
                            {
                                e = 108; /************  Code Position #108  ************/
                                /* Qualify the Incoming Lead */
                                //Guid accId = new Guid(accResultsDupCust[0].Attributes["accountid"].ToString());
                                //Entity entDupAccount = service.Retrieve("account", (Guid)(accResultsDupCust[0].Attributes["accountid"].ToString()), new ColumnSet(true));
                                Account dupAccount = service.Retrieve("account", (Guid)(accResultsDupCust[0].Attributes["accountid"]), new ColumnSet(true)).ToEntity<Account>();
                                e = 109; /************  Code Position #109  ************/
                                //entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                                //entity.Attributes["statuscode"] = new OptionSetValue(3);// "Qualified because Active Rental Header";
                                if (dupAccount.new_GenBusPostingGroup == "MAJORACCT")
                                    entity.Attributes["pearl_map"] = true;
                                entity.Attributes["pearl_leadstatus"] = "Existing Customer";
                                entity.Attributes["pearl_existingcustomer"] = bolDupCustCheck;
                                entity.Attributes["parentaccountid"] = dupAccount.ToEntityReference();// new EntityReference("account", accId);
                                e = 110; /************  Code Position #110  ************/
                            }
                            /*else
                            {
                                if (accResultsDupCust[0].Attributes.Contains("accountid"))
                                {
                                    context.SharedVariables.Add("accountId", accResultsDupCust[0].Attributes["accountid"].ToString());
                                }

                                Guid accId = new Guid(accResultsDupCust[0].Attributes["accountid"].ToString());
                                entity.Attributes["pearl_leadstatus"] = "Previous Customer";
                                entity.Attributes["pearl_existingcustomer"] = bolDupCustCheck;
                                entity.Attributes["parentaccountid"] = new EntityReference("account", accId);
                            }*/
                        }
                    }

                    e = 111; /************  Code Position #111  ************/
                    // Nullify some variables for reuse or GC
                    resultsDupCust = null;
                    accResultsDupCust = null;
                    //listCompanyAbbreviations = null;
                    bolDupCustCheck = false;
                    strChk4DupCust = null;

                    #endregion

                    #region STEP 3 – Duplicate Check and Retention Check

                    e = 201; /************  Code Position #201  ************/

                    /* Process duplicate check based on following criteria:
                
                    All SELLFP.COM leads are NOT Duplicates. Create new record for them and they will go SellFP email distribution.
                    Check for any duplicate leads. If the lead is NOT a duplicate, it will move to Step 4: Qualify stage.
                    CRM to search duplicate leads by the same Customer Check Rules above except Category Match will be a new field to check. 

                    o	First Name
                    o	Last Name
                    o	Company Name (Strip abbreviations: CO, INC, THE, OF and CORP)
                    o	Email Address
                    o	Address Number (123 Main St. Use Only 123)
                    o	Zip Code
                    o	Category Match (meter, folder/inserter, etc)

                    In the rules below, active duplicate match is defined below:

                    •	If an exact match lead comes in but with two different Categories, it is NOT a duplicate.
                    •	Email Address is an exact match then there is a duplicate match.s
                    •	First Name, Last name, Address Number and Zip Code is a duplicate match.
                    •	First Name, Last Name, Zip Code and Company Name is a duplicate match
                    •	Company Name, Address Number and Zip Code is a duplicate match.

                    Any lead marked as duplicate will link to active lead for reference. 
                    */

                    // Check existing Organization Service alongside the Lead Object and check for:

                    // FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Lead Object
                    // (if equal then check if  Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strAddressNumber))
                        if (DuplicateLeadCheck(service, "lead", new string[] { "firstname", "lastname", "address1_line1", "address1_postalcode" }, new string[] { strFirstName ?? "", strLastName ?? "", strAddressNumber, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                        {
                            e = 2002; /************  Code Position #2002  ************/
                            bolDupLeadCheck = true;
                            accResultsDupLead = resultsDupLead;
                        }

                    // FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Lead Object
                    // (if equal then heck if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strCompanyName))
                        if (DuplicateLeadCheck(service, "lead", new string[] { "firstname", "lastname", "companyname", "address1_postalcode" }, new string[] { strFirstName ?? "", strLastName ?? "", strCompanyName, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                        {
                            e = 2003; /************  Code Position #2003  ************/
                            bolDupLeadCheck = true;
                            accResultsDupLead = resultsDupLead;
                        }

                    // Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Lead Object
                    // (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    if (!string.IsNullOrEmpty(strZipCode) && !string.IsNullOrEmpty(strAddressNumber) && !string.IsNullOrEmpty(strCompanyName))
                        if (DuplicateLeadCheck(service, "lead", new string[] { "companyname", "address1_line1", "address1_postalcode" }, new string[] { strCompanyName, strAddressNumber, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                        {
                            e = 2004; /************  Code Position #2004  ************/
                            bolDupLeadCheck = true;
                            accResultsDupLead = resultsDupLead;
                        }
                    // Email form Organization Service == Email from Lead Object
                    // (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no) 
                    if (!string.IsNullOrEmpty(strEmailAddress))
                        if (DuplicateLeadCheck(service, "lead", new string[] { "emailaddress1" }, new string[] { strEmailAddress }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                        {
                            e = 2001; /************  Code Position #2001  ************/
                            bolDupLeadCheck = true;
                            accResultsDupLead = resultsDupLead;
                        }
                    /* Process retention check based on following criteria:
                
                    If incoming lead has retention check, then disable incoming lead. End of Workflow Process.
                    If original lead has the retention check, then disable the original lead 

                    */

                    e = 202; /************  Code Position #202  ************/

                    // Check if Raw Object Lead is of Retention Check == TRUE, if equal set the Raw Object Lead as disabled, otherwise,
                    // check if Organization Service Lead is of Retention Check == TRUE, if equal set the Organization Service Lead as disabled.
                    if (accResultsDupLead.Entities.Count > 0)
                    {
                        e = 203; /************  Code Position #203  ************/
                        foreach (Entity item in accResultsDupLead.Entities)
                        {
                            e = 204; /************  Code Position #204  ************/
                            if (entity.Attributes.Contains("pearl_retention"))
                            {
                                e = 205; /************  Code Position #205  ************/
                                /* Incoming lead no retention */
                                if (!Convert.ToBoolean(entity.Attributes["pearl_retention"]))
                                {
                                    e = 206; /************  Code Position #206  ************/
                                    /* Checking Original Lead for retention Check */
                                    if (item.Attributes["pearl_retention"] != null)
                                    {
                                        e = 207; /************  Code Position #207  ************/
                                        if (Convert.ToBoolean(item.Attributes["pearl_retention"]))
                                        {
                                            e = 208; /************  Code Position #208  ************/
                                            // Pass the data to the post event plug-in in an execution context shared variable named orginalLeadID
                                            context.SharedVariables.Add("orginalLeadID", item.Attributes["leadid"].ToString());
                                        }
                                        else
                                        {
                                            e = 209; /************  Code Position #209  ************/
                                            // pearl_daytoexpire > today ==> disable incoming Lead keep original
                                            double exipre = 0.0;
                                            if (item.Attributes.Contains("pearl_daystoexpire"))
                                                exipre = Convert.ToDouble(item.Attributes["pearl_daystoexpire"]);
                                            else
                                                exipre = 180;
                                            

                                            dDaysToExpire = Convert.ToDateTime(item.Attributes["createdon"]).AddDays(exipre);

                                            if (dDaysToExpire > dToday)
                                            {
                                                /* NOT WORKING CORRECTLY??? */
                                                e = 210; /************  Code Position #210  ************/
                                                /* Disable Incoming Lead */
                                                //new_Qualified && updLead.new_NotQualifiedReason
                                                /*entity.Attributes["statecode"] = new OptionSetValue(2); // Disqualified
                                                entity.Attributes["statuscode"] = new OptionSetValue(7); //"Disqualified - DaysToExpire greater than Today";
                                                entity.Attributes["pearl_leadstatus"] = "Disqualified - DaysToExpire greater than Today";*/
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    e = 212; /************  Code Position #212  ************/
                                    entity.Attributes["pearl_linktooriginallead"] = new EntityReference("lead", (Guid)item.Attributes["leadid"]);

                                    /* Disable Incoming Lead */
                                    entity.Attributes["statecode"] = new OptionSetValue(2); // Disqualified
                                    entity.Attributes["statuscode"] = new OptionSetValue(7); // "Disqualified - Lead No Retention";
                                    entity.Attributes["pearl_leadstatus"] = "Disqualified - Lead No Retention";
                                }
                            }
                        }
                    }

                    // Nullify some variables for reuse or GC
                    resultsDupLead = null;
                    accResultsDupLead = null;
                    bolDupLeadCheck = false;

                    #endregion

                    #region STEP 4 – Pre Map Check
                    e = 401; /************  Code Position #401  ************/
                    /* Process Pre Map Check based on following criteria:

                    Leads need to be checked if they are a MAP lead, opportunity or existing customer. See separate “MAP Accounts” spreadsheet.

                    MAP check will also include one or all fields:
                    o	Emails ending with: .edu, .gov, .mil
                    o	Emails ending with FP’s existing MAP accounts (i.e. @Walgreens.com, @Lowes.com, etc.)
                    o	MAP Customer Names (The UPS Store, Hilton, etc.)
                
                    If any criteria above is met, check the MAP tracking field.
                
                    */
                    // Create an object that will contain a list of identifying MAP emails fields like: .edu, .gov, .mil, existing MAP accounts, MAP Customer Names, etc. - TO DO Based on CRM
                    // Use the pearl_anytwokey table in action = 1
                    PearlAnyTwoKeyValue pATK1 = new PearlAnyTwoKeyValue();
                    entityCollection = pATK1.Execute(serviceProvider, 1);

                    List<String> listMAPEmails = new List<String>();

                    foreach (pearl_anytwokey plD in entityCollection)
                    {
                        listMAPEmails.Add(plD.Attributes["pearl_text1"].ToString());
                    }
                    e++; /************  Code Position #402  ************/

                    // Create a process that will take the Raw Lead Object and check against the object containing a list of identifying MAP email fields,
                    // if found mark the MAP tracking field.

                    // Loop through List with foreach.
                    foreach (String strMAPEmail in listMAPEmails)
                    {
                        // Obtain Raw Lead Object
                        if (entity.Attributes.Contains("emailaddress1"))
                        {
                            string emailString = entity.Attributes["emailaddress1"].ToString().ToLower();
                            if (emailString.Contains(strMAPEmail))
                            {
                                entity.Attributes["pearl_map"] = true;
                            }
                        }
                    }
                    e++; /************  Code Position #403  ************/
                    // Create an object that will contain a list of identifying MAP names fields like existing MAP accounts, MAP Customer Names, etc. - TO DO Based on CRM
                    // Use the pearl_anytwokey table in action = 2
                    PearlAnyTwoKeyValue pATK2 = new PearlAnyTwoKeyValue();
                    entityCollection = pATK2.Execute(serviceProvider, 2);

                    List<string> listMAPNames = new List<string>();

                    foreach (pearl_anytwokey plD in entityCollection)
                    {
                        listMAPNames.Add(plD.Attributes["pearl_text1"].ToString());
                    }
                    e++; /************  Code Position #404  ************/

                    // Create a proess that will take the Raw Lead Object and check against the object containing a list of identifying MAP names,
                    // if found mark the MAP tracking field.

                    // Loop through List with foreach.
                    foreach (String strMAPName in listMAPNames)
                    {
                        // Obtain Raw Lead Object
                        if (entity.Attributes.Contains("company"))
                        {
                            string emailString = entity.Attributes["company"].ToString().ToLower();
                            if (emailString.Contains(strMAPName))
                            {
                                entity.Attributes["pearl_map"] = true;
                            }
                        }
                    }
                    e++; /************  Code Position #405  ************/
                    #endregion

                    #region STEP 5 – Qualified

                    e = 301; /************  Code Position #301  ************/

                    /* Qualification process based on the following criteria:
                     * 
                     *   If pearl_AutoQualified equal true qualify the incoming Lead
                     *   Else
                     *   Assign Round Robin to a DDR                    
                     * 
                    */

                    // GMC - 04/14/16 - Change code by Alex Patrickus email
                    /*
                    Lead entityLead = null;
                    bool manualMode = false;
                    bool distributed = false;
                    bool autoQualify = false;

                    Entity entLead = service.Retrieve("lead", entity.Id, new ColumnSet(true));
                    entityLead = entLead.ToEntity<Lead>();

                    if (entity.Attributes.Contains("pearl_manualmode"))
                        manualMode = (bool)entity.Attributes["pearl_manualmode"];
                    eText = "manual check";

                    if (entity.Attributes.Contains("pearl_distributed"))
                        distributed = (bool)entity.Attributes["pearl_distributed"];
                    eText += "distribution check";

                    if (entity.Attributes.Contains("pearl_autoqualified"))
                        autoQualify = (bool)entity.Attributes["pearl_autoqualified"];
                    eText += "qualify check";
                    */

                    Lead entityLead = null;
                    bool manualMode = false;
                    bool distributed = false;
                    bool autoQualify = false;
                    bool eCatch = false;
                    try
                    {
                        Entity entLead = service.Retrieve("lead", entity.Id, new ColumnSet(true));
                        entityLead = entLead.ToEntity<Lead>();

                        if (entityLead != null)
                        {
                            autoQualify = (bool)entityLead.pearl_AutoQualified;
                            manualMode = (bool)entityLead.pearl_ManualMode;
                            distributed = (bool)entityLead.pearl_Distributed;                            
                        }
                    }
                    catch
                    {
                        eCatch = true;
                        if (entity.Attributes.Contains("pearl_manualmode"))
                            manualMode = (bool)entity.Attributes["pearl_manualmode"];
                        eText = "manual check";

                        if (entity.Attributes.Contains("pearl_distributed"))
                            distributed = (bool)entity.Attributes["pearl_distributed"];
                        eText += "distribution check";

                        if (entity.Attributes.Contains("pearl_autoqualified"))
                            autoQualify = (bool)entity.Attributes["pearl_autoqualified"];
                        eText += "qualify check";
                    }
                    entity.Attributes["new_datereceived"] = dToday;
                    if (entity.Attributes.Contains("pearl_routingtype") && entity.Attributes.Contains("pearl_leadtype"))
                    {
                        pearl_leadroutingtype plrtype = service.Retrieve("pearl_leadroutingtype", ((EntityReference)entity.Attributes["pearl_routingtype"]).Id, new ColumnSet(true)).ToEntity<pearl_leadroutingtype>();
                        entity.Attributes["pearl_daystoexpire"] = plrtype.pearl_DaysToExpire;
                    }
                    if (!(manualMode))
                    {
                        if (!(distributed))
                        {
                            if ((autoQualify))
                            {
                                //AutoQualify is completeted on update of lead in InsertLeadProcess Plugin

                                //e = 3101; /************  Code Position #3101  ************/
                                ///* Qualify the Incoming Lead */
                                //Guid lId = entity.Id;
                                //eText = lId.ToString();
                                //eText += " " + entity.LogicalName + " ";
                                //e = 3102; /************  Code Position #3102  ************/
                                ////service.Update(entity);
                                //e = 3103; /************  Code Position #3103  ************/
                                ////entity.Attributes["statuscode"] = new OptionSetValue(3);
                                //var qlreq = new QualifyLeadRequest
                                //{
                                //    CreateOpportunity = true,
                                //    OpportunityCurrencyId = currencyId,
                                //    CreateAccount = false,
                                //    CreateContact = false,
                                //    //OpportunityCustomerId = null,
                                //    Status = new OptionSetValue(3),//.Qualified),                                
                                //    //Status = new OptionSetValue(1),      
                                //    SourceCampaignId = null,
                                //    LeadId = new EntityReference(entity.LogicalName, lId)//entity.Id)//entity.ToEntityReference(),// ((EntityReference)entity.Id),                                

                                //};
                                //e++; /************  Code Position #3104  ************/
                                //var qlres = (QualifyLeadResponse)service.Execute(qlreq);
                                ////service.Execute(qlreq);

                                //e++; /************  Code Position #3103  ************/

                                ////entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                                ////entity.Attributes["statuscode"] = new OptionSetValue(3); //"Qualified because pearl_AutoQualified is true";
                            }
                            else
                            {
                                if (entity.Attributes.Contains("pearl_routingtype") && entity.Attributes.Contains("pearl_leadtype"))
                                {
                                    e++; /************  Code Position #302  ************/

                                    pearl_leadroutingtype plrt = service.Retrieve("pearl_leadroutingtype", ((EntityReference)entity.Attributes["pearl_routingtype"]).Id, new ColumnSet(true)).ToEntity<pearl_leadroutingtype>();
                                    eText = "1.1";
                                    Contact defaultContact = service.Retrieve("contact", plrt.pearl_DefaultContact.Id, new ColumnSet(true)).ToEntity<Contact>();
                                    eText = "1.2";
                                    EntityReference sDefaultAccount = plrt.pearl_DefaultAccount;
                                    // Query Expression to get from Account all Contacts with Default Distribution Account and "pearl_AssignLead" == true,
                                    AllContactsDefaultAccount aCDA = new AllContactsDefaultAccount();
                                    eText = "1.3";
                                    entityCollection2 = aCDA.Execute(serviceProvider, sDefaultAccount);


                                    // Round Robin
                                    string sContactID = "";
                                    int i = 0;
                                    EntityReference sDefaultContact = null;
                                    e++; /************  Code Position #303  ************/
                                    bool first = true;
                                    foreach (Contact cT in entityCollection2)
                                    {
                                        t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                        dDateDiffCompare = t.TotalMilliseconds;
                                        if (dDateDiffCompare > dDateDiff || first)
                                        {
                                            first = false;
                                            sDefaultContact = entityCollection2[i].ToEntityReference();
                                            sContactID = sDefaultContact.Id.ToString();
                                            dDateDiff = dDateDiffCompare;
                                        }
                                        i++;
                                    }
                                    // Assign the Lead to the Contact

                                    e++; /************  Code Position #304  ************/
                                    e = 305; /************  Code Position #305  ************/

                                    SystemUser sysUser = null;
                                    if (sContactID != "")
                                    {
                                        Contact assignContact = service.Retrieve(Contact.EntityLogicalName, sDefaultContact.Id, new ColumnSet(true)).ToEntity<Contact>();
                                        eText = assignContact.FullName + " catch: " + eCatch +" m "+manualMode+" d "+distributed+" a "+autoQualify;
                                        e = 306; /************  Code Position #306  ************/
                                        if (assignContact.Attributes.Contains("adx_systemuserid"))
                                            if (assignContact.Attributes["adx_systemuserid"] != null)
                                            {
                                                sysUser = service.Retrieve("systemuser", assignContact.Id, new ColumnSet(true)).ToEntity<SystemUser>();
                                                entity.Attributes["ownerid"] = sysUser.ToEntityReference();
                                            }

                                        entity.Attributes["pearl_ddr"] = true;
                                        entity.Attributes["pearl_distributed"] = true;
                                        entity.Attributes["pearl_routingtype"] = null;
                                        entity.Attributes["pearl_leadtype"] = null;
                                        
                                        //entity.Attributes["pearl_daystoexpire"] = plrt.pearl_DaysToExpire;
                                        eText = "12"; assignContact.pearl_LastOpportunityAssignment = dToday;
                                        eText = "13"; service.Update(assignContact);
                                        //eText = "14"; service.Update(entityLead);
                                    }
                                    else
                                    {
                                        e = 307; /************  Code Position #307  ************/
                                        sysUser = service.Retrieve("systemuser", defaultContact.adx_systemuserid.Id, new ColumnSet(true)).ToEntity<SystemUser>();
                                        entity.Attributes["ownerid"] = sysUser.ToEntityReference();
                                        entity.Attributes["pearl_ddr"] = true;
                                        entity.Attributes["pearl_distributed"] = true;
                                        entity.Attributes["pearl_routingtype"] = null;
                                        entity.Attributes["pearl_leadtype"] = null;
                                        //entity.Attributes["new_datereceived"] = dToday;
                                        //entity.Attributes["pearl_daystoexpire"] = plrt.pearl_DaysToExpire;
                                        //sysUser = entSysUser.ToEntity<SystemUser>();
                                        defaultContact.pearl_LastOpportunityAssignment = dToday;
                                        service.Update(defaultContact);
                                        //service.Update(entityLead);
                                    }
                                }
                                e = 308; /************  Code Position #308  ************/
                            }
                        }
                    }

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("** Error in position: " + e + " ** ***Etext: " + eText + "*** An error occurred in the InsertUpdateLead plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("** Error in position: " + e + " ** ***Etext: " + eText + "*** InsertUpdateLead : {0}", ex.ToString());
                    throw;
                }
            }
        }

        /*
        Create a Function that will take the existing Organization Service alongside the Raw Lead Object and check for:
        -	Email form Organization Service == Email from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no) 
        -	FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        -	FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        -	Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        */
        public EntityCollection DuplicateCustomerCheck
            (
            IOrganizationService service,
            string checkFor,
            string[] filterFields,
            string[] fieldValues,
            int productCategory = -1
            )
        {
            int e = 1101;
            try
            {
                int i = 0;
                ColumnSet cs = new ColumnSet();
                QueryExpression QE = new QueryExpression(checkFor);
                cs = new ColumnSet(true);//"name", "accountid", "statuscode", "statecode", "emailaddress1");
                QE.ColumnSet = cs;
                e++;  //1102
                foreach (string item in filterFields)
                {
                    if (item.Equals("address1_line1"))
                    {
                        if (!string.IsNullOrEmpty(fieldValues[i]))
                            QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                    }
                    else if (item.Equals("address1_postalcode"))
                    {
                        if (!string.IsNullOrEmpty(fieldValues[i]))
                            QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                    }
                    else/* if (item.Equals("companyname"))*/
                    {
                        if (!string.IsNullOrEmpty(fieldValues[i]))
                            QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                    }
                    /*
                    else if (item.Equals("name"))
                    {
                        QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                    }
                    else
                    {
                        QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                    }*/
                    i++;
                }
                e++; //1103
                resultsDupCust = service.RetrieveMultiple(QE);
                e++; //1104
                return resultsDupCust;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("** Error in position: " + e + " **  An error occurred in the InsertUpdateLead plug-in.", ex);
            }
            catch
            {
                throw;
            }
        }

        /*
        Create a Function that will take the existing Organization Service alongside the Lead Object and check for:
        -	Email form Organization Service == Email from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no) 
        -	FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Lead Object
        (if equal then check if  Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        -	FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        -	Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        */
        public EntityCollection DuplicateLeadCheck
            (
            IOrganizationService service,
            string checkFor,
            string[] filterFields,
            string[] fieldValues,
            int productCategory = -1
            )
        {
            int i = 0;
            ColumnSet cs = new ColumnSet();
            QueryExpression QE = new QueryExpression(checkFor);
            cs = new ColumnSet(true);//"companyname", "pearl_retention", "statuscode", "pearl_daystoexpire", "parentaccountid", "leadid");
            QE.ColumnSet = cs;

            foreach (string item in filterFields)
            {
                if (item.Equals("address1_line1"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else if (item.Equals("address1_postalcode"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else/* if (item.Equals("companyname"))*/
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                /*
                else if (item.Equals("name"))
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                else
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }*/
                i++;
            }

            if (productCategory != -1)
            {
                QE.Criteria.AddCondition("new_productcategory", ConditionOperator.Equal, productCategory);
            }

            QE.Criteria.AddCondition("statecode", ConditionOperator.LessEqual, 1);  // Look in to open and Qualified.

            resultsDupLead = service.RetrieveMultiple(QE);
            return resultsDupLead;
        }

        // Assign Entity to FieldNames
        public string getFieldData(Entity entity, string fieldName)
        {
            if (entity.Attributes.Contains(fieldName))
            {
                return entity.Attributes[fieldName].ToString();
            }
            else
            {
                return "";
            }
        }

        public Boolean checkLead(EntityCollection ent)
        {
            Boolean bolResult = false;
            if (ent.Entities.Count >= 1)
            {
                bolResult = true;
            }
            return bolResult;
        }
    }

    public static class StringExtensions
    {
        // Create a Function to Strip Company Name abbreviations
        public static string SafeReplace(this string input, string find, string replace, bool matchWholeWord)
        {
            string textToFind = matchWholeWord ? string.Format(@"\b{0}\b", find) : find;
            return Regex.Replace(input, textToFind, replace);
        }

        // Create a Function to Strip Address Number from Addresses
        public static string StripAddressNumber(this string input)
        {
            string[] strTemp = input.Split(' ');
            string strResult = null;
            if (strTemp.Length > 0)
            {
                strResult = strTemp[0];
            }
            return strResult;
        }

        // Create a Function to Strip last 4 numbers from Zip Code
        public static string StripZipCode(this string input)
        {
            string zipcode = input;
            if (zipcode.Length > 4)
                zipcode = zipcode.Substring(0, 5);
            return zipcode;
        }
    }
}
