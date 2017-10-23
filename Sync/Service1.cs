using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using Microsoft.Xrm.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Data.SqlClient;
using System.IO;

namespace Sync
{
    public partial class Service1 : ServiceBase
    {


        System.Timers.Timer myTimer = new System.Timers.Timer();

        public Service1()
        {
            InitializeComponent();


        }

        protected override void OnStart(string[] args)
        {
            myTimer.Interval = 60000;

            myTimer.Enabled = true;

            myTimer.Elapsed += new System.Timers.ElapsedEventHandler(Sync);
            LogWriter.WriteErrorLog("The Questnet to CRM windows service started");
        }

        protected override void OnStop()
        {
            myTimer.Stop();
        }



////*************************************************************************************/////

        // Insert Log record in ServiceCallLog table
        private static void InsertData(SqlConnection sqlConnectionDetails, int callid, DateTime date, bool transfer, char ch)
        {
            // define INSERT query with parameters
            string query = "INSERT INTO dbo.ServiceCallsLog (CallId,date, transfered, action) " +
                           "VALUES (@CallId, @Date, @Transfer, @Char) ";

            // create connection and command

            using (SqlCommand cmmd = new SqlCommand(query, sqlConnectionDetails))
            {
                // define parameters and their values
                cmmd.Parameters.Add("@CallId", SqlDbType.BigInt).Value = callid;
                cmmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = date;
                cmmd.Parameters.Add("@Transfer", SqlDbType.Bit).Value = transfer;
                cmmd.Parameters.Add("@Char", SqlDbType.Char).Value = ch;


                // open connection, execute INSERT, close connection

                //sqlConnection1.Open();
                cmmd.ExecuteNonQuery();
                // sqlConnection1.Close();
            }
        }


        private static string ReturnVersion(string connectionDetails)
        {

            //******** Establish a connection to the organization web service using CrmConnection.
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);

            //********* Obtain an organization service proxy.
            //********* The using statement assures that the service proxy will be properly disposed.
            using (OrganizationService _orgService = new OrganizationService(connection))
            {

                //******** Obtain information about the logged on user from the web service.
                //  Guid userid = ((WhoAmIResponse)_orgService.Execute(new WhoAmIRequest())).UserId;
                // SystemUser systemUser = (SystemUser)_orgService.Retrieve("systemuser", userid,
                //     new ColumnSet(new string[] { "firstname", "lastname" }));
                //MessageBox.Show("Logged on user is {0} {1}.", systemUser.FirstName, systemUser.LastName);

                //*********** Retrieve the version of Microsoft Dynamics CRM.
                RetrieveVersionRequest versionRequest = new RetrieveVersionRequest();
                RetrieveVersionResponse versionResponse = (RetrieveVersionResponse)_orgService.Execute(versionRequest);

                return versionResponse.Version;
            }
        }



        private static string ReturnUser(string connectionDetails)
        {//********** Return Name of Logged in person
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                WhoAmIResponse whoAmI = (WhoAmIResponse)_orgService.Execute(new WhoAmIRequest());
                return whoAmI.ToString();
            }
        }


        private static EntityCollection ReturnAvailableCaseId(string connectionDetails, int callID)
        {// ************Return caseID if case for courrent callID is available
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);

            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                QueryExpression queryIncident = new QueryExpression { EntityName = "incident" };
                queryIncident.Criteria.AddCondition("new_callid", ConditionOperator.Equal, callID);
                queryIncident.ColumnSet = new ColumnSet("incidentid");
                //   queryIncident.ColumnSet.AllColumns = true;
                EntityCollection returnIncident = _orgService.RetrieveMultiple(queryIncident);


                return returnIncident;
            }
        }

        private static Guid CreateCase(string connectionDetails, int callID, string companyNo, string severity, string serialNumber)
        {//************** create Case for callID
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                Guid caseid = new Guid();
                QueryExpression queryAccount = new QueryExpression { EntityName = "account" };
                queryAccount.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, companyNo);
                queryAccount.ColumnSet = new ColumnSet("accountnumber", "accountid");
                EntityCollection returnAccount = _orgService.RetrieveMultiple(queryAccount);

                if (returnAccount.Entities[0].Attributes.Contains("accountid"))
                {
                    Incident incident = new Incident();
                    incident.Title = "QuestNet new Service Call add at " + System.DateTime.Now + "  and Call ID is :" + callID;
                    incident.CustomerId = new EntityReference("account", returnAccount.Entities[0].Id);
                    incident.new_CallID = callID;
                    // incident.new_ServiceLevel = new EntityReference("new_servicelevel", new Guid(serviceLevelCRMID));
                    incident.SubjectId = new EntityReference("subject", new Guid("F1B1CD44-80AD-E511-95A2-001B78BA3150"));
                    incident.new_QuestNetURL = "https://questnet.questinc.com/servicecalldetail.aspx?callno=" + callID + "&CompanyNumber=" + companyNo;
                    incident.CaseTypeCode = new OptionSetValue(1);
                    incident.CaseOriginCode = new OptionSetValue(3);
                    incident.ProductSerialNumber = serialNumber;

                    switch (severity)
                    {
                        case "Critical":
                            incident.PriorityCode = new OptionSetValue(1);

                            break;
                        case "High":
                            incident.PriorityCode = new OptionSetValue(2);
                            break;
                        case "Medium":
                            incident.PriorityCode = new OptionSetValue(3);
                            break;
                        case "Low":
                            incident.PriorityCode = new OptionSetValue(4);
                            break;
                        case "IMAC":
                            incident.PriorityCode = new OptionSetValue(5);
                            break;
                        case "Track Only":
                            incident.PriorityCode = new OptionSetValue(6);
                            break;

                    }

                    caseid = _orgService.Create(incident);
                }
                return caseid;

            }

        }


        private static Guid CheckResource(string connectionDetails, Guid caseID)
        {//*************** assign resource
            Guid _appointmentID = new Guid();
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
              AppointmentRequest appointmentRequest = new AppointmentRequest()
                   {
                       Direction = SearchDirection.Forward,
                       Duration = 5,
                       NumberOfResults = 1,
                       ServiceId = new Guid("49B00170-DBBE-E511-B0BC-001B78BA3150"),
                       SearchWindowStart = DateTime.Now,
                       SearchWindowEnd = DateTime.Now.AddMinutes(30)
                   };

                SearchRequest search = new SearchRequest()
                {
                    AppointmentRequest = appointmentRequest
                };
                SearchResponse searched = new SearchResponse();
                searched = (SearchResponse)_orgService.Execute(search);

                if (searched.SearchResults.Proposals.Length > 0)
                {
                    //booking Action
                    ActivityParty[] activityparties = new ActivityParty[searched.SearchResults.Proposals[0].ProposalParties.Length];
                    for (int i = 0; i <= searched.SearchResults.Proposals[0].ProposalParties.Length - 1; i++)
                    {
                        activityparties[i] = new ActivityParty { PartyId = new EntityReference(searched.SearchResults.Proposals[0].ProposalParties[i].EntityName, searched.SearchResults.Proposals[0].ProposalParties[i].ResourceId) };
                    }

                    ServiceAppointment serviceAppointment = new ServiceAppointment()
                    {
                        Subject = "New Service Call Handle",
                        Description = "Testing New Service Call Handle",
                        ScheduledStart = searched.SearchResults.Proposals[0].Start,
                        ScheduledEnd = searched.SearchResults.Proposals[0].End,
                        Resources = activityparties,
                        ServiceId = new EntityReference("service", new Guid("49B00170-DBBE-E511-B0BC-001B78BA3150")),
                        RegardingObjectId = new EntityReference("incident", caseID)


                    };

                    BookRequest book = new BookRequest
                    {
                        Target = serviceAppointment
                    };

                    BookResponse booked = (BookResponse)_orgService.Execute(book);
                    _appointmentID = booked.ValidationResult.ActivityId;



                }
                return _appointmentID;

            }

        }

        private static void SendEmail(string connectionDetails, Guid appID)
        {
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionDetails);
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                QueryExpression queryFromUser = new QueryExpression { EntityName = "systemuser" };
                queryFromUser.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, "crmadmin@questinc.com");
                queryFromUser.ColumnSet = new ColumnSet("systemuserid");
                EntityCollection returnFromUser = _orgService.RetrieveMultiple(queryFromUser);


                QueryExpression queryAppointment = new QueryExpression { EntityName = "serviceappointment" };
                queryAppointment.Criteria.AddCondition("activityid", ConditionOperator.Equal, appID);
                queryAppointment.ColumnSet = new ColumnSet("resources", "activityid");
                EntityCollection returnAppointment = _orgService.RetrieveMultiple(queryAppointment);
                var resources = (EntityCollection)returnAppointment.Entities[0].Attributes["resources"];
                var resource = (EntityReference)resources[0].Attributes["partyid"];


                Guid FromUserId = (Guid)returnFromUser.Entities[0].Attributes["systemuserid"];

                //    Create the 'From:' activity party for the email
                ActivityParty fromParty = new ActivityParty
                {
                    PartyId = new EntityReference(SystemUser.EntityLogicalName, FromUserId)
                };


                ////    Create the 'To:' activity party for the email
                ActivityParty toParty = new ActivityParty
                {
                    PartyId = resource
                };

                //   Create an e-mail message.
                Email email = new Email
                {
                    To = new ActivityParty[] { toParty },
                    From = new ActivityParty[] { fromParty },
                    Subject = " New Task Test",
                    DirectionCode = true,
                    Description = " You have new task!!!      "
                };


                Guid _emailId = _orgService.Create(email);
                SendEmailRequest sendEmailreq = new SendEmailRequest();
                Microsoft.Crm.Sdk.Messages.SendEmailRequest req = new Microsoft.Crm.Sdk.Messages.SendEmailRequest();
                req.EmailId = _emailId;

                req.TrackingToken = "";
                req.IssueSend = true;

                //   Finally Send the email message.

                SendEmailResponse res = (SendEmailResponse)_orgService.Execute(req);

                if (res != null) LogWriter.WriteErrorLog("message sent");
            }
        }

        public void Sync(object sender, System.Timers.ElapsedEventArgs e)
        {


            String connectionString = "Url=https://qcrm.questinc.com; Username=axie; Password=8291;";
            SqlConnection sqlConnection = new SqlConnection("Data Source=10.10.1.100;Initial Catalog=QuestNetTestDB; UID=quest_www;PWD=1234");
            SqlCommand cmd = new SqlCommand();

            Guid caseId = new Guid();
            Guid userId = new Guid();
            int callID;
            string companyNo;
            string serialNumber;
            //string serviceLevelCRMID;
            string serviceProviderNumber;
            string severity;
            char status = 'U';
            DateTime lastUpdate;

            #region query Service Call from QuestNet

            cmd.CommandText = "SELECT     dbo.ServiceCalls.CallID,dbo.ServiceCalls.CompanyNo, dbo.ServiceCalls.ItemSerialNumber,dbo.ServiceCalls.ItemSupportProvider,dbo.ServiceCalls.Severity,dbo.ServiceCalls.callID, dbo.ServiceCalls.LastUpdate, T0.LastUpdateDate FROM    dbo.ServiceCalls LEFT OUTER JOIN (SELECT CallId, MAX(date) AS LastUpdateDate FROM  dbo.ServiceCallsLog  GROUP BY CallId) AS T0 ON T0.CallId = dbo.ServiceCalls.CallID WHERE    (dbo.ServiceCalls.LastUpdate > T0.LastUpdateDate) OR (dbo.ServiceCalls.LastUpdate > @date) AND (T0.LastUpdateDate IS NULL)";
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue("@date", SqlDbType.DateTime);
            cmd.Parameters["@date"].Value = System.DateTime.Now.AddHours(-20);

            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection;
            DataSet serviceCalls = new DataSet();
            SqlDataAdapter sda = new SqlDataAdapter();
            sqlConnection.Open();


            try
            {
                sda.SelectCommand = cmd;
                sda.Fill(serviceCalls, "Service_Calls");
                int counter = serviceCalls.Tables["Service_Calls"].Rows.Count;
                do
                {
                    if (counter == 0) break;
                    DataRow row = serviceCalls.Tables["Service_Calls"].Rows[counter - 1];
                    companyNo = row["CompanyNo"].ToString();
                    //  serviceLevelCRMID = row["CRMid"].ToString();
                    serialNumber = row["ItemSerialNumber"].ToString();
                    serviceProviderNumber = row["ItemSupportProvider"].ToString();
                    severity = row["Severity"].ToString();
                    callID = System.Convert.ToInt32(row["callID"]);
                    //   status = System.Convert.ToChar(row["action"]);
                    lastUpdate = System.Convert.ToDateTime(row["LastUpdate"]);


                    if (ReturnAvailableCaseId(connectionString, callID).Entities.Count == 0)
                    {

                        caseId = CreateCase(connectionString, callID, companyNo, severity, serialNumber);

                    }
                    else
                    {
                        caseId = ReturnAvailableCaseId(connectionString, callID).Entities[0].Id;
                    }

                    counter--;

                    InsertData(sqlConnection, callID, System.DateTime.Now, false, status);
                    userId = CheckResource(connectionString, caseId);
                    SendEmail(connectionString, userId);

                    LogWriter.WriteErrorLog("Microsoft Dynamics CRM version " + ReturnVersion(connectionString) + "   Service Call ID: " + callID + "Company No" + companyNo);
                    
                } while (counter > 0);


            }
            catch (Exception ex)
            {
                LogWriter.WriteErrorLog(ex);
               
            }

            sqlConnection.Close();
           
            #endregion


        }


    }
}


