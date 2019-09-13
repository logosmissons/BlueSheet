using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BlueSheetAppExcel = Microsoft.Office.Interop.Excel;
using BlueSheetAppExcelTools = Microsoft.Office.Tools.Excel;
using MySql.Data.MySqlClient;
using BlueSheetSForce = BlueSheetApp.Salesforce;
using MigraDocDOM = MigraDoc.DocumentObjectModel;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Charting;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using MigraDoc.RtfRendering;

namespace BlueSheetApp
{
    public partial class frmBlueSheet : Form
    {
        //private String strConnectionString = String.Empty;
        //private MySqlConnection mySqlConn = null;

        public enum EnumPaidTo { Member, MedicalProvider };

        EnumPaidTo PaidTo = EnumPaidTo.Member;

        public SortedField paidSortedField = null;
        public SortedField cmmPendingPaymentSortedField = null;
        public SortedField pendingSortedField = null;
        public SortedField ineligibleSortedField = null;

        public SortedField paidInPaidTabSortedField = null;
        public SortedField cmmCMMPendingPaymentInTabSortedField = null;
        public SortedField pendingInTabSortedField = null;
        public SortedField ineligibleInTabSortedField = null;

        public SortedField prSortedField = null;


        public String strIndividualID = null;
        public String strCheckNo = null;
        public List<Incident> lstIncidents = null;

        public CheckInfo ChkInfoEntered = null;
        public ACHInfo ACHInfoEntered = null;
        public CreditCardPaymentInfo CreditCardPaymentEntered = null;
        public PersonalResponsibilityTotalInfo PersonalResponsibilityTotalEntered = null;

        public Boolean bPaidHasRow = false;
        public Boolean bCMMPendingPaymentHasRow = false;
        public Boolean bPendingHasRow = false;
        public Boolean bIneligibleHasRow = false;

        public String strPrimaryName = String.Empty;
        public String strMembershipId = String.Empty;
        public String strIndividualName = String.Empty;
        public String strIndividualLastName = String.Empty;
        public String strIndividualMiddleName = String.Empty;
        public String strIndiviaualFirstName = String.Empty;
        public String strStreetAddress = String.Empty;
        public String strCity = String.Empty;
        public String strState = String.Empty;
        public String strZip = String.Empty;

        //public List<String> lstINCD = null;
        //public List<Incident> lstIncident = null;
        //public List<Incident> lstIncidentDistinct = null;

        public List<PersonalResponsibilityInfo> lstPersonalResponsibilityInfo;
        public List<SettlementIneligibleInfo> lstSettlementIneligibleInfo;

        //private string strUserName = "harrispark@kcj777.com";
        //private string strPassword = "%Speed5of2Light5%";

        private BlueSheetSForce.SforceService Sfdcbinding = null;
        private BlueSheetSForce.LoginResult CurrentLoginResult = null;

        //private String strGreetingMessage = "귀 회원이 제출한 의료비가 정산되었음을 알려 드립니다. 아래의 내용 중에 (1)잔액 또는 (2)보류(보류되는 사유)되고 있는 의료비가 " +
        //                                    "있다면 해당되는 자료를 정산서 발행일로부터 30일 이내에 제출하여 주시고, 만일 제출한 의료비의 정산에 착오가 있다면 사무실로 " +
        //                                    "연락하여 도움을 받으시기 바랍니다.\n";
        //private String strGreetingMessage2 = "기독의료상조회의 의료비 나눔 사역에 참여하여 주시는 귀회원의 가정과 사업에 우리 주 예수 그리스도의 은총이 가득하시기를 " +
        //                                     "기도드립니다.\n\n 감사합니다.";

        // Check, ACH, Credit Card
        private String strGreetingMessagePara1 = "귀 회원이 신청한 의료비가 정산되어 알려 드립니다. 아래의 ‘의료비 정산 내역서’를 확인하여 주시고, " +
                                                 "만일 정산된 금액이나 내용에 오류가 있다면 의료비지원부(NPD)로 알려 주십시오.";

        private String strGreetingMessagePara2 = "현재 ‘보류 중인 의료비’에 '잔액/보류'가 있다면 미비서류로 인해 의료비 진행이 지연되지 않도록 요청된 미비서류를 신속히 보내 주시기 바랍니다. " +
                                                 "기타 문의사항은 월-금(오전 9시- 오후 5:30 중부시간)까지 NPD 사무실로 연락하여 주십시오.";

        private String strGreetingMessagePara3 = "CMM 사역에 참여하여 주심에 진심으로 감사를 드리며, 오늘도 주 예수 그리스도의 평강이 함께 하시기를 기도합니다. ";
        private String strGreetingMessagePara4 = "감사합니다.";

        // Personal Responsibility Only
        private String strPRGreetingMessagePara1 = "신청하신 의료비가 정산되어 알려드립니다.";
        private String strPRGreetingMessagePara2 = "성경 말씀에 따라 의료비 나눔 사역을 실천하고 있는 기독의료상조회에는 서로의 짐을 나누어 지는(갈 6:2) " +
                                                   "‘의료비 나눔’과 각각 자기의 짐을 지는(갈 6:5) ‘본인 부담금’이 있습니다. ";
        private String strPRGreetingMessagePara3 = "현재 신청하신 의료비는 회원님이 가입하신 프로그램의  본인 부담금 또는 지원 불가 의료비에 해당하는 금액으로 본회의 가이드라인에 따라" +
                                                   " 지원이 되지 않음을 알려드립니다. ";
        private String strPRGreetingMessagePara4 = "\n아래  의료비 정산 내역을 확인하시고 정산된 금액이나 내용에 문의사항이 있으면 의료비 지원부 (773-777-8889 Ext. 5003)로 연락주시기 바랍니다.";
        private String strPRGreetingMessagePara5 = "CMM 사역에 참여하여 주셔서 진심으로 감사드리며,  주 예수 그리스도의 평강이 늘 함께 하시길 기도합니다.";
        private String strPRGreetingMessagePara6 = "감사합니다.";

        //private String strPRGreetingMessagePara2 = "현재 ‘보류 중인 의료비’에 '잔액/보류'가 있다면 미비서류로 인해 의료비 진행이 지연되지 않도록 요청된 미비서류를 신속히 보내 주시기 바랍니다. " +
        //                                         "기타 문의사항은 월-금(오전 9시- 오후 5:30 중부시간)까지 NPD 사무실로 연락하여 주십시오.";

        //private String strPRGreetingMessagePara3 = "CMM 사역에 참여하여 주심에 진심으로 감사를 드리며, 오늘도 주 예수 그리스도의 평강이 함께 하시기를 기도합니다. ";
        //private String strPRGreetingMessagePara4 = "감사합니다.";

        // English message
        private String strDearMember = "Dear ";
        private String strEnglishGreetingMessage1 = "We thank you for participating in our health care sharing ministry and pray that the grace of our Lord " +
                                                    "Jesus Christ overflows in you and your family as we continue to pray for your recovery.";

        private String strEnglishGreetingMessage2 = "Your medical bills that were submitted have been processed. Please carefully review your summary below." +
                                                    "\n\nIf there are any discrepancies, please notify our Needs Processing Department(NPD) immediately. " +
                                                    "If you have been notified by our NPD about incomplete documentation or outstanding balances, " +
                                                    "please submit the requested documents as soon as possible to avoid further delay of your needs sharing process.";

        private String strEnglishGreetingMessage3 = "\nIf you have any questions or concerns, please contact our NPD at 773-777-8889, Monday through Friday, from 9 AM to 5:30 PM (CST).";

        private String strEnglishGreetingMessage4 = "\nSincerely,";

        private String strEnglishPRGreetingMessage1 = "Christian Mutual Med-Aid(CMM) follows the Word of God.Together we share our brothers’ and sisters’ burdens as the Bible says " +
                                                      "in Galatian 6:2. Also, the Bible states that each person must carry their own load in Galatian 6:5.\n" +
                                                      "We define that load as Personal Responsibility (CMM Guidelines, Section VII. Needs Processing and Sharing, B. Personal Responsibility).\n" +
                                                      "Please see the table below for your convenience.";
        private String strEnglishPRGreetingMessage2 = "\nAccording to CMM Guidelines, your medical needs share request amount does not exceed the program’s Personal Responsibility. " +
                                                      "Therefore, we are not able to share your medical needs.";
        private String strEnglishPRGreetingMessage3 = "Should you have any questions regarding this, please contact the Needs Processing Department at 773-777-8889, Monday through Friday, " +
                                                      "from 9:00 AM to 5:30 PM, CST.";

        private String strEnglishPRGreetingMessage4 = "We thank you for participating in our health care sharing ministry and pray that the overflowing grace of our Lord Jesus be with you " +
                                                      "and your family.";

        private String strEnglishPRGreetingMessage5 = "\n\nSincerely,";

        private String strCMM_NeedProcessing = "NEEDS PROCESSING DEPARTMENT";
        //private String strNP_Phone_Fax_Email = "T.773-777-8889(EXT5003)\nF.773-777-0004 EMAIL:NPD@CMMLOGOS.ORG";
        //private String strNP_Phone_Fax_Email = "T.773-777-8889(EXT5003)\nnpd@cmmlogos.org";
        private String strNP_Phone_Fax_Email = "T.773-777-8889\nnpd@cmmlogos.org";

        public List<PaidMedicalExpenseTableRow> lstPaidMedicalExpenseTableRow = new List<PaidMedicalExpenseTableRow>();
        public List<CMMPendingPaymentTableRow> lstCMMPendingPaymentTableRow = new List<CMMPendingPaymentTableRow>();
        public List<PendingTableRow> lstPendingTableRow = new List<PendingTableRow>();
        public List<BillIneligibleTableRow> lstBillIneligibleTableRow = new List<BillIneligibleTableRow>();

        //public void InitializeSfdcbinding(String strUserName, String strPassword)
        //{
        //    Sfdcbinding = new BlueSheetSForce.SforceService();
        //    CurrentLoginResult = Sfdcbinding.login(strUserName, strPassword);
        //    Sfdcbinding.Url = CurrentLoginResult.serverUrl;
        //    Sfdcbinding.SessionHeaderValue = new BlueSheetSForce.SessionHeader();
        //    Sfdcbinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
        //}

        public frmBlueSheet()
        {
            InitializeComponent();

            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;

            paidSortedField = new SortedField();
            cmmPendingPaymentSortedField = new SortedField();
            pendingSortedField = new SortedField();
            ineligibleSortedField = new SortedField();

            paidInPaidTabSortedField = new SortedField();
            cmmCMMPendingPaymentInTabSortedField = new SortedField();
            pendingInTabSortedField = new SortedField();
            ineligibleInTabSortedField = new SortedField();

            prSortedField = new SortedField();

            //ChkInfoEntered = new CheckInfo();

            //lstIncidents = new List<Incident>();

            //dtpCreditCardPaymentDate.Format = DateTimePickerFormat.Custom;
            //dtpCreditCardPaymentDate.CustomFormat = " ";
            this.StartPosition = FormStartPosition.CenterScreen;
            rbCheck.Checked = true;
            txtCheckNo.Enabled = true;
            dtpCheckIssueDate.Enabled = true;

            rbACH.Checked = false;
            txtACH_No.Enabled = false;
            dtpACHDate.Enabled = false;

            rbCreditCard.Checked = false;
            txtCreditCardNo.Enabled = false;
            dtpCreditCardPaymentDate.Enabled = false;

            lstPersonalResponsibilityInfo = new List<PersonalResponsibilityInfo>();
            lstSettlementIneligibleInfo = new List<SettlementIneligibleInfo>();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;

            if (txtIndividualID.Text.Trim() == String.Empty)
            {
                MessageBox.Show("You haven't entered Individual ID.", "Error");
                return;
            }

            if (txtIncidentNo.Text.Trim() == String.Empty && rbCheck.Checked && txtCheckNo.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Please Enter Check No", "Error");
                return;
            }
            else strCheckNo = txtCheckNo.Text;

            if (txtIncidentNo.Text.Trim() == String.Empty && rbACH.Checked && txtACH_No.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Please Enter ACH No", "Error");
                return;
            }


            strPrimaryName = String.Empty;
            strMembershipId = String.Empty;
            strIndividualID = String.Empty;
            strIndividualName = String.Empty;
            strStreetAddress = String.Empty;
            strCity = String.Empty;
            strState = String.Empty;
            strZip = String.Empty;

            // Initialize HasRow boolean variables
            bPaidHasRow = false;
            bCMMPendingPaymentHasRow = false;
            bPendingHasRow = false;
            bIneligibleHasRow = false;

            gvBillPaid.DataSource = null;
            gvPaidInTabPaid.DataSource = null;    
            
            gvCMMPendingPayment.DataSource = null;
            gvCMMPendingInTab.DataSource = null;

            gvPending.DataSource = null;
            gvPendingInTab.DataSource = null;

            gvIneligible.DataSource = null;
            gvIneligibleInTab.DataSource = null;

            gvPersonalResponsibility.DataSource = null;
            gvIneligibleNoSharing.DataSource = null;

            BlueSheetSForce.Settlement__c settlement = null;
            List<BlueSheetSForce.Settlement__c> lstSettlement = new List<BlueSheetSForce.Settlement__c>();
            List<BlueSheetSForce.Settlement__c> lstDistinctSettlements = new List<BlueSheetSForce.Settlement__c>();

            lstIncidents = new List<Incident>();

            List<String> lstMedBillNames = new List<String>();
            List<String> lstDistinctMedBillNames = new List<String>();
            List<String> lstDistinctMedBillNoSharing = new List<String>();

            List<String> lstIncdNames = new List<String>();
            List<String> lstDistinctIncdNames = new List<String>();

            ChkInfoEntered = null;
            ACHInfoEntered = null;
            CreditCardPaymentEntered = null;

            if (rbNoSharingOnly.Checked)
            {

                DateTime? dtStartDate = null;
                DateTime? dtEndDate = null;

                String IndividualId = String.Empty;
                String IncidentNo = String.Empty;

                if (txtIndividualID.Text.Trim() != String.Empty && txtIncidentNo.Text.Trim() != String.Empty)
                {
                    IndividualId = txtIndividualID.Text.Trim();
                    IncidentNo = txtIncidentNo.Text.Trim();
                }
                else
                {
                    MessageBox.Show("You didn't entered Individual Id or Incident No.");
                    return;
                }

                frmStartEndDate frmStartEndDate = new frmStartEndDate();

                frmStartEndDate.StartPosition = FormStartPosition.CenterParent;

                if (frmStartEndDate.ShowDialog() == DialogResult.OK)
                {
                    dtStartDate = frmStartEndDate.StartDate;
                    dtEndDate = frmStartEndDate.EndDate;

                    String strSoqlMedicalBill = "select c4g_Incident__r.Name, c4g_Incident__r.Incident_Occurrence_Date__c, Name, Bill_Date__c, Medical_Provider__c, " +
                                                "c4g_Incident__r.c4g_ICD10_Code__r.Name " +
                                                "from Medical_Bill__c where c4g_Incident__r.Name like '%" + IncidentNo + "' and " +
                                                "c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + IndividualId + "' and " +
                                                "Bill_Date__c >= " + dtStartDate.Value.ToString("yyyy-MM-dd") + " and " +
                                                "Bill_Date__c <= " + dtEndDate.Value.ToString("yyyy-MM-dd");

                    BlueSheetSForce.QueryResult qrMedBill = Sfdcbinding.query(strSoqlMedicalBill);

                    if (qrMedBill.size > 0)
                    {

                        //savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_Ko";

                        //strIndividudalId = IndividualId;

                        strIndividualID = IndividualId;



                        //String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Medical_Provider__c, Check_Number__c, Check_Date__c, c4g_Amount__c, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName, " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c, " +
                        //    "c4g_Type__c from Settlement__c where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'CMM Provider Payment') and " +
                        //    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //    "Check_Number__c = '" + txtCheckNo.Text.Trim() + "' and " +
                        //    "Check_Date__c != null and " +
                        //    "Check_Date__c = " + dtpCheckIssueDate.Value.ToString("yyyy-MM-dd");

                        //if (strMembershipId != String.Empty) paraMembershipInfo.AddFormattedText(strMembershipId + " (" + strIndividualID + ")\n");
                        //else paraMembershipInfo.AddFormattedText(strIndividualID + "\n");
                        //paraMembershipInfo.AddFormattedText(strIndividualName + "\n");
                        //paraMembershipInfo.AddFormattedText(strStreetAddress + "\n");
                        //paraMembershipInfo.AddFormattedText(strCity + ", " + strState + " " + strZip + "\n");


                        //strPrimaryName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c;
                        //if (settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__c != null)
                        //{
                        //    strMembershipId = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name;
                        //}
                        //else strMembershipId = String.Empty;
                        //strIndividualName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                        //strIndividualLastName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName;
                        //strIndividualMiddleName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName;
                        //strIndiviaualFirstName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName;
                        //strIndividualID = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c;
                        //strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.street;
                        //strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.city;
                        //strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.state;
                        //strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.postalCode;






                        String strSoqlIndividualInfo = "select c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                                                       //"c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                                                       "c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress, " +
                                                       "c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                                                       "c4g_Incident__r.c4g_Contact__r.Name, " +
                                                       "c4g_Incident__r.c4g_Contact__r.LastName, " +
                                                       "c4g_Incident__r.c4g_Contact__r.MiddleName, " +
                                                       "c4g_Incident__r.c4g_Contact__r.FirstName, " +
                                                       "c4g_Incident__r.c4g_Contact__r.Individual_ID__c " +
                                                       "from Medical_Bill__c where " +
                                                       "c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + strIndividualID + "'";

                        BlueSheetSForce.QueryResult qrMedBillIndInfo = Sfdcbinding.query(strSoqlIndividualInfo);

                        if (qrMedBillIndInfo.size > 0)
                        {
                            BlueSheetSForce.Medical_Bill__c medbillIndInfo = qrMedBillIndInfo.records[0] as BlueSheetSForce.Medical_Bill__c;

                            strIndividualName = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Name;
                            strIndividualLastName = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.LastName;
                            strIndividualMiddleName = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.MiddleName;
                            strIndiviaualFirstName = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.FirstName;
                            strIndividualID = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Individual_ID__c;
                            //strStreetAddress = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.OtherAddress.street;
                            //strCity = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.OtherAddress.city;
                            //strState = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.OtherAddress.state;
                            //strZip = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.OtherAddress.postalCode;
                            strStreetAddress = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.street;
                            strCity = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.city;
                            strState = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.state;
                            strZip = medbillIndInfo.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.postalCode;

                        }

                        Cursor.Current = Cursors.WaitCursor;

                        lstPersonalResponsibilityInfo.Clear();

                        for (int i = 0; i < qrMedBill.size; i++)
                        {
                            BlueSheetSForce.Medical_Bill__c medbill = qrMedBill.records[i] as BlueSheetSForce.Medical_Bill__c;

                            if (i == 0)
                            {
                                PersonalResponsibilityTotalEntered = new PersonalResponsibilityTotalInfo();
                                PersonalResponsibilityTotalEntered.IncidentNo = medbill.c4g_Incident__r.Name;
                                //if (medbill.c4g_Incident__r.c4g_ICD10_Code__c != null)
                                if (medbill.c4g_Incident__r.c4g_ICD10_Code__r != null)
                                    PersonalResponsibilityTotalEntered.ICD10CodeDescription = medbill.c4g_Incident__r.c4g_ICD10_Code__r.Name;
                                else
                                    PersonalResponsibilityTotalEntered.ICD10CodeDescription = String.Empty;

                                PersonalResponsibilityTotalEntered.IncidentOccurrenceDate = medbill.c4g_Incident__r.Incident_Occurrence_Date__c.Value;
                            }
                            lstMedBillNames.Add(medbill.Name);
                        }

                        foreach (String medbillName in lstMedBillNames.Distinct())
                        {
                            lstDistinctMedBillNames.Add(medbillName);
                        }

                        Boolean bNoSharedAmount = false;

                        foreach (String medbillName in lstDistinctMedBillNames)
                        {
                            String strSoqlMedBillNoSharedAmount = "select Name from Medical_Bill__c where Name = '" + medbillName + "' and c4g_Total_Shared_Amount__c <= 0 and " +
                                                                  "Bill_Status__c != 'Ineligible' and Bill_Status__c != 'Pending' and Bill_Status__c != 'CMM Pending Payment'";

                            BlueSheetSForce.QueryResult qrMedBillNoSharing = Sfdcbinding.query(strSoqlMedBillNoSharedAmount);

                            if (qrMedBillNoSharing.size > 0)
                            {
                                bNoSharedAmount = true;

                                BlueSheetSForce.Medical_Bill__c medbillNoSharing = qrMedBillNoSharing.records[0] as BlueSheetSForce.Medical_Bill__c;


                                String strSoqlSettlementNoSharing = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Bill_Date__c, c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                                                    "c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Type__c, c4g_Personal_Responsibility_Type__c, c4g_Amount__c " +
                                                                    "from Settlement__c " +
                                                                    "where c4g_Medical_Bill__r.Name = '" + medbillNoSharing.Name + "' and c4g_Type__c = 'Personal Responsibility' and " +
                                                                    "(c4g_Personal_Responsibility_Type__c = 'Member Payment' or " +
                                                                    "c4g_Personal_Responsibility_Type__c = 'Member Discount' or " +
                                                                    "c4g_Personal_Responsibility_Type__c= 'Third-Party Discount')";

                                PersonalResponsibilityInfo prInfo = new PersonalResponsibilityInfo();

                                BlueSheetSForce.QueryResult qrSettlementNoSharing = Sfdcbinding.query(strSoqlSettlementNoSharing);

                                if (qrSettlementNoSharing.size > 0)
                                {
                                    for (int j = 0; j < qrSettlementNoSharing.size; j++)
                                    {
                                        BlueSheetSForce.Settlement__c settlementPersonalResponsibility = qrSettlementNoSharing.records[j] as BlueSheetSForce.Settlement__c;

                                        if (j == 0)
                                        {
                                            prInfo.MedBillName = settlementPersonalResponsibility.c4g_Medical_Bill__r.Name;
                                            prInfo.BillDate = settlementPersonalResponsibility.c4g_Medical_Bill__r.Bill_Date__c;
                                            prInfo.MedicalProvider = settlementPersonalResponsibility.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                            prInfo.BillAmount = settlementPersonalResponsibility.c4g_Medical_Bill__r.c4g_Bill_Amount__c.Value;
                                            prInfo.Type = settlementPersonalResponsibility.c4g_Type__c;
                                        }
                                        prInfo.PersonalResponsibilityType = settlementPersonalResponsibility.c4g_Personal_Responsibility_Type__c;
                                        if (prInfo.PersonalResponsibilityType == "Member Payment")
                                        {
                                            prInfo.MemberPayment = settlementPersonalResponsibility.c4g_Amount__c.Value;
                                            prInfo.PersonalResponsibilityTotal += prInfo.MemberPayment.Value;
                                        }
                                        if (prInfo.PersonalResponsibilityType == "Member Discount")
                                        {
                                            prInfo.MemberDiscount = settlementPersonalResponsibility.c4g_Amount__c.Value;
                                            prInfo.PersonalResponsibilityTotal += prInfo.MemberDiscount.Value;
                                        }
                                        if (prInfo.PersonalResponsibilityType == "Third-Party Discount")
                                        {
                                            prInfo.ThirdPartyDiscount = settlementPersonalResponsibility.c4g_Amount__c.Value;
                                            prInfo.PersonalResponsibilityTotal += prInfo.ThirdPartyDiscount.Value;
                                        }
                                    }
                                    lstPersonalResponsibilityInfo.Add(prInfo);
                                }

                            }
                        }

                        Boolean bIneligibleMedBill = false;

                        lstSettlementIneligibleInfo.Clear();

                        foreach (String medbillName in lstDistinctMedBillNames)
                        {
                            //String strSoqlMedBillIneligible = "select Name from Medical_Bill__c where Name = '" + medbillName + "' and c4g_Total_Shared_Amount__c <= 0 and " +
                            //                                      "Bill_Status__c = 'Ineligible'";

                            String strSoqlMedBillIneligible = "select Name from Medical_Bill__c where Name = '" + medbillName + "' and c4g_Total_Shared_Amount__c <= 0 and " +
                                                              "(Bill_Status__c = 'Closed' or Bill_Status__c = 'Ineligible')";

                            BlueSheetSForce.QueryResult qrMedBillIneligible = Sfdcbinding.query(strSoqlMedBillIneligible);

                            if (qrMedBillIneligible.size > 0)
                            {
                                bIneligibleMedBill = true;

                                for (int i = 0; i < qrMedBillIneligible.size; i++)
                                {

                                    BlueSheetSForce.Medical_Bill__c medBillIneligible = qrMedBillIneligible.records[i] as BlueSheetSForce.Medical_Bill__c;

                                    //String strSoqlIneligibleSettlements = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Bill_Date__c, " +
                                    //                                      "c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                    //                                      "c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Amount__c, " +
                                    //                                      "c4g_Medical_Bill__r.Ineligible_Reason__c " +
                                    //                                      "from Settlement__c " +
                                    //                                      "where c4g_Medical_Bill__r.Name = '" + medBillIneligible.Name + "' and " +
                                    //                                      "c4g_Medical_Bill__r.Bill_Status__c = 'Ineligible' and " +
                                    //                                      "c4g_Amount__c > 0";

                                    String strSoqlIneligibleSettlements = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Bill_Date__c, " +
                                                                              "c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                                                              "c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Amount__c, " +
                                                                              "c4g_Medical_Bill__r.Ineligible_Reason__c " +
                                                                              "from Settlement__c " +
                                                                              "where c4g_Type__c = 'Ineligible' and " +
                                                                              //"where " +
                                                                              "c4g_Medical_Bill__r.Name = '" + medBillIneligible.Name + "' and " +
                                                                              "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' or " +
                                                                              "c4g_Medical_Bill__r.Bill_Status__c = 'Ineligible') and " +
                                                                              "c4g_Amount__c >= 0";

                                    BlueSheetSForce.QueryResult qrSettlementMedBillIneligible = Sfdcbinding.query(strSoqlIneligibleSettlements);

                                    if (qrSettlementMedBillIneligible.size > 0)
                                    {
                                        BlueSheetSForce.Settlement__c settlementIneligible = qrSettlementMedBillIneligible.records[0] as BlueSheetSForce.Settlement__c;

                                        SettlementIneligibleInfo info = new SettlementIneligibleInfo();
                                        info.MedBillName = settlementIneligible.c4g_Medical_Bill__r.Name;
                                        info.BillDate = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c;
                                        info.MedicalProvider = settlementIneligible.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                        info.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c.Value;
                                        info.IneligibleAmount = settlementIneligible.c4g_Amount__c;
                                        info.IneligibleReason = settlementIneligible.c4g_Medical_Bill__r.Ineligible_Reason__c;

                                        lstSettlementIneligibleInfo.Add(info);
                                    }
                                }
                            }

                            
                        }

                        //foreach (String medbillName in lstDistinctMedBillNames)
                        //{
                        //    String strSoqlMedBillNoSharedAmount = "select Name from Medical_Bill__c where Name = '" + medbillName + "' and c4g_Total_Shared_Amount__c <= 0 and " +
                        //                                          "Bill_Status__c != 'Ineligible' and Bill_Status__c != 'Pending' and Bill_Status__c != 'CMM Pending Payment'";

                        //    BlueSheetSForce.QueryResult qrMedBillNoSharing = Sfdcbinding.query(strSoqlMedBillNoSharedAmount);

                        //    if (qrMedBillNoSharing.size > 0)
                        //    {
                        //        bNoSharedAmount = true;

                        //        BlueSheetSForce.Medical_Bill__c medbillNoSharing = qrMedBillNoSharing.records[0] as BlueSheetSForce.Medical_Bill__c;

                        //        String strSoqlSettlementIneligible = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Bill_Date__c, c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                        //                                             "c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Type__c, c4g_Personal_Responsibility_Type__c, c4g_Amount__c, " +
                        //                                             "c4g_Medical_Bill__r.Ineligible_Reason__c " +
                        //                                             "from Settlement__c " +
                        //                                             "where c4g_Medical_Bill__r.Name = '" + medbillNoSharing.Name + "' and c4g_Type__c = 'Ineligible'";

                        //        SettlementIneligibleInfo settlementInfo = new SettlementIneligibleInfo();

                        //        BlueSheetSForce.QueryResult qrSettlementIneligible = Sfdcbinding.query(strSoqlSettlementIneligible);

                        //        if (qrSettlementIneligible.size > 0)
                        //        {
                        //            for(int j = 0; j < qrSettlementIneligible.size; j++)
                        //            {
                        //                BlueSheetSForce.Settlement__c settlementIneligible = qrSettlementIneligible.records[j] as BlueSheetSForce.Settlement__c;

                        //                if (j == 0)
                        //                {
                        //                    settlementInfo.MedBillName = settlementIneligible.c4g_Medical_Bill__r.Name;
                        //                    settlementInfo.BillDate = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c;
                        //                    settlementInfo.MedicalProvider = settlementIneligible.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                        //                    settlementInfo.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c.Value;
                        //                    settlementInfo.Type = settlementIneligible.c4g_Type__c;
                        //                    settlementInfo.IneligibleReason = settlementIneligible.c4g_Medical_Bill__r.Ineligible_Reason__c;
                        //                }

                        //                settlementInfo.IneligibleAmount += settlementIneligible.c4g_Amount__c.Value;
                        //            }
                        //            lstSettlementIneligibleInfo.Add(settlementInfo);
                        //        }
                        //    }
                        //}

                        if (!bNoSharedAmount)
                        {
                            MessageBox.Show("These Bills have shared amount. Choose differet start date or end date, or choose to paid BlueSheet.");
                            return;
                        }

                        ///////////////////////////////////////////////////////////
                        ///

                        if (lstPersonalResponsibilityInfo.Count > 0)
                        {
                            DataTable dtPersonalResponsibility = new DataTable();
                            dtPersonalResponsibility.Columns.Add("MEDBILL", typeof(String));
                            dtPersonalResponsibility.Columns.Add("서비스 날짜", typeof(String));
                            dtPersonalResponsibility.Columns.Add("의료기관명", typeof(String));
                            dtPersonalResponsibility.Columns.Add("청구액(원금)", typeof(String));
                            dtPersonalResponsibility.Columns.Add("Type", typeof(String));
                            dtPersonalResponsibility.Columns.Add("PR Type: Member Payment", typeof(String));
                            dtPersonalResponsibility.Columns.Add("PR Type: Member Discount", typeof(String));
                            dtPersonalResponsibility.Columns.Add("PR Type: 3rd Party Discount", typeof(String));
                            dtPersonalResponsibility.Columns.Add("Personal Responsibility Total", typeof(String));

                            Double Zero = 0;

                            foreach (PersonalResponsibilityInfo info in lstPersonalResponsibilityInfo)
                            {
                                DataRow row = dtPersonalResponsibility.NewRow();

                                row["MEDBILL"] = info.MedBillName;
                                row["서비스 날짜"] = info.BillDate.Value.ToString("MM/dd/yyyy");
                                row["의료기관명"] = info.MedicalProvider;
                                row["청구액(원금)"] = info.BillAmount.ToString("C");
                                row["Type"] = info.Type;
                                if (info.MemberPayment != null) row["PR Type: Member Payment"] = info.MemberPayment.Value.ToString("C");
                                else row["PR Type: Member Payment"] = Zero.ToString("C");
                                if (info.MemberDiscount != null) row["PR Type: Member Discount"] = info.MemberDiscount.Value.ToString("C");
                                else row["PR Type: Member Discount"] = Zero.ToString("C");
                                if (info.ThirdPartyDiscount != null) row["PR Type: 3rd Party Discount"] = info.ThirdPartyDiscount.Value.ToString("C");
                                else row["PR Type: 3rd Party Discount"] = Zero.ToString("C");
                                row["Personal Responsibility Total"] = info.PersonalResponsibilityTotal.ToString("C");

                                dtPersonalResponsibility.Rows.Add(row);
                            }

                            //if (lstPersonalResponsibilityInfo.Count > 0)
                            //{
                            double sumBillAmount = 0;
                            double sumMemberPayment = 0;
                            double sumMemberDiscount = 0;
                            double sumThirdPartyDiscount = 0;
                            double sumPersonalResponsibilityTotal = 0;

                            for (int k = 0; k < lstPersonalResponsibilityInfo.Count; k++)
                            {
                                sumBillAmount += lstPersonalResponsibilityInfo[k].BillAmount;
                                if (lstPersonalResponsibilityInfo[k].MemberPayment != null) sumMemberPayment += lstPersonalResponsibilityInfo[k].MemberPayment.Value;
                                if (lstPersonalResponsibilityInfo[k].MemberDiscount != null) sumMemberDiscount += lstPersonalResponsibilityInfo[k].MemberDiscount.Value;
                                if (lstPersonalResponsibilityInfo[k].ThirdPartyDiscount != null) sumThirdPartyDiscount += lstPersonalResponsibilityInfo[k].ThirdPartyDiscount.Value;
                                sumPersonalResponsibilityTotal += lstPersonalResponsibilityInfo[k].PersonalResponsibilityTotal;
                            }

                            DataRow sumRow = dtPersonalResponsibility.NewRow();
                            sumRow["의료기관명"] = "합계";
                            sumRow["청구액(원금)"] = sumBillAmount.ToString("C");
                            sumRow["PR Type: Member Payment"] = sumMemberPayment.ToString("C");
                            sumRow["PR Type: Member Discount"] = sumMemberDiscount.ToString("C");
                            sumRow["PR Type: 3rd Party Discount"] = sumThirdPartyDiscount.ToString("C");
                            sumRow["Personal Responsibility Total"] = sumPersonalResponsibilityTotal.ToString("C");

                            dtPersonalResponsibility.Rows.Add(sumRow);

                            PersonalResponsibilityTotalEntered.PersonalResponsibilityTotal = Decimal.Parse(sumPersonalResponsibilityTotal.ToString());

                            //}


                            //gvBillPaid.Rows.Clear();
                            //gvCMMPendingPayment.Rows.Clear();
                            //gvPending.Rows.Clear();
                            //gvIneligible.Rows.Clear();
                            //gvSummary.Rows.Clear();
                            //gvPaidInTabPaid.Rows.Clear();
                            //gvCMMPendingInTab.Rows.Clear();
                            //gvPendingInTab.Rows.Clear();
                            //gvIneligibleInTab.Rows.Clear();

                            //gvPaidInTabPaid.Rows.Clear();
                            //gvCMMPendingInTab.Rows.Clear();
                            //gvPendingInTab.Rows.Clear();
                            //gvIneligibleInTab.Rows.Clear();

                            gvPersonalResponsibility.DataSource = dtPersonalResponsibility;

                            gvPersonalResponsibility.Columns["MEDBILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvPersonalResponsibility.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvPersonalResponsibility.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvPersonalResponsibility.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable; ;
                            gvPersonalResponsibility.Columns["Type"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvPersonalResponsibility.Columns["PR Type: Member Payment"].SortMode = DataGridViewColumnSortMode.NotSortable;
                            gvPersonalResponsibility.Columns["PR Type: Member Discount"].SortMode = DataGridViewColumnSortMode.NotSortable;
                            gvPersonalResponsibility.Columns["PR Type: 3rd Party Discount"].SortMode = DataGridViewColumnSortMode.NotSortable;
                            gvPersonalResponsibility.Columns["Personal Responsibility Total"].SortMode = DataGridViewColumnSortMode.NotSortable;

                            gvPersonalResponsibility.Columns["MEDBILL"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["서비스 날짜"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["의료기관명"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["청구액(원금)"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["Type"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["PR Type: Member Payment"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["PR Type: Member Discount"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["PR Type: 3rd Party Discount"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["Personal Responsibility Total"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                            gvPersonalResponsibility.Columns["MEDBILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvPersonalResponsibility.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            gvPersonalResponsibility.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            gvPersonalResponsibility.Columns["Type"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            gvPersonalResponsibility.Columns["PR Type: Member Payment"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            gvPersonalResponsibility.Columns["PR Type: Member Discount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            gvPersonalResponsibility.Columns["PR Type: 3rd Party Discount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            gvPersonalResponsibility.Columns["Personal Responsibility Total"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                            gvPersonalResponsibility.Columns["MEDBill"].Width = 150;
                            gvPersonalResponsibility.Columns["서비스 날짜"].Width = 100;
                            gvPersonalResponsibility.Columns["의료기관명"].Width = 200;
                            gvPersonalResponsibility.Columns["청구액(원금)"].Width = 100;
                            gvPersonalResponsibility.Columns["Type"].Width = 180;
                            gvPersonalResponsibility.Columns["PR Type: Member Payment"].Width = 180;
                            gvPersonalResponsibility.Columns["PR Type: Member Discount"].Width = 180;
                            gvPersonalResponsibility.Columns["PR Type: 3rd Party Discount"].Width = 180;
                            gvPersonalResponsibility.Columns["Personal Responsibility Total"].Width = 180;
                        }


                        if (lstSettlementIneligibleInfo.Count > 0)
                        {

                            DataTable dtMedicalBillNoSharingNoPR = new DataTable();
                            dtMedicalBillNoSharingNoPR.Columns.Add("MEDBILL", typeof(String));
                            dtMedicalBillNoSharingNoPR.Columns.Add("서비스 날짜", typeof(String));
                            dtMedicalBillNoSharingNoPR.Columns.Add("의료기관명", typeof(String));
                            dtMedicalBillNoSharingNoPR.Columns.Add("청구액(원금)", typeof(String));
                            //dtMedicalBillNoSharingNoPR.Columns.Add("Type", typeof(String));
                            dtMedicalBillNoSharingNoPR.Columns.Add("지원불가 의료비", typeof(String));
                            dtMedicalBillNoSharingNoPR.Columns.Add("지원불가 사유", typeof(String));

                            Double ZeroNoPRNoSharing = 0;

                            //if (lstSettlementIneligibleInfo.Count > 0)
                            //{
                            foreach (SettlementIneligibleInfo settlementInfo in lstSettlementIneligibleInfo)
                            {
                                DataRow row = dtMedicalBillNoSharingNoPR.NewRow();

                                row["MEDBILL"] = settlementInfo.MedBillName;
                                row["서비스 날짜"] = settlementInfo.BillDate.Value.ToString("MM/dd/yyyy");
                                row["의료기관명"] = settlementInfo.MedicalProvider;
                                row["청구액(원금)"] = settlementInfo.BillAmount.ToString("C");
                                //row["Type"] = settlementInfo.Type;
                                if (settlementInfo.IneligibleAmount != null) row["지원불가 의료비"] = settlementInfo.IneligibleAmount.Value.ToString("C");
                                else row["지원불가 의료비"] = ZeroNoPRNoSharing.ToString("C");
                                row["지원불가 사유"] = settlementInfo.IneligibleReason;

                                dtMedicalBillNoSharingNoPR.Rows.Add(row);
                            }
                            //}

                            //if (lstSettlementIneligibleInfo.Count > 0)
                            //{
                            Double? sumBillAmount = 0;
                            Double? sumIneligibleAmount = 0;

                            foreach (SettlementIneligibleInfo settlementInfo in lstSettlementIneligibleInfo)
                            {
                                sumBillAmount += settlementInfo.BillAmount;
                                sumIneligibleAmount += settlementInfo.IneligibleAmount;
                            }

                            DataRow sumRow = dtMedicalBillNoSharingNoPR.NewRow();
                            sumRow["의료기관명"] = "합계";
                            sumRow["청구액(원금)"] = sumBillAmount.Value.ToString("C");
                            sumRow["지원불가 의료비"] = sumIneligibleAmount.Value.ToString("C");

                            dtMedicalBillNoSharingNoPR.Rows.Add(sumRow);
                            //}

                            gvIneligibleNoSharing.DataSource = dtMedicalBillNoSharingNoPR;

                            gvIneligibleNoSharing.Columns["MEDBILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvIneligibleNoSharing.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvIneligibleNoSharing.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvIneligibleNoSharing.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                            //gvIneligibleNoSharing.Columns["Type"].SortMode = DataGridViewColumnSortMode.Programmatic;
                            gvIneligibleNoSharing.Columns["지원불가 의료비"].SortMode = DataGridViewColumnSortMode.NotSortable;
                            gvIneligibleNoSharing.Columns["지원불가 사유"].SortMode = DataGridViewColumnSortMode.NotSortable;

                            gvIneligibleNoSharing.Columns["MEDBILL"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["서비스 날짜"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["의료기관명"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["청구액(원금)"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            //gvIneligibleNoSharing.Columns["Type"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["지원불가 의료비"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["지원불가 사유"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                            gvIneligibleNoSharing.Columns["MEDBILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            gvIneligibleNoSharing.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            gvIneligibleNoSharing.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            //gvIneligibleNoSharing.Columns["Type"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            gvIneligibleNoSharing.Columns["지원불가 의료비"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            gvIneligibleNoSharing.Columns["지원불가 사유"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                            gvIneligibleNoSharing.Columns["MEDBILL"].Width = 150;
                            gvIneligibleNoSharing.Columns["서비스 날짜"].Width = 100;
                            gvIneligibleNoSharing.Columns["의료기관명"].Width = 200;
                            gvIneligibleNoSharing.Columns["청구액(원금)"].Width = 100;
                            //gvIneligibleNoSharing.Columns["Type"].Width = 100;
                            gvIneligibleNoSharing.Columns["지원불가 의료비"].Width = 100;
                            gvIneligibleNoSharing.Columns["지원불가 사유"].Width = 200;
                        }

                        tabMedicalExpense.SelectedIndex = 5;

                        frmLoadingFinished loadingFinished = new frmLoadingFinished();
                        loadingFinished.StartPosition = FormStartPosition.CenterParent;
                        loadingFinished.ShowDialog();
                        Cursor.Current = Cursors.Default;
                        return;
                    }
                    else
                    {
                        MessageBox.Show("No med bill found. Try different INCD, IndNo, Start Date, or End Date.");
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            if (!rbNoSharingOnly.Checked)
            {
                if (rbCheck.Checked)
                {
                    ChkInfoEntered = new CheckInfo();

                    txtACH_No.Text = String.Empty;
                    dtpCreditCardPaymentDate.Value = DateTime.Today;

                    if (txtCheckNo.Text.Trim() != String.Empty)
                    {
                        String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Medical_Provider__c, Check_Number__c, Check_Date__c, c4g_Amount__c, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                                                    //"c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c, " +
                                                    "c4g_Type__c from Settlement__c " +
                                                    "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement' or c4g_Type__c = 'CMM Provider Payment') and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "Check_Number__c = '" + txtCheckNo.Text.Trim() + "' and " +
                                                    "Check_Date__c != null and " +
                                                    "Check_Date__c = " + dtpCheckIssueDate.Value.ToString("yyyy-MM-dd");


                        //String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, Check_Number__c, Check_Date__c, c4g_Amount__c, c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c " +
                        //                            "from Settlement__c where c4g_Type__c = 'CMM Member Reimbursement' and " +
                        //                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //                            "Check_Number__c = '" + txtCheckNo.Text.Trim() + "' and " +
                        //                            "Approved__c = true and Check_Date__c != null";

                        BlueSheetSForce.QueryResult qrSettlement = Sfdcbinding.query(strSoqlMedicalBill);

                        if (qrSettlement.size > 0)
                        {
                            for (int i = 0; i < qrSettlement.size; i++)
                            {
                                settlement = qrSettlement.records[i] as BlueSheetSForce.Settlement__c;
                                ChkInfoEntered.CheckAmount += settlement.c4g_Amount__c;

                                if (i == 0)
                                {
                                    ChkInfoEntered.CheckNumber = settlement.Check_Number__c;
                                    ChkInfoEntered.dtCheckIssueDate = settlement.Check_Date__c.Value;
                                    //ChkInfoEntered.PaidTo = settlement.c4g_Medical_Bill__r.Medical_Provider__c;

                                    //txtCheckIssueDate.Text = settlement.Check_Date__c.Value.ToLongDateString();
                                    dtpCheckIssueDate.Text = settlement.Check_Date__c.Value.ToString("MM/dd/yyyy");

                                    strPrimaryName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c;
                                    if (settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__c != null)
                                    {
                                        strMembershipId = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name;
                                    }
                                    else strMembershipId = String.Empty;
                                    strIndividualName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                    strIndividualLastName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName;
                                    strIndividualMiddleName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName;
                                    strIndiviaualFirstName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName;
                                    strIndividualID = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c;
                                    //strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.street;
                                    //strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.city;
                                    //strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.state;
                                    //strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.postalCode;

                                    strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.street;
                                    strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.city;
                                    strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.state;
                                    strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.postalCode;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Enter Check No", "Error");
                        return;
                    }


                    if (txtCheckNo.Text.Trim() != String.Empty)
                    {
                        String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name, c4g_Type__c, c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name from Settlement__c " +
                                                    "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement' or c4g_Type__c = 'CMM Provider Payment') and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "Check_Number__c = '" + txtCheckNo.Text.Trim() + "' and " +
                                                    "Check_Date__c != null and " +
                                                    "Check_Date__c = " + dtpCheckIssueDate.Value.ToString("yyyy-MM-dd");

                        //String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name from Settlement__c where " +
                        //                          "c4g_Type__c = 'CMM Member Reimbursement' and " +
                        //                          "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //                          "Check_Number__c = '" + txtCheckNo.Text.Trim() + "' and " +
                        //                          "Approved__c = true and Check_Date__c != null";

                        BlueSheetSForce.QueryResult qrIncidents = Sfdcbinding.query(strSoqlIncidents);

                        if (qrIncidents.size > 0)
                        {
                            for (int i = 0; i < qrIncidents.size; i++)
                            {

                                BlueSheetSForce.Settlement__c settlementIncident = qrIncidents.records[i] as BlueSheetSForce.Settlement__c;
                                if (settlementIncident.c4g_Type__c == "CMM Member Reimbursement" ||
                                    settlementIncident.c4g_Type__c == "PR reimbursement")
                                {
                                    PaidTo = EnumPaidTo.Member;
                                    ChkInfoEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                }
                                else if (settlementIncident.c4g_Type__c == "CMM Provider Payment")
                                {
                                    PaidTo = EnumPaidTo.MedicalProvider;
                                    ChkInfoEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                }
                                lstIncdNames.Add(settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.Name);
                            }
                        }
                        else
                        {
                            MessageBox.Show("No Incident Found");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Enter Check No", "Error");
                        return;
                    }

                    if (lstIncdNames.Count > 0)
                    {
                        foreach (String strIncdName in lstIncdNames.Distinct())
                        {
                            lstDistinctIncdNames.Add(strIncdName);
                        }
                        lstDistinctIncdNames.Sort();
                    }

                    if (lstDistinctIncdNames.Count > 0)
                    {
                        foreach (String IncdName in lstDistinctIncdNames)
                        {
                            String strSoqlMedBills = "select Name from Medical_Bill__c where c4g_Incident__r.Name = '" + IncdName + "'";

                            BlueSheetSForce.QueryResult qrMedBills = Sfdcbinding.query(strSoqlMedBills);

                            if (qrMedBills.size > 0)
                            {
                                for (int i = 0; i < qrMedBills.size; i++)
                                {
                                    BlueSheetSForce.Medical_Bill__c med_bill = qrMedBills.records[i] as BlueSheetSForce.Medical_Bill__c;

                                    lstMedBillNames.Add(med_bill.Name);
                                }
                            }
                        }
                    }
                    if (lstMedBillNames.Count > 0)
                    {
                        foreach (String strMedBillName in lstMedBillNames.Distinct())
                        {
                            lstDistinctMedBillNames.Add(strMedBillName);
                        }
                        lstDistinctMedBillNames.Sort();
                    }
                }

                if (rbACH.Checked)
                {
                    ACHInfoEntered = new ACHInfo();

                    //if (txtCheckNo.Text.Trim() != String.Empty)
                    if (txtACH_No.Text.Trim() != String.Empty)
                    {
                        String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Medical_Provider__c,  ACH_Number__c, ACH_Date__c, c4g_Amount__c, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                                                    //"c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c, " +
                                                    "c4g_Type__c from Settlement__c " +
                                                    "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement' or c4g_Type__c = 'CMM Provider Payment') and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "ACH_Number__c = '" + txtACH_No.Text.Trim() + "' and " +
                                                    "ACH_Date__c != null and " +
                                                    "ACH_Date__c = " + dtpACHDate.Value.ToString("yyyy-MM-dd");
                        //"ACH_Date__c != null";

                        //String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, ACH_Number__c, ACH_Date__c, c4g_Amount__c, c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c from Settlement__c " +
                        //        "where c4g_Type__c = 'CMM Member Reimbursement' and " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //        "ACH_Number__c = '" + txtACH_No.Text.Trim() + "' and " +
                        //        "Approved__c = true and ACH_Date__c != null";

                        BlueSheetSForce.QueryResult qrSettlement = Sfdcbinding.query(strSoqlMedicalBill);

                        if (qrSettlement.size > 0)
                        {
                            for (int i = 0; i < qrSettlement.size; i++)
                            {
                                settlement = qrSettlement.records[i] as BlueSheetSForce.Settlement__c;
                                //ChkInfoEntered.CheckAmount += settlement.c4g_Amount__c;
                                ACHInfoEntered.ACHAmount += settlement.c4g_Amount__c;

                                if (i == 0)
                                {
                                    //ChkInfoEntered.CheckNumber = settlement.Check_Number__c;
                                    //ChkInfoEntered.dtCheckIssueDate = settlement.Check_Date__c.Value;
                                    ACHInfoEntered.ACHNumber = settlement.ACH_Number__c;
                                    ACHInfoEntered.dtACHDate = settlement.ACH_Date__c.Value;
                                    //ACHInfoEntered.PaidTo = settlement.c4g_Medical_Bill__r.Medical_Provider__r.Name;

                                    //txtCheckIssueDate.Text = settlement.Check_Date__c.Value.ToLongDateString();
                                    //txtTransactionDate.Text = settlement.ACH_Date__c.Value.ToLongDateString();
                                    dtpACHDate.Text = settlement.ACH_Date__c.Value.ToString("MM/dd/yyyy");

                                    strPrimaryName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c;
                                    if (settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__c != null)
                                    {
                                        strMembershipId = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name;
                                    }
                                    else strMembershipId = String.Empty;
                                    strIndividualName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                    strIndividualLastName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName;
                                    strIndividualMiddleName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName;
                                    strIndiviaualFirstName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName;
                                    strIndividualID = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c;
                                    //strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.street;
                                    //strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.city;
                                    //strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.state;
                                    //strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.postalCode;

                                    strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.street;
                                    strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.city;
                                    strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.state;
                                    strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.postalCode;

                                }
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please Enter ACH No", "Error");
                        return;
                    }

                    if (txtACH_No.Text.Trim() != String.Empty)
                    {
                        String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name, c4g_Type__c, c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name from Settlement__c " +
                                                    "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement' or c4g_Type__c = 'CMM Provider Payment') and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "ACH_Number__c = '" + txtACH_No.Text.Trim() + "' and " +
                                                    "ACH_Date__c != null and " +
                                                    "ACH_Date__c = " + dtpACHDate.Value.ToString("yyyy-MM-dd");
                        //"ACH_Date__c != null";


                        //String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name from Settlement__c " +
                        //      "where c4g_Type__c = 'CMM Member Reimbursement' and " +
                        //      "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //      "ACH_Number__c = '" + txtACH_No.Text.Trim() + "' and " +
                        //      "Approved__c = true and ACH_Date__c != null";


                        BlueSheetSForce.QueryResult qrIncidents = Sfdcbinding.query(strSoqlIncidents);

                        if (qrIncidents.size > 0)
                        {
                            for (int i = 0; i < qrIncidents.size; i++)
                            {

                                BlueSheetSForce.Settlement__c settlementIncident = qrIncidents.records[i] as BlueSheetSForce.Settlement__c;
                                if ((settlementIncident.c4g_Type__c == "CMM Member Reimbursement")||
                                    (settlementIncident.c4g_Type__c == "PR reimbursement"))
                                {
                                    PaidTo = EnumPaidTo.Member;
                                    ACHInfoEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                }
                                else if (settlementIncident.c4g_Type__c == "CMM Provider Payment")
                                {
                                    PaidTo = EnumPaidTo.MedicalProvider;
                                    ACHInfoEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                }
                                lstIncdNames.Add(settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.Name);
                            }
                        }
                        else
                        {
                            MessageBox.Show("No Incident Found");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Enter ACH No", "Error");
                        return;
                    }

                    if (lstIncdNames.Count > 0)
                    {
                        foreach (String strIncdName in lstIncdNames.Distinct())
                        {
                            lstDistinctIncdNames.Add(strIncdName);
                        }
                        lstDistinctIncdNames.Sort();
                    }

                    if (lstDistinctIncdNames.Count > 0)
                    {
                        foreach (String IncdName in lstDistinctIncdNames)
                        {
                            String strSoqlMedBills = "select Name from Medical_Bill__c where c4g_Incident__r.Name = '" + IncdName + "'";

                            BlueSheetSForce.QueryResult qrMedBills = Sfdcbinding.query(strSoqlMedBills);

                            if (qrMedBills.size > 0)
                            {
                                for (int i = 0; i < qrMedBills.size; i++)
                                {
                                    BlueSheetSForce.Medical_Bill__c med_bill = qrMedBills.records[i] as BlueSheetSForce.Medical_Bill__c;

                                    lstMedBillNames.Add(med_bill.Name);
                                }
                            }
                        }
                    }
                    if (lstMedBillNames.Count > 0)
                    {
                        foreach (String strMedBillName in lstMedBillNames.Distinct())
                        {
                            lstDistinctMedBillNames.Add(strMedBillName);
                        }
                        lstDistinctMedBillNames.Sort();
                    }
                }

                if (rbCreditCard.Checked)
                {
                    CreditCardPaymentEntered = new CreditCardPaymentInfo();

                    DateTime dtYesterday = DateTime.Parse(dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd"));
                    DateTime dtTomorrow = DateTime.Parse(dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd")).AddDays(1);

                    if (dtpCreditCardPaymentDate.Value.ToString() != String.Empty)
                    {
                        String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, c4g_Medical_Bill__r.Medical_Provider__c, CMM_Credit_Card_Paid_day__c, c4g_Amount__c, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                                                    //"c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c, " +
                                                    "c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "CMM_Credit_Card_Paid_day__c != null and " +
                                                    "CMM_Credit_Card_Paid_day__c = " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd");
                        //"CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                        //"CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd");

                        //String strSoqlMedicalBill = "select c4g_Medical_Bill__r.Name, CMM_Credit_Card_Paid_day__c, c4g_Amount__c, c4g_Medical_Bill__r.c4g_Incident__r.Name, Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c " +
                        //        "from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                        //        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //        "Approved__c = true and " +
                        //        "CMM_Credit_Card_Paid_day__c != null and " +
                        //        "CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                        //        "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd");




                        BlueSheetSForce.QueryResult qrSettlement = Sfdcbinding.query(strSoqlMedicalBill);

                        if (qrSettlement.size > 0)
                        {
                            for (int i = 0; i < qrSettlement.size; i++)
                            {
                                settlement = qrSettlement.records[i] as BlueSheetSForce.Settlement__c;
                                CreditCardPaymentEntered.CCPaymentAmount += settlement.c4g_Amount__c;

                                if (i == 0)
                                {
                                    //CreditCardPaymentEntered.dtPaymentDate = settlement.CMM_Credit_Card_Paid_day__c.Value;
                                    CreditCardPaymentEntered.dtPaymentDate = dtpCreditCardPaymentDate.Value;
                                    //CreditCardPaymentEntered.PaidTo = settlement.c4g_Medical_Bill__r.Medical_Provider__c;

                                    //txtCheckIssueDate.Text = settlement.Check_Date__c.Value.ToLongDateString();
                                    //txtTransactionDate.Text = settlement.ACH_Date__c.Value.ToLongDateString();

                                    strPrimaryName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Primary_Name__c;
                                    if (settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__c != null)
                                    {
                                        strMembershipId = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name;
                                    }
                                    else strMembershipId = String.Empty;
                                    strIndividualName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                    strIndividualLastName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.LastName;
                                    strIndividualMiddleName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.MiddleName;
                                    strIndiviaualFirstName = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.FirstName;
                                    strIndividualID = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c;
                                    //strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.street;
                                    //strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.city;
                                    //strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.state;
                                    //strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.OtherAddress.postalCode;

                                    strStreetAddress = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.street;
                                    strCity = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.city;
                                    strState = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.state;
                                    strZip = settlement.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Account.ShippingAddress.postalCode;

                                }
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please Enter Credit Card Payment Date", "Error");
                        return;
                    }

                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    if (dtpCreditCardPaymentDate.Value.ToString() != String.Empty)
                    {
                        String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name, c4g_Type__c, c4g_Medical_Bill__r.Medical_Provider__r.Name, " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name from Settlement__c " +
                                                    "where c4g_Type__c = 'CMM Provider Payment' and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                                                    "CMM_Credit_Card_Paid_day__c != null and " +
                                                    "CMM_Credit_Card_Paid_day__c = " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd");
                        //"CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                        //"CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd");

                        //String strSoqlIncidents = "select c4g_Medical_Bill__r.c4g_Incident__r.Name from Settlement__c " +
                        //      "where c4g_Type__c = 'CMM Provider Payment' and " +
                        //      "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Individual_ID__c like '%" + txtIndividualID.Text.Trim() + "' and " +
                        //      "Approved__c = true and " +
                        //      "CMM_Credit_Card_Paid_day__c != null and " +
                        //      "CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                        //      "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd");

                        BlueSheetSForce.QueryResult qrIncidents = Sfdcbinding.query(strSoqlIncidents);

                        if (qrIncidents.size > 0)
                        {
                            for (int i = 0; i < qrIncidents.size; i++)
                            {

                                BlueSheetSForce.Settlement__c settlementIncident = qrIncidents.records[i] as BlueSheetSForce.Settlement__c;
                                if (settlementIncident.c4g_Type__c == "CMM Member Reimbursement")
                                {
                                    PaidTo = EnumPaidTo.Member;
                                    CreditCardPaymentEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                }
                                else if (settlementIncident.c4g_Type__c == "CMM Provider Payment")
                                {
                                    PaidTo = EnumPaidTo.MedicalProvider;
                                    CreditCardPaymentEntered.PaidTo = settlementIncident.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                }
                                lstIncdNames.Add(settlementIncident.c4g_Medical_Bill__r.c4g_Incident__r.Name);
                            }
                        }
                        else
                        {
                            MessageBox.Show("No Incident Found");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please Enter Credit Card Payment Date", "Error");
                        return;
                    }

                    if (lstIncdNames.Count > 0)
                    {
                        foreach (String strIncdName in lstIncdNames.Distinct())
                        {
                            lstDistinctIncdNames.Add(strIncdName);
                        }
                        lstDistinctIncdNames.Sort();
                    }

                    if (lstDistinctIncdNames.Count > 0)
                    {
                        foreach (String IncdName in lstDistinctIncdNames)
                        {
                            String strSoqlMedBills = "select Name from Medical_Bill__c where c4g_Incident__r.Name = '" + IncdName + "'";

                            BlueSheetSForce.QueryResult qrMedBills = Sfdcbinding.query(strSoqlMedBills);

                            if (qrMedBills.size > 0)
                            {
                                for (int i = 0; i < qrMedBills.size; i++)
                                {
                                    BlueSheetSForce.Medical_Bill__c med_bill = qrMedBills.records[i] as BlueSheetSForce.Medical_Bill__c;

                                    lstMedBillNames.Add(med_bill.Name);
                                }
                            }
                        }
                    }
                    if (lstMedBillNames.Count > 0)
                    {
                        foreach (String strMedBillName in lstMedBillNames.Distinct())
                        {
                            lstDistinctMedBillNames.Add(strMedBillName);
                        }
                        lstDistinctMedBillNames.Sort();
                    }
                }
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                lstIncidents.Clear();

                foreach (String strIncdName in lstDistinctIncdNames)
                {
                    String strSoqlICD10Codes = "select Name, c4g_Contact__r.Name , c4g_ICD10_Code__r.Name from Incident__c where Name = '" + strIncdName + "'";

                    BlueSheetSForce.QueryResult qrICD10Codes = Sfdcbinding.query(strSoqlICD10Codes);

                    if (qrICD10Codes.size > 0)
                    {
                        for (int i = 0; i < qrICD10Codes.size; i++)
                        {
                            BlueSheetSForce.Incident__c incd = qrICD10Codes.records[i] as BlueSheetSForce.Incident__c;

                            if (incd.c4g_ICD10_Code__r != null)
                            {
                                lstIncidents.Add(new Incident(incd.Name, incd.c4g_Contact__r.Name, incd.c4g_ICD10_Code__r.Name));
                            }
                            else if (incd.c4g_ICD10_Code__r == null)
                            {
                                lstIncidents.Add(new Incident(incd.Name, incd.c4g_Contact__r.Name, ""));
                            }
                        }
                    }
                }

                lstIncidents.Sort(delegate (Incident incd1, Incident incd2)
                {
                    if (incd1.Name == null && incd2.Name == null) return 0;
                    else if (incd1.Name == null) return -1;
                    else if (incd2.Name == null) return 1;
                    else return incd1.Name.CompareTo(incd2.Name);
                });

                DataTable dtMedicalBillPaid = new DataTable();

                dtMedicalBillPaid.Columns.Add("INCD", typeof(String));
                dtMedicalBillPaid.Columns.Add("회원 이름", typeof(String));
                dtMedicalBillPaid.Columns.Add("MED_BILL", typeof(String));
                dtMedicalBillPaid.Columns.Add("서비스 날짜", typeof(String));
                dtMedicalBillPaid.Columns.Add("의료기관명", typeof(String));
                dtMedicalBillPaid.Columns.Add("청구액(원금)", typeof(String));
                dtMedicalBillPaid.Columns.Add("본인 부담금", typeof(String));
                dtMedicalBillPaid.Columns.Add("회원할인", typeof(String));
                dtMedicalBillPaid.Columns.Add("CMM 할인", typeof(String));
                dtMedicalBillPaid.Columns.Add("의료기관 지불금", typeof(String));
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    dtMedicalBillPaid.Columns.Add("기지급액", typeof(String));
                    dtMedicalBillPaid.Columns.Add("회원 환불금", typeof(String));
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    dtMedicalBillPaid.Columns.Add("기지급액 (의료기관)", typeof(String));
                    dtMedicalBillPaid.Columns.Add("기지급액 (회원)", typeof(String));
                }
                dtMedicalBillPaid.Columns.Add("잔액/보류", typeof(String));

                DataRow drClosed = null;

                List<String> lstIncidentNames = new List<String>();

                List<MedicalExpense> lstMedicalExpense = new List<MedicalExpense>();

                List<String> lstMedicalBillPaid = new List<String>();


                foreach (String strMedBillName in lstDistinctMedBillNames)
                {

                    String strToday = dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd");

                    String strSoqlSettlementPaid = String.Empty;

                    if (rbCheck.Checked)
                    {
                        strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                "((Check_Number__c != null and Check_Date__c != null) or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";

                        //strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                        //                        "((Approved__c = true and Check_Number__c != null and Check_Date__c != null) or " +
                        //                        "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                        //                        "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";
                    }
                    if (rbACH.Checked)
                    {
                        strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                "((ACH_Number__c != null and ACH_Date__c != null) or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";

                        //strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                        //    "((Approved__c = true and ACH_Number__c != null and ACH_Date__c != null) or " +
                        //    "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                        //    "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";
                    }
                    if (rbCreditCard.Checked)
                    {
                        strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                "(CMM_Credit_Card_Paid_day__c != null or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                                                "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";

                        //strSoqlSettlementPaid = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                        //    "((Approved__c = true and CMM_Credit_Card_Paid_day__c != null) or " +
                        //    "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and c4g_Type__c = 'Personal Responsibility') or " +
                        //    "(c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and c4g_Type__c = 'Personal Responsibility'))";
                    }

                    BlueSheetSForce.QueryResult qrSettlementPaid = Sfdcbinding.query(strSoqlSettlementPaid);

                    if (qrSettlementPaid.size > 0)
                    {
                        for (int i = 0; i < qrSettlementPaid.size; i++)
                        {
                            BlueSheetSForce.Settlement__c settlementPaid = qrSettlementPaid.records[i] as BlueSheetSForce.Settlement__c;

                            lstMedicalBillPaid.Add(settlementPaid.c4g_Medical_Bill__r.Name);
                        }
                    }
                }

                List<String> lstDistinctMedicalBillPaid = new List<String>();

                foreach (String MedicalBillName in lstMedicalBillPaid.Distinct())
                {
                    lstDistinctMedicalBillPaid.Add(MedicalBillName);
                }


                foreach (String strMedBillName in lstDistinctMedicalBillPaid)
                {

                    String strSoqlBillPaid = "select c4g_Incident__r.Name, c4g_Incident__r.c4g_Contact__r.Name, c4g_Incident__r.c4g_ICD10_Code__r.Name, " +
                                                "Name, Bill_Date__c, Medical_Provider__r.Name, c4g_Bill_Amount__c, Personal_Responsibility_Credit__c, c4g_Balance__c " +
                                                "from Medical_Bill__c where Name = '" + strMedBillName + "'";

                    BlueSheetSForce.QueryResult qrMedicalBillsForCheck = Sfdcbinding.query(strSoqlBillPaid);

                    if (qrMedicalBillsForCheck.size > 0)
                    {
                        bPaidHasRow = true;

                        for (int j = 0; j < qrMedicalBillsForCheck.size; j++)
                        {
                            BlueSheetSForce.Medical_Bill__c medBill = qrMedicalBillsForCheck.records[j] as BlueSheetSForce.Medical_Bill__c;

                            lstIncidentNames.Add(medBill.c4g_Incident__r.Name);  // get incident name to retrieve cmm pending, pending, and ineligible

                            MedicalExpense expense = new MedicalExpense();

                            drClosed = dtMedicalBillPaid.NewRow();
                            drClosed["INCD"] = medBill.c4g_Incident__r.Name.Substring(5);
                            drClosed["회원 이름"] = medBill.c4g_Incident__r.c4g_Contact__r.Name;
                            drClosed["MED_BILL"] = medBill.Name.Substring(8);
                            drClosed["서비스 날짜"] = medBill.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                            drClosed["의료기관명"] = medBill.Medical_Provider__r.Name;
                            expense.BillAmount = medBill.c4g_Bill_Amount__c.Value;
                            drClosed["청구액(원금)"] = expense.BillAmount.Value.ToString("C");
                            //expense.PersonalResponsibility = medBill.Personal_Responsibility_Credit__c.Value;
                            //drClosed["본인 부담금"] = expense.PersonalResponsibility.Value.ToString("C");
                            expense.Balance = medBill.c4g_Balance__c.Value;
                            drClosed["잔액/보류"] = expense.Balance.Value.ToString("C");


                            String strSoqlPersonalResponsibility = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'Personal Responsibility' and " +
                                                                    "(c4g_Personal_Responsibility_Type__c = 'Member Payment' or c4g_Personal_Responsibility_Type__c = 'Third-Party Discount') and " +
                                                                    "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "'";

                            BlueSheetSForce.QueryResult qrPersonalResponsibility = Sfdcbinding.query(strSoqlPersonalResponsibility);

                            double? PersonalResponsibility = 0;

                            if (qrPersonalResponsibility.size > 0)
                            {
                                for (int k = 0; k < qrPersonalResponsibility.size; k++)
                                {
                                    BlueSheetSForce.Settlement__c personal_responsibility = qrPersonalResponsibility.records[k] as BlueSheetSForce.Settlement__c;
                                    PersonalResponsibility += personal_responsibility.c4g_Amount__c.Value;
                                }

                                expense.PersonalResponsibility = PersonalResponsibility.Value;
                                drClosed["본인 부담금"] = expense.PersonalResponsibility.Value.ToString("C");
                            }
                            else if (qrPersonalResponsibility.size == 0)
                            {
                                expense.PersonalResponsibility = 0;
                                drClosed["본인 부담금"] = expense.PersonalResponsibility.Value.ToString("C");
                            }


                            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            String strSqlMemberDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where ((c4g_Type__c = 'Member Discount' or c4g_Type__c = 'Third-Party Discount') or " +
                                                            "(c4g_Type__c = 'Personal Responsibility' and c4g_Personal_Responsibility_Type__c = 'Member Discount')) and " +
                                                            "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                            "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "'";
                            //"Approved__c = true and Check_Date__c <> null";

                            BlueSheetSForce.QueryResult qrMemberDiscount = Sfdcbinding.query(strSqlMemberDiscount);

                            double? MemberDiscount = 0;

                            if (qrMemberDiscount.size > 0)
                            {
                                for (int k = 0; k < qrMemberDiscount.size; k++)
                                {
                                    BlueSheetSForce.Settlement__c member_discount = qrMemberDiscount.records[k] as BlueSheetSForce.Settlement__c;
                                    MemberDiscount += member_discount.c4g_Amount__c.Value;
                                }

                                expense.MemberDiscount = MemberDiscount.Value;
                                drClosed["회원할인"] = expense.MemberDiscount.Value.ToString("C");
                            }
                            else if (qrMemberDiscount.size == 0)
                            {
                                expense.MemberDiscount = 0;
                                drClosed["회원할인"] = expense.MemberDiscount.Value.ToString("C");
                            }

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            String strSqlCMMDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Discount' and " +
                                                        "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                        "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "'";

                            BlueSheetSForce.QueryResult qrCMMDiscount = Sfdcbinding.query(strSqlCMMDiscount);

                            double? CMMDiscount = 0;

                            if (qrCMMDiscount.size > 0)
                            {
                                for (int k = 0; k < qrCMMDiscount.size; k++)
                                {
                                    BlueSheetSForce.Settlement__c cmm_discount = qrCMMDiscount.records[k] as BlueSheetSForce.Settlement__c;
                                    CMMDiscount += cmm_discount.c4g_Amount__c.Value;
                                }

                                expense.CMMDiscount = CMMDiscount.Value;
                                drClosed["CMM 할인"] = expense.CMMDiscount.Value.ToString("C");
                            }
                            else if (qrCMMDiscount.size == 0)
                            {
                                expense.CMMDiscount = 0;
                                drClosed["CMM 할인"] = expense.CMMDiscount.Value.ToString("C");
                            }

                            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            String strSqlCMMProviderPayment = String.Empty;

                            BlueSheetSForce.QueryResult qrCMMProviderPayment = null;

                            if (rbCheck.Checked || rbACH.Checked)
                            {
                                strSqlCMMProviderPayment = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                                            "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                            "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                            "((Check_Date__c != null and Check_Number__c != null) or " +
                                                            "(ACH_Date__c != null and ACH_Number__c != null) or " +
                                                            "CMM_Credit_Card_Paid_day__c != null)";

                                qrCMMProviderPayment = Sfdcbinding.query(strSqlCMMProviderPayment);

                                // strSqlCMMProviderPayment = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                //"c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                //"c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                //"((Approved__c = true and (Check_Date__c != null and Check_Number__c != null)) or " +
                                //"(Approved__c = true and (ACH_Date__c != null and ACH_Number__c != null)) or " +
                                //"(Approved__c = true and (CMM_Credit_Card_Paid_day__c != null)))";
                            }

                            //String strTmpCreditCardPaymentDate = String.Empty;
                            if (rbCreditCard.Checked)
                            {

                                strSqlCMMProviderPayment = "select c4g_Amount__c, c4g_Type__c, CMM_Credit_Card_Paid_day__c from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                                            "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                            "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                            "(CMM_Credit_Card_Paid_day__c != null and " +
                                                            "CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                                                            "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd") + ")";

                                qrCMMProviderPayment = Sfdcbinding.query(strSqlCMMProviderPayment);


                                //strSqlCMMProviderPayment = "select c4g_Amount__c, c4g_Type__c, CMM_Credit_Card_Paid_day__c from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                //                           "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                //                           "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                //                           "((Check_Date__c != null and Check_Number__c != null) or " +
                                //                           "(ACH_Date__c != null and ACH_Number__c != null) or " +
                                //                           "(CMM_Credit_Card_Paid_day__c != null and " +
                                //                           "CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + " and " +
                                //                           "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.AddDays(1).ToString("yyyy-MM-dd") + "))";
                            }


                            double? CMMProviderPayment = 0;

                            if (qrCMMProviderPayment.size > 0)
                            {
                                for (int k = 0; k < qrCMMProviderPayment.size; k++)
                                {
                                    BlueSheetSForce.Settlement__c cmm_provider_payment = qrCMMProviderPayment.records[k] as BlueSheetSForce.Settlement__c;
                                    CMMProviderPayment += cmm_provider_payment.c4g_Amount__c;
                                }

                                expense.CMMProviderPayment = CMMProviderPayment.Value;
                                drClosed["의료기관 지불금"] = expense.CMMProviderPayment.Value.ToString("C");
                            }
                            else if (qrCMMProviderPayment.size == 0)
                            {
                                expense.CMMProviderPayment = 0;
                                drClosed["의료기관 지불금"] = expense.CMMProviderPayment.Value.ToString("C");
                            }
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            //String strSqlPastReimbursement = String.Empty;

                            //if (rbCheck.Checked || rbACH.Checked)
                            if (PaidTo == EnumPaidTo.Member)
                            {
                                String strSqlPastReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c " +
                                                                 "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement') and " +
                                                                 "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                 "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                                 "(Check_Date__c != null or ACH_Date__c != null) and " +
                                                                 "(Check_Number__c != '" + txtCheckNo.Text.Trim() + "' or " +
                                                                 "ACH_Number__c != '" + txtACH_No.Text.Trim() + "')";

                                BlueSheetSForce.QueryResult qrPastMemberReimbursement = Sfdcbinding.query(strSqlPastReimbursement);

                                double? PastMemberReimbursement = 0;

                                if (qrPastMemberReimbursement.size > 0)
                                {
                                    for (int k = 0; k < qrPastMemberReimbursement.size; k++)
                                    {
                                        BlueSheetSForce.Settlement__c past_member_reimbursement = qrPastMemberReimbursement.records[k] as BlueSheetSForce.Settlement__c;
                                        PastMemberReimbursement += past_member_reimbursement.c4g_Amount__c;
                                    }
                                    // add code to pub past member reimbursement to table row
                                    expense.PastReimbursement = PastMemberReimbursement.Value;
                                    drClosed["기지급액"] = expense.PastReimbursement.Value.ToString("C");
                                }
                                else if (qrPastMemberReimbursement.size == 0)
                                {
                                    expense.PastReimbursement = 0;
                                    drClosed["기지급액"] = expense.PastReimbursement.Value.ToString("C");
                                }

                                //  strSqlPastReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c where c4g_Type__c = 'CMM Member Reimbursement' and " +
                                //"(CMM_Payment_Method__c = 'Check' or CMM_Payment_Method__c = 'ACH/Banking') and " +
                                //"c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                //"c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                //"(Check_Date__c != null or ACH_Date__c != null or CMM_Credit_Card_Paid_day__c != null) and " +
                                //"(Check_Number__c != '" + txtCheckNo.Text.Trim() + "' or " +
                                //"ACH_Number__c != '" + txtACH_No.Text.Trim() + "')";

                            }

                            //if (rbCheck.Checked || rbACH.Checked)
                            if (PaidTo == EnumPaidTo.Member)
                            {
                                String strSqlReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c " +
                                                             "where (c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement') and " +
                                                             "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                             "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                             "((Check_Date__c != null and Check_Number__c = '" + txtCheckNo.Text.Trim() + "') or " +
                                                             "(ACH_Date__c != null and ACH_Number__c = '" + txtACH_No.Text.Trim() + "'))";

                                //String strSqlReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c where c4g_Type__c = 'CMM Member Reimbursement' and " +
                                //                             "(CMM_Payment_Method__c = 'Check' or CMM_Payment_Method__c = 'ACH/Banking') and " +
                                //                             "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                //                             "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                //                             "((Check_Date__c != null and Check_Number__c = '" + txtCheckNo.Text.Trim() + "') or " +
                                //                             "(ACH_Date__c != null and ACH_Number__c = '" + txtACH_No.Text.Trim() + "'))";

                                BlueSheetSForce.QueryResult qrMemberReimbursement = Sfdcbinding.query(strSqlReimbursement);

                                double? MemberReimbursement = 0;

                                if (qrMemberReimbursement.size > 0)
                                {
                                    for (int k = 0; k < qrMemberReimbursement.size; k++)
                                    {
                                        BlueSheetSForce.Settlement__c member_reimbursement = qrMemberReimbursement.records[k] as BlueSheetSForce.Settlement__c;
                                        MemberReimbursement += member_reimbursement.c4g_Amount__c;
                                    }

                                    expense.Reimbursement = MemberReimbursement.Value;
                                    drClosed["회원 환불금"] = expense.Reimbursement.Value.ToString("C");
                                }
                                else if (qrMemberReimbursement.size == 0)
                                {
                                    expense.Reimbursement = 0;
                                    drClosed["회원 환불금"] = expense.Reimbursement.Value.ToString("C");
                                }
                            }

                            //if (rbCreditCard.Checked)
                            if (PaidTo == EnumPaidTo.MedicalProvider)
                            {
                                String strSqlPastCMMProviderPayment = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                                                        "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                        "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                                        "(Check_Date__c != null or ACH_Date__c != null or CMM_Credit_Card_Paid_day__c != null) and " +
                                                                        "(Check_Number__c != '" + txtCheckNo.Text.Trim() + "' or " +
                                                                        "ACH_Number__c != '" + txtACH_No.Text.Trim() + "' or " +
                                                                        "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + ")";

                                BlueSheetSForce.QueryResult qrPastCMMProviderPayment = Sfdcbinding.query(strSqlPastCMMProviderPayment);

                                double? PastCMMProviderPayment = 0;

                                if (qrPastCMMProviderPayment.size > 0)
                                {
                                    for (int k = 0; k < qrPastCMMProviderPayment.size; k++)
                                    {
                                        BlueSheetSForce.Settlement__c past_member_reimbursement = qrPastCMMProviderPayment.records[k] as BlueSheetSForce.Settlement__c;
                                        PastCMMProviderPayment += past_member_reimbursement.c4g_Amount__c;
                                    }
                                    // add code to pub past member reimbursement to table row

                                    //dtMedicalBillPaid.Columns.Add("기지급액 (의료기관)", typeof(String));
                                    //dtMedicalBillPaid.Columns.Add("기지급액 (회원)", typeof(String));

                                    //expense.PastReimbursement = PastCMMProviderPayment.Value;
                                    expense.PastCMMProviderPayment = PastCMMProviderPayment.Value;
                                    //drClosed["기지급액 (의료기관)"] = expense.PastReimbursement.Value.ToString("C");
                                    drClosed["기지급액 (의료기관)"] = expense.PastCMMProviderPayment.Value.ToString("C");
                                }
                                else if (qrPastCMMProviderPayment.size == 0)
                                {
                                    //expense.PastReimbursement = 0;
                                    expense.PastCMMProviderPayment = 0;
                                    //drClosed["기지급액 (의료기관)"] = expense.PastReimbursement.Value.ToString("C");
                                    drClosed["기지급액 (의료기관)"] = expense.PastCMMProviderPayment.Value.ToString("C");
                                }

                                String strSqlPastReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c where " +
                                                                 "(c4g_Type__c = 'CMM Member Reimbursement' or c4g_Type__c = 'PR reimbursement') and " +
                                                                 "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                 "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                                                 "(Check_Date__c != null or ACH_Date__c != null) and " +
                                                                 "(Check_Number__c != '" + txtCheckNo.Text.Trim() + "' or " +
                                                                 "ACH_Number__c != '" + txtACH_No.Text.Trim() + "')";

                                BlueSheetSForce.QueryResult qrPastMemberReimbursement = Sfdcbinding.query(strSqlPastReimbursement);

                                double? PastMemberReimbursement = 0;

                                if (qrPastMemberReimbursement.size > 0)
                                {
                                    for (int k = 0; k < qrPastMemberReimbursement.size; k++)
                                    {
                                        BlueSheetSForce.Settlement__c past_member_reimbursement = qrPastMemberReimbursement.records[k] as BlueSheetSForce.Settlement__c;
                                        PastMemberReimbursement += past_member_reimbursement.c4g_Amount__c;
                                    }
                                    // add code to pub past member reimbursement to table row
                                    expense.PastReimbursement = PastMemberReimbursement.Value;
                                    drClosed["기지급액 (회원)"] = expense.PastReimbursement.Value.ToString("C");
                                }
                                else if (qrPastMemberReimbursement.size == 0)
                                {
                                    expense.PastReimbursement = 0;
                                    drClosed["기지급액 (회원)"] = expense.PastReimbursement.Value.ToString("C");
                                }

                                //strSqlPastReimbursement = "select c4g_Amount__c, c4g_Type__c, Name from Settlement__c where c4g_Type__c = 'CMM Provider Payment' and " +
                                //                          "c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                //                          "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + medBill.c4g_Incident__r.Name + "' and " +
                                //                          "(Check_Date__c != null or ACH_Date__c != null or CMM_Credit_Card_Paid_day__c != null) and " +
                                //                          "(Check_Number__c != '" + txtCheckNo.Text.Trim() + "' or " +
                                //                          "ACH_Number__c != '" + txtACH_No.Text.Trim() + "' or " +
                                //                          "CMM_Credit_Card_Paid_day__c < " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + ")";

                            }


                            //BlueSheetSForce.QueryResult qrPastMemberReimbursement = Sfdcbinding.query(strSqlPastReimbursement);

                            //double? PastMemberReimbursement = 0;

                            //if (qrPastMemberReimbursement.size > 0)
                            //{
                            //    for (int k = 0; k < qrPastMemberReimbursement.size; k++)
                            //    {
                            //        BlueSheetSForce.Settlement__c past_member_reimbursement = qrPastMemberReimbursement.records[k] as BlueSheetSForce.Settlement__c;
                            //        PastMemberReimbursement += past_member_reimbursement.c4g_Amount__c;
                            //    }
                            //    // add code to pub past member reimbursement to table row
                            //    expense.PastReimbursement = PastMemberReimbursement.Value;
                            //    drClosed["기지급액"] = expense.PastReimbursement.Value.ToString("C");
                            //}
                            //else if (qrPastMemberReimbursement.size == 0)
                            //{
                            //    expense.PastReimbursement = 0;
                            //    drClosed["기지급액"] = expense.PastReimbursement.Value.ToString("C");
                            //}

                            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



                            dtMedicalBillPaid.Rows.Add(drClosed);
                            lstMedicalExpense.Add(expense);
                        }
                    }
                }
                if (bPaidHasRow)
                {
                    double sumClosedBillAmount = 0;
                    double sumClosedPersonalResponsibility = 0;
                    double sumClosedMemberDiscount = 0;
                    double sumClosedCMMDiscount = 0;
                    double sumClosedCMMProviderPayment = 0;
                    double sumClosedPastCMMProviderPayment = 0;
                    double sumClosedPastReimbursement = 0;
                    double sumClosedReimbursement = 0;
                    double sumClosedBalance = 0;

                    foreach (MedicalExpense expense in lstMedicalExpense)
                    {
                        sumClosedBillAmount += expense.BillAmount.Value;
                        sumClosedPersonalResponsibility += expense.PersonalResponsibility.Value;
                        sumClosedMemberDiscount += expense.MemberDiscount.Value;
                        sumClosedCMMDiscount += expense.CMMDiscount.Value;
                        sumClosedCMMProviderPayment += expense.CMMProviderPayment.Value;
                        sumClosedPastCMMProviderPayment += expense.PastCMMProviderPayment.Value;
                        sumClosedPastReimbursement += expense.PastReimbursement.Value;
                        sumClosedReimbursement += expense.Reimbursement.Value;
                        sumClosedBalance += expense.Balance.Value;
                    }

                    DataRow drClosedSum = dtMedicalBillPaid.NewRow();

                    drClosedSum["의료기관명"] = "합계";

                    drClosedSum["청구액(원금)"] = sumClosedBillAmount.ToString("C");
                    drClosedSum["본인 부담금"] = sumClosedPersonalResponsibility.ToString("C");
                    drClosedSum["회원할인"] = sumClosedMemberDiscount.ToString("C");
                    drClosedSum["CMM 할인"] = sumClosedCMMDiscount.ToString("C");
                    drClosedSum["의료기관 지불금"] = sumClosedCMMProviderPayment.ToString("C");
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        drClosedSum["기지급액"] = sumClosedPastReimbursement.ToString("C");
                        drClosedSum["회원 환불금"] = sumClosedReimbursement.ToString("C");
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        drClosedSum["기지급액 (의료기관)"] = sumClosedPastCMMProviderPayment.ToString("C");
                        drClosedSum["기지급액 (회원)"] = sumClosedPastReimbursement.ToString("C");
                    }
                    drClosedSum["잔액/보류"] = sumClosedBalance.ToString("C");

                    dtMedicalBillPaid.Rows.Add(drClosedSum);

                    gvBillPaid.DataSource = null;
                    gvBillPaid.DataSource = dtMedicalBillPaid;

                    gvBillPaid.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvBillPaid.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvBillPaid.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvBillPaid.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvBillPaid.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvBillPaid.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvBillPaid.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvBillPaid.Columns["본인 부담금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvBillPaid.Columns["회원할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvBillPaid.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvBillPaid.Columns["의료기관 지불금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvBillPaid.Columns["기지급액"].SortMode = DataGridViewColumnSortMode.NotSortable;
                        gvBillPaid.Columns["회원 환불금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvBillPaid.Columns["기지급액 (의료기관)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                        gvBillPaid.Columns["기지급액 (회원)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    gvBillPaid.Columns["잔액/보류"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    DataTable dt = dtMedicalBillPaid.Copy();
                    gvPaidInTabPaid.DataSource = dt;

                    gvPaidInTabPaid.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPaidInTabPaid.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPaidInTabPaid.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPaidInTabPaid.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPaidInTabPaid.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvPaidInTabPaid.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPaidInTabPaid.Columns["본인 부담금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPaidInTabPaid.Columns["회원할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPaidInTabPaid.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPaidInTabPaid.Columns["의료기관 지불금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvPaidInTabPaid.Columns["기지급액"].SortMode = DataGridViewColumnSortMode.NotSortable;
                        gvPaidInTabPaid.Columns["회원 환불금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvPaidInTabPaid.Columns["기지급액 (의료기관)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                        gvPaidInTabPaid.Columns["기지급액 (회원)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    gvPaidInTabPaid.Columns["잔액/보류"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    foreach (DataGridViewColumn col in gvBillPaid.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    foreach (DataGridViewColumn col in gvPaidInTabPaid.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvBillPaid.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvBillPaid.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvBillPaid.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvBillPaid.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvBillPaid.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvBillPaid.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvBillPaid.Columns["본인 부담금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvBillPaid.Columns["회원할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvBillPaid.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvBillPaid.Columns["의료기관 지불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvBillPaid.Columns["기지급액"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        gvBillPaid.Columns["회원 환불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvBillPaid.Columns["기지급액 (의료기관)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        gvBillPaid.Columns["기지급액 (회원)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    gvBillPaid.Columns["잔액/보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    gvBillPaid.Columns["INCD"].Width = 80;
                    gvBillPaid.Columns["회원 이름"].Width = 150;
                    gvBillPaid.Columns["MED_BILL"].Width = 80;
                    gvBillPaid.Columns["서비스 날짜"].Width = 100;
                    gvBillPaid.Columns["의료기관명"].Width = 200;
                    gvBillPaid.Columns["청구액(원금)"].Width = 100;
                    gvBillPaid.Columns["본인 부담금"].Width = 100;
                    gvBillPaid.Columns["회원할인"].Width = 80;
                    gvBillPaid.Columns["CMM 할인"].Width = 100;
                    gvBillPaid.Columns["의료기관 지불금"].Width = 120;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvBillPaid.Columns["기지급액"].Width = 100;
                        gvBillPaid.Columns["회원 환불금"].Width = 120;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvBillPaid.Columns["기지급액 (의료기관)"].Width = 140;
                        gvBillPaid.Columns["기지급액 (회원)"].Width = 100;
                    }
                    gvBillPaid.Columns["잔액/보류"].Width = 100;

                    gvPaidInTabPaid.Columns["INCD"].Width = 80;
                    gvPaidInTabPaid.Columns["회원 이름"].Width = 150;
                    gvPaidInTabPaid.Columns["MED_BILL"].Width = 80;
                    gvPaidInTabPaid.Columns["서비스 날짜"].Width = 100;
                    gvPaidInTabPaid.Columns["의료기관명"].Width = 200;
                    gvPaidInTabPaid.Columns["청구액(원금)"].Width = 100;
                    gvPaidInTabPaid.Columns["본인 부담금"].Width = 100;
                    gvPaidInTabPaid.Columns["회원할인"].Width = 80;
                    gvPaidInTabPaid.Columns["CMM 할인"].Width = 100;
                    gvPaidInTabPaid.Columns["의료기관 지불금"].Width = 120;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvPaidInTabPaid.Columns["기지급액"].Width = 100;
                        gvPaidInTabPaid.Columns["회원 환불금"].Width = 120;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvPaidInTabPaid.Columns["기지급액 (의료기관)"].Width = 140;
                        gvPaidInTabPaid.Columns["기지급액 (회원)"].Width = 100;
                    }
                    gvPaidInTabPaid.Columns["잔액/보류"].Width = 100;

                    gvPaidInTabPaid.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; ;
                    gvPaidInTabPaid.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPaidInTabPaid.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPaidInTabPaid.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPaidInTabPaid.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    gvPaidInTabPaid.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPaidInTabPaid.Columns["본인 부담금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPaidInTabPaid.Columns["회원할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPaidInTabPaid.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPaidInTabPaid.Columns["의료기관 지불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        gvPaidInTabPaid.Columns["기지급액"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        gvPaidInTabPaid.Columns["회원 환불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        gvPaidInTabPaid.Columns["기지급액 (의료기관)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        gvPaidInTabPaid.Columns["기지급액 (회원)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    gvPaidInTabPaid.Columns["잔액/보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    //gvPaidInTabPaid.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvPaidInTabPaid.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvPaidInTabPaid.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvPaidInTabPaid.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvPaidInTabPaid.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    //gvPaidInTabPaid.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["본인 부담금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["회원할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["의료기관 지불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["기지급액"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["회원 환불금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPaidInTabPaid.Columns["잔액/보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                }

                // The end of Closed - Paid Status

                // The beginning of CMM Pending Payment Status
                //DataSet dsCMMPendingPayment = new DataSet();

                DataTable dtCMMPendingPayment = new DataTable();
                dtCMMPendingPayment.Columns.Add("INCD", typeof(String));
                dtCMMPendingPayment.Columns.Add("회원 이름", typeof(String));
                dtCMMPendingPayment.Columns.Add("MED_BILL", typeof(String));
                dtCMMPendingPayment.Columns.Add("서비스 날짜", typeof(String));
                //dtCMMPendingPayment.Columns.Add("접수 날짜", typeof(String));
                dtCMMPendingPayment.Columns.Add("의료기관명", typeof(String));
                dtCMMPendingPayment.Columns.Add("청구액(원금)", typeof(String));
                dtCMMPendingPayment.Columns.Add("회원할인", typeof(String));
                dtCMMPendingPayment.Columns.Add("CMM 할인", typeof(String));
                dtCMMPendingPayment.Columns.Add("본인 부담금", typeof(String));
                dtCMMPendingPayment.Columns.Add("정산 완료", typeof(String));
                dtCMMPendingPayment.Columns.Add("지원 예정", typeof(String));

                DataRow drCMMPendingPayment = null;

                List<CMMPendingPayment> lstCMMPendingPayment = new List<CMMPendingPayment>();

                foreach (String incdName in lstDistinctIncdNames)
                {

                    String strSoqlBillCMMPendingPayment = "select c4g_Medical_Bill__r.Name, Approved__c, c4g_Type__c, c4g_Amount__c from Settlement__c " +
                                                            "where c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "' and " +
                                                            "Approved__c = true and c4g_Amount__c > 0 and c4g_Amount__c != null and " +
                                                            "(Check_Number__c = null and ACH_Number__c = null and CMM_Credit_Card_Paid_day__c = null)";

                    BlueSheetSForce.QueryResult qrCMMPendingPayment = Sfdcbinding.query(strSoqlBillCMMPendingPayment);

                    if (qrCMMPendingPayment.size > 0)
                    {
                        bCMMPendingPaymentHasRow = true;

                        List<String> lstCMMPendingMedBillNames = new List<String>();

                        for (int i = 0; i < qrCMMPendingPayment.size; i++)
                        {
                            BlueSheetSForce.Settlement__c cmm_pending_settlement = qrCMMPendingPayment.records[i] as BlueSheetSForce.Settlement__c;

                            lstCMMPendingMedBillNames.Add(cmm_pending_settlement.c4g_Medical_Bill__r.Name);
                        }

                        List<String> lstCMMPendingDistinctMedBillNames = new List<string>();

                        foreach (String MedBillName in lstCMMPendingMedBillNames.Distinct())
                        {
                            lstCMMPendingDistinctMedBillNames.Add(MedBillName);
                        }

                        foreach (String DistinctMedBillName in lstCMMPendingDistinctMedBillNames)
                        {
                            String strSoqlCMMPendingPayment = "select c4g_Incident__r.Name, c4g_Incident__r.c4g_Contact__r.Name, Name, Bill_Date__c, Medical_Provider__r.Name, " +
                                                                "c4g_Bill_Amount__c from Medical_Bill__c " +
                                                                "where Name = '" + DistinctMedBillName + "' and c4g_Incident__r.Name = '" + incdName + "'";

                            BlueSheetSForce.QueryResult qrCMMPendingPaymentMedBill = Sfdcbinding.query(strSoqlCMMPendingPayment);

                            if (qrCMMPendingPaymentMedBill.size > 0)
                            {
                                for (int i = 0; i < qrCMMPendingPaymentMedBill.size; i++)
                                {
                                    BlueSheetSForce.Medical_Bill__c MedBill = qrCMMPendingPaymentMedBill.records[i] as BlueSheetSForce.Medical_Bill__c;

                                    CMMPendingPayment cmm_pending_payment = new CMMPendingPayment();

                                    drCMMPendingPayment = dtCMMPendingPayment.NewRow();
                                    drCMMPendingPayment["INCD"] = MedBill.c4g_Incident__r.Name.Substring(5);
                                    drCMMPendingPayment["회원 이름"] = MedBill.c4g_Incident__r.c4g_Contact__r.Name;
                                    drCMMPendingPayment["MED_BILL"] = MedBill.Name.Substring(8);
                                    drCMMPendingPayment["서비스 날짜"] = MedBill.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                                    drCMMPendingPayment["의료기관명"] = MedBill.Medical_Provider__r.Name;
                                    cmm_pending_payment.BillAmount = MedBill.c4g_Bill_Amount__c.Value;
                                    drCMMPendingPayment["청구액(원금)"] = cmm_pending_payment.BillAmount.Value.ToString("C");

                                    String strSoqlCMMPendingMemberDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'Member Discount' and " +
                                                                                "c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                                                                "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "'";

                                    BlueSheetSForce.QueryResult qrCMMPendingMemberDiscount = Sfdcbinding.query(strSoqlCMMPendingMemberDiscount);

                                    Double? MemberDiscount = 0;

                                    if (qrCMMPendingMemberDiscount.size > 0)
                                    {
                                        for (int j = 0; j < qrCMMPendingMemberDiscount.size; j++)
                                        {
                                            BlueSheetSForce.Settlement__c mem_discount = qrCMMPendingMemberDiscount.records[j] as BlueSheetSForce.Settlement__c;
                                            MemberDiscount += mem_discount.c4g_Amount__c.Value;
                                        }
                                        cmm_pending_payment.MemberDiscount = MemberDiscount.Value;
                                        drCMMPendingPayment["회원할인"] = cmm_pending_payment.MemberDiscount.Value.ToString("C");
                                    }
                                    else if (qrCMMPendingMemberDiscount.size == 0)
                                    {
                                        cmm_pending_payment.MemberDiscount = 0;
                                        drCMMPendingPayment["회원할인"] = cmm_pending_payment.MemberDiscount.Value.ToString("C");
                                    }

                                    String strSoqlCMMPendingCMMDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Discount' and " +
                                                                            "c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                                                            "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "'";

                                    BlueSheetSForce.QueryResult qrCMMPendingCMMDiscount = Sfdcbinding.query(strSoqlCMMPendingCMMDiscount);

                                    Double? CMMDiscount = 0;

                                    if (qrCMMPendingCMMDiscount.size > 0)
                                    {
                                        for (int j = 0; j < qrCMMPendingCMMDiscount.size; j++)
                                        {
                                            BlueSheetSForce.Settlement__c cmm_discount = qrCMMPendingCMMDiscount.records[j] as BlueSheetSForce.Settlement__c;
                                            CMMDiscount += cmm_discount.c4g_Amount__c.Value;
                                        }
                                        cmm_pending_payment.CMMDiscount = CMMDiscount.Value;
                                        drCMMPendingPayment["CMM 할인"] = cmm_pending_payment.CMMDiscount.Value.ToString("C");
                                    }
                                    else if (qrCMMPendingCMMDiscount.size == 0)
                                    {
                                        cmm_pending_payment.CMMDiscount = 0;
                                        drCMMPendingPayment["CMM 할인"] = cmm_pending_payment.CMMDiscount.Value.ToString("C");
                                    }

                                    //String strSoqlCMMPendingPersonalResponsibility = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Personal_Responsibility_Type__c = 'Member Payment' and " +
                                    //                                                 "c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                    //                                                 "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "'";

                                    String strSoqlCMMPendingPersonalResponsibility = "select c4g_Amount__c, c4g_Type__c from Settlement__c where (c4g_Personal_Responsibility_Type__c = 'Member Payment' or " +
                                                                                        "c4g_Personal_Responsibility_Type__c = 'Member Discount') and " +
                                                                                        "c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                                                                        "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "'";

                                    BlueSheetSForce.QueryResult qrPersonalResponsibility = Sfdcbinding.query(strSoqlCMMPendingPersonalResponsibility);

                                    Double? CMMPendingPersonalResponsibility = 0;

                                    if (qrPersonalResponsibility.size > 0)
                                    {
                                        for (int j = 0; j < qrPersonalResponsibility.size; j++)
                                        {
                                            BlueSheetSForce.Settlement__c personal_responsibility = qrPersonalResponsibility.records[j] as BlueSheetSForce.Settlement__c;
                                            CMMPendingPersonalResponsibility += personal_responsibility.c4g_Amount__c.Value;
                                        }
                                        cmm_pending_payment.PersonalResponsibility = CMMPendingPersonalResponsibility;
                                        drCMMPendingPayment["본인 부담금"] = cmm_pending_payment.PersonalResponsibility.Value.ToString("C");
                                    }
                                    else if (qrPersonalResponsibility.size == 0)
                                    {
                                        cmm_pending_payment.PersonalResponsibility = 0;
                                        drCMMPendingPayment["본인 부담금"] = cmm_pending_payment.PersonalResponsibility.Value.ToString("C");
                                    }



                                    String strSoqlCMMPendingSharedAmount = "select c4g_Amount__c from Settlement__c where c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                                                            "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "' and " +
                                                                            "Approved__c = true and (Check_Number__c != null or ACH_Number__c != null or CMM_Credit_Card_Paid_day__c != null)";

                                    BlueSheetSForce.QueryResult qrCMMPendingSharedAmount = Sfdcbinding.query(strSoqlCMMPendingSharedAmount);

                                    Double? SharedAmount = 0;

                                    if (qrCMMPendingSharedAmount.size > 0)
                                    {
                                        for (int j = 0; j < qrCMMPendingSharedAmount.size; j++)
                                        {
                                            BlueSheetSForce.Settlement__c shared_amount = qrCMMPendingSharedAmount.records[j] as BlueSheetSForce.Settlement__c;
                                            SharedAmount += shared_amount.c4g_Amount__c.Value;
                                        }
                                        cmm_pending_payment.SharedAmount = SharedAmount.Value;
                                        drCMMPendingPayment["정산 완료"] = cmm_pending_payment.SharedAmount.Value.ToString("C");
                                    }
                                    else if (qrCMMPendingSharedAmount.size == 0)
                                    {
                                        cmm_pending_payment.SharedAmount = 0;
                                        drCMMPendingPayment["정산 완료"] = cmm_pending_payment.SharedAmount.Value.ToString("C");
                                    }

                                    String strSoqlCMMPendingBalance = "select c4g_Amount__c from Settlement__c where c4g_Medical_Bill__r.Name = '" + MedBill.Name + "' and " +
                                                                        "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "' and " +
                                                                        "Approved__c = true and (Check_Number__c = null and ACH_Number__c = null and CMM_Credit_Card_Paid_day__c = null)";

                                    BlueSheetSForce.QueryResult qrCMMPendingBalance = Sfdcbinding.query(strSoqlCMMPendingBalance);

                                    Double? Balance = 0;

                                    if (qrCMMPendingBalance.size > 0)
                                    {
                                        for (int j = 0; j < qrCMMPendingBalance.size; j++)
                                        {
                                            BlueSheetSForce.Settlement__c balance = qrCMMPendingBalance.records[j] as BlueSheetSForce.Settlement__c;
                                            Balance += balance.c4g_Amount__c.Value;
                                        }
                                        cmm_pending_payment.AmountWillBeShared = Balance.Value;
                                        drCMMPendingPayment["지원 예정"] = cmm_pending_payment.AmountWillBeShared.Value.ToString("C");
                                    }
                                    else if (qrCMMPendingBalance.size == 0)
                                    {
                                        cmm_pending_payment.AmountWillBeShared = 0;
                                        drCMMPendingPayment["지원 예정"] = cmm_pending_payment.AmountWillBeShared.Value.ToString("C");
                                    }

                                    lstCMMPendingPayment.Add(cmm_pending_payment);
                                    dtCMMPendingPayment.Rows.Add(drCMMPendingPayment);
                                }
                            }
                        }
                    }
                }

                if (bCMMPendingPaymentHasRow)
                {
                    double sumCMMPendingPaymentBillAmount = 0;
                    double sumCMMPendingPaymentMemberDiscount = 0;
                    double sumCMMPendingPaymentCMMDiscount = 0;
                    double sumCMMPendingPaymentPersonalResponisiblity = 0;
                    double sumCMMPendingPaymentSharedAmount = 0;
                    double sumCMMPendingPaymentAmountWillBeShared = 0;

                    foreach (CMMPendingPayment cmm_payment in lstCMMPendingPayment)
                    {
                        sumCMMPendingPaymentBillAmount += cmm_payment.BillAmount.Value;
                        sumCMMPendingPaymentMemberDiscount += cmm_payment.MemberDiscount.Value;
                        sumCMMPendingPaymentCMMDiscount += cmm_payment.CMMDiscount.Value;
                        sumCMMPendingPaymentPersonalResponisiblity += cmm_payment.PersonalResponsibility.Value;
                        sumCMMPendingPaymentSharedAmount += cmm_payment.SharedAmount.Value;
                        sumCMMPendingPaymentAmountWillBeShared += cmm_payment.AmountWillBeShared.Value;
                    }

                    DataRow drCMMPendingPaymentSum = dtCMMPendingPayment.NewRow();

                    drCMMPendingPaymentSum["의료기관명"] = "합계";
                    drCMMPendingPaymentSum["청구액(원금)"] = sumCMMPendingPaymentBillAmount.ToString("C");
                    drCMMPendingPaymentSum["회원할인"] = sumCMMPendingPaymentMemberDiscount.ToString("C");
                    drCMMPendingPaymentSum["CMM 할인"] = sumCMMPendingPaymentCMMDiscount.ToString("C");
                    drCMMPendingPaymentSum["본인 부담금"] = sumCMMPendingPaymentPersonalResponisiblity.ToString("C");
                    drCMMPendingPaymentSum["정산 완료"] = sumCMMPendingPaymentSharedAmount.ToString("C");
                    drCMMPendingPaymentSum["지원 예정"] = sumCMMPendingPaymentAmountWillBeShared.ToString("C");

                    dtCMMPendingPayment.Rows.Add(drCMMPendingPaymentSum);

                    //dsCMMPendingPayment.Tables.Add(dtCMMPendingPayment);

                    gvCMMPendingPayment.DataSource = null;
                    gvCMMPendingPayment.DataSource = dtCMMPendingPayment;

                    foreach (DataGridViewColumn col in gvCMMPendingPayment.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }


                    gvCMMPendingPayment.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingPayment.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingPayment.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingPayment.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    //gvCMMPendingPayment.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingPayment.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvCMMPendingPayment.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingPayment.Columns["회원할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingPayment.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingPayment.Columns["본인 부담금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingPayment.Columns["정산 완료"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingPayment.Columns["지원 예정"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    DataTable dtCMMPendingCopy = dtCMMPendingPayment.Copy();

                    gvCMMPendingInTab.DataSource = dtCMMPendingCopy;

                    foreach (DataGridViewColumn col in gvCMMPendingInTab.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvCMMPendingInTab.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingInTab.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingInTab.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingInTab.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    //gvCMMPendingInTab.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvCMMPendingInTab.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvCMMPendingInTab.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingInTab.Columns["회원할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingInTab.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingInTab.Columns["본인 부담금"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingInTab.Columns["정산 완료"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvCMMPendingInTab.Columns["지원 예정"].SortMode = DataGridViewColumnSortMode.NotSortable;


                    gvCMMPendingPayment.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingPayment.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingPayment.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingPayment.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvCMMPendingPayment.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingPayment.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvCMMPendingPayment.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingPayment.Columns["회원할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingPayment.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingPayment.Columns["본인 부담금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingPayment.Columns["정산 완료"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingPayment.Columns["지원 예정"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    gvCMMPendingPayment.Columns["INCD"].Width = 80;
                    gvCMMPendingPayment.Columns["회원 이름"].Width = 150;
                    gvCMMPendingPayment.Columns["MED_BILL"].Width = 80;
                    gvCMMPendingPayment.Columns["서비스 날짜"].Width = 100;
                    //gvCMMPendingPayment.Columns["접수 날짜"].Width = 100;
                    gvCMMPendingPayment.Columns["의료기관명"].Width = 200;
                    gvCMMPendingPayment.Columns["청구액(원금)"].Width = 100;
                    gvCMMPendingPayment.Columns["회원할인"].Width = 80;
                    gvCMMPendingPayment.Columns["CMM 할인"].Width = 100;
                    gvCMMPendingPayment.Columns["본인 부담금"].Width = 100;
                    gvCMMPendingPayment.Columns["정산 완료"].Width = 120;
                    gvCMMPendingPayment.Columns["지원 예정"].Width = 100;

                    gvCMMPendingInTab.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingInTab.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingInTab.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingInTab.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    //gvCMMPendingInTab.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvCMMPendingInTab.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvCMMPendingInTab.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingInTab.Columns["회원할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingInTab.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingInTab.Columns["본인 부담금"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingInTab.Columns["정산 완료"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvCMMPendingInTab.Columns["지원 예정"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }


                // The end of CMM Pending Payment Status

                // The beginning of Pending Status
                //DataSet dsPending = new DataSet();

                DataTable dtPending = new DataTable();
                dtPending.Columns.Add("INCD", typeof(String));
                dtPending.Columns.Add("회원 이름", typeof(String));
                dtPending.Columns.Add("MED_BILL", typeof(String));
                dtPending.Columns.Add("서비스 날짜", typeof(String));
                dtPending.Columns.Add("접수 날짜", typeof(String));
                dtPending.Columns.Add("의료기관명", typeof(String));
                dtPending.Columns.Add("청구액(원금)", typeof(String));
                dtPending.Columns.Add("잔액/보류", typeof(String));
                dtPending.Columns.Add("보류 사유", typeof(String));

                DataRow drPending = null;
                List<Pending> lstPending = new List<Pending>();

                foreach (String incdName in lstDistinctIncdNames)
                {

                    //String strSoqlPending = "select c4g_Incident__r.Name, Name, c4g_Incident__r.c4g_Contact__r.Name, Bill_Date__c, Due_Date__c, Medical_Provider__r.Name, " +
                    //                        "c4g_Balance__c, c4g_Bill_Amount__c, Pending_Reason__c from Medical_Bill__c " +
                    //                        "where c4g_Incident__r.c4g_Contact__r.c4g_Membership__r.Name like '%" + txtMembershipID.Text + "' and " +
                    //                        "c4g_Incident__r.Name = '" + incdName + "' and " +
                    //                        "Bill_Status__c = 'Pending'";

                    String strSoqlPending = "select c4g_Incident__r.Name, Name, c4g_Incident__r.c4g_Contact__r.Name, Bill_Date__c, Due_Date__c, Medical_Provider__r.Name, " +
                                            "c4g_Balance__c, c4g_Bill_Amount__c, Pending_Reason__c from Medical_Bill__c " +
                                            "where c4g_Incident__r.Name = '" + incdName + "' and " +
                                            "Pending_Reason__c <> null and c4g_Balance__c > 0";

                    //"Bill_Status__c = 'Pending'";

                    BlueSheetSForce.QueryResult qrPending = Sfdcbinding.query(strSoqlPending);

                    if (qrPending.size > 0)
                    {
                        bPendingHasRow = true;

                        for (int i = 0; i < qrPending.size; i++)
                        {
                            BlueSheetSForce.Medical_Bill__c medBill = qrPending.records[i] as BlueSheetSForce.Medical_Bill__c;

                            Pending pending_expense = new Pending();

                            drPending = dtPending.NewRow();
                            drPending["INCD"] = medBill.c4g_Incident__r.Name.Substring(5);
                            drPending["회원 이름"] = medBill.c4g_Incident__r.c4g_Contact__r.Name;
                            drPending["MED_BILL"] = medBill.Name.Substring(8);
                            drPending["서비스 날짜"] = medBill.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                            if (medBill.Due_Date__c != null) drPending["접수 날짜"] = medBill.Due_Date__c.Value.ToString("MM/dd/yyyy");
                            drPending["의료기관명"] = medBill.Medical_Provider__r.Name;
                            pending_expense.BillAmount = medBill.c4g_Bill_Amount__c.Value;
                            drPending["청구액(원금)"] = pending_expense.BillAmount.Value.ToString("C");
                            pending_expense.Balance = medBill.c4g_Balance__c.Value;
                            drPending["잔액/보류"] = pending_expense.Balance.Value.ToString("C");
                            //pending_expense.SharedAmount = medBill.c4g_Bill_Amount__c.Value - medBill.c4g_Balance__c.Value;
                            //drPending["정산 완료"] = pending_expense.SharedAmount.Value.ToString("C");
                            //pending_expense.PendingAmount = medBill.c4g_Balance__c.Value;
                            //drPending["보류"] = pending_expense.PendingAmount.Value.ToString("C");
                            if (medBill.Pending_Reason__c != null) drPending["보류 사유"] = medBill.Pending_Reason__c;

                            //String strSoqlPendingMemberDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'Member Discount' and " +
                            //                                      "c4g_Medical_Bill__r.Name = '" + medBill.Name + "' and " +
                            //                                      "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "' and " +
                            //                                      "c4g_Medical_Bill__r.Bill_Status__c = 'Pending'";

                            //BlueSheetSForce.QueryResult qrMemberDiscount = Sfdcbinding.query(strSoqlPendingMemberDiscount);

                            //if (qrMemberDiscount.size > 0)
                            //{
                            //    double MemberDiscount = 0;

                            //    for (int j = 0; j < qrMemberDiscount.size; j++)
                            //    {
                            //        BlueSheetSForce.Settlement__c member_discount = qrMemberDiscount.records[j] as BlueSheetSForce.Settlement__c;
                            //        MemberDiscount += member_discount.c4g_Amount__c.Value;
                            //    }
                            //    pending_expense.MemberDiscount = MemberDiscount;
                            //    drPending["회원 할인"] = pending_expense.MemberDiscount.Value.ToString("C");
                            //}
                            //else if (qrMemberDiscount.size == 0)
                            //{
                            //    pending_expense.MemberDiscount = 0;
                            //    drPending["회원 할인"] = pending_expense.MemberDiscount.Value.ToString("C");
                            //}

                            //String strSoqlPendingCMMDiscount = "select c4g_Amount__c, c4g_Type__c from Settlement__c where c4g_Type__c = 'CMM Discount' and " +
                            //                                    "c4g_Medical_Bill__r.Name = '" + medBill.Name + "' and " +
                            //                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "' and " +
                            //                                    "c4g_Medical_Bill__r.Bill_Status__c = 'Pending'";


                            //BlueSheetSForce.QueryResult qrCMMDiscount = Sfdcbinding.query(strSoqlPendingCMMDiscount);

                            //if (qrCMMDiscount.size > 0)
                            //{
                            //    double CMMDiscount = 0;

                            //    for (int j = 0; j < qrCMMDiscount.size; j++)
                            //    {
                            //        BlueSheetSForce.Settlement__c cmm_discount = qrCMMDiscount.records[j] as BlueSheetSForce.Settlement__c;
                            //        CMMDiscount += cmm_discount.c4g_Amount__c.Value;
                            //    }
                            //    pending_expense.CMMDiscount = CMMDiscount;
                            //    drPending["CMM 할인"] = pending_expense.CMMDiscount.Value.ToString("C");
                            //}
                            //else if (qrCMMDiscount.size == 0)
                            //{
                            //    pending_expense.CMMDiscount = 0;
                            //    drPending["CMM 할인"] = pending_expense.CMMDiscount.Value.ToString("C");
                            //}

                            dtPending.Rows.Add(drPending);
                            lstPending.Add(pending_expense);
                        }
                    }
                }

                if (bPendingHasRow)
                {

                    double sumPendingBillAmount = 0;
                    double sumPendingBalance = 0;
                    //double sumPendingMemberDiscount = 0;
                    //double sumPendingCMMDiscount = 0;
                    //double sumPendingSharedAmount = 0;
                    //double sumPendingPendingAmount = 0;

                    foreach (Pending cmm_pending in lstPending)
                    {
                        sumPendingBillAmount += cmm_pending.BillAmount.Value;
                        sumPendingBalance += cmm_pending.Balance.Value;
                        //sumPendingMemberDiscount += cmm_pending.MemberDiscount.Value;
                        //sumPendingCMMDiscount += cmm_pending.CMMDiscount.Value;
                        //sumPendingSharedAmount += cmm_pending.SharedAmount.Value;
                        //sumPendingPendingAmount += cmm_pending.PendingAmount.Value;
                    }

                    DataRow drPendingSum = dtPending.NewRow();

                    drPendingSum["의료기관명"] = "합계";
                    drPendingSum["청구액(원금)"] = sumPendingBillAmount.ToString("C");
                    drPendingSum["잔액/보류"] = sumPendingBalance.ToString("C");
                    //drPendingSum["회원 할인"] = sumPendingMemberDiscount.ToString("C");
                    //drPendingSum["CMM 할인"] = sumPendingCMMDiscount.ToString("C");
                    //drPendingSum["정산 완료"] = sumPendingSharedAmount.ToString("C");
                    //drPendingSum["보류"] = sumPendingPendingAmount.ToString("C");

                    dtPending.Rows.Add(drPendingSum);

                    gvPending.DataSource = null;
                    gvPending.DataSource = dtPending;


                    foreach (DataGridViewColumn col in gvPending.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvPending.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPending.Columns["보류 사유"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvPending.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPending.Columns["잔액/보류"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPending.Columns["회원 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPending.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPending.Columns["정산 완료"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPending.Columns["보류"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    DataTable dtPendingInTab = dtPending.Copy();

                    gvPendingInTab.DataSource = dtPendingInTab;

                    foreach (DataGridViewColumn col in gvPendingInTab.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvPendingInTab.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvPendingInTab.Columns["보류 사유"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvPendingInTab.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvPendingInTab.Columns["잔액/보류"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPendingInTab.Columns["회원 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPendingInTab.Columns["CMM 할인"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPendingInTab.Columns["정산 완료"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //gvPendingInTab.Columns["보류"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    gvPending.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPending.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPending.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPending.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPending.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPending.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvPending.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPending.Columns["잔액/보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPending.Columns["회원 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPending.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPending.Columns["정산 완료"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPending.Columns["보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPending.Columns["보류 사유"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                    gvPending.Columns["INCD"].Width = 80;
                    gvPending.Columns["회원 이름"].Width = 150;
                    gvPending.Columns["MED_BILL"].Width = 80;
                    gvPending.Columns["서비스 날짜"].Width = 100;
                    gvPending.Columns["접수 날짜"].Width = 100;
                    gvPending.Columns["의료기관명"].Width = 200;
                    gvPending.Columns["청구액(원금)"].Width = 100;
                    gvPending.Columns["잔액/보류"].Width = 80;
                    //gvPending.Columns["회원 할인"].Width = 80;
                    //gvPending.Columns["CMM 할인"].Width = 80;
                    //gvPending.Columns["정산 완료"].Width = 80;
                    //gvPending.Columns["보류"].Width = 80;
                    gvPending.Columns["보류 사유"].Width = 520;

                    gvPendingInTab.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPendingInTab.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPendingInTab.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPendingInTab.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPendingInTab.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvPendingInTab.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                    gvPendingInTab.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPendingInTab.Columns["잔액/보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPendingInTab.Columns["회원 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPendingInTab.Columns["CMM 할인"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPendingInTab.Columns["정산 완료"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    //gvPendingInTab.Columns["보류"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvPendingInTab.Columns["보류 사유"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;


                }

                //////////////////////////////////////////////////////////////////////////////////////////////
                // Beginning of ineligible bill

                //DataSet dsMedicalBillIneligible = new DataSet();

                DataTable dtMedicalBillIneligible = new DataTable();
                dtMedicalBillIneligible.Columns.Add("INCD", typeof(String));
                dtMedicalBillIneligible.Columns.Add("회원 이름", typeof(String));
                dtMedicalBillIneligible.Columns.Add("MED_BILL", typeof(String));
                dtMedicalBillIneligible.Columns.Add("서비스 날짜", typeof(String));
                dtMedicalBillIneligible.Columns.Add("접수 날짜", typeof(String));
                dtMedicalBillIneligible.Columns.Add("의료기관명", typeof(String));
                dtMedicalBillIneligible.Columns.Add("청구액(원금)", typeof(String));
                dtMedicalBillIneligible.Columns.Add("전액/일부 지원불가 금액", typeof(String));
                dtMedicalBillIneligible.Columns.Add("지원되지않는 사유", typeof(String));

                DataRow drIneligible = null;



                List<MedicalExpenseIneligible> lstMedicalBillIneligible = new List<MedicalExpenseIneligible>();

                foreach (String strIncdName in lstDistinctIncdNames)
                {

                    String strSoqlMedBillIneligible = "select c4g_Incident__r.Name, Name, c4g_Incident__r.c4g_Contact__r.Name, Bill_Date__c, Due_Date__c, " +
                                                        "Medical_Provider__r.Name, c4g_Bill_Amount__c, Ineligible_Reason__c " +
                                                        "from Medical_Bill__c where Bill_Status__c = 'Ineligible' and Ineligible_Reason__c <> null and " +
                                                        "c4g_Incident__r.Name = '" + strIncdName + "'";

                    BlueSheetSForce.QueryResult qrMedBillIneligible = Sfdcbinding.query(strSoqlMedBillIneligible);

                    if (qrMedBillIneligible.size > 0)
                    {
                        bIneligibleHasRow = true;
                        for (int i = 0; i < qrMedBillIneligible.size; i++)
                        {
                            BlueSheetSForce.Medical_Bill__c med_bill_ineligible = qrMedBillIneligible.records[i] as BlueSheetSForce.Medical_Bill__c;

                            MedicalExpenseIneligible expenseIneligible = new MedicalExpenseIneligible();

                            drIneligible = dtMedicalBillIneligible.NewRow();

                            drIneligible["INCD"] = med_bill_ineligible.c4g_Incident__r.Name.Substring(5);
                            drIneligible["회원 이름"] = med_bill_ineligible.c4g_Incident__r.c4g_Contact__r.Name;
                            drIneligible["MED_BILL"] = med_bill_ineligible.Name.Substring(8);
                            drIneligible["서비스 날짜"] = med_bill_ineligible.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                            if (med_bill_ineligible.Due_Date__c != null)
                            {
                                drIneligible["접수 날짜"] = med_bill_ineligible.Due_Date__c.Value.ToString("MM/dd/yyyy");
                            }
                            else drIneligible["접수 날짜"] = null;
                            //drIneligible["서비스 날짜"] = med_bill_ineligible.Bill_Date__c.Value;
                            drIneligible["의료기관명"] = med_bill_ineligible.Medical_Provider__r.Name;
                            expenseIneligible.BillAmount = med_bill_ineligible.c4g_Bill_Amount__c;
                            drIneligible["청구액(원금)"] = expenseIneligible.BillAmount.Value.ToString("C");
                            expenseIneligible.AmountIneligible = med_bill_ineligible.c4g_Bill_Amount__c;
                            drIneligible["전액/일부 지원불가 금액"] = expenseIneligible.AmountIneligible.Value.ToString("C");
                            drIneligible["지원되지않는 사유"] = med_bill_ineligible.Ineligible_Reason__c;

                            lstMedicalBillIneligible.Add(expenseIneligible);
                            dtMedicalBillIneligible.Rows.Add(drIneligible);
                        }

                    }
                }
                foreach (String strIncdName in lstDistinctIncdNames)
                {

                    String strSoqlSettlementPartiallyIneligible = "select c4g_Medical_Bill__r.Name from Settlement__c where " +
                                                                    "c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and " +      // Bill_Status = 'CMM_Pending_Payment might be added
                                                                    "c4g_Medical_Bill__r.Ineligible_Reason__c <> null and " +
                                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + strIncdName + "'";


                    BlueSheetSForce.QueryResult qrSettlementPartiallyIneligible = Sfdcbinding.query(strSoqlSettlementPartiallyIneligible);

                    if (qrSettlementPartiallyIneligible.size > 0)
                    {
                        bIneligibleHasRow = true;
                        List<String> lstMedicalBillNames = new List<String>();
                        List<String> lstDistinceMedicalBillNames = new List<String>();

                        for (int i = 0; i < qrSettlementPartiallyIneligible.size; i++)
                        {
                            BlueSheetSForce.Settlement__c settlementPartiallyIneligible = qrSettlementPartiallyIneligible.records[i] as BlueSheetSForce.Settlement__c;
                            lstMedicalBillNames.Add(settlementPartiallyIneligible.c4g_Medical_Bill__r.Name);
                        }

                        foreach (String MedicalBillName in lstMedicalBillNames.Distinct())
                        {
                            lstDistinceMedicalBillNames.Add(MedicalBillName);
                        }

                        if (lstDistinceMedicalBillNames.Count > 0)
                        {
                            foreach (String strMedBillName in lstDistinceMedicalBillNames)
                            {
                                String strSoqlMedBillPartiallyIneligible = "select c4g_Medical_Bill__r.c4g_Incident__r.Name, c4g_Medical_Bill__r.Name, " +
                                                                            "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, c4g_Medical_Bill__r.Bill_Date__c, c4g_Medical_Bill__r.Due_Date__c, " +
                                                                            "c4g_Medical_Bill__r.Medical_Provider__r.Name, c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Amount__c, " +
                                                                            "c4g_Medical_Bill__r.Bill_Status__c, c4g_Type__c, c4g_Medical_Bill__r.Ineligible_Reason__c " +
                                                                            "from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                            "c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' and " +
                                                                            "c4g_Type__c = 'Ineligible' and " +
                                                                            "c4g_Medical_Bill__r.Ineligible_Reason__c <> null";

                                BlueSheetSForce.QueryResult qrMedBillPartiallyIneligible = Sfdcbinding.query(strSoqlMedBillPartiallyIneligible);

                                if (qrMedBillPartiallyIneligible.size > 0)
                                {
                                    bIneligibleHasRow = true;

                                    List<MedicalExpensePartiallyIneligible> lstMedBillPartiallyIneligible = new List<MedicalExpensePartiallyIneligible>();

                                    for (int i = 0; i < qrMedBillPartiallyIneligible.size; i++)
                                    {
                                        BlueSheetSForce.Settlement__c settlementIneligible = qrMedBillPartiallyIneligible.records[i] as BlueSheetSForce.Settlement__c;

                                        MedicalExpensePartiallyIneligible medbill = new MedicalExpensePartiallyIneligible();

                                        MedicalExpenseIneligible expenseIneligible = new MedicalExpenseIneligible();

                                        medbill.INCD = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.Name.Substring(5);
                                        medbill.PatientName = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                        medbill.MedBill = settlementIneligible.c4g_Medical_Bill__r.Name.Substring(8);
                                        if (settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c != null)
                                        {
                                            medbill.ServiceDate = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c;
                                        }
                                        else medbill.ServiceDate = null;
                                        if (settlementIneligible.c4g_Medical_Bill__r.Due_Date__c != null)
                                        {
                                            medbill.ReceiveDate = settlementIneligible.c4g_Medical_Bill__r.Due_Date__c;
                                        }
                                        else medbill.ReceiveDate = null;
                                        medbill.MedicalProvider = settlementIneligible.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                        if (i == 0) expenseIneligible.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c;
                                        medbill.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c;
                                        medbill.IneligibleAmount = settlementIneligible.c4g_Amount__c;
                                        expenseIneligible.AmountIneligible = settlementIneligible.c4g_Amount__c; ;
                                        medbill.IneligibleReason = settlementIneligible.c4g_Medical_Bill__r.Ineligible_Reason__c;

                                        lstMedicalBillIneligible.Add(expenseIneligible);
                                        lstMedBillPartiallyIneligible.Add(medbill);

                                    }

                                    drIneligible = dtMedicalBillIneligible.NewRow();

                                    drIneligible["INCD"] = lstMedBillPartiallyIneligible[0].INCD;
                                    drIneligible["회원 이름"] = lstMedBillPartiallyIneligible[0].PatientName;
                                    drIneligible["MED_BILL"] = lstMedBillPartiallyIneligible[0].MedBill;
                                    drIneligible["서비스 날짜"] = lstMedBillPartiallyIneligible[0].ServiceDate.Value.ToString("MM/dd/yyyy");
                                    if (lstMedBillPartiallyIneligible[0].ReceiveDate != null) drIneligible["접수 날짜"] = lstMedBillPartiallyIneligible[0].ReceiveDate.Value.ToString("MM/dd/yyyy");
                                    drIneligible["의료기관명"] = lstMedBillPartiallyIneligible[0].MedicalProvider;
                                    drIneligible["청구액(원금)"] = lstMedBillPartiallyIneligible[0].BillAmount.Value.ToString("C");


                                    Double? IneligibleAmount = 0;
                                    for (int i = 0; i < lstMedBillPartiallyIneligible.Count; i++)
                                    {
                                        IneligibleAmount += lstMedBillPartiallyIneligible[i].IneligibleAmount;
                                    }
                                    drIneligible["전액/일부 지원불가 금액"] = IneligibleAmount.Value.ToString("C");
                                    drIneligible["지원되지않는 사유"] = lstMedBillPartiallyIneligible[0].IneligibleReason;

                                    dtMedicalBillIneligible.Rows.Add(drIneligible);

                                    //for (int i = 0; i < qrMedBillPartiallyIneligible.size; i++)
                                    //{
                                    //    BlueSheetSForce.Settlement__c settlementIneligible = qrMedBillPartiallyIneligible.records[i] as BlueSheetSForce.Settlement__c;

                                    //    MedicalExpenseIneligible expenseIneligible = new MedicalExpenseIneligible();

                                    //    drIneligible = dtMedicalBillIneligible.NewRow();

                                    //    drIneligible["INCD"] = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.Name.Substring(5);
                                    //    drIneligible["회원 이름"] = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                    //    drIneligible["MED_BILL"] = settlementIneligible.c4g_Medical_Bill__r.Name.Substring(8);
                                    //    drIneligible["서비스 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                                    //    if (settlementIneligible.c4g_Medical_Bill__r.Due_Date__c != null)
                                    //    {
                                    //        drIneligible["접수 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Due_Date__c.Value.ToString("MM/dd/yyyy");
                                    //    }
                                    //    else drIneligible["접수 날짜"] = null;
                                    //    //drIneligible["서비스 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c.Value;
                                    //    drIneligible["의료기관명"] = settlementIneligible.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                    //    expenseIneligible.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c;
                                    //    drIneligible["청구액(원금)"] = expenseIneligible.BillAmount.Value.ToString("C");
                                    //    expenseIneligible.AmountIneligible = settlementIneligible.c4g_Amount__c;
                                    //    drIneligible["전액/일부 지원불가 금액"] = expenseIneligible.AmountIneligible.Value.ToString("C");
                                    //    drIneligible["지원되지않는 사유"] = settlementIneligible.c4g_Medical_Bill__r.Ineligible_Reason__c;

                                    //    lstMedicalBillIneligible.Add(expenseIneligible);
                                    //    dtMedicalBillIneligible.Rows.Add(drIneligible);
                                    //}
                                }
                            }
                        }
                    }
                }

                foreach (String incdName in lstDistinctIncdNames)
                {
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    String strSoqlMedBillClosed = "select c4g_Medical_Bill__r.Name from Settlement__c where c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and " +
                                                    "(Check_Number__c = '" + txtCheckNo.Text.Trim() + "' or ACH_Number__c = '" + txtACH_No.Text.Trim() + "' or " +
                                                    "CMM_Credit_Card_Paid_day__c >= " + dtpCreditCardPaymentDate.Value.ToString("yyyy-MM-dd") + ") and " +
                                                    "c4g_Medical_Bill__r.c4g_Incident__r.Name = '" + incdName + "'";

                    BlueSheetSForce.QueryResult qrMedBillClosed = Sfdcbinding.query(strSoqlMedBillClosed);

                    if (qrMedBillClosed.size > 0)
                    {

                        List<String> lstMedicalBillNames = new List<String>();

                        for (int i = 0; i < qrMedBillClosed.size; i++)
                        {
                            BlueSheetSForce.Settlement__c settlementClosed = qrMedBillClosed.records[i] as BlueSheetSForce.Settlement__c;
                            lstMedicalBillNames.Add(settlementClosed.c4g_Medical_Bill__r.Name);
                        }

                        List<String> lstDistinctMedicalBillNames = new List<String>();

                        foreach (String MedicalBillName in lstMedicalBillNames.Distinct())
                        {
                            //lstDistinctMedBillNames.Add(MedicalBillName);
                            lstDistinctMedicalBillNames.Add(MedicalBillName);
                        }


                        //if (lstMedicalBillNames.Count > 0)
                        //if (lstDistinctMedBillNames.Count > 0)
                        if (lstDistinctMedicalBillNames.Count > 0)
                        {
                            //foreach (String strMedBillName in lstMedicalBillNames)
                            foreach (String strMedBillName in lstDistinctMedicalBillNames)
                            {
                                String strSoqlMedBillClosedIneligible = "select c4g_Medical_Bill__r.c4g_Incident__r.Name, c4g_Medical_Bill__r.Name, " +
                                                                        "c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name, c4g_Medical_Bill__r.Bill_Date__c, c4g_Medical_Bill__r.Due_Date__c, " +
                                                                        "c4g_Medical_Bill__r.Medical_Provider__r.Name, c4g_Medical_Bill__r.c4g_Bill_Amount__c, c4g_Amount__c, " +
                                                                        "c4g_Medical_Bill__r.Bill_Status__c, c4g_Type__c, c4g_Medical_Bill__r.Ineligible_Reason__c " +
                                                                        "from Settlement__c where c4g_Medical_Bill__r.Name = '" + strMedBillName + "' and " +
                                                                        "c4g_Medical_Bill__r.Bill_Status__c = 'Closed' and " +
                                                                        "c4g_Type__c = 'Ineligible' and " +
                                                                        "c4g_Medical_Bill__r.Ineligible_Reason__c <> null and " +
                                                                        "(c4g_Medical_Bill__r.Bill_Status__c = 'Partially Ineligible' or c4g_Medical_Bill__r.Bill_Status__c = 'Closed')";

                                BlueSheetSForce.QueryResult qrMedBillClosedIneligible = Sfdcbinding.query(strSoqlMedBillClosedIneligible);

                                if (qrMedBillClosedIneligible.size > 0)
                                {
                                    bIneligibleHasRow = true;
                                    for (int i = 0; i < qrMedBillClosedIneligible.size; i++)
                                    {
                                        BlueSheetSForce.Settlement__c settlementIneligible = qrMedBillClosedIneligible.records[i] as BlueSheetSForce.Settlement__c;

                                        MedicalExpenseIneligible expenseIneligible = new MedicalExpenseIneligible();

                                        drIneligible = dtMedicalBillIneligible.NewRow();

                                        drIneligible["INCD"] = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.Name.Substring(5);
                                        drIneligible["회원 이름"] = settlementIneligible.c4g_Medical_Bill__r.c4g_Incident__r.c4g_Contact__r.Name;
                                        drIneligible["MED_BILL"] = settlementIneligible.c4g_Medical_Bill__r.Name.Substring(8);
                                        drIneligible["서비스 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c.Value.ToString("MM/dd/yyyy");
                                        if (settlementIneligible.c4g_Medical_Bill__r.Due_Date__c != null)
                                        {
                                            drIneligible["접수 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Due_Date__c.Value.ToString("MM/dd/yyyy");
                                        }
                                        else drIneligible["접수 날짜"] = null;
                                        //drIneligible["서비스 날짜"] = settlementIneligible.c4g_Medical_Bill__r.Bill_Date__c.Value;
                                        drIneligible["의료기관명"] = settlementIneligible.c4g_Medical_Bill__r.Medical_Provider__r.Name;
                                        expenseIneligible.BillAmount = settlementIneligible.c4g_Medical_Bill__r.c4g_Bill_Amount__c;
                                        drIneligible["청구액(원금)"] = expenseIneligible.BillAmount.Value.ToString("C");
                                        expenseIneligible.AmountIneligible = settlementIneligible.c4g_Amount__c;
                                        drIneligible["전액/일부 지원불가 금액"] = expenseIneligible.AmountIneligible.Value.ToString("C");
                                        drIneligible["지원되지않는 사유"] = settlementIneligible.c4g_Medical_Bill__r.Ineligible_Reason__c;

                                        lstMedicalBillIneligible.Add(expenseIneligible);
                                        dtMedicalBillIneligible.Rows.Add(drIneligible);
                                    }
                                }

                            }
                        }
                    }
                }

                if (bIneligibleHasRow)
                {
                    double sumIneligibleBillAmount = 0;
                    double sumIneligibleAmountIneligible = 0;

                    foreach (MedicalExpenseIneligible expenseIneligible in lstMedicalBillIneligible)
                    {
                        sumIneligibleBillAmount += expenseIneligible.BillAmount.Value;
                        sumIneligibleAmountIneligible += expenseIneligible.AmountIneligible.Value;
                    }

                    DataRow drSumIneligible = dtMedicalBillIneligible.NewRow();

                    drSumIneligible["의료기관명"] = "합계";
                    drSumIneligible["청구액(원금)"] = sumIneligibleBillAmount.ToString("C");
                    drSumIneligible["전액/일부 지원불가 금액"] = sumIneligibleAmountIneligible.ToString("C");

                    dtMedicalBillIneligible.Rows.Add(drSumIneligible);

                    gvIneligible.DataSource = null;
                    gvIneligible.DataSource = dtMedicalBillIneligible;

                    foreach (DataGridViewColumn col in gvIneligible.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvIneligible.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligible.Columns["지원되지않는 사유"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvIneligible.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvIneligible.Columns["전액/일부 지원불가 금액"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    DataTable dtIneligibleInTab = dtMedicalBillIneligible.Copy();
                    gvIneligibleInTab.DataSource = dtIneligibleInTab;

                    foreach (DataGridViewColumn col in gvIneligibleInTab.Columns)
                    {
                        col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    gvIneligibleInTab.Columns["INCD"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["회원 이름"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["MED_BILL"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["서비스 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["접수 날짜"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["의료기관명"].SortMode = DataGridViewColumnSortMode.Programmatic;
                    gvIneligibleInTab.Columns["지원되지않는 사유"].SortMode = DataGridViewColumnSortMode.Programmatic;

                    gvIneligibleInTab.Columns["청구액(원금)"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    gvIneligibleInTab.Columns["전액/일부 지원불가 금액"].SortMode = DataGridViewColumnSortMode.NotSortable;


                    gvIneligible.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligible.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligible.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligible.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligible.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligible.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvIneligible.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvIneligible.Columns["전액/일부 지원불가 금액"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvIneligible.Columns["지원되지않는 사유"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                    gvIneligible.Columns["의료기관명"].Width = 200;
                    gvIneligible.Columns["청구액(원금)"].Width = 100;
                    gvIneligible.Columns["전액/일부 지원불가 금액"].Width = 180;
                    gvIneligible.Columns["지원되지않는 사유"].Width = 200;

                    gvIneligibleInTab.Columns["INCD"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligibleInTab.Columns["회원 이름"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligibleInTab.Columns["MED_BILL"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligibleInTab.Columns["서비스 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligibleInTab.Columns["접수 날짜"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    gvIneligibleInTab.Columns["의료기관명"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    gvIneligibleInTab.Columns["청구액(원금)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvIneligibleInTab.Columns["전액/일부 지원불가 금액"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    gvIneligibleInTab.Columns["지원되지않는 사유"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                }

                if (bPaidHasRow == true || bCMMPendingPaymentHasRow == true || bPaidHasRow == true || bIneligibleHasRow == true)
                {
                    frmLoadingFinished loadingFinished = new frmLoadingFinished();
                    loadingFinished.StartPosition = FormStartPosition.CenterParent;
                    loadingFinished.ShowDialog();
                    Cursor.Current = Cursors.Default;
                }

                tabMedicalExpense.SelectedIndex = 0;                
                //}
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {

            frmExit frm = new frmExit();

            frm.StartPosition = FormStartPosition.CenterParent;
            frm.ShowDialog();

            if (frm.DialogResult == DialogResult.Yes)
            {
                this.Close();
            }
            else if (frm.DialogResult == DialogResult.Cancel)
            {
                return;
            }
        }

        private void btnGeneratePDF_Click(object sender, EventArgs e)
        {

            if ((gvBillPaid.RowCount > 0) || (gvCMMPendingPayment.RowCount > 0) || (gvPending.RowCount > 0) || (gvIneligible.RowCount > 0))
            {

                //DateTime? dtDocReceivedDate = null;

                //frmDocReceivedDate frmDocumentReceivedDate = new frmDocReceivedDate();

                //frmDocumentReceivedDate.StartPosition = FormStartPosition.CenterParent;

                //var dlgResultDocReceivedDate = frmDocumentReceivedDate.ShowDialog();

                //if (dlgResultDocReceivedDate == DialogResult.OK)
                //{
                //dtDocReceivedDate = frmDocumentReceivedDate.ReceivedDate;

                Document pdfDoc = new Document();

                Section section = pdfDoc.AddSection();
                pdfDoc.UseCmykColor = true;

                section.PageSetup.PageFormat = PageFormat.Letter;
                section.PageSetup.HeaderDistance = "0.25in";
                section.PageSetup.TopMargin = "1.5in";
                //section.PageSetup.LeftMargin = "0.3in";
                //section.PageSetup.RightMargin = "0.3in";
                section.PageSetup.LeftMargin = "0.8in";
                section.PageSetup.RightMargin = "0.8in";
                section.PageSetup.BottomMargin = "0.5in";

                section.PageSetup.DifferentFirstPageHeaderFooter = false;
                section.Headers.Primary.Format.SpaceBefore = "0.25in";

                //MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\Program Files (x86)\\CMM\\BlueSheet\\cmmlogo.jpg");
                //MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\cmmlogo.png");
                MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\Program Files (x86)\\CMM\\BlueSheet\\cmmlogo.png");


                //savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_Ko";


                image.Height = "0.8in";
                image.LockAspectRatio = true;
                image.RelativeVertical = MigraDocDOM.Shapes.RelativeVertical.Line;
                image.RelativeHorizontal = MigraDocDOM.Shapes.RelativeHorizontal.Margin;
                image.Top = MigraDocDOM.Shapes.ShapePosition.Top;
                image.Left = MigraDocDOM.Shapes.ShapePosition.Center;
                image.WrapFormat.Style = MigraDocDOM.Shapes.WrapStyle.TopBottom;

                Paragraph paraCMMAddress = section.Headers.Primary.AddParagraph();
                paraCMMAddress.Format.Font.Name = "Arial";
                paraCMMAddress.Format.Font.Size = 8;
                paraCMMAddress.Format.SpaceBefore = "0.15in";
                paraCMMAddress.Format.SpaceAfter = "0.25in";
                //paraCMMAddress.Format.LeftIndent = "0.5in";
                //paraCMMAddress.Format.RightIndent = "0.5in";
                paraCMMAddress.Format.Alignment = ParagraphAlignment.Center;

                String strVerticalBar = " | ";
                String strStreet = "5235 N. Elston Ave.";
                String strCityStateZip = "Chicago, IL 60630";
                String strPhone = "Phone 773.777.8889";
                String strFax = "Fax 773.777.0004";
                String strWebsiteAddr = "www.cmmlogos.org";

                paraCMMAddress.AddFormattedText(strStreet, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strCityStateZip, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strPhone, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strFax, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strWebsiteAddr, TextFormat.NotBold);

                Paragraph paraToday = section.AddParagraph();
                paraToday.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraToday.Format.Font.Name = "Arial";
                paraToday.Format.Alignment = ParagraphAlignment.Left;
                paraToday.Format.SpaceBefore = "0.25in";
                paraToday.Format.SpaceAfter = "0.25in";
                //paraToday.Format.LeftIndent = "0.5in";
                //paraToday.Format.RightIndent = "0.5in";
                paraToday.AddFormattedText(DateTime.Today.ToString("MM/dd/yyyy"));

                Paragraph paraMembershipInfo = section.AddParagraph();

                paraMembershipInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraMembershipInfo.Format.Font.Name = "Arial";
                //paraMembershipInfo.Format.SpaceBefore = "0.70in";
                paraMembershipInfo.Format.SpaceBefore = "0.20in";
                paraMembershipInfo.Format.SpaceAfter = "0.20in";
                paraMembershipInfo.Format.LeftIndent = "0.75in";
                //paraMembershipInfo.Format.LeftIndent = "0.5in";
                //paraMembershipInfo.Format.RightIndent = "0.5in";
                paraMembershipInfo.Format.Alignment = ParagraphAlignment.Left;
                //paraMembershipInfo.AddFormattedText("Primary Name: " + strPrimaryName + "\n");
                if (strMembershipId != String.Empty) paraMembershipInfo.AddFormattedText(strMembershipId + " (" + strIndividualID + ")\n");
                else paraMembershipInfo.AddFormattedText(strIndividualID + "\n");
                paraMembershipInfo.AddFormattedText(strIndividualName + "\n");
                paraMembershipInfo.AddFormattedText(strStreetAddress + "\n");
                paraMembershipInfo.AddFormattedText(strCity + ", " + strState + " " + strZip + "\n");


                Paragraph paraDearMember = section.AddParagraph();
                paraDearMember.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraDearMember.Format.Font.Name = "Malgun Gothic";
                paraDearMember.Format.Font.Size = 8;
                paraDearMember.Format.Alignment = ParagraphAlignment.Left;
                //paraDearMember.Format.LeftIndent = "0.5in";
                //paraDearMember.Format.RightIndent = "0.5in";
                paraDearMember.Format.SpaceBefore = "0.1in";
                paraDearMember.Format.SpaceAfter = "0.1in";
                if (strIndividualMiddleName != String.Empty) paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " 회원께,");
                else paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " " + strIndividualMiddleName + " 회원께,");


                Paragraph paraGreetingMessage = section.AddParagraph();

                paraGreetingMessage.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage.Format.Font.Name = "Malgun Gothic";
                paraGreetingMessage.Format.Font.Size = 8;
                paraGreetingMessage.Format.SpaceAfter = "5pt";
                //paraGreetingMessage.Format.LeftIndent = "0.5in";
                //paraGreetingMessage.Format.RightIndent = "0.5in";
                paraGreetingMessage.Format.Alignment = ParagraphAlignment.Justify;
                //paraGreetingMessage.AddFormattedText(strGreetingMessage, TextFormat.NotBold);
                paraGreetingMessage.AddFormattedText(strGreetingMessagePara1, TextFormat.NotBold);

                Paragraph paraGreetingMessage2 = section.AddParagraph();
                paraGreetingMessage2.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage2.Format.Font.Name = "Malgun Gothic";
                paraGreetingMessage2.Format.Font.Size = 8;
                paraGreetingMessage2.Format.SpaceAfter = "5pt";
                //paraGreetingMessage2.Format.LeftIndent = "0.5in";
                //paraGreetingMessage2.Format.RightIndent = "0.5in";
                paraGreetingMessage2.Format.Alignment = ParagraphAlignment.Justify;
                paraGreetingMessage2.AddFormattedText(strGreetingMessagePara2, TextFormat.NotBold);

                Paragraph paraGreetingMessage3 = section.AddParagraph();
                paraGreetingMessage3.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage3.Format.Font.Name = "Malgun Gothic";
                paraGreetingMessage3.Format.Font.Size = 8;
                paraGreetingMessage3.Format.SpaceAfter = "5pt";
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraGreetingMessage3.Format.Alignment = ParagraphAlignment.Justify;
                paraGreetingMessage3.AddFormattedText(strGreetingMessagePara3, TextFormat.NotBold);

                Paragraph paraGreetingMessage4 = section.AddParagraph();
                paraGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage4.Format.Font.Name = "Malgun Gothic";
                paraGreetingMessage4.Format.Font.Size = 8;
                //paraGreetingMessage4.Format.LeftIndent = "0.5in";
                //paraGreetingMessage4.Format.RightIndent = "0.5in";
                paraGreetingMessage4.Format.Alignment = ParagraphAlignment.Justify;
                paraGreetingMessage4.AddFormattedText(strGreetingMessagePara4, TextFormat.NotBold);

                Paragraph paraNeedsProcessing = section.AddParagraph();

                paraNeedsProcessing.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraNeedsProcessing.Format.Font.Name = "Arial";
                paraNeedsProcessing.Format.Font.Size = 8;
                paraNeedsProcessing.Format.Font.Bold = true;
                paraNeedsProcessing.Format.Alignment = ParagraphAlignment.Left;
                paraNeedsProcessing.Format.SpaceBefore = "0.2in";
                //paraNeedsProcessing.Format.LeftIndent = "0.5in";
                //paraNeedsProcessing.Format.RightIndent = "0.5in";
                paraNeedsProcessing.AddFormattedText(strCMM_NeedProcessing + "\n");

                Paragraph paraPhoneFaxEmail = section.AddParagraph();
                paraPhoneFaxEmail.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPhoneFaxEmail.Format.Font.Size = 8;
                paraPhoneFaxEmail.Format.Alignment = ParagraphAlignment.Left;
                //paraPhoneFaxEmail.Format.LeftIndent = "0.5in";
                paraPhoneFaxEmail.Format.RightIndent = "0.5in";
                paraPhoneFaxEmail.AddFormattedText(strNP_Phone_Fax_Email + "\n");

                Paragraph paraHorizontalLine = section.AddParagraph();

                paraHorizontalLine.Format.SpaceBefore = "0.05in";
                paraHorizontalLine.Format.SpaceAfter = "0.05in";
                paraHorizontalLine.Format.Borders.Top.Width = 0;
                paraHorizontalLine.Format.Borders.Left.Width = 0;
                paraHorizontalLine.Format.Borders.Right.Width = 0;
                paraHorizontalLine.Format.Borders.Bottom.Width = 1;
                paraHorizontalLine.Format.Borders.Bottom.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraHorizontalLine.Format.Borders.Style = MigraDocDOM.BorderStyle.DashDot;

                Paragraph paraNPStatement = section.AddParagraph();
                paraNPStatement.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                paraNPStatement.Format.Font.Name = "Malgun Gothic";
                paraNPStatement.Format.Font.Size = 12;
                paraNPStatement.Format.Font.Bold = true;
                paraNPStatement.Format.Alignment = ParagraphAlignment.Center;
                paraNPStatement.Format.SpaceAfter = "0.1in";

                paraNPStatement.AddFormattedText("의료비 정산 내역서\n", TextFormat.Bold);

                Paragraph paraCheckInfo = section.AddParagraph();
                paraCheckInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraCheckInfo.Format.Font.Name = "Arial";
                paraCheckInfo.Format.Font.Size = 8;
                paraCheckInfo.Format.Font.Bold = true;
                paraCheckInfo.Format.Alignment = ParagraphAlignment.Left;
                paraCheckInfo.Format.SpaceBefore = "0.1in";
                //paraCheckInfo.Format.SpaceAfter = "0.1in";

                if (rbCheck.Checked)
                {
                    paraCheckInfo.AddFormattedText("Issue Date: " + ChkInfoEntered.dtCheckIssueDate.ToString("MM/dd/yyyy") +
                                                    "\tCheck No: " + ChkInfoEntered.CheckNumber +
                                                    "\tCheck Amount: " + ChkInfoEntered.CheckAmount.Value.ToString("C") +
                                                    "\tPaid To: " + ChkInfoEntered.PaidTo);

                }
                else if (rbACH.Checked)
                {
                    paraCheckInfo.AddFormattedText("Issue Date: " + ACHInfoEntered.dtACHDate.ToString("MM/dd/yyyy") +
                                                    "\tACH No: " + ACHInfoEntered.ACHNumber +
                                                    "\tACH Amount: " + ACHInfoEntered.ACHAmount.Value.ToString("C") +
                                                    "\tPaid To: " + ACHInfoEntered.PaidTo);
                }
                else if (rbCreditCard.Checked)
                {
                    paraCheckInfo.AddFormattedText("Date: " + CreditCardPaymentEntered.dtPaymentDate.ToString("MM/dd/yyyy") +
                                                    "\tCredit Card Payment Amount: " + CreditCardPaymentEntered.CCPaymentAmount.Value.ToString("C") +
                                                    "\tPaid To: " + CreditCardPaymentEntered.PaidTo);
                }

                //int nRowHeight = 302;
                int nRowHeight = 296;

                if (lstIncidents.Count > 0)
                {
                    Paragraph paraIncd = section.AddParagraph();

                    paraIncd.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                    paraIncd.Format.Font.Name = "Malgun Gothic";
                    paraIncd.Format.Font.Size = 8;
                    paraIncd.Format.Font.Bold = true;

                    MigraDocDOM.Tables.Table tableIncd = new MigraDocDOM.Tables.Table();
                    tableIncd.Borders.Width = 0;
                    tableIncd.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Column colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(0.85));
                    colINCD.Format.Alignment = ParagraphAlignment.Left;
                    //colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(1.1));
                    //colINCD.Format.Alignment = ParagraphAlignment.Left;
                    colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(4.5));
                    colINCD.Format.Alignment = ParagraphAlignment.Left;

                    foreach (Incident incd in lstIncidents)
                    {
                        nRowHeight += 18;

                        MigraDocDOM.Tables.Row IncdRow = tableIncd.AddRow();
                        IncdRow.Height = "0.15in";
                        IncdRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        MigraDocDOM.Tables.Cell cellIncdName = IncdRow.Cells[0];
                        cellIncdName.Format.Font.Bold = true;
                        cellIncdName.Format.Font.Size = 8;
                        //cellIncdName.Format.Font.Name = "Malgun Gothic";
                        cellIncdName.Format.Font.Name = "Arial";
                        cellIncdName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cellIncdName.AddParagraph(incd.Name + ": ");

                        //MigraDocDOM.Tables.Cell cellPatientName = IncdRow.Cells[1];
                        //cellPatientName.Format.Font.Bold = true;
                        //cellPatientName.Format.Font.Size = 8;
                        ////cellPatientName.Format.Font.Name = "Malgun Gothic";
                        //cellIncdName.Format.Font.Name = "Arial";
                        //cellPatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        //if (incd.PatientName.Length > 11)
                        //{
                        //    cellPatientName.AddParagraph(incd.PatientName.Substring(0, 11) + " ...");
                        //}
                        //else cellPatientName.AddParagraph(incd.PatientName);

                        MigraDocDOM.Tables.Cell cellICD10Code = IncdRow.Cells[1];
                        cellICD10Code.Format.Font.Bold = true;
                        cellICD10Code.Format.Font.Size = 8;
                        //cellICD10Code.Format.Font.Name = "Malgun Gothic";
                        cellIncdName.Format.Font.Name = "Arial";
                        cellICD10Code.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cellICD10Code.AddParagraph(incd.ICD10_Code);
                    }

                    pdfDoc.LastSection.Add(tableIncd);
                }

                //section.AddParagraph();
                //lstPaidMedicalExpenseTableRow.Clear();

                if (gvBillPaid.RowCount > 0)
                {

                    section.AddParagraph();
                    lstPaidMedicalExpenseTableRow.Clear();

                    //nRowHeight += 30;
                    nRowHeight += 22;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    //paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceBefore = "0.08in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                    paraSpaceBefore.Format.Font.Name = "Malgun Gothic";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("CMM이 회원 및 의료기관으로 지불한 의료비", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvBillPaid.RowCount; nRow++)
                    {

                        PaidMedicalExpenseTableRow expenseRow = new PaidMedicalExpenseTableRow();
                        expenseRow.PatientName = gvBillPaid[1, nRow].Value.ToString();
                        expenseRow.MED_BILL = gvBillPaid[2, nRow].Value.ToString();
                        //String strBillDate = gvBillPaid[3, nRow].Value.ToString();
                        //DateTime? dtBillDate = DateTime.Parse(strBillDate);
                        //expenseRow.Bill_Date = dtBillDate
                        if (gvBillPaid[3, nRow].Value.ToString() != String.Empty) expenseRow.Bill_Date = DateTime.Parse(gvBillPaid[3, nRow].Value.ToString());
                        else expenseRow.Bill_Date = null;
                        expenseRow.Medical_Provider = gvBillPaid[4, nRow].Value.ToString();
                        expenseRow.Bill_Amount = gvBillPaid[5, nRow].Value.ToString();
                        expenseRow.Personal_Responsibility = gvBillPaid[6, nRow].Value.ToString();
                        expenseRow.Member_Discount = gvBillPaid[7, nRow].Value.ToString();
                        expenseRow.CMM_Discount = gvBillPaid[8, nRow].Value.ToString();
                        expenseRow.CMM_Provider_Payment = gvBillPaid[9, nRow].Value.ToString();
                        //if (rbCheck.Checked || rbACH.Checked)
                        if (PaidTo == EnumPaidTo.Member)
                        {
                            expenseRow.PastReimbursement = gvBillPaid[10, nRow].Value.ToString();
                            expenseRow.Reimbursement = gvBillPaid[11, nRow].Value.ToString();
                        }
                        //if (rbCreditCard.Checked)
                        if (PaidTo == EnumPaidTo.MedicalProvider)
                        {
                            expenseRow.PastCMM_Provider_Payment = gvBillPaid[10, nRow].Value.ToString();
                            expenseRow.PastReimbursement = gvBillPaid[11, nRow].Value.ToString();
                            //expenseRow.Reimbursement = gvBillPaid[12, nRow].Value.ToString();
                        }
                        expenseRow.Balance = gvBillPaid[12, nRow].Value.ToString();
                        lstPaidMedicalExpenseTableRow.Add(expenseRow);
                    }



                    MigraDocDOM.Tables.Table table = new MigraDocDOM.Tables.Table();
                    table.Borders.Width = 0.1;
                    table.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column col = table.AddColumn(Unit.FromInch(0.8));
                    //col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //col = table.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    //col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(1.2));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //if (rbCreditCard.Checked)
                    //{
                    //    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    //}

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row row = table.AddRow();

                    nRowHeight += 22;
                    row.Height = "0.3in";
                    row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellTitlePatientName = row.Cells[0];
                    //cellTitlePatientName.AddParagraph("회원 이름");
                    //cellTitlePatientName.Format.Font.Bold = true;
                    //cellTitlePatientName.Format.Font.Size = 7;
                    //cellTitlePatientName.Format.Font.Name = "Malgun Gothic";
                    //cellTitlePatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleMedBill = row.Cells[0];
                    cellTitleMedBill.AddParagraph("MEDBILL");
                    cellTitleMedBill.Format.Font.Bold = true;
                    cellTitleMedBill.Format.Font.Size = 7;
                    cellTitleMedBill.Format.Font.Name = "Malgun Gothic";
                    cellTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleBillDate = row.Cells[1];
                    cellTitleBillDate.AddParagraph("서비스 날짜");
                    cellTitleBillDate.Format.Font.Bold = true;
                    cellTitleBillDate.Format.Font.Size = 7;
                    cellTitleBillDate.Format.Font.Name = "Malgun Gothic";
                    cellTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleMedicalProvider = row.Cells[2];
                    cellTitleMedicalProvider.AddParagraph("의료기관명");
                    cellTitleMedicalProvider.Format.Font.Bold = true;
                    cellTitleMedicalProvider.Format.Font.Size = 7;
                    cellTitleMedicalProvider.Format.Font.Name = "Malgun Gothic";
                    cellTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleBillAmount = row.Cells[3];
                    cellTitleBillAmount.AddParagraph("청구 (원금)");
                    cellTitleBillAmount.Format.Font.Bold = true;
                    cellTitleBillAmount.Format.Font.Size = 7;
                    cellTitleBillAmount.Format.Font.Name = "Malgun Gothic";
                    cellTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitlePersonalResponsibility = row.Cells[4];
                    cellTitlePersonalResponsibility.AddParagraph("본인 부담금");
                    cellTitlePersonalResponsibility.Format.Font.Bold = true;
                    cellTitlePersonalResponsibility.Format.Font.Size = 7;
                    cellTitlePersonalResponsibility.Format.Font.Name = "Malgun Gothic";
                    cellTitlePersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleMemberDiscount = row.Cells[5];
                    cellTitleMemberDiscount.AddParagraph("회원 (할인)");
                    cellTitleMemberDiscount.Format.Font.Bold = true;
                    cellTitleMemberDiscount.Format.Font.Size = 7;
                    cellTitleMemberDiscount.Format.Font.Name = "Malgun Gothic";
                    cellTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleCMMDiscount = row.Cells[6];
                    cellTitleCMMDiscount.AddParagraph("CMM (할인)");
                    cellTitleCMMDiscount.Format.Font.Bold = true;
                    cellTitleCMMDiscount.Format.Font.Size = 7;
                    cellTitleCMMDiscount.Format.Font.Name = "Malgun Gothic";
                    cellTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleCMMProviderDiscount = row.Cells[7];
                    cellTitleCMMProviderDiscount.AddParagraph("의료기관 지불금");
                    cellTitleCMMProviderDiscount.Format.Font.Bold = true;
                    cellTitleCMMProviderDiscount.Format.Font.Size = 7;
                    cellTitleCMMProviderDiscount.Format.Font.Name = "Malgun Gothic";
                    cellTitleCMMProviderDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //if (rbCheck.Checked || rbACH.Checked)
                    if (PaidTo == EnumPaidTo.Member)
                    {
                        MigraDocDOM.Tables.Cell cellTitlePastReimbursement = row.Cells[8];
                        cellTitlePastReimbursement.AddParagraph("기지급액");
                        cellTitlePastReimbursement.Format.Font.Bold = true;
                        cellTitlePastReimbursement.Format.Font.Size = 7;
                        cellTitlePastReimbursement.Format.Font.Name = "Malgun Gothic";
                        cellTitlePastReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        MigraDocDOM.Tables.Cell cellTitleReimbursement = row.Cells[9];
                        cellTitleReimbursement.AddParagraph("회원 환불금");
                        cellTitleReimbursement.Format.Font.Bold = true;
                        cellTitleReimbursement.Format.Font.Size = 7;
                        cellTitleReimbursement.Format.Font.Name = "Malgun Gothic";
                        cellTitleReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100); ;


                    }
                    //if (rbCreditCard.Checked)
                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        MigraDocDOM.Tables.Cell cellTitlePastReimbursement = row.Cells[8];
                        cellTitlePastReimbursement.AddParagraph("기지급액 (의료기관)");
                        cellTitlePastReimbursement.Format.Font.Bold = true;
                        cellTitlePastReimbursement.Format.Font.Size = 7;
                        cellTitlePastReimbursement.Format.Font.Name = "Malgun Gothic";
                        cellTitlePastReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        MigraDocDOM.Tables.Cell cellTitlePastCMMProviderPayment = row.Cells[9];
                        cellTitlePastCMMProviderPayment.AddParagraph("기지급액 (회원)");
                        cellTitlePastCMMProviderPayment.Format.Font.Bold = true;
                        cellTitlePastCMMProviderPayment.Format.Font.Size = 7;
                        cellTitlePastCMMProviderPayment.Format.Font.Name = "Malgun Gothic";
                        cellTitlePastCMMProviderPayment.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        //MigraDocDOM.Tables.Cell cellTitleReimbursement = row.Cells[10];
                        //cellTitleReimbursement.AddParagraph("회원 환불금");
                        //cellTitleReimbursement.Format.Font.Bold = true;
                        //cellTitleReimbursement.Format.Font.Size = 7;
                        //cellTitleReimbursement.Format.Font.Name = "Malgun Gothic";
                        //cellTitleReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100); ;


                    }
                    MigraDocDOM.Tables.Cell cellTitleBalance = row.Cells[10];
                    cellTitleBalance.AddParagraph("잔액/보류");
                    cellTitleBalance.Format.Font.Bold = true;
                    cellTitleBalance.Format.Font.Size = 7;
                    cellTitleBalance.Format.Font.Name = "Malgun Gothic";
                    cellTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);


                    for (int i = 0; i < lstPaidMedicalExpenseTableRow.Count; i++)
                    {
                        if (nRowHeight > 645) nRowHeight = 0;
                        nRowHeight += 18;
                        MigraDocDOM.Tables.Row rowData = table.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        if (i < lstPaidMedicalExpenseTableRow.Count - 1)
                        {
                            rowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstPaidMedicalExpenseTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Date.Value.ToString("MM/dd/yy"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            if (lstPaidMedicalExpenseTableRow[i].Medical_Provider.Length > 16)
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider.Substring(0, 16) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;


                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Personal_Responsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Provider_Payment);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //if (rbCheck.Checked || rbACH.Checked)
                            if (PaidTo == EnumPaidTo.Member)
                            {
                                cell = rowData.Cells[8];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                                cell.Format.Font.Bold = false;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                cell = rowData.Cells[9];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                                cell.Format.Font.Bold = false;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                            }
                            //if (rbCreditCard.Checked)
                            if (PaidTo == EnumPaidTo.MedicalProvider)
                            {
                                cell = rowData.Cells[8];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastCMM_Provider_Payment);
                                cell.Format.Font.Bold = false;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                cell = rowData.Cells[9];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                                cell.Format.Font.Bold = false;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                //cell = rowData.Cells[10];
                                //cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                                //cell.Format.Font.Bold = false;
                                //cell.Format.Font.Name = "Malgun Gothic";
                                //cell.Format.Font.Size = 7;
                                //cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                                //cell.Format.Alignment = ParagraphAlignment.Right;
                            }

                            cell = rowData.Cells[10];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                        }
                        if (i == lstPaidMedicalExpenseTableRow.Count - 1)
                        {
                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstPaidMedicalExpenseTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[1];
                            //cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Date.Value.ToString("MM/dd/yy"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            if (lstPaidMedicalExpenseTableRow[i].Medical_Provider.Length > 22)
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider.Substring(0, 22) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider);
                            }

                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Center;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Personal_Responsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Provider_Payment);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //if (rbCheck.Checked || rbACH.Checked)
                            if (PaidTo == EnumPaidTo.Member)
                            {
                                cell = rowData.Cells[8];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                                cell.Format.Font.Bold = true;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                cell = rowData.Cells[9];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                                cell.Format.Font.Bold = true;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                            }
                            //if (rbCreditCard.Checked)
                            if (PaidTo == EnumPaidTo.MedicalProvider)
                            {
                                cell = rowData.Cells[8];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastCMM_Provider_Payment);
                                cell.Format.Font.Bold = true;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                cell = rowData.Cells[9];
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                                cell.Format.Font.Bold = true;
                                cell.Format.Font.Name = "Malgun Gothic";
                                cell.Format.Font.Size = 7;
                                cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                                cell.Format.Alignment = ParagraphAlignment.Right;

                                //cell = rowData.Cells[10];
                                //cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                                //cell.Format.Font.Bold = true;
                                //cell.Format.Font.Name = "Malgun Gothic";
                                //cell.Format.Font.Size = 7;
                                //cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                                //cell.Format.Alignment = ParagraphAlignment.Right;
                            }

                            cell = rowData.Cells[10];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                    }

                    pdfDoc.LastSection.Add(table);
                }

                int nHeightAfterCMMPendingPayment = 0;

                if (gvCMMPendingPayment.RowCount > 0)
                {
                    nHeightAfterCMMPendingPayment += 22;
                    for (int nRow = 0; nRow < gvCMMPendingPayment.RowCount; nRow++)
                    {
                        nHeightAfterCMMPendingPayment += 15;
                    }

                    if ((nRowHeight > 645) ||
                        (nRowHeight + nHeightAfterCMMPendingPayment) > 645)
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }



                //////////////////////////////////////////////////////////////////////////////////////////////////

                // The beginning of CMM Pending Payment table


                if (gvCMMPendingPayment.RowCount > 0)
                {
                    lstCMMPendingPaymentTableRow.Clear();
                    nRowHeight += 30;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                    paraSpaceBefore.Format.Font.Name = "Malgun Gothic";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("지원 예정 의료비", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvCMMPendingPayment.RowCount; nRow++)
                    {
                        CMMPendingPaymentTableRow cmmPendingRow = new CMMPendingPaymentTableRow();

                        cmmPendingRow.PatientName = gvCMMPendingPayment[1, nRow].Value.ToString();
                        cmmPendingRow.MED_BILL = gvCMMPendingPayment[2, nRow].Value.ToString();
                        cmmPendingRow.Bill_Date = gvCMMPendingPayment[3, nRow].Value.ToString();
                        //cmmPendingRow.Due_Date = gvCMMPendingPayment[4, nRow].Value.ToString();
                        cmmPendingRow.Medical_Provider = gvCMMPendingPayment[4, nRow].Value.ToString();
                        cmmPendingRow.Bill_Amount = gvCMMPendingPayment[5, nRow].Value.ToString();
                        cmmPendingRow.Member_Discount = gvCMMPendingPayment[6, nRow].Value.ToString();
                        cmmPendingRow.CMM_Discount = gvCMMPendingPayment[7, nRow].Value.ToString();
                        cmmPendingRow.PersonalResponsibility = gvCMMPendingPayment[8, nRow].Value.ToString();
                        //cmmPendingRow.Member_Discount = gvCMMPendingPayment[7, nRow].Value.ToString();
                        //cmmPendingRow.CMM_Discount = gvCMMPendingPayment[8, nRow].Value.ToString();
                        cmmPendingRow.Shared_Amount = gvCMMPendingPayment[9, nRow].Value.ToString();
                        cmmPendingRow.Balance = gvCMMPendingPayment[10, nRow].Value.ToString();

                        lstCMMPendingPaymentTableRow.Add(cmmPendingRow);
                    }

                    MigraDocDOM.Tables.Table tableCMMPendingPayment = new MigraDocDOM.Tables.Table();

                    tableCMMPendingPayment.Borders.Width = 0.1;
                    tableCMMPendingPayment.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    //colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row cmm_pending_row = tableCMMPendingPayment.AddRow();

                    cmm_pending_row.Height = "0.3in";
                    cmm_pending_row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    nRowHeight += 22;

                    //MigraDocDOM.Tables.Cell cellCMMPendingTitlePatientName = cmm_pending_row.Cells[0];
                    //cellCMMPendingTitlePatientName.AddParagraph("회원 이름");
                    //cellCMMPendingTitlePatientName.Format.Font.Bold = true;
                    //cellCMMPendingTitlePatientName.Format.Font.Size = 7;
                    //cellCMMPendingTitlePatientName.Format.Font.Name = "Malgun Gothic";
                    //cellCMMPendingTitlePatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMedBill = cmm_pending_row.Cells[0];
                    cellCMMPendingTitleMedBill.AddParagraph("MEDBILL");
                    cellCMMPendingTitleMedBill.Format.Font.Bold = true;
                    cellCMMPendingTitleMedBill.Format.Font.Size = 7;
                    cellCMMPendingTitleMedBill.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBillDate = cmm_pending_row.Cells[1];
                    cellCMMPendingTitleBillDate.AddParagraph("서비스 날짜");
                    cellCMMPendingTitleBillDate.Format.Font.Bold = true;
                    cellCMMPendingTitleBillDate.Format.Font.Size = 7;
                    cellCMMPendingTitleBillDate.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellCMMPendingTitleDueDate = cmm_pending_row.Cells[3];
                    //cellCMMPendingTitleDueDate.AddParagraph("접수 날짜");
                    //cellCMMPendingTitleDueDate.Format.Font.Bold = true;
                    //cellCMMPendingTitleDueDate.Format.Font.Size = 7;
                    //cellCMMPendingTitleDueDate.Format.Font.Name = "Malgun Gothic";
                    //cellCMMPendingTitleDueDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMedicalProvider = cmm_pending_row.Cells[2];
                    cellCMMPendingTitleMedicalProvider.AddParagraph("의료기관명");
                    cellCMMPendingTitleMedicalProvider.Format.Font.Bold = true;
                    cellCMMPendingTitleMedicalProvider.Format.Font.Size = 7;
                    cellCMMPendingTitleMedicalProvider.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBillAmount = cmm_pending_row.Cells[3];
                    cellCMMPendingTitleBillAmount.AddParagraph("청구액(원금)");
                    cellCMMPendingTitleBillAmount.Format.Font.Bold = true;
                    cellCMMPendingTitleBillAmount.Format.Font.Size = 7;
                    cellCMMPendingTitleBillAmount.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMemberDiscount = cmm_pending_row.Cells[4];
                    cellCMMPendingTitleMemberDiscount.AddParagraph("회원할인");
                    cellCMMPendingTitleMemberDiscount.Format.Font.Bold = true;
                    cellCMMPendingTitleMemberDiscount.Format.Font.Size = 7;
                    cellCMMPendingTitleMemberDiscount.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleCMMDiscount = cmm_pending_row.Cells[5];
                    cellCMMPendingTitleCMMDiscount.AddParagraph("CMM 할인");
                    cellCMMPendingTitleCMMDiscount.Format.Font.Bold = true;
                    cellCMMPendingTitleCMMDiscount.Format.Font.Size = 7;
                    cellCMMPendingTitleCMMDiscount.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitlePersonalResponsibility = cmm_pending_row.Cells[6];
                    cellCMMPendingTitlePersonalResponsibility.AddParagraph("본인 부담금");
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Bold = true;
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Size = 7;
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleSharedAmount = cmm_pending_row.Cells[7];
                    cellCMMPendingTitleSharedAmount.AddParagraph("정산 완료");
                    cellCMMPendingTitleSharedAmount.Format.Font.Bold = true;
                    cellCMMPendingTitleSharedAmount.Format.Font.Size = 7;
                    cellCMMPendingTitleSharedAmount.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleSharedAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBalance = cmm_pending_row.Cells[8];
                    cellCMMPendingTitleBalance.AddParagraph("지원 예정");
                    cellCMMPendingTitleBalance.Format.Font.Bold = true;
                    cellCMMPendingTitleBalance.Format.Font.Size = 7;
                    cellCMMPendingTitleBalance.Format.Font.Name = "Malgun Gothic";
                    cellCMMPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    for (int i = 0; i < lstCMMPendingPaymentTableRow.Count; i++)
                    {
                        nRowHeight += 18;
                        MigraDocDOM.Tables.Row rowData = tableCMMPendingPayment.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        if (i < lstCMMPendingPaymentTableRow.Count - 1)
                        {
                            rowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstCMMPendingPaymentTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            //cell = rowData.Cells[3];
                            //cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Due_Date);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            if (lstCMMPendingPaymentTableRow[i].Medical_Provider.Length > 23)
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider.Substring(0, 23) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PersonalResponsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                                

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Shared_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                        if (i == lstCMMPendingPaymentTableRow.Count - 1)
                        {
                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstCMMPendingPaymentTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            //cell = rowData.Cells[3];
                            //cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Due_Date);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            if (lstCMMPendingPaymentTableRow[i].Medical_Provider.Length > 25)
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider);
                            }

                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PersonalResponsibility);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Shared_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                    }

                    pdfDoc.LastSection.Add(tableCMMPendingPayment);
                }

                int nHeightAfterPending = 0;

                if (gvPending.RowCount > 0)
                {
                    nHeightAfterPending += 22;
                    for (int nRow = 0; nRow < gvPending.RowCount; nRow++)
                    {
                        nHeightAfterPending += 15;
                    }

                    if ((nRowHeight > 645) || ((nRowHeight + nHeightAfterPending) > 645))
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }

                // Pending table
                if (gvPending.RowCount > 0)
                {
                    lstPendingTableRow.Clear();

                    nRowHeight += 18;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                    paraSpaceBefore.Format.Font.Name = "Malgun Gothic";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("보류 중인 의료비", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvPending.RowCount; nRow++)
                    {
                        nRowHeight += 12;

                        PendingTableRow pendingRow = new PendingTableRow();
                        pendingRow.PatientName = gvPending[1, nRow].Value.ToString();
                        pendingRow.MED_BILL = gvPending[2, nRow].Value.ToString();
                        pendingRow.Bill_Date = gvPending[3, nRow].Value.ToString();
                        pendingRow.Due_Date = gvPending[4, nRow].Value.ToString();
                        pendingRow.Medical_Provider = gvPending[5, nRow].Value.ToString();
                        pendingRow.Bill_Amount = gvPending[6, nRow].Value.ToString();
                        pendingRow.Balance = gvPending[7, nRow].Value.ToString();
                        //pendingRow.Member_Discount = gvPending[7, nRow].Value.ToString();
                        //pendingRow.CMM_Discount = gvPending[8, nRow].Value.ToString();
                        //pendingRow.Shared_Amount = gvPending[9, nRow].Value.ToString();
                        //pendingRow.Balance = gvPending[10, nRow].Value.ToString();
                        pendingRow.Pending_Reason = gvPending[8, nRow].Value.ToString();

                        lstPendingTableRow.Add(pendingRow);
                    }

                    MigraDocDOM.Tables.Table tablePending = new MigraDocDOM.Tables.Table();
                    tablePending.Borders.Width = 0.1;
                    tablePending.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(1.3));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;


                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(3));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row pending_Row = tablePending.AddRow();

                    nRowHeight += 22;

                    pending_Row.Height = "0.3in";
                    pending_Row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellPendingTitleINCD = pending_Row.Cells[0];
                    //cellPendingTitleINCD.AddParagraph("회원 이름");
                    //cellPendingTitleINCD.Format.Font.Bold = true;
                    //cellPendingTitleINCD.Format.Font.Size = 7;
                    //cellPendingTitleINCD.Format.Font.Name = "Malgun Gothic";
                    //cellPendingTitleINCD.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleMED_BILL = pending_Row.Cells[0];
                    cellPendingTitleMED_BILL.AddParagraph("MEDBILL");
                    cellPendingTitleMED_BILL.Format.Font.Bold = true;
                    cellPendingTitleMED_BILL.Format.Font.Size = 7;
                    cellPendingTitleMED_BILL.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleMED_BILL.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBill_Date = pending_Row.Cells[1];
                    cellPendingTitleBill_Date.AddParagraph("서비스 날짜");
                    cellPendingTitleBill_Date.Format.Font.Bold = true;
                    cellPendingTitleBill_Date.Format.Font.Size = 7;
                    cellPendingTitleBill_Date.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleBill_Date.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleDue_Date = pending_Row.Cells[2];
                    cellPendingTitleDue_Date.AddParagraph("접수 날짜");
                    cellPendingTitleDue_Date.Format.Font.Bold = true;
                    cellPendingTitleDue_Date.Format.Font.Size = 7;
                    cellPendingTitleDue_Date.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleDue_Date.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleMedicalProvider = pending_Row.Cells[3];
                    cellPendingTitleMedicalProvider.AddParagraph("의료기관명");
                    cellPendingTitleMedicalProvider.Format.Font.Bold = true;
                    cellPendingTitleMedicalProvider.Format.Font.Size = 7;
                    cellPendingTitleMedicalProvider.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBillAmount = pending_Row.Cells[4];
                    cellPendingTitleBillAmount.AddParagraph("청구 (원금)");
                    cellPendingTitleBillAmount.Format.Font.Bold = true;
                    cellPendingTitleBillAmount.Format.Font.Size = 7;
                    cellPendingTitleBillAmount.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBalance = pending_Row.Cells[5];
                    cellPendingTitleBalance.AddParagraph("잔액/보류");
                    cellPendingTitleBalance.Format.Font.Bold = true;
                    cellPendingTitleBalance.Format.Font.Size = 7;
                    cellPendingTitleBalance.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleMemberDiscount = pending_Row.Cells[6];
                    //cellPendingTitleMemberDiscount.AddParagraph("회원 (할인)");
                    //cellPendingTitleMemberDiscount.Format.Font.Bold = true;
                    //cellPendingTitleMemberDiscount.Format.Font.Size = 7;
                    //cellPendingTitleMemberDiscount.Format.Font.Name = "Malgun Gothic";
                    //cellPendingTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleCMMDiscount = pending_Row.Cells[7];
                    //cellPendingTitleCMMDiscount.AddParagraph("CMM (할인)");
                    //cellPendingTitleCMMDiscount.Format.Font.Bold = true;
                    //cellPendingTitleCMMDiscount.Format.Font.Size = 7;
                    //cellPendingTitleCMMDiscount.Format.Font.Name = "Malgun Gothic";
                    //cellPendingTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleSharedAmount = pending_Row.Cells[8];
                    //cellPendingTitleSharedAmount.AddParagraph("정산 완료");
                    //cellPendingTitleSharedAmount.Format.Font.Bold = true;
                    //cellPendingTitleSharedAmount.Format.Font.Size = 7;
                    //cellPendingTitleSharedAmount.Format.Font.Name = "Malgun Gothic";
                    //cellPendingTitleSharedAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleBalance = pending_Row.Cells[9];
                    //cellPendingTitleBalance.AddParagraph("보류");
                    //cellPendingTitleBalance.Format.Font.Bold = true;
                    //cellPendingTitleBalance.Format.Font.Size = 7;
                    //cellPendingTitleBalance.Format.Font.Name = "Malgun Gothic";
                    //cellPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitlePendingReason = pending_Row.Cells[6];
                    cellPendingTitlePendingReason.AddParagraph("보류 사유");
                    cellPendingTitlePendingReason.Format.Font.Bold = true;
                    cellPendingTitlePendingReason.Format.Font.Size = 7;
                    cellPendingTitlePendingReason.Format.Font.Name = "Malgun Gothic";
                    cellPendingTitlePendingReason.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    for (int i = 0; i < lstPendingTableRow.Count; i++)
                    {
                        MigraDocDOM.Tables.Row pendingRowData = tablePending.AddRow();
                        pendingRowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        if (i < lstPendingTableRow.Count - 1)
                        {
                            pendingRowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            //if (lstPendingTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            cell.AddParagraph(lstPendingTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[1];
                            DateTime dtBillDate = DateTime.Parse(lstPendingTableRow[i].Bill_Date);
                            cell.AddParagraph(dtBillDate.ToString("MM/dd/yy"));
                            //cell.AddParagraph(lstPendingTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[2];
                            DateTime dtDueDate = new DateTime();
                            if (lstPendingTableRow[i].Due_Date != String.Empty)
                            {
                                dtDueDate = DateTime.Parse(lstPendingTableRow[i].Due_Date);
                                cell.AddParagraph(dtDueDate.ToString("MM/dd/yy"));
                            }
                            //cell.AddParagraph(lstPendingTableRow[i].Due_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[3];

                            if (lstPendingTableRow[i].Medical_Provider.Length > 20)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider.Substring(0, 20) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = pendingRowData.Cells[4];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[5];
                            cell.AddParagraph(lstPendingTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[6];
                            //cell.AddParagraph(lstPendingTableRow[i].Member_Discount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[7];
                            //cell.AddParagraph(lstPendingTableRow[i].CMM_Discount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[8];
                            //cell.AddParagraph(lstPendingTableRow[i].Shared_Amount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[9];
                            //cell.AddParagraph(lstPendingTableRow[i].Balance);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[6];
                            if (lstPendingTableRow[i].Pending_Reason.Length > 40)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason.Substring(0, 40) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                        }
                        if (i == lstPendingTableRow.Count - 1)
                        {
                            pendingRowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            //if (lstPendingTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName.Substring(0, 11));
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            cell.AddParagraph(lstPendingTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[1];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[2];
                            cell.AddParagraph(lstPendingTableRow[i].Due_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[3];
                            if (lstPendingTableRow[i].Medical_Provider.Length > 20)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider.Substring(0, 20) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider);
                            }

                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[4];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[5];
                            cell.AddParagraph(lstPendingTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[6];
                            //cell.AddParagraph(lstPendingTableRow[i].Member_Discount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[7];
                            //cell.AddParagraph(lstPendingTableRow[i].CMM_Discount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[8];
                            //cell.AddParagraph(lstPendingTableRow[i].Shared_Amount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[9];
                            //cell.AddParagraph(lstPendingTableRow[i].Balance);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[6];
                            if (lstPendingTableRow[i].Pending_Reason.Length > 40)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason.Substring(0, 40) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            }
                            //cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                        }

                    }

                    pdfDoc.LastSection.Add(tablePending);
                }

                //int nHeightAfterIneligible = 0;

                //if (gvIneligible.RowCount > 0)
                //{
                //    nHeightAfterIneligible += 22;
                //    for (int nRow = 0; nRow < gvIneligible.RowCount; nRow++)
                //    {
                //        nHeightAfterIneligible += 15;
                //    }
                //}

                //if ((nRowHeight > 645) ||
                //    (nRowHeight + nHeightAfterIneligible) > 645)
                //{
                //    nRowHeight = 0;
                //    section.AddPageBreak();
                //}

                lstBillIneligibleTableRow.Clear();

                for (int nRow = 0; nRow < gvIneligible.RowCount; nRow++)
                {
                    //if ((gvIneligible[4, nRow].Value.ToString() != "") &&
                    //    (DateTime.Parse(gvIneligible[4, nRow].Value.ToString()) > dtDocReceivedDate.Value))
                    if (gvIneligible[4, nRow].Value.ToString() != String.Empty)
                    {
                        BillIneligibleTableRow ineligibleRow = new BillIneligibleTableRow();
                        ineligibleRow.PatientName = gvIneligible[1, nRow].Value.ToString();
                        ineligibleRow.MED_BILL = gvIneligible[2, nRow].Value.ToString();
                        ineligibleRow.Bill_Date = gvIneligible[3, nRow].Value.ToString();
                        ineligibleRow.Received_Date = gvIneligible[4, nRow].Value.ToString();
                        ineligibleRow.Medical_Provider = gvIneligible[5, nRow].Value.ToString();
                        ineligibleRow.Bill_Amount = gvIneligible[6, nRow].Value.ToString();
                        ineligibleRow.Amount_Ineligible = gvIneligible[7, nRow].Value.ToString();
                        ineligibleRow.Ineligible_Reason = gvIneligible[8, nRow].Value.ToString();

                        lstBillIneligibleTableRow.Add(ineligibleRow);
                    }
                    if (gvIneligible[4, nRow].Value.ToString() == "")
                    {
                        BillIneligibleTableRow ineligibleRow = new BillIneligibleTableRow();
                        ineligibleRow.PatientName = gvIneligible[1, nRow].Value.ToString();
                        ineligibleRow.MED_BILL = gvIneligible[2, nRow].Value.ToString();
                        ineligibleRow.Bill_Date = gvIneligible[3, nRow].Value.ToString();
                        ineligibleRow.Received_Date = gvIneligible[4, nRow].Value.ToString();
                        ineligibleRow.Medical_Provider = gvIneligible[5, nRow].Value.ToString();
                        ineligibleRow.Bill_Amount = gvIneligible[6, nRow].Value.ToString();
                        ineligibleRow.Amount_Ineligible = gvIneligible[7, nRow].Value.ToString();
                        ineligibleRow.Ineligible_Reason = gvIneligible[8, nRow].Value.ToString();

                        lstBillIneligibleTableRow.Add(ineligibleRow);
                    }
                }

                //if (lstBillIneligibleTableRow.Count > 0)
                //{
                //    double? SumBillAmount = 0;
                //    double? SumAmountIneligible = 0;

                //    for (int nRow = 0; nRow < lstBillIneligibleTableRow.Count; nRow++)
                //    {
                //        SumBillAmount += Double.Parse(lstBillIneligibleTableRow[nRow].Bill_Amount.Substring(1));
                //        SumAmountIneligible += Double.Parse(lstBillIneligibleTableRow[nRow].Amount_Ineligible.Substring(1));
                //    }
                //    BillIneligibleTableRow SumIneligibleRow = new BillIneligibleTableRow();

                //    SumIneligibleRow.Medical_Provider = "합계";
                //    SumIneligibleRow.Bill_Amount = SumBillAmount.Value.ToString("C");
                //    SumIneligibleRow.Amount_Ineligible = SumAmountIneligible.Value.ToString("C");

                //    lstBillIneligibleTableRow.Add(SumIneligibleRow);
                //}

                int nHeightAfterIneligible = 0;

                if (lstBillIneligibleTableRow.Count > 0)
                {
                    nHeightAfterIneligible += 22;

                    for (int nRow = 0; nRow < lstBillIneligibleTableRow.Count; nRow++)
                    {
                        nHeightAfterIneligible += 15;
                    }

                    if ((nRowHeight > 645) ||
                        (nRowHeight + nHeightAfterIneligible) > 645)
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }

                if (lstBillIneligibleTableRow.Count > 0)
                {

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                    paraSpaceBefore.Format.Font.Name = "Malgun Gothic";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("지원 불가한 의료비", TextFormat.Bold);

                    MigraDocDOM.Tables.Table tableIneligible = new MigraDocDOM.Tables.Table();
                    tableIneligible.Borders.Width = 0.1;
                    tableIneligible.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(1.5));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.9));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(1.2));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(2.1));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row ineligible_Row = tableIneligible.AddRow();

                    nRowHeight += 22;
                    ineligible_Row.Height = "0.31in";
                    ineligible_Row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellIneligibleTitleICND = ineligible_Row.Cells[0];
                    //cellIneligibleTitleICND.AddParagraph("회원 이름");
                    //cellIneligibleTitleICND.Format.Font.Bold = true;
                    //cellIneligibleTitleICND.Format.Font.Size = 7;
                    //cellIneligibleTitleICND.Format.Font.Name = "Malgun Gothic";
                    //cellIneligibleTitleICND.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleMedBill = ineligible_Row.Cells[0];
                    cellIneligibleTitleMedBill.AddParagraph("MEDBILL");
                    cellIneligibleTitleMedBill.Format.Font.Bold = true;
                    cellIneligibleTitleMedBill.Format.Font.Size = 7;
                    cellIneligibleTitleMedBill.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleBillDate = ineligible_Row.Cells[1];
                    cellIneligibleTitleBillDate.AddParagraph("서비스 날짜");
                    cellIneligibleTitleBillDate.Format.Font.Bold = true;
                    cellIneligibleTitleBillDate.Format.Font.Size = 7;
                    cellIneligibleTitleBillDate.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleMedicalProvider = ineligible_Row.Cells[2];
                    cellIneligibleTitleMedicalProvider.AddParagraph("의료기관명");
                    cellIneligibleTitleMedicalProvider.Format.Font.Bold = true;
                    cellIneligibleTitleMedicalProvider.Format.Font.Size = 7;
                    cellIneligibleTitleMedicalProvider.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleBillAmount = ineligible_Row.Cells[3];
                    cellIneligibleTitleBillAmount.AddParagraph("청구 (원금)");
                    cellIneligibleTitleBillAmount.Format.Font.Bold = true;
                    cellIneligibleTitleBillAmount.Format.Font.Size = 7;
                    cellIneligibleTitleBillAmount.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleAmountIneligible = ineligible_Row.Cells[4];
                    cellIneligibleTitleAmountIneligible.AddParagraph("전액/일부 지원불가 금액");
                    cellIneligibleTitleAmountIneligible.Format.Font.Bold = true;
                    cellIneligibleTitleAmountIneligible.Format.Font.Size = 7;
                    cellIneligibleTitleAmountIneligible.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleAmountIneligible.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleIneligibleReason = ineligible_Row.Cells[5];
                    cellIneligibleTitleIneligibleReason.AddParagraph("지원되지 않는 사유");
                    cellIneligibleTitleIneligibleReason.Format.Font.Bold = true;
                    cellIneligibleTitleIneligibleReason.Format.Font.Size = 7;
                    cellIneligibleTitleIneligibleReason.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleTitleIneligibleReason.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    List<BillIneligibleRow> lstBillIneligible = new List<BillIneligibleRow>();

                    for (int i = 0; i < lstBillIneligibleTableRow.Count; i++)
                    {
                        if (i < lstBillIneligibleTableRow.Count - 1)
                        {

                            MigraDocDOM.Tables.Row rowData = tableIneligible.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            BillIneligibleRow ineligible = new BillIneligibleRow();

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];

                            //if (lstBillIneligibleTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName);
                            //}

                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];

                            if (lstBillIneligibleTableRow[i].Medical_Provider.Length > 24)
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider.Substring(0, 24) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider);
                            }

                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Amount);
                            ineligible.Bill_Amount = Double.Parse(lstBillIneligibleTableRow[i].Bill_Amount.Substring(1));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Amount_Ineligible);
                            ineligible.Amount_Ineligible = Double.Parse(lstBillIneligibleTableRow[i].Amount_Ineligible.Substring(1));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;



                            cell = rowData.Cells[5];
                            if (lstBillIneligibleTableRow[i].Ineligible_Reason.Length > 33)
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason.Substring(0, 33) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            lstBillIneligible.Add(ineligible);

                            PdfDocument doc = new PdfDocument();

                                
                        }

                        if (i == lstBillIneligibleTableRow.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligible.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //cell.AddParagraph();
                            //if (lstBillIneligibleTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //cell.AddParagraph(lstBillIneligibleTableRow[i].MED_BILL);
                            cell.AddParagraph();
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[1];
                            //cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Date);
                            cell.AddParagraph();
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            //if (lstBillIneligibleTableRow[i].Medical_Provider.Length > 25)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider);
                            //}
                            cell.AddParagraph("합계");
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            Double? BillAmount = 0;
                            foreach (BillIneligibleRow row in lstBillIneligible)
                            {
                                BillAmount += row.Bill_Amount;
                            }

                            cell = rowData.Cells[3];
                            //cell.AddParagraph(BillAmount.Value.ToString("C"));
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //Double? IneligibleAmount = 0;
                            //foreach (BillIneligibleRow row in lstBillIneligible)
                            //{
                            //    IneligibleAmount += row.Amount_Ineligible;
                            //}

                            cell = rowData.Cells[4];
                            //cell.AddParagraph(IneligibleAmount.Value.ToString("C"));
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Amount_Ineligible);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            //if (lstBillIneligibleTableRow[i].Ineligible_Reason.Length > 33)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason.Substring(0, 33) + " ...");
                            //}
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason);
                            //}
                            cell.AddParagraph();
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                    }

                    pdfDoc.LastSection.Add(tableIneligible);

                }
                // The end of tables

                const bool unicode = true;
                const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);
                pdfRenderer.Document = pdfDoc;
                pdfRenderer.RenderDocument();


                if (rbCheck.Checked)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + ChkInfoEntered.dtCheckIssueDate.ToString("MM-dd-yyyy") + "_Ko";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                        //finally
                        //{
                        //    ChkInfoEntered = null;
                        //}
                    }
                }
                if (rbACH.Checked)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + ACHInfoEntered.dtACHDate.ToString("MM-dd-yyyy") + "_Ko";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                        //finally
                        //{
                        //    ACHInfoEntered = null;
                        //}
                    }
                }
                if (rbCreditCard.Checked)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + CreditCardPaymentEntered.dtPaymentDate.ToString("MM-dd-yyyy") + "_Ko";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                        //finally
                        //{
                        //    ACHInfoEntered = null;
                        //}
                    }
                }
                //}
                //else if (dlgResultDocReceivedDate == DialogResult.Cancel)
                //{
                //    return;
                //}

            }
            else if ((gvPersonalResponsibility.RowCount > 0)||(gvIneligibleNoSharing.RowCount > 0))
            {

                List<PersonalResponsibilityInfo> lstPersonalResponsibilityInfo = new List<PersonalResponsibilityInfo>();

                for (int i = 0; i < gvPersonalResponsibility.Rows.Count; i++)
                {
                    PersonalResponsibilityInfo prInfo = new PersonalResponsibilityInfo();

                    prInfo.MedBillName = gvPersonalResponsibility[0, i]?.Value.ToString();
                    String BillDate = gvPersonalResponsibility[1, i]?.Value.ToString();
                    if (BillDate != String.Empty) prInfo.BillDate = DateTime.Parse(BillDate);
                    prInfo.MedicalProvider = gvPersonalResponsibility[2, i]?.Value.ToString();
                    if (gvPersonalResponsibility[3, i] != null) prInfo.BillAmount = (Double)Decimal.Parse(gvPersonalResponsibility[3, i].Value.ToString().Substring(1));                        
                    prInfo.Type = gvPersonalResponsibility[4, i]?.Value.ToString();
                    if (gvPersonalResponsibility[5, i] != null) prInfo.PersonalResponsibilityTotal = (Double)Decimal.Parse(gvPersonalResponsibility[8, i].Value.ToString().Substring(1));

                    lstPersonalResponsibilityInfo.Add(prInfo);
                }

                Document pdfPersonalResponsibilityDoc = new Document();

                Section section = pdfPersonalResponsibilityDoc.AddSection();

                section.PageSetup.PageFormat = PageFormat.Letter;
                section.PageSetup.HeaderDistance = "0.25in";
                section.PageSetup.TopMargin = "1.5in";
                section.PageSetup.LeftMargin = "0.8in";
                section.PageSetup.RightMargin = "0.8in";
                section.PageSetup.BottomMargin = "0.5in";

                section.PageSetup.DifferentFirstPageHeaderFooter = false;
                section.Headers.Primary.Format.SpaceBefore = "0.25in";

                MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\Program Files (x86)\\CMM\\BlueSheet\\cmmlogo.png");

                image.Height = "0.8in";
                image.LockAspectRatio = true;
                image.RelativeVertical = MigraDocDOM.Shapes.RelativeVertical.Line;
                image.RelativeHorizontal = MigraDocDOM.Shapes.RelativeHorizontal.Margin;
                image.Top = MigraDocDOM.Shapes.ShapePosition.Top;
                image.Left = MigraDocDOM.Shapes.ShapePosition.Center;
                image.WrapFormat.Style = MigraDocDOM.Shapes.WrapStyle.TopBottom;

                Paragraph paraCMMAddress = section.Headers.Primary.AddParagraph();
                paraCMMAddress.Format.Font.Name = "Arial";
                paraCMMAddress.Format.Font.Size = 8;
                paraCMMAddress.Format.SpaceBefore = "0.15in";
                paraCMMAddress.Format.SpaceAfter = "0.25in";
                paraCMMAddress.Format.Alignment = ParagraphAlignment.Center;

                String strVerticalBar = " | ";
                String strStreet = "5235 N. Elston Ave.";
                String strCityStateZip = "Chicago, IL 60630";
                String strPhone = "Phone 773.777.8889";
                String strFax = "Fax 773.777.0004";
                String strWebsiteAddr = "www.cmmlogos.org";

                paraCMMAddress.AddFormattedText(strStreet, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strCityStateZip, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strPhone, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strFax, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strWebsiteAddr, TextFormat.NotBold);

                Paragraph paraToday = section.AddParagraph();
                paraToday.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraToday.Format.Font.Name = "Arial";
                paraToday.Format.Alignment = ParagraphAlignment.Left;
                paraToday.Format.SpaceBefore = "0.25in";
                paraToday.Format.SpaceAfter = "0.25in";
                paraToday.AddFormattedText(DateTime.Today.ToString("MM/dd/yyyy"));

                Paragraph paraMembershipInfo = section.AddParagraph();

                paraMembershipInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraMembershipInfo.Format.Font.Name = "Arial";
                //paraMembershipInfo.Format.SpaceBefore = "0.70in";
                paraMembershipInfo.Format.SpaceBefore = "0.20in";
                paraMembershipInfo.Format.SpaceAfter = "0.20in";
                paraMembershipInfo.Format.LeftIndent = "0.75in";
                //paraMembershipInfo.Format.LeftIndent = "1.00in";
                //paraMembershipInfo.Format.LeftIndent = "0.5in";
                //paraMembershipInfo.Format.RightIndent = "0.5in";
                paraMembershipInfo.Format.Alignment = ParagraphAlignment.Left;
                //paraMembershipInfo.AddFormattedText("Primary Name: " + strPrimaryName + "\n");
                if (strMembershipId != String.Empty) paraMembershipInfo.AddFormattedText(strMembershipId + " (" + strIndividualID + ")\n");
                else paraMembershipInfo.AddFormattedText(strIndividualID + "\n");
                paraMembershipInfo.AddFormattedText(strIndividualName + "\n");
                paraMembershipInfo.AddFormattedText(strStreetAddress + "\n");
                paraMembershipInfo.AddFormattedText(strCity + ", " + strState + " " + strZip + "\n");


                Paragraph paraDearMember = section.AddParagraph();
                paraDearMember.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraDearMember.Format.Font.Name = "Malgun Gothic";
                paraDearMember.Format.Font.Size = 8;
                paraDearMember.Format.Alignment = ParagraphAlignment.Left;
                //paraDearMember.Format.LeftIndent = "0.5in";
                //paraDearMember.Format.RightIndent = "0.5in";
                paraDearMember.Format.SpaceBefore = "0.1in";
                paraDearMember.Format.SpaceAfter = "0.1in";
                //if (strIndividualMiddleName != String.Empty) paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " 회원께,");
                //else paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " " + strIndividualMiddleName + " 회원께,");
                paraDearMember.AddFormattedText(strIndividualName + " 회원께,");

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Paragraph paraPRGreetingMessage1 = section.AddParagraph();

                paraPRGreetingMessage1.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage1.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage1.Format.Font.Size = 8;
                paraPRGreetingMessage1.Format.SpaceAfter = "5pt";
                //paraGreetingMessage.Format.LeftIndent = "0.5in";
                //paraGreetingMessage.Format.RightIndent = "0.5in";
                paraPRGreetingMessage1.Format.Alignment = ParagraphAlignment.Justify;
                //paraGreetingMessage.AddFormattedText(strGreetingMessage, TextFormat.NotBold);
                paraPRGreetingMessage1.AddFormattedText(strPRGreetingMessagePara1, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage2 = section.AddParagraph();
                paraPRGreetingMessage2.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage2.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage2.Format.Font.Size = 8;
                paraPRGreetingMessage2.Format.SpaceAfter = "5pt";
                //paraGreetingMessage2.Format.LeftIndent = "0.5in";
                //paraGreetingMessage2.Format.RightIndent = "0.5in";
                paraPRGreetingMessage2.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage2.AddFormattedText(strPRGreetingMessagePara2, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage3 = section.AddParagraph();
                paraPRGreetingMessage3.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage3.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage3.Format.Font.Size = 8;
                paraPRGreetingMessage3.Format.SpaceAfter = "5pt";
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraPRGreetingMessage3.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage3.AddFormattedText(strPRGreetingMessagePara3, TextFormat.NotBold);

                //////////////////////////////////////////////////////////////////////////////////////////////////////
                /// Program personal responsibility table
                /// 

                MigraDocDOM.Tables.Table tableProgramPRGuide = new MigraDocDOM.Tables.Table();
                tableProgramPRGuide.Borders.Width = 0.1;
                tableProgramPRGuide.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Column colProgram = tableProgramPRGuide.AddColumn(MigraDocDOM.Unit.FromInch(2));
                colProgram.Format.Alignment = ParagraphAlignment.Left;
                MigraDocDOM.Tables.Column colPersonalResponsibility = tableProgramPRGuide.AddColumn(MigraDocDOM.Unit.FromInch(4.8));
                colPersonalResponsibility.Format.Alignment = ParagraphAlignment.Left;

                MigraDocDOM.Tables.Row rowHeader = tableProgramPRGuide.AddRow();
                rowHeader.Height = "0.2in";
                rowHeader.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellProgramHeader = rowHeader.Cells[0];
                cellProgramHeader.Format.Font.Bold = true;
                cellProgramHeader.Format.Font.Size = 8;
                cellProgramHeader.Format.Font.Name = "Malgun Gothic";
                cellProgramHeader.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellProgramHeader.AddParagraph("프로그램");

                MigraDocDOM.Tables.Cell cellPersonalResponsibility = rowHeader.Cells[1];
                cellPersonalResponsibility.Format.Font.Bold = true;
                cellPersonalResponsibility.Format.Font.Size = 8;
                cellPersonalResponsibility.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellPersonalResponsibility.AddParagraph("본인부담금");

                MigraDocDOM.Tables.Row rowBronze = tableProgramPRGuide.AddRow();
                rowBronze.Height = "0.2in";
                rowBronze.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellBronze = rowBronze.Cells[0];
                cellBronze.Format.Font.Bold = false;
                cellBronze.Format.Font.Size = 8;
                cellBronze.Format.Font.Name = "Malgun Gothic";
                cellBronze.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellBronze.AddParagraph("브론즈 Bronze");

                MigraDocDOM.Tables.Cell cellBronzePR = rowBronze.Cells[1];
                cellBronzePR.Format.Font.Bold = false;
                cellBronzePR.Format.Font.Size = 8;
                cellBronzePR.Format.Font.Name = "Malgun Gothic";
                cellBronzePR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellBronzePR.AddParagraph("$5,000 Per Incident");

                MigraDocDOM.Tables.Row rowSilver = tableProgramPRGuide.AddRow();
                rowSilver.Height = "0.2in";
                rowSilver.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellSilver = rowSilver.Cells[0];
                cellSilver.Format.Font.Bold = false;
                cellSilver.Format.Font.Size = 8;
                cellSilver.Format.Font.Name = "Malgun Gothic";
                cellSilver.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellSilver.AddParagraph("실버 Silver");

                MigraDocDOM.Tables.Cell cellSilverPR = rowSilver.Cells[1];
                cellSilverPR.Format.Font.Bold = false;
                cellSilverPR.Format.Font.Size = 8;
                cellSilverPR.Format.Font.Name = "Malgun Gothic";
                cellSilverPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellSilverPR.AddParagraph("$1,000 Per Incident");

                MigraDocDOM.Tables.Row rowGold = tableProgramPRGuide.AddRow();
                rowGold.Height = "0.2in";
                rowGold.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellGold = rowGold.Cells[0];
                cellGold.Format.Font.Bold = false;
                cellGold.Format.Font.Size = 8;
                cellGold.Format.Font.Name = "Malgun Gothic";
                cellGold.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGold.AddParagraph("골드 Gold");

                MigraDocDOM.Tables.Cell cellGoldPR = rowGold.Cells[1];
                cellGoldPR.Format.Font.Bold = false;
                cellGoldPR.Format.Font.Size = 8;
                cellGoldPR.Format.Font.Name = "Malgun Gothic";
                cellGoldPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPR.AddParagraph("$500 Per Incident");

                MigraDocDOM.Tables.Row rowGoldPlus = tableProgramPRGuide.AddRow();
                rowGoldPlus.Height = "0.2in";
                rowGoldPlus.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellGoldPlus = rowGoldPlus.Cells[0];
                cellGoldPlus.Format.Font.Bold = false;
                cellGoldPlus.Format.Font.Size = 8;
                cellGoldPlus.Format.Font.Name = "Malgun Gothic";
                cellGoldPlus.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPlus.AddParagraph("골드플러스 Gold Plus");

                MigraDocDOM.Tables.Cell cellGoldPlusPR = rowGoldPlus.Cells[1];
                cellGoldPlusPR.Format.Font.Bold = false;
                cellGoldPlusPR.Format.Font.Size = 8;
                cellGoldPlusPR.Format.Font.Name = "Malgun Gothic";
                cellGoldPlusPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPlusPR.AddParagraph("$500 Per Membership Anniversary");

                pdfPersonalResponsibilityDoc.LastSection.Add(tableProgramPRGuide);

                Paragraph paraPRGreetingMessage4 = section.AddParagraph();
                paraPRGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage4.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage4.Format.Font.Size = 8;
                //paraGreetingMessage4.Format.LeftIndent = "0.5in";
                //paraGreetingMessage4.Format.RightIndent = "0.5in";
                paraPRGreetingMessage4.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage4.AddFormattedText(strPRGreetingMessagePara4, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage5 = section.AddParagraph();
                paraPRGreetingMessage5.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage5.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage5.Format.Font.Size = 8;
                //paraGreetingMessage4.Format.LeftIndent = "0.5in";
                //paraGreetingMessage4.Format.RightIndent = "0.5in";
                paraPRGreetingMessage5.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage5.AddFormattedText(strPRGreetingMessagePara5, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage6 = section.AddParagraph();
                paraPRGreetingMessage6.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage6.Format.Font.Name = "Malgun Gothic";
                paraPRGreetingMessage6.Format.Font.Size = 8;
                //paraGreetingMessage4.Format.LeftIndent = "0.5in";
                //paraGreetingMessage4.Format.RightIndent = "0.5in";
                paraPRGreetingMessage6.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage6.AddFormattedText(strPRGreetingMessagePara6, TextFormat.NotBold);




                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                Paragraph paraNeedsProcessing = section.AddParagraph();

                paraNeedsProcessing.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraNeedsProcessing.Format.Font.Name = "Arial";
                paraNeedsProcessing.Format.Font.Size = 8;
                paraNeedsProcessing.Format.Font.Bold = true;
                paraNeedsProcessing.Format.Alignment = ParagraphAlignment.Left;
                paraNeedsProcessing.Format.SpaceBefore = "0.2in";
                //paraNeedsProcessing.Format.LeftIndent = "0.5in";
                //paraNeedsProcessing.Format.RightIndent = "0.5in";
                paraNeedsProcessing.AddFormattedText(strCMM_NeedProcessing + "\n");

                Paragraph paraPhoneFaxEmail = section.AddParagraph();
                paraPhoneFaxEmail.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPhoneFaxEmail.Format.Font.Size = 8;
                paraPhoneFaxEmail.Format.Alignment = ParagraphAlignment.Left;
                //paraPhoneFaxEmail.Format.LeftIndent = "0.5in";
                paraPhoneFaxEmail.Format.RightIndent = "0.5in";
                paraPhoneFaxEmail.AddFormattedText(strNP_Phone_Fax_Email + "\n");

                Paragraph paraHorizontalLine = section.AddParagraph();

                paraHorizontalLine.Format.SpaceBefore = "0.05in";
                paraHorizontalLine.Format.SpaceAfter = "0.05in";
                paraHorizontalLine.Format.Borders.Top.Width = 0;
                paraHorizontalLine.Format.Borders.Left.Width = 0;
                paraHorizontalLine.Format.Borders.Right.Width = 0;
                paraHorizontalLine.Format.Borders.Bottom.Width = 1;
                paraHorizontalLine.Format.Borders.Bottom.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraHorizontalLine.Format.Borders.Style = MigraDocDOM.BorderStyle.DashDot;

                Paragraph paraNPStatement = section.AddParagraph();
                paraNPStatement.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                paraNPStatement.Format.Font.Name = "Malgun Gothic";
                paraNPStatement.Format.Font.Size = 12;
                paraNPStatement.Format.Font.Bold = true;
                paraNPStatement.Format.Alignment = ParagraphAlignment.Center;
                paraNPStatement.Format.SpaceAfter = "0.1in";

                //paraNPStatement.AddFormattedText("본인 부담금 또는 지원 불가 의료비 내역서\n", TextFormat.Bold);
                paraNPStatement.AddFormattedText("본인 부담금 내역서\n", TextFormat.Bold);

                Paragraph paraPersonalResponsibilityTotal = section.AddParagraph();
                paraPersonalResponsibilityTotal.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPersonalResponsibilityTotal.Format.Font.Name = "Arial";
                paraPersonalResponsibilityTotal.Format.Font.Size = 8;
                paraPersonalResponsibilityTotal.Format.Font.Bold = true;
                paraPersonalResponsibilityTotal.Format.Alignment = ParagraphAlignment.Left;
                paraPersonalResponsibilityTotal.Format.SpaceBefore = "0.2in";
                //paraPersonalResponsibilityTotal.Format.SpaceAfter = "0.2in";

                paraPersonalResponsibilityTotal.AddFormattedText("Incident Occurrence Date: " + PersonalResponsibilityTotalEntered.IncidentOccurrenceDate.Value.ToString("MM/dd/yyyy"));

                Paragraph paraINCDNumber = section.AddParagraph();
                paraINCDNumber.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraINCDNumber.Format.Font.Name = "Arial";
                paraINCDNumber.Format.Font.Size = 8;
                paraINCDNumber.Format.Font.Bold = true;
                paraINCDNumber.Format.Alignment = ParagraphAlignment.Left;
                paraINCDNumber.Format.SpaceBefore = "0.05in";

                paraINCDNumber.AddFormattedText(PersonalResponsibilityTotalEntered.IncidentNo + ": " + PersonalResponsibilityTotalEntered.ICD10CodeDescription);

                //paraINCDNumber.AddFormattedText(PersonalResponsibilityTotalEntered.IncidentNo + ": " + PersonalResponsibilityTotalEntered.)



                //paraPersonalResponsibilityTotal.AddFormattedText("Incident Occurrence Date: " + PersonalResponsibilityTotalEntered.IncidentOccurrenceDate.Value.ToString("MM/dd/yyyy") + "\t" +
                //                                 "Personal Responsibility Total: " + PersonalResponsibilityTotalEntered.PersonalResponsibilityTotal.ToString("C"));


                paraPersonalResponsibilityTotal.Format.SpaceAfter = "0.05in";

                Paragraph paraPersonalResponsibilityTitle = section.AddParagraph();
                paraPersonalResponsibilityTitle.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPersonalResponsibilityTitle.Format.Font.Name = "Malgun Gothic";
                paraPersonalResponsibilityTitle.Format.Font.Size = 7;
                paraPersonalResponsibilityTitle.Format.Font.Bold = true;
                paraPersonalResponsibilityTitle.Format.Alignment = ParagraphAlignment.Left;
                paraPersonalResponsibilityTitle.Format.SpaceBefore = "0.18in";
                paraPersonalResponsibilityTitle.Format.SpaceAfter = "0.05in";
                paraPersonalResponsibilityTitle.AddFormattedText("본인부담금", TextFormat.Bold);


                //Paragraph paraIncidentInfo = section.AddParagraph();
                //paraIncidentInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //paraIncidentInfo.Format.Font.Name = "Arial";
                //paraIncidentInfo.Format.Font.Size = 8;
                //paraIncidentInfo.Format.Font.Bold = true;
                //paraIncidentInfo.Format.Alignment = ParagraphAlignment.Left;
                //paraIncidentInfo.Format.SpaceBefore = "0.2in";
                //paraIncidentInfo.Format.SpaceAfter = "0.2in";

                //if (lstIncidents.Count > 0)
                //{

                //    // 09/18/18 begin here
                //    Paragraph paraIncd = section.AddParagraph();

                //    paraIncd.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                //    paraIncd.Format.Font.Name = "Malgun Gothic";
                //    paraIncd.Format.Font.Size = 8;
                //    paraIncd.Format.Font.Bold = true;
                //    paraIncd.Format.SpaceAfter = "1in";

                //    MigraDocDOM.Tables.Table tableIncd = new MigraDocDOM.Tables.Table();
                //    tableIncd.Borders.Width = 0;
                //    tableIncd.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //    tableIncd.Format.SpaceAfter = "0.05in";

                //    MigraDocDOM.Tables.Column colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(0.85));
                //    colINCD.Format.Alignment = ParagraphAlignment.Left;
                //    colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(4.5));
                //    colINCD.Format.Alignment = ParagraphAlignment.Left;

                //    foreach (Incident incd in lstIncidents)
                //    {
                //        //nRowHeight += 18;

                //        MigraDocDOM.Tables.Row IncdRow = tableIncd.AddRow();
                //        IncdRow.Height = "0.1in";
                //        IncdRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                //        MigraDocDOM.Tables.Cell cellIncdName = IncdRow.Cells[0];
                //        cellIncdName.Format.Font.Bold = true;
                //        cellIncdName.Format.Font.Size = 8;
                //        //cellIncdName.Format.Font.Name = "Malgun Gothic";
                //        cellIncdName.Format.Font.Name = "Arial";
                //        cellIncdName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //        cellIncdName.AddParagraph(incd.Name + ": ");

                //        MigraDocDOM.Tables.Cell cellICD10Code = IncdRow.Cells[1];
                //        cellICD10Code.Format.Font.Bold = true;
                //        cellICD10Code.Format.Font.Size = 8;
                //        //cellICD10Code.Format.Font.Name = "Malgun Gothic";
                //        cellIncdName.Format.Font.Name = "Arial";
                //        cellICD10Code.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //        cellICD10Code.AddParagraph(incd.ICD10_Code);
                //    }

                //    pdfPersonalResponsibilityDoc.LastSection.Add(tableIncd);
                //}

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////

                MigraDocDOM.Tables.Table tablePersonalResponsibility = new MigraDocDOM.Tables.Table();
                tablePersonalResponsibility.Borders.Width = 0.1;
                tablePersonalResponsibility.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Column colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.6));  // Med bill
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.7));        // 서비스 날짜
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.8));        // Medical Provider
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.8));        // Bill Amount
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.3));        // Type
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.5));        // Personal Responsibility Total
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                // generate row for personal responsibility table
                MigraDocDOM.Tables.Row prRow = tablePersonalResponsibility.AddRow();
                prRow.Height = "0.31in";
                prRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityMedBill = prRow.Cells[0];
                cellPersonalResponsibilityMedBill.AddParagraph("MEDBILL");
                cellPersonalResponsibilityMedBill.Format.Font.Bold = true;
                cellPersonalResponsibilityMedBill.Format.Font.Size = 7;
                cellPersonalResponsibilityMedBill.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityBillDate = prRow.Cells[1];
                cellPersonalResponsibilityBillDate.AddParagraph("서비스 날짜");
                cellPersonalResponsibilityBillDate.Format.Font.Bold = true;
                cellPersonalResponsibilityBillDate.Format.Font.Size = 7;
                cellPersonalResponsibilityBillDate.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityMedicalProvider = prRow.Cells[2];
                cellPersonalResponsibilityMedicalProvider.AddParagraph("의료기관명");
                cellPersonalResponsibilityMedicalProvider.Format.Font.Bold = true;
                cellPersonalResponsibilityMedicalProvider.Format.Font.Size = 7;
                cellPersonalResponsibilityMedicalProvider.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityBillAmount = prRow.Cells[3];
                cellPersonalResponsibilityBillAmount.AddParagraph("청구액(원금)");
                cellPersonalResponsibilityBillAmount.Format.Font.Bold = true;
                cellPersonalResponsibilityBillAmount.Format.Font.Size = 7;
                cellPersonalResponsibilityBillAmount.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityType = prRow.Cells[4];
                cellPersonalResponsibilityType.AddParagraph("Type");
                cellPersonalResponsibilityType.Format.Font.Bold = true;
                cellPersonalResponsibilityType.Format.Font.Size = 7;
                cellPersonalResponsibilityType.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityType.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityTotal = prRow.Cells[5];
                cellPersonalResponsibilityTotal.AddParagraph("Personal Responsibility Total");
                cellPersonalResponsibilityTotal.Format.Font.Bold = true;
                cellPersonalResponsibilityTotal.Format.Font.Size = 7;
                cellPersonalResponsibilityTotal.Format.Font.Name = "Malgun Gothic";
                cellPersonalResponsibilityTotal.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);


                for (int i = 0; i < lstPersonalResponsibilityInfo.Count; i++)
                {
                    if (i < lstPersonalResponsibilityInfo.Count - 1)
                    {
                        MigraDocDOM.Tables.Row rowData = tablePersonalResponsibility.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                        rowData.Height = "0.18in";

                        MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedBillName.Substring(8));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[1];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillDate.Value.ToString("MM/dd/yyyy"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[2];
                        if (lstPersonalResponsibilityInfo[i].MedicalProvider.Length > 30) cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider.Substring(0, 29) + "...");
                        else cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider);
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[3];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillAmount.ToString("C"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;

                        cell = rowData.Cells[4];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].Type);
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[5];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].PersonalResponsibilityTotal.ToString("C"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;
                    }
                    else if (i == lstPersonalResponsibilityInfo.Count - 1)
                    {
                        MigraDocDOM.Tables.Row rowData = tablePersonalResponsibility.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                        rowData.Height = "0.2in";

                        MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedBillName);
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[1];
                        if (lstPersonalResponsibilityInfo[i].BillDate != null) cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillDate.Value.ToString("MM/dd/yyyy"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[2];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider);
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[3];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillAmount.ToString("C"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;

                        cell = rowData.Cells[4];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].Type);
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[5];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].PersonalResponsibilityTotal.ToString("C"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Malgun Gothic";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;
                    }
                }

                pdfPersonalResponsibilityDoc.LastSection.Add(tablePersonalResponsibility);

                Paragraph verticalSpace = section.AddParagraph();

                verticalSpace.Format.SpaceBefore = "0.05in";
                //verticalSpace.Format.SpaceAfter = "0.1in";

                // put No PR, No Sharing code for generating pdf here

                if (gvIneligibleNoSharing.Rows.Count > 0)
                {
                    List<SettlementIneligibleInfo> lstMedBillNoPRNoSharing = new List<SettlementIneligibleInfo>();

                    for(int i = 0; i < gvIneligibleNoSharing.Rows.Count; i++)
                    {
                        SettlementIneligibleInfo info = new SettlementIneligibleInfo();

                        //String BillDate = gvPersonalResponsibility[1, i]?.Value.ToString();
                        //if (BillDate != String.Empty) prInfo.BillDate = DateTime.Parse(BillDate);


                        info.MedBillName = gvIneligibleNoSharing["MEDBILL", i]?.Value.ToString();
                        //info.BillDate = DateTime.Parse(gvIneligibleNoSharing["서비스 날짜", i].Value.ToString());
                        String BillDate = gvIneligibleNoSharing["서비스 날짜", i]?.Value.ToString();
                        if (BillDate != String.Empty) info.BillDate = DateTime.Parse(BillDate);

                        info.MedicalProvider = gvIneligibleNoSharing["의료기관명", i]?.Value.ToString();
                        info.BillAmount = Double.Parse(gvIneligibleNoSharing["청구액(원금)", i]?.Value.ToString().Substring(1));
                        //info.Type = gvIneligibleNoSharing["Type", i]?.Value.ToString();
                        info.IneligibleAmount = Double.Parse(gvIneligibleNoSharing["지원불가 의료비", i]?.Value.ToString().Substring(1));
                        info.IneligibleReason = gvIneligibleNoSharing["지원불가 사유", i]?.Value.ToString();

                        lstMedBillNoPRNoSharing.Add(info);
                    }

                    Paragraph paraIneligibleTitle = section.AddParagraph();
                    paraIneligibleTitle.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                    paraIneligibleTitle.Format.Font.Name = "Malgun Gothic";
                    paraIneligibleTitle.Format.Font.Size = 7;
                    paraIneligibleTitle.Format.Font.Bold = true;
                    paraIneligibleTitle.Format.Alignment = ParagraphAlignment.Left;
                    paraIneligibleTitle.Format.SpaceAfter = "0.05in";
                    paraIneligibleTitle.AddFormattedText("지원불가 의료비", TextFormat.Bold);


                    MigraDocDOM.Tables.Table tableIneligibleNoPR = new MigraDocDOM.Tables.Table();
                    tableIneligibleNoPR.Borders.Width = 0.1;
                    tableIneligibleNoPR.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Column colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(2.5));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    //colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row ineligibleRow = tableIneligibleNoPR.AddRow();
                    ineligibleRow.Height = "0.31in";
                    ineligibleRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRMedBill = ineligibleRow.Cells[0];
                    cellIneligibleNoPRMedBill.AddParagraph("MEDBILL");
                    cellIneligibleNoPRMedBill.Format.Font.Bold = true;
                    cellIneligibleNoPRMedBill.Format.Font.Size = 7;
                    cellIneligibleNoPRMedBill.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleNoPRMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRBillDate = ineligibleRow.Cells[1];
                    cellIneligibleNoPRBillDate.AddParagraph("서비스 날짜");
                    cellIneligibleNoPRBillDate.Format.Font.Bold = true;
                    cellIneligibleNoPRBillDate.Format.Font.Size = 7;
                    cellIneligibleNoPRBillDate.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleNoPRBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRMedicalProvider = ineligibleRow.Cells[2];
                    cellIneligibleNoPRMedicalProvider.AddParagraph("의료기관명");
                    cellIneligibleNoPRMedicalProvider.Format.Font.Bold = true;
                    cellIneligibleNoPRMedicalProvider.Format.Font.Size = 7;
                    cellIneligibleNoPRMedicalProvider.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRBillAmount = ineligibleRow.Cells[3];
                    cellIneligibleNoPRBillAmount.AddParagraph("청구액(원금)");
                    cellIneligibleNoPRBillAmount.Format.Font.Bold = true;
                    cellIneligibleNoPRBillAmount.Format.Font.Size = 7;
                    cellIneligibleNoPRBillAmount.Format.Font.Name = "Malgun Gothic";

                    //MigraDocDOM.Tables.Cell cellIneligibleNoPRType = ineligibleRow.Cells[4];
                    //cellIneligibleNoPRType.AddParagraph("Type");
                    //cellIneligibleNoPRType.Format.Font.Bold = true;
                    //cellIneligibleNoPRType.Format.Font.Size = 7;
                    //cellIneligibleNoPRType.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRIneligibleAmount = ineligibleRow.Cells[4];
                    cellIneligibleNoPRIneligibleAmount.AddParagraph("지원불가 의료비");
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Bold = true;
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Size = 7;
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRIneligibleReason = ineligibleRow.Cells[5];
                    cellIneligibleNoPRIneligibleReason.AddParagraph("지원불가 사유");
                    cellIneligibleNoPRIneligibleReason.Format.Font.Bold = true;
                    cellIneligibleNoPRIneligibleReason.Format.Font.Size = 7;
                    cellIneligibleNoPRIneligibleReason.Format.Font.Name = "Malgun Gothic";

                    for(int i = 0; i < lstMedBillNoPRNoSharing.Count; i++)
                    {
                        if (i < lstMedBillNoPRNoSharing.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligibleNoPR.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedBillName.Substring(8));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillDate.Value.ToString("MM/dd/yyyy"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedicalProvider);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillAmount.ToString("C"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = rowData.Cells[4];
                            //cell.AddParagraph(lstMedBillNoPRNoSharing[i].Type);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleReason);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                        else if (i == lstMedBillNoPRNoSharing.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligibleNoPR.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedBillName);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            if (lstMedBillNoPRNoSharing[i].BillDate != null) cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillDate.Value.ToString("MM/dd/yyyy"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedicalProvider);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillAmount.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = rowData.Cells[4];
                            //cell.AddParagraph(lstMedBillNoPRNoSharing[i].Type);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleReason);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                    }
                    pdfPersonalResponsibilityDoc.LastSection.Add(tableIneligibleNoPR);

                }



                const bool unicode = true;
                const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);
                pdfRenderer.Document = pdfPersonalResponsibilityDoc;
                pdfRenderer.RenderDocument();


                if (txtIncidentNo.Text.Trim() != String.Empty)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_Ko";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                        //finally
                        //{
                        //    ChkInfoEntered = null;
                        //}
                    }
                }

            }
            else
            {
                MessageBox.Show("No table is populated", "Error");
            }
        }

        private void btnGenerateEnPDF_Click(object sender, EventArgs e)
        {
            if ((gvBillPaid.RowCount > 0) || (gvCMMPendingPayment.RowCount > 0) || (gvPending.RowCount > 0) || (gvIneligible.RowCount > 0))
            {

                //DateTime? dtDocReceivedDate = null;

                //frmDocReceivedDate frmDocumentReceivedDate = new frmDocReceivedDate();

                //frmDocumentReceivedDate.StartPosition = FormStartPosition.CenterParent;
                
                //var dlgResultDocReceivedDate = frmDocumentReceivedDate.ShowDialog();

                //if (dlgResultDocReceivedDate == DialogResult.OK)
                //{
                //dtDocReceivedDate = frmDocumentReceivedDate.ReceivedDate;

                Document pdfDoc = new Document();

                Section section = pdfDoc.AddSection();
                pdfDoc.UseCmykColor = true;

                section.PageSetup.PageFormat = PageFormat.Letter;
                section.PageSetup.HeaderDistance = "0.25in";
                section.PageSetup.TopMargin = "1.5in";
                //section.PageSetup.LeftMargin = "0.3in";
                //section.PageSetup.RightMargin = "0.3in";
                section.PageSetup.LeftMargin = "0.8in";
                section.PageSetup.RightMargin = "0.8in";

                section.PageSetup.BottomMargin = "0.5in";

                section.PageSetup.DifferentFirstPageHeaderFooter = false;
                section.Headers.Primary.Format.SpaceBefore = "0.25in";

                MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\Program Files (x86)\\CMM\\BlueSheet\\cmmlogo.png");

                image.Height = "0.8in";
                image.LockAspectRatio = true;
                image.RelativeVertical = MigraDocDOM.Shapes.RelativeVertical.Line;
                image.RelativeHorizontal = MigraDocDOM.Shapes.RelativeHorizontal.Margin;
                image.Top = MigraDocDOM.Shapes.ShapePosition.Top;
                image.Left = MigraDocDOM.Shapes.ShapePosition.Center;
                image.WrapFormat.Style = MigraDocDOM.Shapes.WrapStyle.TopBottom;

                Paragraph paraCMMAddress = section.Headers.Primary.AddParagraph();
                paraCMMAddress.Format.Font.Name = "Arial";
                paraCMMAddress.Format.Font.Size = 8;
                paraCMMAddress.Format.SpaceBefore = "0.15in";
                paraCMMAddress.Format.SpaceAfter = "0.25in";
                //paraCMMAddress.Format.LeftIndent = "0.5in";
                //paraCMMAddress.Format.RightIndent = "0.5in";
                paraCMMAddress.Format.Alignment = ParagraphAlignment.Center;

                String strVerticalBar = " | ";
                String strStreet = "5235 N. Elston Ave.";
                String strCityStateZip = "Chicago, IL 60630";
                String strPhone = "Phone 773.777.8889";
                String strFax = "773.777.0004";
                String strWebsiteAddr = "www.cmmlogos.org";

                paraCMMAddress.AddFormattedText(strStreet, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strCityStateZip, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strPhone, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strFax, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strWebsiteAddr, TextFormat.NotBold);

                Paragraph paraToday = section.AddParagraph();
                paraToday.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraToday.Format.Font.Name = "Arial";
                paraToday.Format.Alignment = ParagraphAlignment.Left;
                paraToday.Format.SpaceBefore = "0.25in";
                paraToday.Format.SpaceAfter = "0.25in";
                //paraToday.Format.LeftIndent = "0.5in";
                //paraToday.Format.RightIndent = "0.5in";
                paraToday.AddFormattedText(DateTime.Today.ToString("MM/dd/yyyy"));

                Paragraph paraMembershipInfo = section.AddParagraph();

                paraMembershipInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraMembershipInfo.Format.Font.Name = "Arial";
                //paraMembershipInfo.Format.SpaceBefore = "0.7in";
                paraMembershipInfo.Format.SpaceBefore = "0.2in";
                paraMembershipInfo.Format.SpaceAfter = "0.2in";
                paraMembershipInfo.Format.LeftIndent = "0.75in";
                //paraMembershipInfo.Format.LeftIndent = "0.5in";
                //paraMembershipInfo.Format.RightIndent = "0.5in";
                paraMembershipInfo.Format.Alignment = ParagraphAlignment.Left;
                if (strMembershipId != String.Empty) paraMembershipInfo.AddFormattedText(strMembershipId + " (" + strIndividualID + ")\n");
                else paraMembershipInfo.AddFormattedText(strIndividualID + "\n");
                paraMembershipInfo.AddFormattedText(strIndividualName + "\n");
                paraMembershipInfo.AddFormattedText(strStreetAddress + "\n");
                paraMembershipInfo.AddFormattedText(strCity + ", " + strState + " " + strZip + "\n");


                Paragraph paraDearMember = section.AddParagraph();
                paraDearMember.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraDearMember.Format.Font.Name = "Arial";
                paraDearMember.Format.Font.Size = 9;
                paraDearMember.Format.Alignment = ParagraphAlignment.Left;
                //paraDearMember.Format.LeftIndent = "0.5in";
                //paraDearMember.Format.RightIndent = "0.5in";
                paraDearMember.Format.SpaceBefore = "0.1in";
                paraDearMember.Format.SpaceAfter = "0.1in";
                //paraDearMember.AddFormattedText(strPrimaryName + "회원(가족)께,");
                paraDearMember.AddFormattedText(strDearMember + " " + strIndividualName.Trim() + ", ");


                Paragraph paraGreetingMessage = section.AddParagraph();

                paraGreetingMessage.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage.Format.Font.Name = "Arial";
                paraGreetingMessage.Format.Font.Size = 9;
                //paraGreetingMessage.Format.LeftIndent = "0.5in";
                //paraGreetingMessage.Format.RightIndent = "0.5in";
                paraGreetingMessage.Format.Alignment = ParagraphAlignment.Justify;
                //paraGreetingMessage.AddFormattedText(strGreetingMessage, TextFormat.NotBold);
                //paraGreetingMessage.AddFormattedText(strGreetingMessagePara1, TextFormat.NotBold);
                paraGreetingMessage.AddFormattedText(strEnglishGreetingMessage1, TextFormat.NotBold);

                Paragraph paraGreetingMessage2 = section.AddParagraph();
                paraGreetingMessage2.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage2.Format.Font.Name = "Arial";
                paraGreetingMessage2.Format.Font.Size = 9;
                //paraGreetingMessage2.Format.LeftIndent = "0.5in";
                //paraGreetingMessage2.Format.RightIndent = "0.5in";
                paraGreetingMessage2.Format.Alignment = ParagraphAlignment.Left;
                //paraGreetingMessage2.AddFormattedText(strGreetingMessagePara2, TextFormat.NotBold);
                paraGreetingMessage2.AddFormattedText(strEnglishGreetingMessage2, TextFormat.NotBold);

                Paragraph paraGreetingMessage3 = section.AddParagraph();
                paraGreetingMessage3.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage3.Format.Font.Name = "Arial";
                paraGreetingMessage3.Format.Font.Size = 9;
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraGreetingMessage3.Format.Alignment = ParagraphAlignment.Justify;
                //paraGreetingMessage3.AddFormattedText(strGreetingMessagePara3, TextFormat.NotBold);
                paraGreetingMessage3.AddFormattedText(strEnglishGreetingMessage3, TextFormat.NotBold);

                Paragraph paraGreetingMessage4 = section.AddParagraph();
                paraGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraGreetingMessage4.Format.Font.Name = "Arial";
                paraGreetingMessage4.Format.Font.Size = 9;
                //paraGreetingMessage4.Format.LeftIndent = "0.5in";
                //paraGreetingMessage4.Format.RightIndent = "0.5in";
                paraGreetingMessage4.Format.Alignment = ParagraphAlignment.Justify;
                //paraGreetingMessage4.AddFormattedText(strGreetingMessagePara4, TextFormat.NotBold);
                paraGreetingMessage4.AddFormattedText(strEnglishGreetingMessage4, TextFormat.NotBold);



                Paragraph paraNeedsProcessing = section.AddParagraph();

                paraNeedsProcessing.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraNeedsProcessing.Format.Font.Name = "Arial";
                paraNeedsProcessing.Format.Font.Size = 9;
                paraNeedsProcessing.Format.Font.Bold = true;
                paraNeedsProcessing.Format.Alignment = ParagraphAlignment.Left;
                paraNeedsProcessing.Format.SpaceBefore = "0.1in";
                //paraNeedsProcessing.Format.LeftIndent = "0.5in";
                //paraNeedsProcessing.Format.RightIndent = "0.5in";
                paraNeedsProcessing.AddFormattedText(strCMM_NeedProcessing + "\n");

                Paragraph paraPhoneFaxEmail = section.AddParagraph();
                paraPhoneFaxEmail.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPhoneFaxEmail.Format.Font.Size = 9;
                paraPhoneFaxEmail.Format.Alignment = ParagraphAlignment.Left;
                //paraPhoneFaxEmail.Format.LeftIndent = "0.5in";
                //paraPhoneFaxEmail.Format.RightIndent = "0.5in";
                paraPhoneFaxEmail.AddFormattedText(strNP_Phone_Fax_Email + "\n");

                Paragraph paraHorizontalLine = section.AddParagraph();

                paraHorizontalLine.Format.SpaceBefore = "0.05in";
                paraHorizontalLine.Format.SpaceAfter = "0.05in";
                paraHorizontalLine.Format.Borders.Top.Width = 0;
                paraHorizontalLine.Format.Borders.Left.Width = 0;
                paraHorizontalLine.Format.Borders.Right.Width = 0;
                paraHorizontalLine.Format.Borders.Bottom.Width = 1;
                paraHorizontalLine.Format.Borders.Bottom.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraHorizontalLine.Format.Borders.Style = MigraDocDOM.BorderStyle.DashDot;

                Paragraph paraNPStatement = section.AddParagraph();
                paraNPStatement.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                paraNPStatement.Format.Font.Name = "Arial";
                paraNPStatement.Format.Font.Size = 9;
                paraNPStatement.Format.Font.Bold = true;
                paraNPStatement.Format.Alignment = ParagraphAlignment.Center;
                paraNPStatement.Format.SpaceAfter = "0.1in";

                paraNPStatement.AddFormattedText("Needs Processing Statement\n", TextFormat.Bold);

                Paragraph paraCheckInfo = section.AddParagraph();
                paraCheckInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraCheckInfo.Format.Font.Name = "Arial";
                paraCheckInfo.Format.Font.Size = 9;
                paraCheckInfo.Format.Font.Bold = true;
                paraCheckInfo.Format.Alignment = ParagraphAlignment.Left;
                paraCheckInfo.Format.SpaceBefore = "0.1in";
                //paraCheckInfo.Format.SpaceAfter = "0.1in";

                //if (ChkInfoEntered != null)
                if (rbCheck.Checked)
                {
                    paraCheckInfo.AddFormattedText("Issue Date: " + ChkInfoEntered.dtCheckIssueDate.ToString("MM/dd/yyyy") +
                                                    "\tCheck No: " + ChkInfoEntered.CheckNumber +
                                                    "\tCheck Amount: " + ChkInfoEntered.CheckAmount.Value.ToString("C") +
                                                    "\tPaid To: " + ChkInfoEntered.PaidTo);

                    //ChkInfoEntered = null;
                }
                else if (rbACH.Checked)
                {
                    paraCheckInfo.AddFormattedText("Issue Date: " + ACHInfoEntered.dtACHDate.ToString("MM/dd/yyyy") +
                                                    "\tACH No: " + ACHInfoEntered.ACHNumber +
                                                    "\tACH Amount: " + ACHInfoEntered.ACHAmount.Value.ToString("C") +
                                                    "\tPaid To: " + ACHInfoEntered.PaidTo);
                }
                else if (rbCreditCard.Checked)
                {
                    paraCheckInfo.AddFormattedText("Date:" + CreditCardPaymentEntered.dtPaymentDate.ToString("MM/dd/yyyy") +
                                                    "\tCredit Card Payment Amount: " + CreditCardPaymentEntered.CCPaymentAmount.Value.ToString("C") +
                                                    "\tPaid To: " + CreditCardPaymentEntered.PaidTo);
                }

                //int nRowHeight = 338;
                //int nRowHeight = 302;
                int nRowHeight = 296;

                if (lstIncidents.Count > 0)
                {
                    Paragraph paraIncd = section.AddParagraph();

                    paraIncd.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                    paraIncd.Format.Font.Name = "Arial";
                    paraIncd.Format.Font.Size = 8;
                    paraIncd.Format.Font.Bold = true;

                    MigraDocDOM.Tables.Table tableIncd = new MigraDocDOM.Tables.Table();
                    tableIncd.Borders.Width = 0;
                    tableIncd.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Column colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(0.85));
                    colINCD.Format.Alignment = ParagraphAlignment.Left;
                    //colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(1));
                    //colINCD.Format.Alignment = ParagraphAlignment.Left;
                    colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(3.5));
                    colINCD.Format.Alignment = ParagraphAlignment.Left;

                    foreach (Incident incd in lstIncidents)
                    {
                        nRowHeight += 18;

                        MigraDocDOM.Tables.Row IncdRow = tableIncd.AddRow();
                        IncdRow.Height = "0.15in";
                        IncdRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        MigraDocDOM.Tables.Cell cellIncdName = IncdRow.Cells[0];
                        cellIncdName.Format.Font.Bold = true;
                        cellIncdName.Format.Font.Size = 8;
                        cellIncdName.Format.Font.Name = "Arial";
                        cellIncdName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cellIncdName.AddParagraph(incd.Name + ": ");

                        //MigraDocDOM.Tables.Cell cellPatientName = IncdRow.Cells[1];
                        //cellPatientName.Format.Font.Bold = true;
                        //cellPatientName.Format.Font.Size = 8;
                        //cellPatientName.Format.Font.Name = "Arial";
                        //cellPatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        //if (incd.PatientName.Length > 11)
                        //{
                        //    cellPatientName.AddParagraph(incd.PatientName.Substring(0, 11) + " ...");
                        //}
                        //else cellPatientName.AddParagraph(incd.PatientName);

                        MigraDocDOM.Tables.Cell cellICD10Code = IncdRow.Cells[1];
                        cellICD10Code.Format.Font.Bold = true;
                        cellICD10Code.Format.Font.Size = 8;
                        cellICD10Code.Format.Font.Name = "Arial";
                        cellICD10Code.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cellICD10Code.AddParagraph(incd.ICD10_Code);
                    }

                    pdfDoc.LastSection.Add(tableIncd);
                }

                //section.AddParagraph();
                //lstPaidMedicalExpenseTableRow.Clear();

                if (gvBillPaid.RowCount > 0)
                {

                    section.AddParagraph();
                    lstPaidMedicalExpenseTableRow.Clear();

                    //nRowHeight += 30;
                    nRowHeight += 22;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    //paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceBefore = "0.08in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                    paraSpaceBefore.Format.Font.Name = "Arial";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("Processed Medical Bill(s)", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvBillPaid.RowCount; nRow++)
                    {
                        PaidMedicalExpenseTableRow expenseRow = new PaidMedicalExpenseTableRow();

                        //if (nRow < gvBillPaid.RowCount - 1)
                        //{
                        //    expenseRow.PatientName = gvBillPaid[1, nRow].Value.ToString();
                        //    expenseRow.MED_BILL = gvBillPaid[2, nRow].Value.ToString();
                        //    //expenseRow.Bill_Date = gvBillPaid[3, nRow].Value.ToString();
                        //    expenseRow.Bill_Date = DateTime.Parse(gvBillPaid[3, nRow].Value.ToString());
                        //    expenseRow.Medical_Provider = gvBillPaid[4, nRow].Value.ToString();
                        //    expenseRow.Bill_Amount = gvBillPaid[5, nRow].Value.ToString();
                        //    expenseRow.Personal_Responsibility = gvBillPaid[6, nRow].Value.ToString();
                        //    expenseRow.Member_Discount = gvBillPaid[7, nRow].Value.ToString();
                        //    expenseRow.CMM_Discount = gvBillPaid[8, nRow].Value.ToString();
                        //    expenseRow.CMM_Provider_Payment = gvBillPaid[9, nRow].Value.ToString();
                        //    expenseRow.PastReimbursement = gvBillPaid[10, nRow].Value.ToString();
                        //    expenseRow.Reimbursement = gvBillPaid[11, nRow].Value.ToString();
                        //    expenseRow.Balance = gvBillPaid[12, nRow].Value.ToString();

                        //}
                        //if (nRow == gvBillPaid.RowCount - 1)
                        //{
                        expenseRow.PatientName = gvBillPaid[1, nRow].Value.ToString();
                        expenseRow.MED_BILL = gvBillPaid[2, nRow].Value.ToString();
                        if (gvBillPaid[3, nRow].Value.ToString() != String.Empty) expenseRow.Bill_Date = DateTime.Parse(gvBillPaid[3, nRow].Value.ToString());
                        expenseRow.Medical_Provider = gvBillPaid[4, nRow].Value.ToString();
                        expenseRow.Bill_Amount = gvBillPaid[5, nRow].Value.ToString();
                        expenseRow.Personal_Responsibility = gvBillPaid[6, nRow].Value.ToString();
                        expenseRow.Member_Discount = gvBillPaid[7, nRow].Value.ToString();
                        expenseRow.CMM_Discount = gvBillPaid[8, nRow].Value.ToString();
                        expenseRow.CMM_Provider_Payment = gvBillPaid[9, nRow].Value.ToString();
                        if (PaidTo == EnumPaidTo.Member)
                        {
                            expenseRow.PastReimbursement = gvBillPaid[10, nRow].Value.ToString();
                            expenseRow.Reimbursement = gvBillPaid[11, nRow].Value.ToString();
                        }
                        if (PaidTo == EnumPaidTo.MedicalProvider)
                        {
                            expenseRow.PastCMM_Provider_Payment = gvBillPaid[10, nRow].Value.ToString();
                            expenseRow.Reimbursement = gvBillPaid[11, nRow].Value.ToString();
                        }
                        expenseRow.Balance = gvBillPaid[12, nRow].Value.ToString();
                        //}
                        lstPaidMedicalExpenseTableRow.Add(expenseRow);
                    }



                    MigraDocDOM.Tables.Table table = new MigraDocDOM.Tables.Table();
                    table.Borders.Width = 0.1;
                    table.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    //col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(1.0));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    col = table.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    col.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;


                    MigraDocDOM.Tables.Row row = table.AddRow();

                    nRowHeight += 22;
                    row.Height = "0.3in";
                    row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellTitlePatientName = row.Cells[0];
                    //cellTitlePatientName.AddParagraph("Member Name");
                    //cellTitlePatientName.Format.Font.Bold = true;
                    //cellTitlePatientName.Format.Font.Size = 7;
                    //cellTitlePatientName.Format.Font.Name = "Arial";
                    //cellTitlePatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleMedBill = row.Cells[0];
                    cellTitleMedBill.AddParagraph("MEDBILL");
                    cellTitleMedBill.Format.Font.Bold = true;
                    cellTitleMedBill.Format.Font.Size = 7;
                    cellTitleMedBill.Format.Font.Name = "Arial";
                    cellTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleBillDate = row.Cells[1];
                    cellTitleBillDate.AddParagraph("Date of Service");
                    cellTitleBillDate.Format.Font.Bold = true;
                    cellTitleBillDate.Format.Font.Size = 7;
                    cellTitleBillDate.Format.Font.Name = "Arial";
                    cellTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleMedicalProvider = row.Cells[2];
                    cellTitleMedicalProvider.AddParagraph("Medical Provider");
                    cellTitleMedicalProvider.Format.Font.Bold = true;
                    cellTitleMedicalProvider.Format.Font.Size = 7;
                    cellTitleMedicalProvider.Format.Font.Name = "Arial";
                    cellTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleBillAmount = row.Cells[3];
                    cellTitleBillAmount.AddParagraph("Original Amount");
                    cellTitleBillAmount.Format.Font.Bold = true;
                    cellTitleBillAmount.Format.Font.Size = 7;
                    cellTitleBillAmount.Format.Font.Name = "Arial";
                    cellTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitlePersonalResponsibility = row.Cells[4];
                    cellTitlePersonalResponsibility.AddParagraph("Personal Responsibility");
                    cellTitlePersonalResponsibility.Format.Font.Bold = true;
                    cellTitlePersonalResponsibility.Format.Font.Size = 7;
                    cellTitlePersonalResponsibility.Format.Font.Name = "Arial";
                    cellTitlePersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);


                    MigraDocDOM.Tables.Cell cellTitleMemberDiscount = row.Cells[5];
                    cellTitleMemberDiscount.AddParagraph("Member Discount");
                    cellTitleMemberDiscount.Format.Font.Bold = true;
                    cellTitleMemberDiscount.Format.Font.Size = 7;
                    cellTitleMemberDiscount.Format.Font.Name = "Arial";
                    cellTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleCMMDiscount = row.Cells[6];
                    cellTitleCMMDiscount.AddParagraph("CMM Discount");
                    cellTitleCMMDiscount.Format.Font.Bold = true;
                    cellTitleCMMDiscount.Format.Font.Size = 7;
                    cellTitleCMMDiscount.Format.Font.Name = "Arial";
                    cellTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellTitleCMMProviderDiscount = row.Cells[7];
                    //cellTitleCMMProviderDiscount.AddParagraph("CMM Provider Payment");
                    //AddParagraph("CMM Provider Payment");
                    cellTitleCMMProviderDiscount.AddParagraph("Paid to Provider");
                    cellTitleCMMProviderDiscount.Format.Font.Bold = true;
                    cellTitleCMMProviderDiscount.Format.Font.Size = 7;
                    cellTitleCMMProviderDiscount.Format.Font.Name = "Arial";
                    cellTitleCMMProviderDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    if (PaidTo == EnumPaidTo.Member)
                    {
                        MigraDocDOM.Tables.Cell cellTitleSharedAmount = row.Cells[8];
                        cellTitleSharedAmount.AddParagraph("Shared Amount");
                        cellTitleSharedAmount.Format.Font.Bold = true;
                        cellTitleSharedAmount.Format.Font.Size = 7;
                        cellTitleSharedAmount.Format.Font.Name = "Arial";
                        cellTitleSharedAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        MigraDocDOM.Tables.Cell cellTitleReimbursement = row.Cells[9];
                        cellTitleReimbursement.AddParagraph("Reimbursement Amount");
                        cellTitleReimbursement.Format.Font.Bold = true;
                        cellTitleReimbursement.Format.Font.Size = 7;
                        cellTitleReimbursement.Format.Font.Name = "Arial";
                        cellTitleReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                    }

                    if (PaidTo == EnumPaidTo.MedicalProvider)
                    {
                        MigraDocDOM.Tables.Cell cellTitlePastReimbursement = row.Cells[8];
                        cellTitlePastReimbursement.AddParagraph("Shared To Provider");
                        cellTitlePastReimbursement.Format.Font.Bold = true;
                        cellTitlePastReimbursement.Format.Font.Size = 7;
                        cellTitlePastReimbursement.Format.Font.Name = "Arial";
                        cellTitlePastReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        MigraDocDOM.Tables.Cell cellTitlePastCMMProviderPayment = row.Cells[9];
                        cellTitlePastCMMProviderPayment.AddParagraph("Shared To Member");
                        cellTitlePastCMMProviderPayment.Format.Font.Bold = true;
                        cellTitlePastCMMProviderPayment.Format.Font.Size = 7;
                        cellTitlePastCMMProviderPayment.Format.Font.Name = "Arial";
                        cellTitlePastCMMProviderPayment.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                        //MigraDocDOM.Tables.Cell cellTitleReimbursement = row.Cells[10];
                        //cellTitleReimbursement.AddParagraph("회원 환불금");
                        //cellTitleReimbursement.Format.Font.Bold = true;
                        //cellTitleReimbursement.Format.Font.Size = 7;
                        //cellTitleReimbursement.Format.Font.Name = "Malgun Gothic";
                        //cellTitleReimbursement.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100); ;


                    }



                    MigraDocDOM.Tables.Cell cellTitleBalance = row.Cells[10];
                    cellTitleBalance.AddParagraph("Balance");
                    cellTitleBalance.Format.Font.Bold = true;
                    cellTitleBalance.Format.Font.Size = 7;
                    cellTitleBalance.Format.Font.Name = "Arial";
                    cellTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);


                    for (int i = 0; i < lstPaidMedicalExpenseTableRow.Count; i++)
                    {
                        if (nRowHeight > 645) nRowHeight = 0;
                        nRowHeight += 18;
                        MigraDocDOM.Tables.Row rowData = table.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        if (i < lstPaidMedicalExpenseTableRow.Count - 1)
                        {
                            rowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstPaidMedicalExpenseTableRow[i].PatientName.Length > 9)
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName.Substring(0, 9) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Date.Value.ToString("MM/dd/yy"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            if (lstPaidMedicalExpenseTableRow[i].Medical_Provider.Length > 14)
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider.Substring(0, 14) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;


                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Personal_Responsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Provider_Payment);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            //cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastCMM_Provider_Payment);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[9];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[10];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                        }
                        if (i == lstPaidMedicalExpenseTableRow.Count - 1)
                        {
                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstPaidMedicalExpenseTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[1];
                            if (lstPaidMedicalExpenseTableRow[i].Bill_Date != null) cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Date.Value.ToString("MM/dd/yy"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            //if (lstPaidMedicalExpenseTableRow[i].Medical_Provider.Length > 25)
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Medical_Provider);
                            //}
                            cell.AddParagraph("Total");

                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Center;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Personal_Responsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].CMM_Provider_Payment);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            //cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastReimbursement);
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].PastCMM_Provider_Payment);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[9];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Reimbursement);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                            
                            cell = rowData.Cells[10];
                            cell.AddParagraph(lstPaidMedicalExpenseTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                    }

                    pdfDoc.LastSection.Add(table);
                }

                int nHeightAfterCMMPendingPayment = 0;

                if (gvCMMPendingPayment.RowCount > 0)
                {
                    nHeightAfterCMMPendingPayment += 22;
                    for (int nRow = 0; nRow < gvCMMPendingPayment.RowCount; nRow++)
                    {
                        nHeightAfterCMMPendingPayment += 15;
                    }

                    if ((nRowHeight > 645) ||
                        (nRowHeight + nHeightAfterCMMPendingPayment) > 645)
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }



                //////////////////////////////////////////////////////////////////////////////////////////////////

                // The beginning of CMM Pending Payment table


                if (gvCMMPendingPayment.RowCount > 0)
                {
                    lstCMMPendingPaymentTableRow.Clear();
                    nRowHeight += 30;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                    paraSpaceBefore.Format.Font.Name = "Arial";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("Pending Payment(s)", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvCMMPendingPayment.RowCount; nRow++)
                    {
                        CMMPendingPaymentTableRow cmmPendingRow = new CMMPendingPaymentTableRow();

                        cmmPendingRow.PatientName = gvCMMPendingPayment[1, nRow].Value.ToString();
                        cmmPendingRow.MED_BILL = gvCMMPendingPayment[2, nRow].Value.ToString();
                        cmmPendingRow.Bill_Date = gvCMMPendingPayment[3, nRow].Value.ToString();
                        //cmmPendingRow.Due_Date = gvCMMPendingPayment[4, nRow].Value.ToString();
                        cmmPendingRow.Medical_Provider = gvCMMPendingPayment[4, nRow].Value.ToString();
                        cmmPendingRow.Bill_Amount = gvCMMPendingPayment[5, nRow].Value.ToString();
                        cmmPendingRow.Member_Discount = gvCMMPendingPayment[6, nRow].Value.ToString();
                        cmmPendingRow.CMM_Discount = gvCMMPendingPayment[7, nRow].Value.ToString();
                        cmmPendingRow.PersonalResponsibility = gvCMMPendingPayment[8, nRow].Value.ToString();
                        cmmPendingRow.Shared_Amount = gvCMMPendingPayment[9, nRow].Value.ToString();
                        cmmPendingRow.Balance = gvCMMPendingPayment[10, nRow].Value.ToString();

                        lstCMMPendingPaymentTableRow.Add(cmmPendingRow);
                    }

                    MigraDocDOM.Tables.Table tableCMMPendingPayment = new MigraDocDOM.Tables.Table();

                    tableCMMPendingPayment.Borders.Width = 0.1;
                    tableCMMPendingPayment.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    //colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colCMMPendingPayment = tableCMMPendingPayment.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colCMMPendingPayment.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row cmm_pending_row = tableCMMPendingPayment.AddRow();

                    cmm_pending_row.Height = "0.3in";
                    cmm_pending_row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    nRowHeight += 22;

                    //MigraDocDOM.Tables.Cell cellCMMPendingTitlePatientName = cmm_pending_row.Cells[0];
                    //cellCMMPendingTitlePatientName.AddParagraph("Member Name");
                    //cellCMMPendingTitlePatientName.Format.Font.Bold = true;
                    //cellCMMPendingTitlePatientName.Format.Font.Size = 7;
                    //cellCMMPendingTitlePatientName.Format.Font.Name = "Arial";
                    //cellCMMPendingTitlePatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMedBill = cmm_pending_row.Cells[0];
                    cellCMMPendingTitleMedBill.AddParagraph("MEDBILL");
                    cellCMMPendingTitleMedBill.Format.Font.Bold = true;
                    cellCMMPendingTitleMedBill.Format.Font.Size = 7;
                    cellCMMPendingTitleMedBill.Format.Font.Name = "Arial";
                    cellCMMPendingTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBillDate = cmm_pending_row.Cells[1];
                    cellCMMPendingTitleBillDate.AddParagraph("Date of Service");
                    cellCMMPendingTitleBillDate.Format.Font.Bold = true;
                    cellCMMPendingTitleBillDate.Format.Font.Size = 7;
                    cellCMMPendingTitleBillDate.Format.Font.Name = "Arial";
                    cellCMMPendingTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellCMMPendingTitleDueDate = cmm_pending_row.Cells[3];
                    //cellCMMPendingTitleDueDate.AddParagraph("Received Date");
                    //cellCMMPendingTitleDueDate.Format.Font.Bold = true;
                    //cellCMMPendingTitleDueDate.Format.Font.Size = 7;
                    //cellCMMPendingTitleDueDate.Format.Font.Name = "Arial";
                    //cellCMMPendingTitleDueDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMedicalProvider = cmm_pending_row.Cells[2];
                    cellCMMPendingTitleMedicalProvider.AddParagraph("Medical Provider");
                    cellCMMPendingTitleMedicalProvider.Format.Font.Bold = true;
                    cellCMMPendingTitleMedicalProvider.Format.Font.Size = 7;
                    cellCMMPendingTitleMedicalProvider.Format.Font.Name = "Arial";
                    cellCMMPendingTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBillAmount = cmm_pending_row.Cells[3];
                    cellCMMPendingTitleBillAmount.AddParagraph("Original Amount");
                    cellCMMPendingTitleBillAmount.Format.Font.Bold = true;
                    cellCMMPendingTitleBillAmount.Format.Font.Size = 7;
                    cellCMMPendingTitleBillAmount.Format.Font.Name = "Arial";
                    cellCMMPendingTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleMemberDiscount = cmm_pending_row.Cells[4];
                    cellCMMPendingTitleMemberDiscount.AddParagraph("Member Discount");
                    cellCMMPendingTitleMemberDiscount.Format.Font.Bold = true;
                    cellCMMPendingTitleMemberDiscount.Format.Font.Size = 7;
                    cellCMMPendingTitleMemberDiscount.Format.Font.Name = "Arial";
                    cellCMMPendingTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleCMMDiscount = cmm_pending_row.Cells[5];
                    cellCMMPendingTitleCMMDiscount.AddParagraph("CMM Discount");
                    cellCMMPendingTitleCMMDiscount.Format.Font.Bold = true;
                    cellCMMPendingTitleCMMDiscount.Format.Font.Size = 7;
                    cellCMMPendingTitleCMMDiscount.Format.Font.Name = "Arial";
                    cellCMMPendingTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitlePersonalResponsibility = cmm_pending_row.Cells[6];
                    cellCMMPendingTitlePersonalResponsibility.AddParagraph("Personal Responsibility");
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Bold = true;
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Size = 7;
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Name = "Arial";
                    cellCMMPendingTitlePersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleSharedAmount = cmm_pending_row.Cells[7];
                    cellCMMPendingTitleSharedAmount.AddParagraph("Total Shared Amount");
                    cellCMMPendingTitleSharedAmount.Format.Font.Bold = true;
                    cellCMMPendingTitleSharedAmount.Format.Font.Size = 7;
                    cellCMMPendingTitleSharedAmount.Format.Font.Name = "Arial";
                    cellCMMPendingTitleSharedAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellCMMPendingTitleBalance = cmm_pending_row.Cells[8];
                    cellCMMPendingTitleBalance.AddParagraph("Balance");
                    cellCMMPendingTitleBalance.Format.Font.Bold = true;
                    cellCMMPendingTitleBalance.Format.Font.Size = 7;
                    cellCMMPendingTitleBalance.Format.Font.Name = "Arial";
                    cellCMMPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    for (int i = 0; i < lstCMMPendingPaymentTableRow.Count; i++)
                    {
                        nRowHeight += 18;
                        MigraDocDOM.Tables.Row rowData = tableCMMPendingPayment.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                        if (i < lstCMMPendingPaymentTableRow.Count - 1)
                        {
                            rowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstCMMPendingPaymentTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            //cell = rowData.Cells[3];
                            //cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Due_Date);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            if (lstCMMPendingPaymentTableRow[i].Medical_Provider.Length > 25)
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PersonalResponsibility);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Shared_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                        if (i == lstCMMPendingPaymentTableRow.Count - 1)
                        {
                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstCMMPendingPaymentTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            //cell = rowData.Cells[3];
                            //cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Due_Date);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            //if (lstCMMPendingPaymentTableRow[i].Medical_Provider.Length > 25)
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Medical_Provider);
                            //}
                            cell.AddParagraph("Total");
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Member_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].CMM_Discount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[6];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].PersonalResponsibility);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[7];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Shared_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[8];
                            cell.AddParagraph(lstCMMPendingPaymentTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(0, 100, 100, 0);
                            cell.Format.Alignment = ParagraphAlignment.Right;
                        }
                    }

                    pdfDoc.LastSection.Add(tableCMMPendingPayment);
                }

                int nHeightAfterPending = 0;

                if (gvPending.RowCount > 0)
                {
                    nHeightAfterPending += 22;
                    for (int nRow = 0; nRow < gvPending.RowCount; nRow++)
                    {
                        nHeightAfterPending += 15;
                    }

                    if ((nRowHeight > 645) || ((nRowHeight + nHeightAfterPending) > 645))
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }





                // Pending table

                if (gvPending.RowCount > 0)
                {
                    lstPendingTableRow.Clear();

                    nRowHeight += 30;

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                    paraSpaceBefore.Format.Font.Name = "Arial";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("Pending Medical Bill(s)", TextFormat.Bold);

                    for (int nRow = 0; nRow < gvPending.RowCount; nRow++)
                    {
                        //nRowHeight += 15;

                        PendingTableRow pendingRow = new PendingTableRow();
                        pendingRow.PatientName = gvPending[1, nRow].Value.ToString();
                        pendingRow.MED_BILL = gvPending[2, nRow].Value.ToString();
                        pendingRow.Bill_Date = gvPending[3, nRow].Value.ToString();
                        pendingRow.Due_Date = gvPending[4, nRow].Value.ToString();
                        pendingRow.Medical_Provider = gvPending[5, nRow].Value.ToString();
                        pendingRow.Bill_Amount = gvPending[6, nRow].Value.ToString();
                        pendingRow.Balance = gvPending[7, nRow].Value.ToString();
                        //pendingRow.Member_Discount = gvPending[7, nRow].Value.ToString();
                        //pendingRow.CMM_Discount = gvPending[8, nRow].Value.ToString();
                        //pendingRow.Shared_Amount = gvPending[9, nRow].Value.ToString();
                        //pendingRow.Balance = gvPending[10, nRow].Value.ToString();
                        pendingRow.Pending_Reason = gvPending[8, nRow].Value.ToString();

                        lstPendingTableRow.Add(pendingRow);
                    }

                    MigraDocDOM.Tables.Table tablePending = new MigraDocDOM.Tables.Table();
                    tablePending.Borders.Width = 0.1;
                    tablePending.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(1.2));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.5));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    //colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    //colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colPending = tablePending.AddColumn(MigraDocDOM.Unit.FromInch(2.8));
                    colPending.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row pending_Row = tablePending.AddRow();

                    nRowHeight += 22;

                    pending_Row.Height = "0.3in";
                    pending_Row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellPendingTitleINCD = pending_Row.Cells[0];
                    //cellPendingTitleINCD.AddParagraph("Member Name");
                    //cellPendingTitleINCD.Format.Font.Bold = true;
                    //cellPendingTitleINCD.Format.Font.Size = 7;
                    //cellPendingTitleINCD.Format.Font.Name = "Arial";
                    //cellPendingTitleINCD.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleMED_BILL = pending_Row.Cells[0];
                    cellPendingTitleMED_BILL.AddParagraph("MEDBILL");
                    cellPendingTitleMED_BILL.Format.Font.Bold = true;
                    cellPendingTitleMED_BILL.Format.Font.Size = 7;
                    cellPendingTitleMED_BILL.Format.Font.Name = "Arial";
                    cellPendingTitleMED_BILL.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBill_Date = pending_Row.Cells[1];
                    cellPendingTitleBill_Date.AddParagraph("Date of Service");
                    cellPendingTitleBill_Date.Format.Font.Bold = true;
                    cellPendingTitleBill_Date.Format.Font.Size = 7;
                    cellPendingTitleBill_Date.Format.Font.Name = "Arial";
                    cellPendingTitleBill_Date.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleDue_Date = pending_Row.Cells[2];
                    cellPendingTitleDue_Date.AddParagraph("Received Date");
                    cellPendingTitleDue_Date.Format.Font.Bold = true;
                    cellPendingTitleDue_Date.Format.Font.Size = 7;
                    cellPendingTitleDue_Date.Format.Font.Name = "Arial";
                    cellPendingTitleDue_Date.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleMedicalProvider = pending_Row.Cells[3];
                    cellPendingTitleMedicalProvider.AddParagraph("Medical Provider");
                    cellPendingTitleMedicalProvider.Format.Font.Bold = true;
                    cellPendingTitleMedicalProvider.Format.Font.Size = 7;
                    cellPendingTitleMedicalProvider.Format.Font.Name = "Arial";
                    cellPendingTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBillAmount = pending_Row.Cells[4];
                    cellPendingTitleBillAmount.AddParagraph("Original Amount");
                    cellPendingTitleBillAmount.Format.Font.Bold = true;
                    cellPendingTitleBillAmount.Format.Font.Size = 7;
                    cellPendingTitleBillAmount.Format.Font.Name = "Arial";
                    cellPendingTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleMemberDiscount = pending_Row.Cells[6];
                    //cellPendingTitleMemberDiscount.AddParagraph("회원 (할인)");
                    //cellPendingTitleMemberDiscount.Format.Font.Bold = true;
                    //cellPendingTitleMemberDiscount.Format.Font.Size = 7;
                    //cellPendingTitleMemberDiscount.Format.Font.Name = "Arial";
                    //cellPendingTitleMemberDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleCMMDiscount = pending_Row.Cells[7];
                    //cellPendingTitleCMMDiscount.AddParagraph("CMM (할인)");
                    //cellPendingTitleCMMDiscount.Format.Font.Bold = true;
                    //cellPendingTitleCMMDiscount.Format.Font.Size = 7;
                    //cellPendingTitleCMMDiscount.Format.Font.Name = "Arial";
                    //cellPendingTitleCMMDiscount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleSharedAmount = pending_Row.Cells[8];
                    //cellPendingTitleSharedAmount.AddParagraph("정산 완료");
                    //cellPendingTitleSharedAmount.Format.Font.Bold = true;
                    //cellPendingTitleSharedAmount.Format.Font.Size = 7;
                    //cellPendingTitleSharedAmount.Format.Font.Name = "Arial";
                    //cellPendingTitleSharedAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Cell cellPendingTitleBalance = pending_Row.Cells[9];
                    //cellPendingTitleBalance.AddParagraph("보류");
                    //cellPendingTitleBalance.Format.Font.Bold = true;
                    //cellPendingTitleBalance.Format.Font.Size = 7;
                    //cellPendingTitleBalance.Format.Font.Name = "Arial";
                    //cellPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitleBalance = pending_Row.Cells[5];
                    cellPendingTitleBalance.AddParagraph("Balance");
                    cellPendingTitleBalance.Format.Font.Bold = true;
                    cellPendingTitleBalance.Format.Font.Size = 7;
                    cellPendingTitleBalance.Format.Font.Name = "Arial";
                    cellPendingTitleBalance.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellPendingTitlePendingReason = pending_Row.Cells[6];
                    cellPendingTitlePendingReason.AddParagraph("Pending Reason");
                    cellPendingTitlePendingReason.Format.Font.Bold = true;
                    cellPendingTitlePendingReason.Format.Font.Size = 7;
                    cellPendingTitlePendingReason.Format.Font.Name = "Arial";
                    cellPendingTitlePendingReason.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    for (int i = 0; i < lstPendingTableRow.Count; i++)
                    {
                        MigraDocDOM.Tables.Row pendingRowData = tablePending.AddRow();
                        pendingRowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                        nRowHeight += 15;

                        if (i < lstPendingTableRow.Count - 1)
                        {
                            pendingRowData.Height = "0.18in";

                            //MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            //if (lstPendingTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName.Substring(0, 11));
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            cell.AddParagraph(lstPendingTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[1];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[2];
                            cell.AddParagraph(lstPendingTableRow[i].Due_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = pendingRowData.Cells[3];

                            if (lstPendingTableRow[i].Medical_Provider.Length > 20)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider.Substring(0, 20) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Medical_Provider);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = pendingRowData.Cells[4];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[5];
                            cell.AddParagraph(lstPendingTableRow[i].Balance);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[6];
                            //cell.AddParagraph(lstPendingTableRow[i].Member_Discount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[7];
                            //cell.AddParagraph(lstPendingTableRow[i].CMM_Discount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[8];
                            //cell.AddParagraph(lstPendingTableRow[i].Shared_Amount);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[9];
                            //cell.AddParagraph(lstPendingTableRow[i].Balance);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[6];
                            if (lstPendingTableRow[i].Pending_Reason.Length > 40)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason.Substring(0, 40) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                        }
                        if (i == lstPendingTableRow.Count - 1)
                        {
                            pendingRowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            //if (lstPendingTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName.Substring(0, 11));
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = pendingRowData.Cells[0];
                            cell.AddParagraph(lstPendingTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[1];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[2];
                            cell.AddParagraph(lstPendingTableRow[i].Due_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[3];
                            //if (lstPendingTableRow[i].Medical_Provider.Length > 20)
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].Medical_Provider.Substring(0, 20) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstPendingTableRow[i].Medical_Provider);
                            //}
                            cell.AddParagraph("Total");

                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = pendingRowData.Cells[4];
                            cell.AddParagraph(lstPendingTableRow[i].Bill_Amount);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[5];
                            cell.AddParagraph(lstPendingTableRow[i].Balance);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[6];
                            //cell.AddParagraph(lstPendingTableRow[i].Member_Discount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[7];
                            //cell.AddParagraph(lstPendingTableRow[i].CMM_Discount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[8];
                            //cell.AddParagraph(lstPendingTableRow[i].Shared_Amount);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = pendingRowData.Cells[9];
                            //cell.AddParagraph(lstPendingTableRow[i].Balance);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = pendingRowData.Cells[6];
                            if (lstPendingTableRow[i].Pending_Reason.Length > 40)
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason.Substring(0, 40) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            }
                            //cell.AddParagraph(lstPendingTableRow[i].Pending_Reason);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                        }

                    }

                    pdfDoc.LastSection.Add(tablePending);
                }

                int nHeightAfterIneligible = 0;

                if (gvIneligible.RowCount > 0)
                {
                    nHeightAfterIneligible += 22;
                    for (int nRow = 0; nRow < gvIneligible.RowCount; nRow++)
                    {
                        nHeightAfterIneligible += 15;
                    }

                    if ((nRowHeight > 645) ||
                        (nRowHeight + nHeightAfterIneligible) > 645)
                    {
                        nRowHeight = 0;
                        section.AddPageBreak();
                    }
                }

                lstBillIneligibleTableRow.Clear();

                for (int nRow = 0; nRow < gvIneligible.RowCount; nRow++)
                {
                    if (gvIneligible[4, nRow].Value.ToString() != "")
                        //(DateTime.Parse(gvIneligible[4, nRow].Value.ToString()) > dtDocReceivedDate.Value))
                    {
                        BillIneligibleTableRow ineligibleRow = new BillIneligibleTableRow();
                        ineligibleRow.PatientName = gvIneligible[1, nRow].Value.ToString();
                        ineligibleRow.MED_BILL = gvIneligible[2, nRow].Value.ToString();
                        ineligibleRow.Bill_Date = gvIneligible[3, nRow].Value.ToString();
                        ineligibleRow.Received_Date = gvIneligible[4, nRow].Value.ToString();
                        ineligibleRow.Medical_Provider = gvIneligible[5, nRow].Value.ToString();
                        ineligibleRow.Bill_Amount = gvIneligible[6, nRow].Value.ToString();
                        ineligibleRow.Amount_Ineligible = gvIneligible[7, nRow].Value.ToString();
                        ineligibleRow.Ineligible_Reason = gvIneligible[8, nRow].Value.ToString();

                        lstBillIneligibleTableRow.Add(ineligibleRow);
                    }
                    if (gvIneligible[4, nRow].Value.ToString() == "")
                    {
                        BillIneligibleTableRow ineligibleRow = new BillIneligibleTableRow();
                        ineligibleRow.PatientName = gvIneligible[1, nRow].Value.ToString();
                        ineligibleRow.MED_BILL = gvIneligible[2, nRow].Value.ToString();
                        ineligibleRow.Bill_Date = gvIneligible[3, nRow].Value.ToString();
                        ineligibleRow.Received_Date = gvIneligible[4, nRow].Value.ToString();
                        ineligibleRow.Medical_Provider = gvIneligible[5, nRow].Value.ToString();
                        ineligibleRow.Bill_Amount = gvIneligible[6, nRow].Value.ToString();
                        ineligibleRow.Amount_Ineligible = gvIneligible[7, nRow].Value.ToString();
                        ineligibleRow.Ineligible_Reason = gvIneligible[8, nRow].Value.ToString();

                        lstBillIneligibleTableRow.Add(ineligibleRow);
                    }
                }

                if (lstBillIneligibleTableRow.Count > 0)
                {

                    Paragraph paraSpaceBefore = section.AddParagraph();
                    paraSpaceBefore.Format.SpaceBefore = "0.18in";
                    paraSpaceBefore.Format.SpaceAfter = "0.05in";
                    paraSpaceBefore.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 100, 100, 0);
                    paraSpaceBefore.Format.Font.Name = "Arial";
                    paraSpaceBefore.Format.Font.Size = 7;
                    paraSpaceBefore.AddFormattedText("Ineligible Medical Bill(s)", TextFormat.Bold);

                    MigraDocDOM.Tables.Table tableIneligible = new MigraDocDOM.Tables.Table();
                    tableIneligible.Borders.Width = 0.1;
                    tableIneligible.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    //MigraDocDOM.Tables.Column colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Column colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(1));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(1.2));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    colIneligible = tableIneligible.AddColumn(MigraDocDOM.Unit.FromInch(2.1));
                    colIneligible.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row ineligible_Row = tableIneligible.AddRow();

                    nRowHeight += 22;
                    ineligible_Row.Height = "0.31in";
                    ineligible_Row.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    //MigraDocDOM.Tables.Cell cellIneligibleTitleICND = ineligible_Row.Cells[0];
                    //cellIneligibleTitleICND.AddParagraph("Member Name");
                    //cellIneligibleTitleICND.Format.Font.Bold = true;
                    //cellIneligibleTitleICND.Format.Font.Size = 7;
                    //cellIneligibleTitleICND.Format.Font.Name = "Arial";
                    //cellIneligibleTitleICND.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleMedBill = ineligible_Row.Cells[0];
                    cellIneligibleTitleMedBill.AddParagraph("MEDBILL");
                    cellIneligibleTitleMedBill.Format.Font.Bold = true;
                    cellIneligibleTitleMedBill.Format.Font.Size = 7;
                    cellIneligibleTitleMedBill.Format.Font.Name = "Arial";
                    cellIneligibleTitleMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleBillDate = ineligible_Row.Cells[1];
                    cellIneligibleTitleBillDate.AddParagraph("Date of Service");
                    cellIneligibleTitleBillDate.Format.Font.Bold = true;
                    cellIneligibleTitleBillDate.Format.Font.Size = 7;
                    cellIneligibleTitleBillDate.Format.Font.Name = "Arial";
                    cellIneligibleTitleBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleMedicalProvider = ineligible_Row.Cells[2];
                    cellIneligibleTitleMedicalProvider.AddParagraph("Medical Provider");
                    cellIneligibleTitleMedicalProvider.Format.Font.Bold = true;
                    cellIneligibleTitleMedicalProvider.Format.Font.Size = 7;
                    cellIneligibleTitleMedicalProvider.Format.Font.Name = "Arial";
                    cellIneligibleTitleMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleBillAmount = ineligible_Row.Cells[3];
                    cellIneligibleTitleBillAmount.AddParagraph("Original Amount");
                    cellIneligibleTitleBillAmount.Format.Font.Bold = true;
                    cellIneligibleTitleBillAmount.Format.Font.Size = 7;
                    cellIneligibleTitleBillAmount.Format.Font.Name = "Arial";
                    cellIneligibleTitleBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleAmountIneligible = ineligible_Row.Cells[4];
                    cellIneligibleTitleAmountIneligible.AddParagraph("Ineligible");
                    cellIneligibleTitleAmountIneligible.Format.Font.Bold = true;
                    cellIneligibleTitleAmountIneligible.Format.Font.Size = 7;
                    cellIneligibleTitleAmountIneligible.Format.Font.Name = "Arial";
                    cellIneligibleTitleAmountIneligible.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleTitleIneligibleReason = ineligible_Row.Cells[5];
                    cellIneligibleTitleIneligibleReason.AddParagraph("Ineligible Reason");
                    cellIneligibleTitleIneligibleReason.Format.Font.Bold = true;
                    cellIneligibleTitleIneligibleReason.Format.Font.Size = 7;
                    cellIneligibleTitleIneligibleReason.Format.Font.Name = "Arial";
                    cellIneligibleTitleIneligibleReason.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    List<BillIneligibleRow> lstBillIneligible = new List<BillIneligibleRow>();

                    for (int i = 0; i < lstBillIneligibleTableRow.Count; i++)
                    {
                        if (i < lstBillIneligibleTableRow.Count - 1)
                        {

                            MigraDocDOM.Tables.Row rowData = tableIneligible.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            BillIneligibleRow ineligible = new BillIneligibleRow();

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];

                            //if (lstBillIneligibleTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName);
                            //}

                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];

                            if (lstBillIneligibleTableRow[i].Medical_Provider.Length > 24)
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider.Substring(0, 24) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider);
                            }

                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Amount);
                            ineligible.Bill_Amount = Double.Parse(lstBillIneligibleTableRow[i].Bill_Amount.Substring(1));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Amount_Ineligible);
                            ineligible.Amount_Ineligible = Double.Parse(lstBillIneligibleTableRow[i].Amount_Ineligible.Substring(1));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];

                            if (lstBillIneligibleTableRow[i].Ineligible_Reason.Length > 33)
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason.Substring(0, 33) + " ...");
                            }
                            else
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason);
                            }
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            lstBillIneligible.Add(ineligible);
                        }

                        if (i == lstBillIneligibleTableRow.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligible.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                            rowData.Height = "0.2in";

                            //MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            //if (lstBillIneligibleTableRow[i].PatientName.Length > 11)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName.Substring(0, 11) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].PatientName);
                            //}
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Arial";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].MED_BILL);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstBillIneligibleTableRow[i].Bill_Date);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            cell = rowData.Cells[2];
                            //if (lstBillIneligibleTableRow[i].Medical_Provider.Length > 25)
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider.Substring(0, 25) + " ...");
                            //}
                            //else
                            //{
                            //    cell.AddParagraph(lstBillIneligibleTableRow[i].Medical_Provider);
                            //}
                            cell.AddParagraph("Total");
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                            Double? BillAmount = 0;
                            foreach (BillIneligibleRow row in lstBillIneligible)
                            {
                                BillAmount += row.Bill_Amount;
                            }

                            cell = rowData.Cells[3];
                            cell.AddParagraph(BillAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            Double? IneligibleAmount = 0;
                            foreach (BillIneligibleRow row in lstBillIneligible)
                            {
                                IneligibleAmount += row.Amount_Ineligible;
                            }

                            cell = rowData.Cells[4];
                            cell.AddParagraph(IneligibleAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            if (lstBillIneligibleTableRow[i].Ineligible_Reason.Length > 33)
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason.Substring(0, 33) + " ...");
                            }
                            {
                                cell.AddParagraph(lstBillIneligibleTableRow[i].Ineligible_Reason);
                            }
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Arial";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                    }

                    pdfDoc.LastSection.Add(tableIneligible);

                }
                // The end of tables

                const bool unicode = true;
                const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);
                pdfRenderer.Document = pdfDoc;
                pdfRenderer.RenderDocument();

                if (rbCheck.Checked)
                {

                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + ChkInfoEntered.dtCheckIssueDate.ToString("MM-dd-yyyy") + "_En";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;
                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                    }
                }

                if (rbACH.Checked)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + ACHInfoEntered.dtACHDate.ToString("MM-dd-yyyy") + "_En";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch(IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                    }
                }
                if (rbCreditCard.Checked)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + CreditCardPaymentEntered.dtPaymentDate.ToString("MM-dd-yyyy") + "_En";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch(IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                    }
                }
                //}
                //else if (dlgResultDocReceivedDate == DialogResult.Cancel)
                //{
                //    return;
                //}

            }
            else if ((gvPersonalResponsibility.Rows.Count > 0)||(gvIneligibleNoSharing.Rows.Count > 0))
            {

                List<PersonalResponsibilityInfo> lstPersonalResponsibilityInfo = new List<PersonalResponsibilityInfo>();

                for (int i = 0; i < gvPersonalResponsibility.Rows.Count; i++)
                {
                    PersonalResponsibilityInfo prInfo = new PersonalResponsibilityInfo();

                    prInfo.MedBillName = gvPersonalResponsibility[0, i]?.Value.ToString();
                    String BillDate = gvPersonalResponsibility[1, i]?.Value.ToString();
                    if (BillDate != String.Empty) prInfo.BillDate = DateTime.Parse(BillDate);
                    prInfo.MedicalProvider = gvPersonalResponsibility[2, i]?.Value.ToString();
                    if (gvPersonalResponsibility[3, i] != null) prInfo.BillAmount = (Double)Decimal.Parse(gvPersonalResponsibility[3, i].Value.ToString().Substring(1));
                    prInfo.Type = gvPersonalResponsibility[4, i]?.Value.ToString();
                    if (gvPersonalResponsibility[5, i] != null) prInfo.PersonalResponsibilityTotal = (Double)Decimal.Parse(gvPersonalResponsibility[8, i].Value.ToString().Substring(1));

                    lstPersonalResponsibilityInfo.Add(prInfo);
                }

                Document pdfPersonalResponsibilityDoc = new Document();

                Section section = pdfPersonalResponsibilityDoc.AddSection();

                section.PageSetup.PageFormat = PageFormat.Letter;
                section.PageSetup.HeaderDistance = "0.25in";
                section.PageSetup.TopMargin = "1.5in";
                section.PageSetup.LeftMargin = "0.8in";
                section.PageSetup.RightMargin = "0.8in";
                section.PageSetup.BottomMargin = "0.5in";

                section.PageSetup.DifferentFirstPageHeaderFooter = false;
                section.Headers.Primary.Format.SpaceBefore = "0.25in";

                MigraDocDOM.Shapes.Image image = section.Headers.Primary.AddImage("C:\\Program Files (x86)\\CMM\\BlueSheet\\cmmlogo.png");

                image.Height = "0.8in";
                image.LockAspectRatio = true;
                image.RelativeVertical = MigraDocDOM.Shapes.RelativeVertical.Line;
                image.RelativeHorizontal = MigraDocDOM.Shapes.RelativeHorizontal.Margin;
                image.Top = MigraDocDOM.Shapes.ShapePosition.Top;
                image.Left = MigraDocDOM.Shapes.ShapePosition.Center;
                image.WrapFormat.Style = MigraDocDOM.Shapes.WrapStyle.TopBottom;

                Paragraph paraCMMAddress = section.Headers.Primary.AddParagraph();
                paraCMMAddress.Format.Font.Name = "Arial";
                paraCMMAddress.Format.Font.Size = 8;
                paraCMMAddress.Format.SpaceBefore = "0.15in";
                paraCMMAddress.Format.SpaceAfter = "0.25in";
                paraCMMAddress.Format.Alignment = ParagraphAlignment.Center;

                String strVerticalBar = " | ";
                String strStreet = "5235 N. Elston Ave.";
                String strCityStateZip = "Chicago, IL 60630";
                String strPhone = "Phone 773.777.8889";
                String strFax = "Fax 773.777.0004";
                String strWebsiteAddr = "www.cmmlogos.org";

                paraCMMAddress.AddFormattedText(strStreet, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strCityStateZip, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strPhone, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strFax, TextFormat.NotBold);
                paraCMMAddress.AddFormattedText(strVerticalBar, TextFormat.Bold);
                paraCMMAddress.AddFormattedText(strWebsiteAddr, TextFormat.NotBold);

                Paragraph paraToday = section.AddParagraph();
                paraToday.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraToday.Format.Font.Name = "Arial";
                paraToday.Format.Alignment = ParagraphAlignment.Left;
                paraToday.Format.SpaceBefore = "0.25in";
                paraToday.Format.SpaceAfter = "0.25in";
                paraToday.AddFormattedText(DateTime.Today.ToString("MM/dd/yyyy"));

                Paragraph paraMembershipInfo = section.AddParagraph();

                paraMembershipInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraMembershipInfo.Format.Font.Name = "Arial";
                //paraMembershipInfo.Format.SpaceBefore = "0.70in";
                paraMembershipInfo.Format.SpaceBefore = "0.20in";
                paraMembershipInfo.Format.SpaceAfter = "0.20in";
                paraMembershipInfo.Format.LeftIndent = "0.75in";
                //paraMembershipInfo.Format.LeftIndent = "0.5in";
                //paraMembershipInfo.Format.RightIndent = "0.5in";
                paraMembershipInfo.Format.Alignment = ParagraphAlignment.Left;
                //paraMembershipInfo.AddFormattedText("Primary Name: " + strPrimaryName + "\n");
                if (strMembershipId != String.Empty) paraMembershipInfo.AddFormattedText(strMembershipId + " (" + strIndividualID + ")\n");
                else paraMembershipInfo.AddFormattedText(strIndividualID + "\n");
                paraMembershipInfo.AddFormattedText(strIndividualName.Trim() + "\n");
                paraMembershipInfo.AddFormattedText(strStreetAddress + "\n");
                paraMembershipInfo.AddFormattedText(strCity + ", " + strState + " " + strZip + "\n");


                Paragraph paraDearMember = section.AddParagraph();
                paraDearMember.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraDearMember.Format.Font.Name = "Arial";
                paraDearMember.Format.Font.Size = 8;
                paraDearMember.Format.Alignment = ParagraphAlignment.Left;
                //paraDearMember.Format.LeftIndent = "0.5in";
                //paraDearMember.Format.RightIndent = "0.5in";
                paraDearMember.Format.SpaceBefore = "0.1in";
                paraDearMember.Format.SpaceAfter = "0.1in";
                //if (strIndividualMiddleName != String.Empty) paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " 회원께,");
                //else paraDearMember.AddFormattedText(strIndividualLastName + ", " + strIndiviaualFirstName + " " + strIndividualMiddleName + " 회원께,");
                //paraDearMember.AddFormattedText(strIndividualName + " 회원께,");
                paraDearMember.AddFormattedText(strDearMember + strIndividualName.Trim() + ", ");

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Paragraph paraPRGreetingMessage1 = section.AddParagraph();

                paraPRGreetingMessage1.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage1.Format.Font.Name = "Arial";
                paraPRGreetingMessage1.Format.Font.Size = 8;
                paraPRGreetingMessage1.Format.SpaceAfter = "5pt";
                //paraGreetingMessage.Format.LeftIndent = "0.5in";
                //paraGreetingMessage.Format.RightIndent = "0.5in";
                paraPRGreetingMessage1.Format.Alignment = ParagraphAlignment.Left;
                //paraGreetingMessage.AddFormattedText(strGreetingMessage, TextFormat.NotBold);
                paraPRGreetingMessage1.AddFormattedText(strEnglishPRGreetingMessage1, TextFormat.NotBold);




                //////////////////////////////////////////////////////////////////////////////////////////////////////
                /// Program personal responsibility table
                /// 

                MigraDocDOM.Tables.Table tableProgramPRGuide = new MigraDocDOM.Tables.Table();
                tableProgramPRGuide.Borders.Width = 0.1;
                tableProgramPRGuide.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Column colProgram = tableProgramPRGuide.AddColumn(MigraDocDOM.Unit.FromInch(2));
                colProgram.Format.Alignment = ParagraphAlignment.Left;
                MigraDocDOM.Tables.Column colPersonalResponsibility = tableProgramPRGuide.AddColumn(MigraDocDOM.Unit.FromInch(4.8));
                colPersonalResponsibility.Format.Alignment = ParagraphAlignment.Left;

                MigraDocDOM.Tables.Row rowHeader = tableProgramPRGuide.AddRow();
                rowHeader.Height = "0.2in";
                rowHeader.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellProgramHeader = rowHeader.Cells[0];
                cellProgramHeader.Format.Font.Bold = true;
                cellProgramHeader.Format.Font.Size = 8;
                cellProgramHeader.Format.Font.Name = "Arial";
                cellProgramHeader.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellProgramHeader.AddParagraph("Program");

                MigraDocDOM.Tables.Cell cellPersonalResponsibility = rowHeader.Cells[1];
                cellPersonalResponsibility.Format.Font.Bold = true;
                cellPersonalResponsibility.Format.Font.Size = 8;
                cellPersonalResponsibility.Format.Font.Name = "Arial";
                cellPersonalResponsibility.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellPersonalResponsibility.AddParagraph("Personal Responsibility Amount");

                MigraDocDOM.Tables.Row rowBronze = tableProgramPRGuide.AddRow();
                rowBronze.Height = "0.2in";
                rowBronze.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellBronze = rowBronze.Cells[0];
                cellBronze.Format.Font.Bold = false;
                cellBronze.Format.Font.Size = 8;
                cellBronze.Format.Font.Name = "Arial";
                cellBronze.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellBronze.AddParagraph("Bronze");

                MigraDocDOM.Tables.Cell cellBronzePR = rowBronze.Cells[1];
                cellBronzePR.Format.Font.Bold = false;
                cellBronzePR.Format.Font.Size = 8;
                cellBronzePR.Format.Font.Name = "Arial";
                cellBronzePR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellBronzePR.AddParagraph("$5,000 Per Incident");

                MigraDocDOM.Tables.Row rowSilver = tableProgramPRGuide.AddRow();
                rowSilver.Height = "0.2in";
                rowSilver.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellSilver = rowSilver.Cells[0];
                cellSilver.Format.Font.Bold = false;
                cellSilver.Format.Font.Size = 8;
                cellSilver.Format.Font.Name = "Arial";
                cellSilver.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellSilver.AddParagraph("Silver");

                MigraDocDOM.Tables.Cell cellSilverPR = rowSilver.Cells[1];
                cellSilverPR.Format.Font.Bold = false;
                cellSilverPR.Format.Font.Size = 8;
                cellSilverPR.Format.Font.Name = "Arial";
                cellSilverPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellSilverPR.AddParagraph("$1,000 Per Incident");

                MigraDocDOM.Tables.Row rowGold = tableProgramPRGuide.AddRow();
                rowGold.Height = "0.2in";
                rowGold.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellGold = rowGold.Cells[0];
                cellGold.Format.Font.Bold = false;
                cellGold.Format.Font.Size = 8;
                cellGold.Format.Font.Name = "Arial";
                cellGold.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGold.AddParagraph("Gold");

                MigraDocDOM.Tables.Cell cellGoldPR = rowGold.Cells[1];
                cellGoldPR.Format.Font.Bold = false;
                cellGoldPR.Format.Font.Size = 8;
                cellGoldPR.Format.Font.Name = "Arial";
                cellGoldPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPR.AddParagraph("$500 Per Incident");

                MigraDocDOM.Tables.Row rowGoldPlus = tableProgramPRGuide.AddRow();
                rowGoldPlus.Height = "0.2in";
                rowGoldPlus.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellGoldPlus = rowGoldPlus.Cells[0];
                cellGoldPlus.Format.Font.Bold = false;
                cellGoldPlus.Format.Font.Size = 8;
                cellGoldPlus.Format.Font.Name = "Arial";
                cellGoldPlus.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPlus.AddParagraph("Gold Plus");

                MigraDocDOM.Tables.Cell cellGoldPlusPR = rowGoldPlus.Cells[1];
                cellGoldPlusPR.Format.Font.Bold = false;
                cellGoldPlusPR.Format.Font.Size = 8;
                cellGoldPlusPR.Format.Font.Name = "Arial";
                cellGoldPlusPR.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                cellGoldPlusPR.AddParagraph("$500 Per Membership Anniversary");

                pdfPersonalResponsibilityDoc.LastSection.Add(tableProgramPRGuide);


                Paragraph paraPRGreetingMessage2 = section.AddParagraph();
                paraPRGreetingMessage2.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage2.Format.Font.Name = "Arial";
                paraPRGreetingMessage2.Format.Font.Size = 8;
                paraPRGreetingMessage2.Format.SpaceAfter = "5pt";
                //paraGreetingMessage2.Format.LeftIndent = "0.5in";
                //paraGreetingMessage2.Format.RightIndent = "0.5in";
                paraPRGreetingMessage2.Format.Alignment = ParagraphAlignment.Left;
                paraPRGreetingMessage2.AddFormattedText(strEnglishPRGreetingMessage2, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage3 = section.AddParagraph();
                paraPRGreetingMessage3.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage3.Format.Font.Name = "Arial";
                paraPRGreetingMessage3.Format.Font.Size = 8;
                paraPRGreetingMessage3.Format.SpaceAfter = "5pt";
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraPRGreetingMessage3.Format.Alignment = ParagraphAlignment.Left;
                paraPRGreetingMessage3.AddFormattedText(strEnglishPRGreetingMessage3, TextFormat.NotBold);


                Paragraph paraPRGreetingMessage4 = section.AddParagraph();
                paraPRGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage4.Format.Font.Name = "Arial";
                paraPRGreetingMessage4.Format.Font.Size = 8;
                paraPRGreetingMessage4.Format.SpaceAfter = "5pt";
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraPRGreetingMessage4.Format.Alignment = ParagraphAlignment.Left;
                paraPRGreetingMessage4.AddFormattedText(strEnglishPRGreetingMessage4, TextFormat.NotBold);

                Paragraph paraPRGreetingMessage5 = section.AddParagraph();
                paraPRGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                paraPRGreetingMessage4.Format.Font.Name = "Arial";
                paraPRGreetingMessage4.Format.Font.Size = 8;
                paraPRGreetingMessage4.Format.SpaceAfter = "5pt";
                //paraGreetingMessage3.Format.LeftIndent = "0.5in";
                //paraGreetingMessage3.Format.RightIndent = "0.5in";
                paraPRGreetingMessage4.Format.Alignment = ParagraphAlignment.Justify;
                paraPRGreetingMessage4.AddFormattedText(strEnglishPRGreetingMessage5, TextFormat.NotBold);

                //Paragraph paraPRGreetingMessage4 = section.AddParagraph();
                //paraPRGreetingMessage4.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                //paraPRGreetingMessage4.Format.Font.Name = "Arial";
                //paraPRGreetingMessage4.Format.Font.Size = 8;
                ////paraGreetingMessage4.Format.LeftIndent = "0.5in";
                ////paraGreetingMessage4.Format.RightIndent = "0.5in";
                //paraPRGreetingMessage4.Format.Alignment = ParagraphAlignment.Justify;
                //paraPRGreetingMessage4.AddFormattedText(strPRGreetingMessagePara4, TextFormat.NotBold);

                //Paragraph paraPRGreetingMessage5 = section.AddParagraph();
                //paraPRGreetingMessage5.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                //paraPRGreetingMessage5.Format.Font.Name = "Arial";
                //paraPRGreetingMessage5.Format.Font.Size = 8;
                ////paraGreetingMessage4.Format.LeftIndent = "0.5in";
                ////paraGreetingMessage4.Format.RightIndent = "0.5in";
                //paraPRGreetingMessage5.Format.Alignment = ParagraphAlignment.Justify;
                //paraPRGreetingMessage5.AddFormattedText(strPRGreetingMessagePara5, TextFormat.NotBold);

                //Paragraph paraPRGreetingMessage6 = section.AddParagraph();
                //paraPRGreetingMessage6.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                //paraPRGreetingMessage6.Format.Font.Name = "Arial";
                //paraPRGreetingMessage6.Format.Font.Size = 8;
                ////paraGreetingMessage4.Format.LeftIndent = "0.5in";
                ////paraGreetingMessage4.Format.RightIndent = "0.5in";
                //paraPRGreetingMessage6.Format.Alignment = ParagraphAlignment.Justify;
                //paraPRGreetingMessage6.AddFormattedText(strPRGreetingMessagePara6, TextFormat.NotBold);




                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                Paragraph paraNeedsProcessing = section.AddParagraph();

                paraNeedsProcessing.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraNeedsProcessing.Format.Font.Name = "Arial";
                paraNeedsProcessing.Format.Font.Size = 8;
                paraNeedsProcessing.Format.Font.Bold = true;
                paraNeedsProcessing.Format.Alignment = ParagraphAlignment.Left;
                //paraNeedsProcessing.Format.SpaceBefore = "0.1in";
                //paraNeedsProcessing.Format.LeftIndent = "0.5in";
                //paraNeedsProcessing.Format.RightIndent = "0.5in";
                paraNeedsProcessing.AddFormattedText(strCMM_NeedProcessing + "\n");

                Paragraph paraPhoneFaxEmail = section.AddParagraph();
                paraPhoneFaxEmail.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPhoneFaxEmail.Format.Font.Size = 8;
                paraPhoneFaxEmail.Format.Alignment = ParagraphAlignment.Left;
                //paraPhoneFaxEmail.Format.LeftIndent = "0.5in";
                paraPhoneFaxEmail.Format.RightIndent = "0.5in";
                paraPhoneFaxEmail.AddFormattedText(strNP_Phone_Fax_Email + "\n");

                Paragraph paraHorizontalLine = section.AddParagraph();

                paraHorizontalLine.Format.SpaceBefore = "0.05in";
                paraHorizontalLine.Format.SpaceAfter = "0.05in";
                paraHorizontalLine.Format.Borders.Top.Width = 0;
                paraHorizontalLine.Format.Borders.Left.Width = 0;
                paraHorizontalLine.Format.Borders.Right.Width = 0;
                paraHorizontalLine.Format.Borders.Bottom.Width = 1;
                paraHorizontalLine.Format.Borders.Bottom.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraHorizontalLine.Format.Borders.Style = MigraDocDOM.BorderStyle.DashDot;

                Paragraph paraNPStatement = section.AddParagraph();
                paraNPStatement.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 0, 0);
                paraNPStatement.Format.Font.Name = "Arial";
                paraNPStatement.Format.Font.Size = 12;
                paraNPStatement.Format.Font.Bold = true;
                paraNPStatement.Format.Alignment = ParagraphAlignment.Center;
                paraNPStatement.Format.SpaceAfter = "0.1in";

                paraNPStatement.AddFormattedText("Needs Processing Statement\n", TextFormat.Bold);

                Paragraph paraPersonalResponsibilityTotal = section.AddParagraph();
                paraPersonalResponsibilityTotal.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPersonalResponsibilityTotal.Format.Font.Name = "Arial";
                paraPersonalResponsibilityTotal.Format.Font.Size = 8;
                paraPersonalResponsibilityTotal.Format.Font.Bold = true;
                paraPersonalResponsibilityTotal.Format.Alignment = ParagraphAlignment.Left;
                paraPersonalResponsibilityTotal.Format.SpaceBefore = "0.2in";
                //paraPersonalResponsibilityTotal.Format.SpaceAfter = "0.2in";

                //paraPersonalResponsibilityTotal.AddFormattedText("Incident Occurrence Date: " + PersonalResponsibilityTotalEntered.IncidentOccurrenceDate.Value.ToString("MM/dd/yyyy") + "\t" +
                //                                                 "Personal Responsibility Total: " + PersonalResponsibilityTotalEntered.PersonalResponsibilityTotal.ToString("C"));

                paraPersonalResponsibilityTotal.AddFormattedText("Incident Occurrence Date: " + PersonalResponsibilityTotalEntered.IncidentOccurrenceDate.Value.ToString("MM/dd/yyyy"));

                Paragraph paraINCDNumber = section.AddParagraph();
                paraINCDNumber.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraINCDNumber.Format.Font.Name = "Arial";
                paraINCDNumber.Format.Font.Size = 8;
                paraINCDNumber.Format.Font.Bold = true;
                paraINCDNumber.Format.Alignment = ParagraphAlignment.Left;
                paraINCDNumber.Format.SpaceBefore = "0.05in";
                paraINCDNumber.Format.SpaceAfter = "0.15in";

                paraINCDNumber.AddFormattedText(PersonalResponsibilityTotalEntered.IncidentNo + ": " + PersonalResponsibilityTotalEntered.ICD10CodeDescription);

                //Paragraph paraIncidentInfo = section.AddParagraph();
                //paraIncidentInfo.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //paraIncidentInfo.Format.Font.Name = "Arial";
                //paraIncidentInfo.Format.Font.Size = 8;
                //paraIncidentInfo.Format.Font.Bold = true;
                //paraIncidentInfo.Format.Alignment = ParagraphAlignment.Left;
                //paraIncidentInfo.Format.SpaceBefore = "0.2in";
                //paraIncidentInfo.Format.SpaceAfter = "0.2in";

                //if (lstIncidents.Count > 0)
                //{
                //    Paragraph paraIncd = section.AddParagraph();

                //    paraIncd.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);
                //    paraIncd.Format.Font.Name = "Arial";
                //    paraIncd.Format.Font.Size = 8;
                //    paraIncd.Format.Font.Bold = true;

                //    MigraDocDOM.Tables.Table tableIncd = new MigraDocDOM.Tables.Table();
                //    tableIncd.Borders.Width = 0;
                //    tableIncd.Borders.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //    tableIncd.Format.SpaceAfter = "0.05in";

                //    MigraDocDOM.Tables.Column colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(0.85));
                //    colINCD.Format.Alignment = ParagraphAlignment.Left;
                //    //colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(1.1));
                //    //colINCD.Format.Alignment = ParagraphAlignment.Left;
                //    colINCD = tableIncd.AddColumn(MigraDocDOM.Unit.FromInch(4.5));
                //    colINCD.Format.Alignment = ParagraphAlignment.Left;

                //    foreach (Incident incd in lstIncidents)
                //    {
                //        //nRowHeight += 18;

                //        MigraDocDOM.Tables.Row IncdRow = tableIncd.AddRow();
                //        IncdRow.Height = "0.1in";
                //        IncdRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                //        MigraDocDOM.Tables.Cell cellIncdName = IncdRow.Cells[0];
                //        cellIncdName.Format.Font.Bold = true;
                //        cellIncdName.Format.Font.Size = 8;
                //        //cellIncdName.Format.Font.Name = "Malgun Gothic";
                //        cellIncdName.Format.Font.Name = "Arial";
                //        cellIncdName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //        cellIncdName.AddParagraph(incd.Name + ": ");

                //        //MigraDocDOM.Tables.Cell cellPatientName = IncdRow.Cells[1];
                //        //cellPatientName.Format.Font.Bold = true;
                //        //cellPatientName.Format.Font.Size = 8;
                //        ////cellPatientName.Format.Font.Name = "Malgun Gothic";
                //        //cellIncdName.Format.Font.Name = "Arial";
                //        //cellPatientName.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //        //if (incd.PatientName.Length > 11)
                //        //{
                //        //    cellPatientName.AddParagraph(incd.PatientName.Substring(0, 11) + " ...");
                //        //}
                //        //else cellPatientName.AddParagraph(incd.PatientName);

                //        MigraDocDOM.Tables.Cell cellICD10Code = IncdRow.Cells[1];
                //        cellICD10Code.Format.Font.Bold = true;
                //        cellICD10Code.Format.Font.Size = 8;
                //        //cellICD10Code.Format.Font.Name = "Malgun Gothic";
                //        cellIncdName.Format.Font.Name = "Arial";
                //        cellICD10Code.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                //        cellICD10Code.AddParagraph(incd.ICD10_Code);
                //    }

                //    pdfPersonalResponsibilityDoc.LastSection.Add(tableIncd);
                //}

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///

                Paragraph paraPersonalResponsibilityTitle = section.AddParagraph();
                paraPersonalResponsibilityTitle.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                paraPersonalResponsibilityTitle.Format.Font.Name = "Arial";
                paraPersonalResponsibilityTitle.Format.Font.Size = 7;
                paraPersonalResponsibilityTitle.Format.Font.Bold = true;
                paraPersonalResponsibilityTitle.Format.Alignment = ParagraphAlignment.Left;
                paraPersonalResponsibilityTitle.Format.SpaceBefore = "0.18in";
                paraPersonalResponsibilityTitle.Format.SpaceAfter = "0.05in";
                paraPersonalResponsibilityTitle.AddFormattedText("Personal Responsibility", TextFormat.Bold);



                MigraDocDOM.Tables.Table tablePersonalResponsibility = new MigraDocDOM.Tables.Table();
                tablePersonalResponsibility.Borders.Width = 0.1;
                tablePersonalResponsibility.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Column colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.6));  // Med bill
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.7));        // 서비스 날짜
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.8));        // Medical Provider
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(0.9));        // Bill Amount
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.2));        // Type
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                colPersonalResponsibilityColumn = tablePersonalResponsibility.AddColumn(MigraDocDOM.Unit.FromInch(1.5));        // Personal Responsibility Total
                colPersonalResponsibilityColumn.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                // generate row for personal responsibility table
                MigraDocDOM.Tables.Row prRow = tablePersonalResponsibility.AddRow();
                prRow.Height = "0.31in";
                prRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityMedBill = prRow.Cells[0];
                cellPersonalResponsibilityMedBill.AddParagraph("MEDBILL");
                cellPersonalResponsibilityMedBill.Format.Font.Bold = true;
                cellPersonalResponsibilityMedBill.Format.Font.Size = 7;
                cellPersonalResponsibilityMedBill.Format.Font.Name = "Arial";
                cellPersonalResponsibilityMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityBillDate = prRow.Cells[1];
                cellPersonalResponsibilityBillDate.AddParagraph("Date of Service");
                cellPersonalResponsibilityBillDate.Format.Font.Bold = true;
                cellPersonalResponsibilityBillDate.Format.Font.Size = 7;
                cellPersonalResponsibilityBillDate.Format.Font.Name = "Arial";
                cellPersonalResponsibilityBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityMedicalProvider = prRow.Cells[2];
                cellPersonalResponsibilityMedicalProvider.AddParagraph("Medical Provider");
                cellPersonalResponsibilityMedicalProvider.Format.Font.Bold = true;
                cellPersonalResponsibilityMedicalProvider.Format.Font.Size = 7;
                cellPersonalResponsibilityMedicalProvider.Format.Font.Name = "Arial";
                cellPersonalResponsibilityMedicalProvider.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityBillAmount = prRow.Cells[3];
                cellPersonalResponsibilityBillAmount.AddParagraph("Original Amount");
                cellPersonalResponsibilityBillAmount.Format.Font.Bold = true;
                cellPersonalResponsibilityBillAmount.Format.Font.Size = 7;
                cellPersonalResponsibilityBillAmount.Format.Font.Name = "Arial";
                cellPersonalResponsibilityBillAmount.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityType = prRow.Cells[4];
                cellPersonalResponsibilityType.AddParagraph("Type");
                cellPersonalResponsibilityType.Format.Font.Bold = true;
                cellPersonalResponsibilityType.Format.Font.Size = 7;
                cellPersonalResponsibilityType.Format.Font.Name = "Arial";
                cellPersonalResponsibilityType.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                MigraDocDOM.Tables.Cell cellPersonalResponsibilityTotal = prRow.Cells[5];
                cellPersonalResponsibilityTotal.AddParagraph("Personal Responsibility Total");
                cellPersonalResponsibilityTotal.Format.Font.Bold = true;
                cellPersonalResponsibilityTotal.Format.Font.Size = 7;
                cellPersonalResponsibilityTotal.Format.Font.Name = "Arial";
                cellPersonalResponsibilityTotal.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);


                for (int i = 0; i < lstPersonalResponsibilityInfo.Count; i++)
                {
                    if (i < lstPersonalResponsibilityInfo.Count - 1)
                    {
                        MigraDocDOM.Tables.Row rowData = tablePersonalResponsibility.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                        rowData.Height = "0.18in";

                        MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedBillName.Substring(8));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[1];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillDate.Value.ToString("MM/dd/yyyy"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[2];
                        if (lstPersonalResponsibilityInfo[i].MedicalProvider.Length > 30) cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider.Substring(0, 30) + "...");
                        else cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider);
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[3];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillAmount.ToString("C"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;

                        cell = rowData.Cells[4];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].Type);
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[5];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].PersonalResponsibilityTotal.ToString("C"));
                        cell.Format.Font.Bold = false;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;
                    }
                    else if (i == lstPersonalResponsibilityInfo.Count - 1)
                    {
                        MigraDocDOM.Tables.Row rowData = tablePersonalResponsibility.AddRow();
                        rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                        rowData.Height = "0.2in";

                        MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedBillName);
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[1];
                        if (lstPersonalResponsibilityInfo[i].BillDate != null) cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillDate.Value.ToString("MM/dd/yyyy"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[2];
                        //cell.AddParagraph(lstPersonalResponsibilityInfo[i].MedicalProvider);
                        cell.AddParagraph("Total");
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        //cell.Format.Alignment = ParagraphAlignment.Left;
                        cell.Format.Alignment = ParagraphAlignment.Center;

                        cell = rowData.Cells[3];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].BillAmount.ToString("C"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;

                        cell = rowData.Cells[4];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].Type);
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Left;

                        cell = rowData.Cells[5];
                        cell.AddParagraph(lstPersonalResponsibilityInfo[i].PersonalResponsibilityTotal.ToString("C"));
                        cell.Format.Font.Bold = true;
                        cell.Format.Font.Name = "Arial";
                        cell.Format.Font.Size = 7;
                        cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                        cell.Format.Alignment = ParagraphAlignment.Right;
                    }
                }

                pdfPersonalResponsibilityDoc.LastSection.Add(tablePersonalResponsibility);
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///

                if (gvIneligibleNoSharing.Rows.Count > 0)
                {
                    List<SettlementIneligibleInfo> lstMedBillNoPRNoSharing = new List<SettlementIneligibleInfo>();

                    for (int i = 0; i < gvIneligibleNoSharing.Rows.Count; i++)
                    {
                        SettlementIneligibleInfo info = new SettlementIneligibleInfo();

                        //String BillDate = gvPersonalResponsibility[1, i]?.Value.ToString();
                        //if (BillDate != String.Empty) prInfo.BillDate = DateTime.Parse(BillDate);


                        info.MedBillName = gvIneligibleNoSharing["MEDBILL", i]?.Value.ToString();
                        //info.BillDate = DateTime.Parse(gvIneligibleNoSharing["서비스 날짜", i].Value.ToString());
                        String BillDate = gvIneligibleNoSharing["서비스 날짜", i]?.Value.ToString();
                        if (BillDate != String.Empty) info.BillDate = DateTime.Parse(BillDate);

                        info.MedicalProvider = gvIneligibleNoSharing["의료기관명", i]?.Value.ToString();
                        info.BillAmount = Double.Parse(gvIneligibleNoSharing["청구액(원금)", i]?.Value.ToString().Substring(1));
                        //info.Type = gvIneligibleNoSharing["Type", i]?.Value.ToString();
                        info.IneligibleAmount = Double.Parse(gvIneligibleNoSharing["지원불가 의료비", i]?.Value.ToString().Substring(1));
                        info.IneligibleReason = gvIneligibleNoSharing["지원불가 사유", i]?.Value.ToString();

                        lstMedBillNoPRNoSharing.Add(info);
                    }

                    Paragraph paraIneligibleTitle = section.AddParagraph();
                    paraIneligibleTitle.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                    paraIneligibleTitle.Format.Font.Name = "Arial";
                    paraIneligibleTitle.Format.Font.Size = 7;
                    paraIneligibleTitle.Format.Font.Bold = true;
                    paraIneligibleTitle.Format.Alignment = ParagraphAlignment.Left;
                    paraIneligibleTitle.Format.SpaceBefore = "0.15in";
                    paraIneligibleTitle.Format.SpaceAfter = "0.05in";
                    paraIneligibleTitle.AddFormattedText("Ineligible Medical Expenses", TextFormat.Bold);


                    MigraDocDOM.Tables.Table tableIneligibleNoPR = new MigraDocDOM.Tables.Table();
                    tableIneligibleNoPR.Borders.Width = 0.1;
                    tableIneligibleNoPR.Borders.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Column colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.6));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(2.5));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.7));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    //colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(0.8));
                    //colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;
                    colIneligibleNoPR = tableIneligibleNoPR.AddColumn(MigraDocDOM.Unit.FromInch(1.4));
                    colIneligibleNoPR.Format.Alignment = MigraDocDOM.ParagraphAlignment.Center;

                    MigraDocDOM.Tables.Row ineligibleRow = tableIneligibleNoPR.AddRow();
                    ineligibleRow.Height = "0.31in";
                    ineligibleRow.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRMedBill = ineligibleRow.Cells[0];
                    cellIneligibleNoPRMedBill.AddParagraph("MEDBILL");
                    cellIneligibleNoPRMedBill.Format.Font.Bold = true;
                    cellIneligibleNoPRMedBill.Format.Font.Size = 7;
                    cellIneligibleNoPRMedBill.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleNoPRMedBill.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRBillDate = ineligibleRow.Cells[1];
                    cellIneligibleNoPRBillDate.AddParagraph("Date of Service");
                    cellIneligibleNoPRBillDate.Format.Font.Bold = true;
                    cellIneligibleNoPRBillDate.Format.Font.Size = 7;
                    cellIneligibleNoPRBillDate.Format.Font.Name = "Malgun Gothic";
                    cellIneligibleNoPRBillDate.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRMedicalProvider = ineligibleRow.Cells[2];
                    cellIneligibleNoPRMedicalProvider.AddParagraph("Medical Provider");
                    cellIneligibleNoPRMedicalProvider.Format.Font.Bold = true;
                    cellIneligibleNoPRMedicalProvider.Format.Font.Size = 7;
                    cellIneligibleNoPRMedicalProvider.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRBillAmount = ineligibleRow.Cells[3];
                    cellIneligibleNoPRBillAmount.AddParagraph("Original Amount");
                    cellIneligibleNoPRBillAmount.Format.Font.Bold = true;
                    cellIneligibleNoPRBillAmount.Format.Font.Size = 7;
                    cellIneligibleNoPRBillAmount.Format.Font.Name = "Malgun Gothic";

                    //MigraDocDOM.Tables.Cell cellIneligibleNoPRType = ineligibleRow.Cells[4];
                    //cellIneligibleNoPRType.AddParagraph("Type");
                    //cellIneligibleNoPRType.Format.Font.Bold = true;
                    //cellIneligibleNoPRType.Format.Font.Size = 7;
                    //cellIneligibleNoPRType.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRIneligibleAmount = ineligibleRow.Cells[4];
                    cellIneligibleNoPRIneligibleAmount.AddParagraph("Ineligible");
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Bold = true;
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Size = 7;
                    cellIneligibleNoPRIneligibleAmount.Format.Font.Name = "Malgun Gothic";

                    MigraDocDOM.Tables.Cell cellIneligibleNoPRIneligibleReason = ineligibleRow.Cells[5];
                    cellIneligibleNoPRIneligibleReason.AddParagraph("Ineligible Reason");
                    cellIneligibleNoPRIneligibleReason.Format.Font.Bold = true;
                    cellIneligibleNoPRIneligibleReason.Format.Font.Size = 7;
                    cellIneligibleNoPRIneligibleReason.Format.Font.Name = "Malgun Gothic";

                    for (int i = 0; i < lstMedBillNoPRNoSharing.Count; i++)
                    {
                        if (i < lstMedBillNoPRNoSharing.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligibleNoPR.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedBillName.Substring(8));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillDate.Value.ToString("MM/dd/yyyy"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedicalProvider);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillAmount.ToString("C"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = rowData.Cells[4];
                            //cell.AddParagraph(lstMedBillNoPRNoSharing[i].Type);
                            //cell.Format.Font.Bold = false;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Left;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleReason);
                            cell.Format.Font.Bold = false;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                        else if (i == lstMedBillNoPRNoSharing.Count - 1)
                        {
                            MigraDocDOM.Tables.Row rowData = tableIneligibleNoPR.AddRow();
                            rowData.VerticalAlignment = MigraDocDOM.Tables.VerticalAlignment.Center;
                            rowData.Height = "0.18in";

                            MigraDocDOM.Tables.Cell cell = rowData.Cells[0];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedBillName);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[1];
                            if (lstMedBillNoPRNoSharing[i].BillDate != null) cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillDate.Value.ToString("MM/dd/yyyy"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[2];
                            //cell.AddParagraph(lstMedBillNoPRNoSharing[i].MedicalProvider);
                            cell.AddParagraph("Total");
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[3];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].BillAmount.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            //cell = rowData.Cells[4];
                            //cell.AddParagraph(lstMedBillNoPRNoSharing[i].Type);
                            //cell.Format.Font.Bold = true;
                            //cell.Format.Font.Name = "Malgun Gothic";
                            //cell.Format.Font.Size = 7;
                            //cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            //cell.Format.Alignment = ParagraphAlignment.Center;

                            cell = rowData.Cells[4];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleAmount.Value.ToString("C"));
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Right;

                            cell = rowData.Cells[5];
                            cell.AddParagraph(lstMedBillNoPRNoSharing[i].IneligibleReason);
                            cell.Format.Font.Bold = true;
                            cell.Format.Font.Name = "Malgun Gothic";
                            cell.Format.Font.Size = 7;
                            cell.Format.Font.Color = MigraDocDOM.Color.FromCmyk(100, 100, 100, 100);
                            cell.Format.Alignment = ParagraphAlignment.Left;
                        }
                    }
                    pdfPersonalResponsibilityDoc.LastSection.Add(tableIneligibleNoPR);

                }




                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                const bool unicode = true;
                const PdfFontEmbedding embedding = PdfFontEmbedding.Always;

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode, embedding);
                pdfRenderer.Document = pdfPersonalResponsibilityDoc;
                pdfRenderer.RenderDocument();


                if (txtIncidentNo.Text.Trim() != String.Empty)
                {
                    SaveFileDialog savefileDlg = new SaveFileDialog();
                    savefileDlg.FileName = strIndividualID + "_" + strIndividualName + "_" + DateTime.Today.ToString("MM-dd-yyyy") + "_En";
                    savefileDlg.Filter = "PDF Files | *.pdf";
                    savefileDlg.DefaultExt = "pdf";
                    savefileDlg.RestoreDirectory = true;

                    if (savefileDlg.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            pdfRenderer.PdfDocument.Save(savefileDlg.FileName);
                            System.Diagnostics.ProcessStartInfo processInfo = new System.Diagnostics.ProcessStartInfo();
                            processInfo.FileName = savefileDlg.FileName;

                            System.Diagnostics.Process.Start(processInfo);
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(ex.Message, "Error");
                            return;
                        }
                        //finally
                        //{
                        //    ChkInfoEntered = null;
                        //}
                    }
                }


            }
            else
            {
                MessageBox.Show("No table is populated", "Error");
            }
        }

        private void rbCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (rbCheck.Checked)
            {
                // Enable check info controls
                txtACH_No.Text = String.Empty;
                txtACH_No.Enabled = false;
                //txtTransactionDate.Text = String.Empty;
                //txtTransactionDate.Enabled = false;
                dtpACHDate.Text = String.Empty;
                dtpACHDate.Enabled = false;
                txtCreditCardNo.Text = String.Empty;
                txtCreditCardNo.Enabled = false;
                dtpCreditCardPaymentDate.Value = DateTime.Parse(DateTime.Today.ToShortDateString());
                dtpCreditCardPaymentDate.Enabled = false;

                txtIncidentNo.Text = String.Empty;
                txtIncidentNo.Enabled = false;

                txtCheckNo.Enabled = true;
                dtpCheckIssueDate.Enabled = true;
                txtIndividualID.Focus();


                // Disable credit card and ACH info controls
                //txtCreditCardNo.Text = String.Empty;
                //txtCreditCardNo.Enabled = false;
                //dtpCreditCardPaymentDate.Enabled = false;
                //txtACH_No.Text = String.Empty;
                //txtACH_No.Enabled = false;
                //dtpTransactionDate.Enabled = false;
            }
        }

        private void rbCreditCard_CheckedChanged(object sender, EventArgs e)
        {
            if (rbCreditCard.Checked)
            {
                // Enable credit card info controls
                //txtCreditCardNo.Enabled = true;
                //dtpCreditCardPaymentDate.Enabled = true;
                //txtCreditCardNo.Focus();

                // Disable check info and ACH info controls
                txtCheckNo.Text = String.Empty;
                txtCheckNo.Enabled = false;
                //txtCheckIssueDate.Text = String.Empty;
                dtpCheckIssueDate.Text = String.Empty;
                dtpCheckIssueDate.Enabled = false;
                //dtpCheckIssueDate.Enabled = false;
                txtACH_No.Text = String.Empty;
                txtACH_No.Enabled = false;
                //txtTransactionDate.Text = String.Empty;
                //txtTransactionDate.Enabled = false;
                dtpACHDate.Text = String.Empty;
                dtpACHDate.Enabled = false;

                txtIncidentNo.Text = String.Empty;
                txtIncidentNo.Enabled = false;

                //txtCreditCardNo.Text = String.Empty;
                //txtCreditCardNo.Enabled = true;
                dtpCreditCardPaymentDate.Enabled = true;
                txtIndividualID.Focus();

                //dtpTransactionDate.Enabled = false;
            }
        }

        private void rbACH_CheckedChanged(object sender, EventArgs e)
        {
            if (rbACH.Checked)
            {
                // Enable ACH info controls

                // Disable Check info and credit card info controls
                txtCheckNo.Text = String.Empty;
                txtCheckNo.Enabled = false;
                //txtCheckIssueDate.Text = String.Empty;
                //txtCheckIssueDate.Enabled = false;
                dtpCheckIssueDate.Text = String.Empty;
                dtpCheckIssueDate.Enabled = false;

                //dtpCheckIssueDate.Enabled = false;
                txtCreditCardNo.Text = String.Empty;
                txtCreditCardNo.Enabled = false;
                //dtpCreditCardPaymentDate.Value = DateTime.Parse(DateTime.Today.ToShortDateString());
                dtpCreditCardPaymentDate.Enabled = false;
                //dtpCreditCardPaymentDate.Enabled = false;

                txtIncidentNo.Text = String.Empty;
                txtIncidentNo.Enabled = false;

                txtACH_No.Text = String.Empty;
                txtACH_No.Enabled = true;
                dtpACHDate.Enabled = true;
                txtIndividualID.Focus();

            }
        }

        private void SortBillPaidTable(SortedField sf)
        {
            DataTable dtPaidWithSum = (DataTable)gvBillPaid.DataSource;

            DataTable dtSorted = dtPaidWithSum.Clone();

            DataTable dtClone = new DataTable();

            /////////////////////////////////////////

            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));            
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("본인 부담금", typeof(String));
            dtClone.Columns.Add("회원할인", typeof(String));
            dtClone.Columns.Add("CMM 할인", typeof(String));
            dtClone.Columns.Add("의료기관 지불금", typeof(String));
            //if (rbCheck.Checked || rbACH.Checked)
            if (PaidTo == EnumPaidTo.Member)
            {
                dtClone.Columns.Add("기지급액", typeof(String));
                dtClone.Columns.Add("회원 환불금", typeof(String));
            }
            //if (rbCreditCard.Checked)
            if (PaidTo == EnumPaidTo.MedicalProvider)
            {
                dtClone.Columns.Add("기지급액 (의료기관)", typeof(String));
                dtClone.Columns.Add("기지급액 (회원)", typeof(String));
            }
            dtClone.Columns.Add("잔액/보류", typeof(String));

            dtPaidWithSum.Rows.RemoveAt(dtPaidWithSum.Rows.Count - 1);

            foreach(DataRow row in dtPaidWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for(int i = 0; i < dtPaidWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtPaidWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtPaidWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtPaidWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtPaidWithSum.Rows[i]["서비스 날짜"].ToString());
                dtClone.Rows[i]["의료기관명"] = dtPaidWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtPaidWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["본인 부담금"] = dtPaidWithSum.Rows[i]["본인 부담금"];
                dtClone.Rows[i]["회원할인"] = dtPaidWithSum.Rows[i]["회원할인"];
                dtClone.Rows[i]["CMM 할인"] = dtPaidWithSum.Rows[i]["CMM 할인"];
                dtClone.Rows[i]["의료기관 지불금"] = dtPaidWithSum.Rows[i]["의료기관 지불금"];
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    dtClone.Rows[i]["기지급액"] = dtPaidWithSum.Rows[i]["기지급액"];
                    dtClone.Rows[i]["회원 환불금"] = dtPaidWithSum.Rows[i]["회원 환불금"];
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    dtClone.Rows[i]["기지급액 (의료기관)"] = dtPaidWithSum.Rows[i]["기지급액 (의료기관)"];
                    dtClone.Rows[i]["기지급액 (회원)"] = dtPaidWithSum.Rows[i]["기지급액 (회원)"];
                }
                dtClone.Rows[i]["잔액/보류"] = dtPaidWithSum.Rows[i]["잔액/보류"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach (DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["본인 부담금"] = dtCloneSorted.Rows[i]["본인 부담금"];
                dtSorted.Rows[i]["회원할인"] = dtCloneSorted.Rows[i]["회원할인"];
                dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                dtSorted.Rows[i]["의료기관 지불금"] = dtCloneSorted.Rows[i]["의료기관 지불금"];
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    dtSorted.Rows[i]["기지급액"] = dtCloneSorted.Rows[i]["기지급액"];
                    dtSorted.Rows[i]["회원 환불금"] = dtCloneSorted.Rows[i]["회원 환불금"];
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    dtSorted.Rows[i]["기지급액 (의료기관)"] = dtCloneSorted.Rows[i]["기지급액 (의료기관)"];
                    dtSorted.Rows[i]["기지급액 (회원)"] = dtCloneSorted.Rows[i]["기지급액 (회원)"];
                }
                dtSorted.Rows[i]["잔액/보류"] = dtCloneSorted.Rows[i]["잔액/보류"];
            }

            DataRow drBillPaidSumNew = dtSorted.NewRow();

            List<MedicalExpense> lstMedicalExpense = new List<MedicalExpense>();

            DataRow rowSelected = dtSorted.Rows[0];

            foreach (DataRow row in dtSorted.Rows)
            {
                if (PaidTo == EnumPaidTo.Member)
                {
                    //lstMedicalExpense.Add(new MedicalExpense(Double.Parse(row[5].ToString().Substring(1)),
                    //                                         Double.Parse(row[6].ToString().Substring(1)),
                    //                                         Double.Parse(row[7].ToString().Substring(1)),
                    //                                         Double.Parse(row[8].ToString().Substring(1)),
                    //                                         Double.Parse(row[9].ToString().Substring(1)),
                    //                                         0,
                    //                                         Double.Parse(row[10].ToString().Substring(1)),
                    //                                         Double.Parse(row[11].ToString().Substring(1)),
                    //                                         Double.Parse(row[12].ToString().Substring(1))));

                    MedicalExpense expense = new MedicalExpense();
                    expense.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                    expense.PersonalResponsibility = Double.Parse(row[6].ToString().Substring(1));
                    expense.MemberDiscount = Double.Parse(row[7].ToString().Substring(1));
                    expense.CMMDiscount = Double.Parse(row[8].ToString().Substring(1));
                    expense.CMMProviderPayment = Double.Parse(row[9].ToString().Substring(1));
                    expense.PastReimbursement = Double.Parse(row[10].ToString().Substring(1));
                    expense.Reimbursement = Double.Parse(row[11].ToString().Substring(1));
                    expense.Balance = Double.Parse(row[12].ToString().Substring(1));

                    lstMedicalExpense.Add(expense);
                }
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    MedicalExpense expense = new MedicalExpense();
                    expense.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                    expense.PersonalResponsibility = Double.Parse(row[6].ToString().Substring(1));
                    expense.MemberDiscount = Double.Parse(row[7].ToString().Substring(1));
                    expense.CMMDiscount = Double.Parse(row[8].ToString().Substring(1));
                    expense.CMMProviderPayment = Double.Parse(row[9].ToString().Substring(1));
                    expense.PastCMMProviderPayment = Double.Parse(row[10].ToString().Substring(1));
                    expense.PastReimbursement = Double.Parse(row[11].ToString().Substring(1));
                    expense.Balance = Double.Parse(row[12].ToString().Substring(1));

                    lstMedicalExpense.Add(expense);
                }
            }

            drBillPaidSumNew["INCD"] = String.Empty;
            drBillPaidSumNew["MED_BILL"] = String.Empty;
            drBillPaidSumNew["서비스 날짜"] = String.Empty;
            drBillPaidSumNew["의료기관명"] = "합계";

            MedicalExpense sumMedicalExpense = new MedicalExpense();

            foreach (MedicalExpense expense in lstMedicalExpense)
            {
                sumMedicalExpense.BillAmount += expense.BillAmount;
                sumMedicalExpense.MemberDiscount += expense.MemberDiscount;
                sumMedicalExpense.CMMDiscount += expense.CMMDiscount;
                sumMedicalExpense.PersonalResponsibility += expense.PersonalResponsibility;
                sumMedicalExpense.CMMProviderPayment += expense.CMMProviderPayment;
                sumMedicalExpense.PastCMMProviderPayment += expense.PastCMMProviderPayment;
                sumMedicalExpense.PastReimbursement += expense.PastReimbursement;
                sumMedicalExpense.Reimbursement += expense.Reimbursement;
                sumMedicalExpense.Balance += expense.Balance;
            }

            drBillPaidSumNew["청구액(원금)"] = sumMedicalExpense.BillAmount.Value.ToString("C");
            drBillPaidSumNew["본인 부담금"] = sumMedicalExpense.PersonalResponsibility.Value.ToString("C");
            drBillPaidSumNew["회원할인"] = sumMedicalExpense.MemberDiscount.Value.ToString("C");
            drBillPaidSumNew["CMM 할인"] = sumMedicalExpense.CMMDiscount.Value.ToString("C");
            drBillPaidSumNew["의료기관 지불금"] = sumMedicalExpense.CMMProviderPayment.Value.ToString("C");
            //if (rbCheck.Checked || rbACH.Checked)
            if (PaidTo == EnumPaidTo.Member)
            {
                drBillPaidSumNew["기지급액"] = sumMedicalExpense.PastReimbursement.Value.ToString("C");
                drBillPaidSumNew["회원 환불금"] = sumMedicalExpense.Reimbursement.Value.ToString("C");
            }
            //if (rbCreditCard.Checked)
            if (PaidTo == EnumPaidTo.MedicalProvider)
            {
                drBillPaidSumNew["기지급액 (의료기관)"] = sumMedicalExpense.PastCMMProviderPayment.Value.ToString("C");
                drBillPaidSumNew["기지급액 (회원)"] = sumMedicalExpense.PastReimbursement.Value.ToString("C");
            }
            drBillPaidSumNew["잔액/보류"] = sumMedicalExpense.Balance.Value.ToString("C");

            dtSorted.Rows.Add(drBillPaidSumNew);
            gvBillPaid.DataSource = dtSorted;
        }

        private void SortBillPaidTableInPaidTab(SortedField sf)
        {

            DataTable dtPaidWithSum = (DataTable)gvPaidInTabPaid.DataSource;

            DataTable dtSorted = dtPaidWithSum.Clone();

            DataTable dtClone = new DataTable();

            /////////////////////////////////////////

            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("본인 부담금", typeof(String));
            dtClone.Columns.Add("회원할인", typeof(String));
            dtClone.Columns.Add("CMM 할인", typeof(String));
            dtClone.Columns.Add("의료기관 지불금", typeof(String));
            //if (rbCheck.Checked || rbACH.Checked)
            if (PaidTo == EnumPaidTo.Member)
            {
                dtClone.Columns.Add("기지급액", typeof(String));
                dtClone.Columns.Add("회원 환불금", typeof(String));
            }
            //if (rbCreditCard.Checked)
            if (PaidTo == EnumPaidTo.MedicalProvider)
            {
                dtClone.Columns.Add("기지급액 (의료기관)", typeof(String));
                dtClone.Columns.Add("기지급액 (회원)", typeof(String));
            }
            dtClone.Columns.Add("잔액/보류", typeof(String));

            dtPaidWithSum.Rows.RemoveAt(dtPaidWithSum.Rows.Count - 1);

            foreach (DataRow row in dtPaidWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for (int i = 0; i < dtPaidWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtPaidWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtPaidWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtPaidWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtPaidWithSum.Rows[i]["서비스 날짜"].ToString());
                dtClone.Rows[i]["의료기관명"] = dtPaidWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtPaidWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["본인 부담금"] = dtPaidWithSum.Rows[i]["본인 부담금"];
                dtClone.Rows[i]["회원할인"] = dtPaidWithSum.Rows[i]["회원할인"];
                dtClone.Rows[i]["CMM 할인"] = dtPaidWithSum.Rows[i]["CMM 할인"];
                dtClone.Rows[i]["의료기관 지불금"] = dtPaidWithSum.Rows[i]["의료기관 지불금"];
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    dtClone.Rows[i]["기지급액"] = dtPaidWithSum.Rows[i]["기지급액"];
                    dtClone.Rows[i]["회원 환불금"] = dtPaidWithSum.Rows[i]["회원 환불금"];
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    dtClone.Rows[i]["기지급액 (의료기관)"] = dtPaidWithSum.Rows[i]["기지급액 (의료기관)"];
                    dtClone.Rows[i]["기지급액 (회원)"] = dtPaidWithSum.Rows[i]["기지급액 (회원)"];
                }
                dtClone.Rows[i]["잔액/보류"] = dtPaidWithSum.Rows[i]["잔액/보류"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach (DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["본인 부담금"] = dtCloneSorted.Rows[i]["본인 부담금"];
                dtSorted.Rows[i]["회원할인"] = dtCloneSorted.Rows[i]["회원할인"];
                dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                dtSorted.Rows[i]["의료기관 지불금"] = dtCloneSorted.Rows[i]["의료기관 지불금"];
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    dtSorted.Rows[i]["기지급액"] = dtCloneSorted.Rows[i]["기지급액"];
                    dtSorted.Rows[i]["회원 환불금"] = dtCloneSorted.Rows[i]["회원 환불금"];
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    dtSorted.Rows[i]["기지급액 (의료기관)"] = dtCloneSorted.Rows[i]["기지급액 (의료기관)"];
                    dtSorted.Rows[i]["기지급액 (회원)"] = dtCloneSorted.Rows[i]["기지급액 (회원)"];
                }
                dtSorted.Rows[i]["잔액/보류"] = dtCloneSorted.Rows[i]["잔액/보류"];
            }

            DataRow drBillPaidSumNew = dtSorted.NewRow();

            List<MedicalExpense> lstMedicalExpense = new List<MedicalExpense>();

            foreach (DataRow row in dtSorted.Rows)
            {
                //if (rbCheck.Checked || rbACH.Checked)
                if (PaidTo == EnumPaidTo.Member)
                {
                    //lstMedicalExpense.Add(new MedicalExpense(Double.Parse(row[5].ToString().Substring(1)),
                    //                                         Double.Parse(row[6].ToString().Substring(1)),
                    //                                         Double.Parse(row[7].ToString().Substring(1)),
                    //                                         Double.Parse(row[8].ToString().Substring(1)),
                    //                                         Double.Parse(row[9].ToString().Substring(1)),
                    //                                         0,
                    //                                         Double.Parse(row[10].ToString().Substring(1)),
                    //                                         Double.Parse(row[11].ToString().Substring(1)),
                    //                                         Double.Parse(row[12].ToString().Substring(1))));

                    MedicalExpense expense = new MedicalExpense();
                    expense.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                    expense.PersonalResponsibility = Double.Parse(row[6].ToString().Substring(1));
                    expense.MemberDiscount = Double.Parse(row[7].ToString().Substring(1));
                    expense.CMMDiscount = Double.Parse(row[8].ToString().Substring(1));
                    expense.CMMProviderPayment = Double.Parse(row[9].ToString().Substring(1));
                    expense.PastReimbursement = Double.Parse(row[10].ToString().Substring(1));
                    expense.Reimbursement = Double.Parse(row[11].ToString().Substring(1));
                    expense.Balance = Double.Parse(row[12].ToString().Substring(1));

                    lstMedicalExpense.Add(expense);
                }
                //if (rbCreditCard.Checked)
                if (PaidTo == EnumPaidTo.MedicalProvider)
                {
                    //lstMedicalExpense.Add(new MedicalExpense(Double.Parse(row[5].ToString().Substring(1)),
                    //                                         Double.Parse(row[6].ToString().Substring(1)),
                    //                                         Double.Parse(row[7].ToString().Substring(1)),
                    //                                         Double.Parse(row[8].ToString().Substring(1)),
                    //                                         Double.Parse(row[9].ToString().Substring(1)),
                    //                                         Double.Parse(row[10].ToString().Substring(1)),
                    //                                         Double.Parse(row[11].ToString().Substring(1)),
                    //                                         0,
                    //                                         Double.Parse(row[12].ToString().Substring(1))));

                    MedicalExpense expense = new MedicalExpense();
                    expense.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                    expense.PersonalResponsibility = Double.Parse(row[6].ToString().Substring(1));
                    expense.MemberDiscount = Double.Parse(row[7].ToString().Substring(1));
                    expense.CMMDiscount = Double.Parse(row[8].ToString().Substring(1));
                    expense.CMMProviderPayment = Double.Parse(row[9].ToString().Substring(1));
                    expense.PastCMMProviderPayment = Double.Parse(row[10].ToString().Substring(1));
                    expense.PastReimbursement = Double.Parse(row[11].ToString().Substring(1));
                    expense.Balance = Double.Parse(row[12].ToString().Substring(1));

                    lstMedicalExpense.Add(expense);

                }
            }

            drBillPaidSumNew["INCD"] = String.Empty;
            drBillPaidSumNew["MED_BILL"] = String.Empty;
            drBillPaidSumNew["서비스 날짜"] = String.Empty;
            drBillPaidSumNew["의료기관명"] = "합계";

            MedicalExpense sumMedicalExpense = new MedicalExpense();

            foreach (MedicalExpense expense in lstMedicalExpense)
            {
                sumMedicalExpense.BillAmount += expense.BillAmount;
                sumMedicalExpense.PersonalResponsibility += expense.PersonalResponsibility;
                sumMedicalExpense.MemberDiscount += expense.MemberDiscount;
                sumMedicalExpense.CMMDiscount += expense.CMMDiscount;
                sumMedicalExpense.CMMProviderPayment += expense.CMMProviderPayment;
                sumMedicalExpense.PastCMMProviderPayment += expense.PastCMMProviderPayment;
                sumMedicalExpense.PastReimbursement += expense.PastReimbursement;
                sumMedicalExpense.Reimbursement += expense.Reimbursement;
                sumMedicalExpense.Balance += expense.Balance;
            }

            drBillPaidSumNew["청구액(원금)"] = sumMedicalExpense.BillAmount.Value.ToString("C");
            drBillPaidSumNew["본인 부담금"] = sumMedicalExpense.PersonalResponsibility.Value.ToString("C");
            drBillPaidSumNew["회원할인"] = sumMedicalExpense.MemberDiscount.Value.ToString("C");
            drBillPaidSumNew["CMM 할인"] = sumMedicalExpense.CMMDiscount.Value.ToString("C");
            drBillPaidSumNew["의료기관 지불금"] = sumMedicalExpense.CMMProviderPayment.Value.ToString("C");
            //if (rbCheck.Checked || rbACH.Checked)
            if (PaidTo == EnumPaidTo.Member)
            {
                drBillPaidSumNew["기지급액"] = sumMedicalExpense.PastReimbursement.Value.ToString("C");
                drBillPaidSumNew["회원 환불금"] = sumMedicalExpense.Reimbursement.Value.ToString("C");
            }
            //if (rbCreditCard.Checked)
            if (PaidTo == EnumPaidTo.MedicalProvider)
            {
                drBillPaidSumNew["기지급액 (의료기관)"] = sumMedicalExpense.PastCMMProviderPayment.Value.ToString("C");
                drBillPaidSumNew["기지급액 (회원)"] = sumMedicalExpense.PastReimbursement.Value.ToString("C");
            }
            drBillPaidSumNew["잔액/보류"] = sumMedicalExpense.Balance.Value.ToString("C");

            dtSorted.Rows.Add(drBillPaidSumNew);
            gvPaidInTabPaid.DataSource = dtSorted;
        }

        private void SortCMMPendingPaymentTable(SortedField sf)
        {
            DataTable dtCMMPendingPaymentWithSum = (DataTable)gvCMMPendingPayment.DataSource;

            DataTable dtSorted = dtCMMPendingPaymentWithSum.Clone();

            DataTable dtClone = new DataTable();
            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            //dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("회원할인", typeof(String));
            dtClone.Columns.Add("CMM 할인", typeof(String));
            dtClone.Columns.Add("본인 부담금", typeof(String));
            dtClone.Columns.Add("정산 완료", typeof(String));
            dtClone.Columns.Add("지원 예정", typeof(String));

            dtCMMPendingPaymentWithSum.Rows.RemoveAt(dtCMMPendingPaymentWithSum.Rows.Count - 1);

            foreach(DataRow row in dtCMMPendingPaymentWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for (int i = 0; i < dtCMMPendingPaymentWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtCMMPendingPaymentWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtCMMPendingPaymentWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtCMMPendingPaymentWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCMMPendingPaymentWithSum.Rows[i]["서비스 날짜"].ToString());

                //String strDate = dtCMMPendingPaymentWithSum.Rows[i]["접수 날짜"].ToString();

                //if (dtCMMPendingPaymentWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                //{
                //    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtCMMPendingPaymentWithSum.Rows[i]["접수 날짜"].ToString());
                //}
                //else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtCMMPendingPaymentWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtCMMPendingPaymentWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["회원할인"] = dtCMMPendingPaymentWithSum.Rows[i]["회원할인"];
                dtClone.Rows[i]["CMM 할인"] = dtCMMPendingPaymentWithSum.Rows[i]["CMM 할인"];
                dtClone.Rows[i]["본인 부담금"] = dtCMMPendingPaymentWithSum.Rows[i]["본인 부담금"];
                dtClone.Rows[i]["정산 완료"] = dtCMMPendingPaymentWithSum.Rows[i]["정산 완료"];
                dtClone.Rows[i]["지원 예정"] = dtCMMPendingPaymentWithSum.Rows[i]["지원 예정"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach(DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                //if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                //{
                //    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                //}
                //else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["회원할인"] = dtCloneSorted.Rows[i]["회원할인"];
                dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                dtSorted.Rows[i]["본인 부담금"] = dtCloneSorted.Rows[i]["본인 부담금"];
                dtSorted.Rows[i]["정산 완료"] = dtCloneSorted.Rows[i]["정산 완료"];
                dtSorted.Rows[i]["지원 예정"] = dtCloneSorted.Rows[i]["지원 예정"];
            }

            DataRow drCMMPendingPaymentSumNew = dtSorted.NewRow();

            List<CMMPendingPayment> lstCMMPendingPayment = new List<CMMPendingPayment>();

            foreach(DataRow row in dtSorted.Rows)
            {
                //lstCMMPendingPayment.Add(new CMMPendingPayment(Double.Parse(row[5].ToString().Substring(1)),
                //                                               Double.Parse(row[6].ToString().Substring(1)),
                //                                               Double.Parse(row[7].ToString().Substring(1)),
                //                                               Double.Parse(row[8].ToString().Substring(1)),
                //                                               Double.Parse(row[9].ToString().Substring(1)),
                //                                               Double.Parse(row[10].ToString().Substring(1))));

                CMMPendingPayment cmm_pending_payment = new CMMPendingPayment();
                cmm_pending_payment.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                cmm_pending_payment.MemberDiscount = Double.Parse(row[6].ToString().Substring(1));
                cmm_pending_payment.CMMDiscount = Double.Parse(row[7].ToString().Substring(1));
                cmm_pending_payment.PersonalResponsibility = Double.Parse(row[8].ToString().Substring(1));
                cmm_pending_payment.SharedAmount = Double.Parse(row[9].ToString().Substring(1));
                cmm_pending_payment.AmountWillBeShared = Double.Parse(row[10].ToString().Substring(1));

                lstCMMPendingPayment.Add(cmm_pending_payment);
            }

            drCMMPendingPaymentSumNew["INCD"] = String.Empty;
            drCMMPendingPaymentSumNew["MED_BILL"] = String.Empty;
            drCMMPendingPaymentSumNew["의료기관명"] = "합계";

            CMMPendingPayment sumCMMPendingPayment = new CMMPendingPayment();

            foreach(CMMPendingPayment cmm_pending in lstCMMPendingPayment)
            {
                sumCMMPendingPayment.BillAmount += cmm_pending.BillAmount;
                sumCMMPendingPayment.MemberDiscount += cmm_pending.MemberDiscount;
                sumCMMPendingPayment.CMMDiscount += cmm_pending.CMMDiscount;
                sumCMMPendingPayment.PersonalResponsibility += cmm_pending.PersonalResponsibility;
                sumCMMPendingPayment.SharedAmount += cmm_pending.SharedAmount;
                sumCMMPendingPayment.AmountWillBeShared += cmm_pending.AmountWillBeShared;
            }

            drCMMPendingPaymentSumNew["청구액(원금)"] = sumCMMPendingPayment.BillAmount.Value.ToString("C");
            drCMMPendingPaymentSumNew["회원할인"] = sumCMMPendingPayment.MemberDiscount.Value.ToString("C");
            drCMMPendingPaymentSumNew["CMM 할인"] = sumCMMPendingPayment.CMMDiscount.Value.ToString("C");
            drCMMPendingPaymentSumNew["본인 부담금"] = sumCMMPendingPayment.PersonalResponsibility.Value.ToString("C");
            drCMMPendingPaymentSumNew["정산 완료"] = sumCMMPendingPayment.SharedAmount.Value.ToString("C");
            drCMMPendingPaymentSumNew["지원 예정"] = sumCMMPendingPayment.AmountWillBeShared.Value.ToString("C");

            dtSorted.Rows.Add(drCMMPendingPaymentSumNew);
            gvCMMPendingPayment.DataSource = dtSorted;
            
        }

        private void SortCMMPendingPaymentTableInTab(SortedField sf)
        {
            DataTable dtCMMPendingPaymentWithSum = (DataTable)gvCMMPendingInTab.DataSource;

            DataTable dtSorted = dtCMMPendingPaymentWithSum.Clone();

            DataTable dtClone = new DataTable();
            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            //dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("회원할인", typeof(String));
            dtClone.Columns.Add("CMM 할인", typeof(String));
            dtClone.Columns.Add("본인 부담금", typeof(String));
            dtClone.Columns.Add("정산 완료", typeof(String));
            dtClone.Columns.Add("지원 예정", typeof(String));

            dtCMMPendingPaymentWithSum.Rows.RemoveAt(dtCMMPendingPaymentWithSum.Rows.Count - 1);

            foreach (DataRow row in dtCMMPendingPaymentWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for (int i = 0; i < dtCMMPendingPaymentWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtCMMPendingPaymentWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtCMMPendingPaymentWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtCMMPendingPaymentWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCMMPendingPaymentWithSum.Rows[i]["서비스 날짜"].ToString());
                //if (dtCMMPendingPaymentWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                //{
                //    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtCMMPendingPaymentWithSum.Rows[i]["접수 날짜"].ToString());
                //}
                //else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtCMMPendingPaymentWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtCMMPendingPaymentWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["회원할인"] = dtCMMPendingPaymentWithSum.Rows[i]["회원할인"];
                dtClone.Rows[i]["CMM 할인"] = dtCMMPendingPaymentWithSum.Rows[i]["CMM 할인"];
                dtClone.Rows[i]["본인 부담금"] = dtCMMPendingPaymentWithSum.Rows[i]["본인 부담금"];
                dtClone.Rows[i]["정산 완료"] = dtCMMPendingPaymentWithSum.Rows[i]["정산 완료"];
                dtClone.Rows[i]["지원 예정"] = dtCMMPendingPaymentWithSum.Rows[i]["지원 예정"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach (DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                //if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                //{
                //    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                //}
                //else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["회원할인"] = dtCloneSorted.Rows[i]["회원할인"];
                dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                dtSorted.Rows[i]["본인 부담금"] = dtCloneSorted.Rows[i]["본인 부담금"];
                dtSorted.Rows[i]["정산 완료"] = dtCloneSorted.Rows[i]["정산 완료"];
                dtSorted.Rows[i]["지원 예정"] = dtCloneSorted.Rows[i]["지원 예정"];
            }

            DataRow drCMMPendingPaymentSumNew = dtSorted.NewRow();

            List<CMMPendingPayment> lstCMMPendingPayment = new List<CMMPendingPayment>();

            foreach (DataRow row in dtSorted.Rows)
            {
                //lstCMMPendingPayment.Add(new CMMPendingPayment(Double.Parse(row[5].ToString().Substring(1)),
                //                                               Double.Parse(row[6].ToString().Substring(1)),
                //                                               Double.Parse(row[7].ToString().Substring(1)),
                //                                               Double.Parse(row[8].ToString().Substring(1)),
                //                                               Double.Parse(row[9].ToString().Substring(1)),
                //                                               Double.Parse(row[10].ToString().Substring(1))));

                CMMPendingPayment cmm_pending_payment = new CMMPendingPayment();
                cmm_pending_payment.BillAmount = Double.Parse(row[5].ToString().Substring(1));
                cmm_pending_payment.MemberDiscount = Double.Parse(row[6].ToString().Substring(1));
                cmm_pending_payment.CMMDiscount = Double.Parse(row[7].ToString().Substring(1));
                cmm_pending_payment.PersonalResponsibility = Double.Parse(row[8].ToString().Substring(1));
                cmm_pending_payment.SharedAmount = Double.Parse(row[9].ToString().Substring(1));
                cmm_pending_payment.AmountWillBeShared = Double.Parse(row[10].ToString().Substring(1));

                lstCMMPendingPayment.Add(cmm_pending_payment);                
            }

            drCMMPendingPaymentSumNew["INCD"] = String.Empty;
            drCMMPendingPaymentSumNew["MED_BILL"] = String.Empty;
            drCMMPendingPaymentSumNew["의료기관명"] = "합계";

            CMMPendingPayment sumCMMPendingPayment = new CMMPendingPayment();

            foreach (CMMPendingPayment cmm_pending in lstCMMPendingPayment)
            {
                sumCMMPendingPayment.BillAmount += cmm_pending.BillAmount;
                sumCMMPendingPayment.MemberDiscount += cmm_pending.MemberDiscount;
                sumCMMPendingPayment.CMMDiscount += cmm_pending.CMMDiscount;
                sumCMMPendingPayment.PersonalResponsibility += cmm_pending.PersonalResponsibility;
                sumCMMPendingPayment.SharedAmount += cmm_pending.SharedAmount;
                sumCMMPendingPayment.AmountWillBeShared += cmm_pending.AmountWillBeShared;
            }

            drCMMPendingPaymentSumNew["청구액(원금)"] = sumCMMPendingPayment.BillAmount.Value.ToString("C");
            drCMMPendingPaymentSumNew["회원할인"] = sumCMMPendingPayment.MemberDiscount.Value.ToString("C");
            drCMMPendingPaymentSumNew["CMM 할인"] = sumCMMPendingPayment.CMMDiscount.Value.ToString("C");
            drCMMPendingPaymentSumNew["본인 부담금"] = sumCMMPendingPayment.PersonalResponsibility.Value.ToString("C");
            drCMMPendingPaymentSumNew["정산 완료"] = sumCMMPendingPayment.SharedAmount.Value.ToString("C");
            drCMMPendingPaymentSumNew["지원 예정"] = sumCMMPendingPayment.AmountWillBeShared.Value.ToString("C");

            dtSorted.Rows.Add(drCMMPendingPaymentSumNew);
            gvCMMPendingInTab.DataSource = dtSorted;
        }

        private void SortPendingTable(SortedField sf)
        {
            DataTable dtPendingWithSum = (DataTable)gvPending.DataSource;

            DataTable dtSorted = dtPendingWithSum.Clone();

            DataTable dtClone = new DataTable();

            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            //dtClone.Columns.Add("회원 할인", typeof(String));
            //dtClone.Columns.Add("CMM 할인", typeof(String));
            //dtClone.Columns.Add("정산 완료", typeof(String));
            dtClone.Columns.Add("잔액/보류", typeof(String));
            dtClone.Columns.Add("보류 사유", typeof(String));

            dtPendingWithSum.Rows.RemoveAt(dtPendingWithSum.Rows.Count - 1);

            foreach (DataRow row in dtPendingWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for(int i = 0; i < dtPendingWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtPendingWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtPendingWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtPendingWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtPendingWithSum.Rows[i]["서비스 날짜"].ToString());
                if (dtPendingWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtPendingWithSum.Rows[i]["접수 날짜"].ToString());
                }
                else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtPendingWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtPendingWithSum.Rows[i]["청구액(원금)"];
                //dtClone.Rows[i]["회원 할인"] = dtPendingWithSum.Rows[i]["회원 할인"];
                //dtClone.Rows[i]["CMM 할인"] = dtPendingWithSum.Rows[i]["CMM 할인"];
                //dtClone.Rows[i]["정산 완료"] = dtPendingWithSum.Rows[i]["정산 완료"];
                dtClone.Rows[i]["잔액/보류"] = dtPendingWithSum.Rows[i]["잔액/보류"];
                dtClone.Rows[i]["보류 사유"] = dtPendingWithSum.Rows[i]["보류 사유"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach(DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                }
                else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["잔액/보류"] = dtCloneSorted.Rows[i]["잔액/보류"];
                //dtSorted.Rows[i]["회원 할인"] = dtCloneSorted.Rows[i]["회원 할인"];
                //dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                //dtSorted.Rows[i]["정산 완료"] = dtCloneSorted.Rows[i]["정산 완료"];
                //dtSorted.Rows[i]["보류"] = dtCloneSorted.Rows[i]["보류"];
                dtSorted.Rows[i]["보류 사유"] = dtCloneSorted.Rows[i]["보류 사유"];
            }

            //dtSorted = dtPendingWithSum.DefaultView.ToTable();
            DataRow drPendingSumNew = dtSorted.NewRow();
            List<Pending> lstPending = new List<Pending>();

            foreach(DataRow row in dtSorted.Rows)
            {
                //lstPending.Add(new Pending(Double.Parse(row[6].ToString().Substring(1)),
                //                           Double.Parse(row[7].ToString().Substring(1)), 0, 0, 0, 0));
                Pending pending = new Pending();
                pending.BillAmount = Double.Parse(row[6].ToString().Substring(1));
                pending.Balance = Double.Parse(row[7].ToString().Substring(1));
                pending.MemberDiscount = 0;
                pending.CMMDiscount = 0;
                pending.SharedAmount = 0;
                pending.PendingAmount = 0;

                lstPending.Add(pending);
            }

            drPendingSumNew["INCD"] = String.Empty;
            drPendingSumNew["MED_BILL"] = String.Empty;
            //drPendingSumNew["서비스 날짜"] = String.Empty;
            //drPendingSumNew["접수 날짜"] = String.Empty;
            drPendingSumNew["의료기관명"] = "합계";
            drPendingSumNew["보류 사유"] = String.Empty;

            Pending sumPending = new Pending();
            foreach (Pending pending in lstPending)
            {
                sumPending.BillAmount += pending.BillAmount;
                sumPending.Balance += pending.Balance;
                //sumPending.MemberDiscount += pending.MemberDiscount;
                //sumPending.CMMDiscount += pending.CMMDiscount;
                //sumPending.SharedAmount += pending.SharedAmount;
                //sumPending.PendingAmount += pending.PendingAmount;
            }

            drPendingSumNew["청구액(원금)"] = sumPending.BillAmount.Value.ToString("C");
            //drPendingSumNew["회원 할인"] = sumPending.MemberDiscount.Value.ToString("C");
            //drPendingSumNew["CMM 할인"] = sumPending.CMMDiscount.Value.ToString("C");
            //drPendingSumNew["정산 완료"] = sumPending.SharedAmount.Value.ToString("C");
            //drPendingSumNew["잔액/보류"] = sumPending.PendingAmount.Value.ToString("C");
            drPendingSumNew["잔액/보류"] = sumPending.Balance.Value.ToString("C");

            dtSorted.Rows.Add(drPendingSumNew);
            gvPending.DataSource = dtSorted;
        }

        private void SortPendingTableInTab(SortedField sf)
        {
            DataTable dtPendingWithSum = (DataTable)gvPendingInTab.DataSource;

            DataTable dtSorted = dtPendingWithSum.Clone();

            DataTable dtClone = new DataTable();

            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            //dtClone.Columns.Add("회원 할인", typeof(String));
            //dtClone.Columns.Add("CMM 할인", typeof(String));
            //dtClone.Columns.Add("정산 완료", typeof(String));
            dtClone.Columns.Add("잔액/보류", typeof(String));
            dtClone.Columns.Add("보류 사유", typeof(String));

            dtPendingWithSum.Rows.RemoveAt(dtPendingWithSum.Rows.Count - 1);

            foreach (DataRow row in dtPendingWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for (int i = 0; i < dtPendingWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtPendingWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtPendingWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtPendingWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtPendingWithSum.Rows[i]["서비스 날짜"].ToString());
                if (dtPendingWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtPendingWithSum.Rows[i]["접수 날짜"].ToString());
                }
                else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtPendingWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtPendingWithSum.Rows[i]["청구액(원금)"];
                //dtClone.Rows[i]["회원 할인"] = dtPendingWithSum.Rows[i]["회원 할인"];
                //dtClone.Rows[i]["CMM 할인"] = dtPendingWithSum.Rows[i]["CMM 할인"];
                //dtClone.Rows[i]["정산 완료"] = dtPendingWithSum.Rows[i]["정산 완료"];
                dtClone.Rows[i]["잔액/보류"] = dtPendingWithSum.Rows[i]["잔액/보류"];
                dtClone.Rows[i]["보류 사유"] = dtPendingWithSum.Rows[i]["보류 사유"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach (DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                }
                else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                //dtSorted.Rows[i]["회원 할인"] = dtCloneSorted.Rows[i]["회원 할인"];
                //dtSorted.Rows[i]["CMM 할인"] = dtCloneSorted.Rows[i]["CMM 할인"];
                //dtSorted.Rows[i]["정산 완료"] = dtCloneSorted.Rows[i]["정산 완료"];
                dtSorted.Rows[i]["잔액/보류"] = dtCloneSorted.Rows[i]["잔액/보류"];
                dtSorted.Rows[i]["보류 사유"] = dtCloneSorted.Rows[i]["보류 사유"];
            }

            //dtSorted = dtPendingWithSum.DefaultView.ToTable();
            DataRow drPendingSumNew = dtSorted.NewRow();
            List<Pending> lstPending = new List<Pending>();

            foreach (DataRow row in dtSorted.Rows)
            {

                //lstPending.Add(new Pending(Double.Parse(row[6].ToString().Substring(1)),
                //                           Double.Parse(row[7].ToString().Substring(1)), 0, 0, 0, 0));

                Pending pending = new Pending();
                pending.BillAmount = Double.Parse(row[6].ToString().Substring(1));
                pending.Balance = Double.Parse(row[7].ToString().Substring(1));
                pending.MemberDiscount = 0;
                pending.CMMDiscount = 0;
                pending.SharedAmount = 0;
                pending.PendingAmount = 0;

                lstPending.Add(pending);
            }

            drPendingSumNew["INCD"] = String.Empty;
            drPendingSumNew["MED_BILL"] = String.Empty;
            //drPendingSumNew["서비스 날짜"] = String.Empty;
            //drPendingSumNew["접수 날짜"] = String.Empty;
            drPendingSumNew["의료기관명"] = "합계";
            drPendingSumNew["보류 사유"] = String.Empty;

            Pending sumPending = new Pending();
            foreach (Pending pending in lstPending)
            {
                sumPending.BillAmount += pending.BillAmount;
                sumPending.Balance += pending.Balance;
                //sumPending.MemberDiscount += pending.MemberDiscount;
                //sumPending.CMMDiscount += pending.CMMDiscount;
                //sumPending.SharedAmount += pending.SharedAmount;
                //sumPending.PendingAmount += pending.PendingAmount;
            }

            drPendingSumNew["청구액(원금)"] = sumPending.BillAmount.Value.ToString("C");
            drPendingSumNew["잔액/보류"] = sumPending.Balance.Value.ToString("C");
            //drPendingSumNew["회원 할인"] = sumPending.MemberDiscount.Value.ToString("C");
            //drPendingSumNew["CMM 할인"] = sumPending.CMMDiscount.Value.ToString("C");
            //drPendingSumNew["정산 완료"] = sumPending.SharedAmount.Value.ToString("C");
            //drPendingSumNew["잔액/보류"] = sumPending.PendingAmount.Value.ToString("C");

            dtSorted.Rows.Add(drPendingSumNew);
            gvPendingInTab.DataSource = dtSorted;
        }

        private void SortIneligibleTable(SortedField sf)
        {
            DataTable dtIneligibleWithSum = (DataTable)gvIneligible.DataSource;

            DataTable dtSorted = dtIneligibleWithSum.Clone();

            DataTable dtClone = new DataTable();
            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            //dtSorted.Columns.Add("서비스 날짜", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("전액/일부 지원불가 금액", typeof(String));
            dtClone.Columns.Add("지원되지않는 사유", typeof(String));

            dtIneligibleWithSum.Rows.RemoveAt(dtIneligibleWithSum.Rows.Count - 1);

            foreach (DataRow row in dtIneligibleWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for(int i = 0; i < dtIneligibleWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtIneligibleWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtIneligibleWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtIneligibleWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtIneligibleWithSum.Rows[i]["서비스 날짜"].ToString());
                if (dtIneligibleWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtIneligibleWithSum.Rows[i]["접수 날짜"].ToString());
                }
                else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtIneligibleWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtIneligibleWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["전액/일부 지원불가 금액"] = dtIneligibleWithSum.Rows[i]["전액/일부 지원불가 금액"];
                dtClone.Rows[i]["지원되지않는 사유"] = dtIneligibleWithSum.Rows[i]["지원되지않는 사유"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    //dtIneligibleWithSum.DefaultView.Sort = sf.Field + " ASC";
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    //dtIneligibleWithSum.DefaultView.Sort = sf.Field + " ASC";
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    //dtIneligibleWithSum.DefaultView.Sort = sf.Field + " DESC";
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach(DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();

                dtSorted.Rows.Add(rowSorted);
            }

            for(int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                }
                else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["전액/일부 지원불가 금액"] = dtCloneSorted.Rows[i]["전액/일부 지원불가 금액"];
                dtSorted.Rows[i]["지원되지않는 사유"] = dtCloneSorted.Rows[i]["지원되지않는 사유"];

            }

            DataRow drIneligibleSumNew = dtSorted.NewRow();

            List<MedicalExpenseIneligible> lstIneligible = new List<MedicalExpenseIneligible>();

            foreach (DataRow row in dtSorted.Rows)
            {
                //lstIneligible.Add(new MedicalExpenseIneligible(Double.Parse(row[6].ToString().Substring(1)),
                //                                               Double.Parse(row[7].ToString().Substring(1))));

                MedicalExpenseIneligible ineligible = new MedicalExpenseIneligible();
                ineligible.BillAmount = Double.Parse(row[6].ToString().Substring(1));
                ineligible.AmountIneligible = Double.Parse(row[7].ToString().Substring(1));

                lstIneligible.Add(ineligible);
            }

            drIneligibleSumNew["INCD"] = String.Empty;
            drIneligibleSumNew["MED_BILL"] = String.Empty;
            //drIneligibleSumNew["서비스 날짜"] = String.Empty;
            drIneligibleSumNew["의료기관명"] = "합계";
            drIneligibleSumNew["지원되지않는 사유"] = String.Empty;

            MedicalExpenseIneligible sumIneligible = new MedicalExpenseIneligible();

            foreach (MedicalExpenseIneligible expense in lstIneligible)
            {
                sumIneligible.BillAmount += expense.BillAmount;
                sumIneligible.AmountIneligible += expense.AmountIneligible;
            }

            drIneligibleSumNew["청구액(원금)"] = sumIneligible.BillAmount.Value.ToString("C");
            drIneligibleSumNew["전액/일부 지원불가 금액"] = sumIneligible.AmountIneligible.Value.ToString("C");

            dtSorted.Rows.Add(drIneligibleSumNew);
            gvIneligible.DataSource = dtSorted;

        }

        private void SortIneligibleTableInTab(SortedField sf)
        {
            DataTable dtIneligibleWithSum = (DataTable)gvIneligibleInTab.DataSource;
          
            DataTable dtSorted = dtIneligibleWithSum.Clone();

            DataTable dtClone = new DataTable();
            dtClone.Columns.Add("INCD", typeof(String));
            dtClone.Columns.Add("회원 이름", typeof(String));
            dtClone.Columns.Add("MED_BILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("접수 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("전액/일부 지원불가 금액", typeof(String));
            dtClone.Columns.Add("지원되지않는 사유", typeof(String));

            dtIneligibleWithSum.Rows.RemoveAt(dtIneligibleWithSum.Rows.Count - 1);

            foreach (DataRow row in dtIneligibleWithSum.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for(int i = 0; i < dtIneligibleWithSum.Rows.Count; i++)
            {
                dtClone.Rows[i]["INCD"] = dtIneligibleWithSum.Rows[i]["INCD"];
                dtClone.Rows[i]["회원 이름"] = dtIneligibleWithSum.Rows[i]["회원 이름"];
                dtClone.Rows[i]["MED_BILL"] = dtIneligibleWithSum.Rows[i]["MED_BILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtIneligibleWithSum.Rows[i]["서비스 날짜"].ToString());
                if (dtIneligibleWithSum.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtClone.Rows[i]["접수 날짜"] = DateTime.Parse(dtIneligibleWithSum.Rows[i]["접수 날짜"].ToString());
                }
                else dtClone.Rows[i]["접수 날짜"] = DBNull.Value;
                dtClone.Rows[i]["의료기관명"] = dtIneligibleWithSum.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtIneligibleWithSum.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["전액/일부 지원불가 금액"] = dtIneligibleWithSum.Rows[i]["전액/일부 지원불가 금액"];
                dtClone.Rows[i]["지원되지않는 사유"] = dtIneligibleWithSum.Rows[i]["지원되지않는 사유"];
            }

            switch (sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach(DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();

                dtSorted.Rows.Add(rowSorted);
            }

            for(int i = 0; i < dtCloneSorted.Rows.Count; i ++)
            {
                dtSorted.Rows[i]["INCD"] = dtCloneSorted.Rows[i]["INCD"];
                dtSorted.Rows[i]["회원 이름"] = dtCloneSorted.Rows[i]["회원 이름"];
                dtSorted.Rows[i]["MED_BILL"] = dtCloneSorted.Rows[i]["MED_BILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                if (dtCloneSorted.Rows[i]["접수 날짜"].ToString() != String.Empty)
                {
                    dtSorted.Rows[i]["접수 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["접수 날짜"].ToString()).ToString("MM/dd/yyyy");
                }
                else dtSorted.Rows[i]["접수 날짜"] = DBNull.Value;
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["전액/일부 지원불가 금액"] = dtCloneSorted.Rows[i]["전액/일부 지원불가 금액"];
                dtSorted.Rows[i]["지원되지않는 사유"] = dtCloneSorted.Rows[i]["지원되지않는 사유"];
            }

            DataRow drIneligibleSumNew = dtSorted.NewRow();

            List<MedicalExpenseIneligible> lstIneligible = new List<MedicalExpenseIneligible>();

            foreach (DataRow row in dtSorted.Rows)
            {
                //lstIneligible.Add(new MedicalExpenseIneligible(Double.Parse(row[6].ToString().Substring(1)),
                //                                               Double.Parse(row[7].ToString().Substring(1))));

                MedicalExpenseIneligible ineligible = new MedicalExpenseIneligible();
                ineligible.BillAmount = Double.Parse(row[6].ToString().Substring(1));
                ineligible.AmountIneligible = Double.Parse(row[7].ToString().Substring(1));

                lstIneligible.Add(ineligible);
            }

            drIneligibleSumNew["INCD"] = String.Empty;
            drIneligibleSumNew["MED_BILL"] = String.Empty;
            //drIneligibleSumNew["서비스 날짜"] = String.Empty;
            drIneligibleSumNew["의료기관명"] = "합계";
            drIneligibleSumNew["지원되지않는 사유"] = String.Empty;

            MedicalExpenseIneligible sumIneligible = new MedicalExpenseIneligible();

            foreach (MedicalExpenseIneligible expense in lstIneligible)
            {
                sumIneligible.BillAmount += expense.BillAmount;
                sumIneligible.AmountIneligible += expense.AmountIneligible;
            }

            drIneligibleSumNew["청구액(원금)"] = sumIneligible.BillAmount.Value.ToString("C");
            drIneligibleSumNew["전액/일부 지원불가 금액"] = sumIneligible.AmountIneligible.Value.ToString("C");

            dtSorted.Rows.Add(drIneligibleSumNew);
            gvIneligibleInTab.DataSource = dtSorted;

        }


        private void gvBillPaid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (paidSortedField.Field != "INCD")
                    {
                        paidSortedField.Field = "INCD";
                        paidSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTable(paidSortedField);
                    break;
                case 1:
                    if (paidSortedField.Field != "회원 이름")
                    {
                        paidSortedField.Field = "회원 이름";
                        paidSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTable(paidSortedField);
                    break;
                case 2:
                    if (paidSortedField.Field != "MED_BILL")
                    {
                        paidSortedField.Field = "MED_BILL";
                        paidSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTable(paidSortedField);
                    break;
                case 3:
                    if (paidSortedField.Field != "서비스 날짜")
                    {
                        paidSortedField.Field = "서비스 날짜";
                        paidSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTable(paidSortedField);
                    break;
                case 4:
                    if (paidSortedField.Field != "의료기관명")
                    {
                        paidSortedField.Field = "의료기관명";
                        paidSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTable(paidSortedField);
                    break;

            }
        }

        private void gvCMMPendingPayment_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (cmmPendingPaymentSortedField.Field != "INCD")
                    {
                        cmmPendingPaymentSortedField.Field = "INCD";
                        cmmPendingPaymentSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTable(cmmPendingPaymentSortedField);
                    break;
                case 1:
                    if (cmmPendingPaymentSortedField.Field != "회원 이름")
                    {
                        cmmPendingPaymentSortedField.Field = "회원 이름";
                        cmmPendingPaymentSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTable(cmmPendingPaymentSortedField);
                    break;
                case 2:
                    if (cmmPendingPaymentSortedField.Field != "MED_BILL")
                    {
                        cmmPendingPaymentSortedField.Field = "MED_BILL";
                        cmmPendingPaymentSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTable(cmmPendingPaymentSortedField);
                    break;
                case 3:
                    if (cmmPendingPaymentSortedField.Field != "서비스 날짜")
                    {
                        cmmPendingPaymentSortedField.Field = "서비스 날짜";
                        cmmPendingPaymentSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTable(cmmPendingPaymentSortedField);
                    break;
                case 4:
                    if (cmmPendingPaymentSortedField.Field != "의료기관명")
                    {
                        cmmPendingPaymentSortedField.Field = "의료기관명";
                        cmmPendingPaymentSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTable(cmmPendingPaymentSortedField);
                    break;
            }
        }

        private void gvPending_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (pendingSortedField.Field != "INCD")
                    {
                        pendingSortedField.Field = "INCD";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
                case 1:
                    if (pendingSortedField.Field != "회원 이름")
                    {
                        pendingSortedField.Field = "회원 이름";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
                case 2:
                    if (pendingSortedField.Field != "MED_BILL")
                    {
                        pendingSortedField.Field = "MED_BILL";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
                case 3:
                    if (pendingSortedField.Field != "서비스 날짜")
                    {
                        pendingSortedField.Field = "서비스 날짜";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
                //case 4:
                //    if (pendingSortedField.Field != "접수 날짜")
                //    {
                //        pendingSortedField.Field = "접수 날짜";
                //        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                //    }
                //    SortPendingTable(pendingSortedField);
                //    break;
                case 5:
                    if (pendingSortedField.Field != "의료기관명")
                    {
                        pendingSortedField.Field = "의료기관명";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
                case 8:
                    if (pendingSortedField.Field != "보류 사유")
                    {
                        pendingSortedField.Field = "보류 사유";
                        pendingSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTable(pendingSortedField);
                    break;
            }
        }

        private void gvIneligible_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch(e.ColumnIndex)
            {
                case 0:
                    if (ineligibleSortedField.Field != "INCD")
                    {
                        ineligibleSortedField.Field = "INCD";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 1:
                    if (ineligibleSortedField.Field != "회원 이름")
                    {
                        ineligibleSortedField.Field = "회원 이름";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 2:
                    if (ineligibleSortedField.Field != "MED_BILL")
                    {
                        ineligibleSortedField.Field = "MED_BILL";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 3:
                    if (ineligibleSortedField.Field != "서비스 날짜")
                    {
                        ineligibleSortedField.Field = "서비스 날짜";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 4:
                    if (ineligibleSortedField.Field != "접수 날짜")
                    {
                        ineligibleSortedField.Field = "접수 날짜";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 5:
                    if (ineligibleSortedField.Field != "의료기관명")
                    {
                        ineligibleSortedField.Field = "의료기관명";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
                case 8:
                    if (ineligibleSortedField.Field != "지원되지않는 사유")
                    {
                        ineligibleSortedField.Field = "지원되지않는 사유";
                        ineligibleSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTable(ineligibleSortedField);
                    break;
            }
        }

        private void gvPaidInTabPaid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (paidInPaidTabSortedField.Field != "INCD")
                    {
                        paidInPaidTabSortedField.Field = "INCD";
                        paidInPaidTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTableInPaidTab(paidInPaidTabSortedField);
                    break;
                case 1:
                    if (paidInPaidTabSortedField.Field != "회원 이름")
                    {
                        paidInPaidTabSortedField.Field = "회원 이름";
                        paidInPaidTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTableInPaidTab(paidInPaidTabSortedField);
                    break;
                case 2:
                    if (paidInPaidTabSortedField.Field != "MED_BILL")
                    {
                        paidInPaidTabSortedField.Field = "MED_BILL";
                        paidInPaidTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTableInPaidTab(paidInPaidTabSortedField);
                    break;
                case 3:
                    if (paidInPaidTabSortedField.Field != "서비스 날짜")
                    {
                        paidInPaidTabSortedField.Field = "서비스 날짜";
                        paidInPaidTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTableInPaidTab(paidInPaidTabSortedField);
                    break;
                case 4:
                    if (paidInPaidTabSortedField.Field != "의료기관명")
                    {
                        paidInPaidTabSortedField.Field = "의료기관명";
                        paidInPaidTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortBillPaidTableInPaidTab(paidInPaidTabSortedField);
                    break;

            }
        }

        private void gvCMMPendingInTab_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (cmmCMMPendingPaymentInTabSortedField.Field != "INCD")
                    {
                        cmmCMMPendingPaymentInTabSortedField.Field = "INCD";
                        cmmCMMPendingPaymentInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTableInTab(cmmCMMPendingPaymentInTabSortedField);
                    break;
                case 1:
                    if (cmmCMMPendingPaymentInTabSortedField.Field != "회원 이름")
                    {
                        cmmCMMPendingPaymentInTabSortedField.Field = "회원 이름";
                        cmmCMMPendingPaymentInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTableInTab(cmmCMMPendingPaymentInTabSortedField);
                    break;
                case 2:
                    if (cmmCMMPendingPaymentInTabSortedField.Field != "MED_BILL")
                    {
                        cmmCMMPendingPaymentInTabSortedField.Field = "MED_BILL";
                        cmmCMMPendingPaymentInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTableInTab(cmmCMMPendingPaymentInTabSortedField);
                    break;
                case 3:
                    if (cmmCMMPendingPaymentInTabSortedField.Field != "서비스 날짜")
                    {
                        cmmCMMPendingPaymentInTabSortedField.Field = "서비스 날짜";
                        cmmCMMPendingPaymentInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTableInTab(cmmCMMPendingPaymentInTabSortedField);
                    break;
                case 4:
                    if (cmmCMMPendingPaymentInTabSortedField.Field != "의료기관명")
                    {
                        cmmCMMPendingPaymentInTabSortedField.Field = "의료기관명";
                        cmmCMMPendingPaymentInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortCMMPendingPaymentTableInTab(cmmCMMPendingPaymentInTabSortedField);
                    break;
            }
        }

        private void gvPendingInTab_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (pendingInTabSortedField.Field != "INCD")
                    {
                        pendingInTabSortedField.Field = "INCD";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
                case 1:
                    if (pendingInTabSortedField.Field != "회원 이름")
                    {
                        pendingInTabSortedField.Field = "회원 이름";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
                case 2:
                    if (pendingInTabSortedField.Field != "MED_BILL")
                    {
                        pendingInTabSortedField.Field = "MED_BILL";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
                case 3:
                    if (pendingInTabSortedField.Field != "서비스 날짜")
                    {
                        pendingInTabSortedField.Field = "서비스 날짜";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
                //case 4:
                //    if (pendingInTabSortedField.Field != "접수 날짜")
                //    {
                //        pendingInTabSortedField.Field = "접수 날짜";
                //        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                //    }
                //    SortPendingTableInTab(pendingInTabSortedField);
                //    break;
                case 5:
                    if (pendingInTabSortedField.Field != "의료기관명")
                    {
                        pendingInTabSortedField.Field = "의료기관명";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
                case 8:
                    if (pendingInTabSortedField.Field != "보류 사유")
                    {
                        pendingInTabSortedField.Field = "보류 사유";
                        pendingInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPendingTableInTab(pendingInTabSortedField);
                    break;
            }
        }

        private void gvIneligibleInTab_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (ineligibleInTabSortedField.Field != "INCD")
                    {
                        ineligibleInTabSortedField.Field = "INCD";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 1:
                    if (ineligibleInTabSortedField.Field != "회원 이름")
                    {
                        ineligibleInTabSortedField.Field = "회원 이름";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 2:
                    if (ineligibleInTabSortedField.Field != "MED_BILL")
                    {
                        ineligibleInTabSortedField.Field = "MED_BILL";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 3:
                    if (ineligibleInTabSortedField.Field != "서비스 날짜")
                    {
                        ineligibleInTabSortedField.Field = "서비스 날짜";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 4:
                    if (ineligibleInTabSortedField.Field != "접수 날짜")
                    {
                        ineligibleInTabSortedField.Field = "접수 날짜";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 5:
                    if (ineligibleInTabSortedField.Field != "의료기관명")
                    {
                        ineligibleInTabSortedField.Field = "의료기관명";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
                case 8:
                    if (ineligibleInTabSortedField.Field != "지원되지않는 사유")
                    {
                        ineligibleInTabSortedField.Field = "지원되지않는 사유";
                        ineligibleInTabSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortIneligibleTableInTab(ineligibleInTabSortedField);
                    break;
            }
        }

        private void frmBlueSheet_Shown(object sender, EventArgs e)
        {
            frmLogin login_form = new frmLogin();

            login_form.StartPosition = FormStartPosition.CenterParent;

            for (int i = 0; i < 3; i++)
            {
                var LoginResult = login_form.ShowDialog();

                if (LoginResult == DialogResult.OK)
                {
                    Sfdcbinding = login_form.SalesforceBinding;
                    CurrentLoginResult = login_form.SalesforceLoginResult;
                    return;
                }
                else if (LoginResult == DialogResult.Cancel)
                {
                    MessageBox.Show("Login Cancelled", "Error");
                    break;
                }
            }
            this.Close();
        }

        private void gvPersonalResponsibility_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    if (prSortedField.Field != "MEDBILL")
                    {
                        prSortedField.Field = "MEDBILL";
                        prSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPersonalResponsibilityTable(prSortedField);
                    break;
                case 1:
                    if (prSortedField.Field != "서비스 날짜")
                    {
                        prSortedField.Field = "서비스 날짜";
                        prSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPersonalResponsibilityTable(prSortedField);
                    break;
                case 2:
                    if (prSortedField.Field != "의료기관명")
                    {
                        prSortedField.Field = "의료기관명";
                        prSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPersonalResponsibilityTable(prSortedField);
                    break;
                case 4:
                    if (prSortedField.Field != "Type")
                    {
                        prSortedField.Field = "Type";
                        prSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPersonalResponsibilityTable(prSortedField);
                    break;
                case 5:
                    if (prSortedField.Field != "PR Type")
                    {
                        prSortedField.Field = "PR Type";
                        prSortedField.Sorted = SortedField.EnumSorted.NotSorted;
                    }
                    SortPersonalResponsibilityTable(prSortedField);
                    break;
            }
        }

        private void SortPersonalResponsibilityTable(SortedField sf)
        {
            DataTable dtPersonalResponsibility = (DataTable)gvPersonalResponsibility.DataSource;

            DataTable dtSorted = dtPersonalResponsibility.Clone();
            DataTable dtClone = new DataTable();

            //////////////////////////////////////////////////////////////////////////////////////////////////
            dtClone.Columns.Add("MEDBILL", typeof(String));
            dtClone.Columns.Add("서비스 날짜", typeof(DateTime));
            dtClone.Columns.Add("의료기관명", typeof(String));
            dtClone.Columns.Add("청구액(원금)", typeof(String));
            dtClone.Columns.Add("Type", typeof(String));
            dtClone.Columns.Add("PR Type: Member Payment", typeof(String));
            dtClone.Columns.Add("PR Type: Member Discount", typeof(String));
            dtClone.Columns.Add("PR Type: 3rd Party Discount", typeof(String));
            dtClone.Columns.Add("Personal Responsibility Total", typeof(String));

            //dtClone.Columns.Add("본인 부담금", typeof(String));


            dtPersonalResponsibility.Rows.RemoveAt(dtPersonalResponsibility.Rows.Count - 1);

            foreach (DataRow row in dtPersonalResponsibility.Rows)
            {
                DataRow rowClone = dtClone.NewRow();
                dtClone.Rows.Add(rowClone);
            }

            for (int i = 0; i < dtPersonalResponsibility.Rows.Count; i++)
            {
                dtClone.Rows[i]["MEDBILL"] = dtPersonalResponsibility.Rows[i]["MEDBILL"];
                dtClone.Rows[i]["서비스 날짜"] = DateTime.Parse(dtPersonalResponsibility.Rows[i]["서비스 날짜"].ToString());
                dtClone.Rows[i]["의료기관명"] = dtPersonalResponsibility.Rows[i]["의료기관명"];
                dtClone.Rows[i]["청구액(원금)"] = dtPersonalResponsibility.Rows[i]["청구액(원금)"];
                dtClone.Rows[i]["Type"] = dtPersonalResponsibility.Rows[i]["Type"];
                dtClone.Rows[i]["PR Type: Member Payment"] = dtPersonalResponsibility.Rows[i]["PR Type: Member Payment"];
                dtClone.Rows[i]["PR Type: Member Discount"] = dtPersonalResponsibility.Rows[i]["PR Type: Member Discount"];
                dtClone.Rows[i]["PR Type: 3rd Party Discount"] = dtPersonalResponsibility.Rows[i]["PR Type: 3rd Party Discount"];
                dtClone.Rows[i]["Personal Responsibility Total"] = dtPersonalResponsibility.Rows[i]["Personal Responsibility Total"];
                //dtClone.Rows[i]["본인 부담금"] = dtPersonalResponsibility.Rows[i]["본인 부담금"];
            }

            switch(sf.Sorted)
            {
                case SortedField.EnumSorted.NotSorted:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedDesc:
                    dtClone.DefaultView.Sort = sf.Field + " ASC";
                    sf.Sorted = SortedField.EnumSorted.SortedAsc;
                    break;
                case SortedField.EnumSorted.SortedAsc:
                    dtClone.DefaultView.Sort = sf.Field + " DESC";
                    sf.Sorted = SortedField.EnumSorted.SortedDesc;
                    break;
            }

            DataTable dtCloneSorted = dtClone.Clone();
            dtCloneSorted = dtClone.DefaultView.ToTable();

            foreach(DataRow row in dtCloneSorted.Rows)
            {
                DataRow rowSorted = dtSorted.NewRow();
                dtSorted.Rows.Add(rowSorted);
            }

            for (int i = 0; i < dtCloneSorted.Rows.Count; i++)
            {
                dtSorted.Rows[i]["MEDBILL"] = dtCloneSorted.Rows[i]["MEDBILL"];
                dtSorted.Rows[i]["서비스 날짜"] = DateTime.Parse(dtCloneSorted.Rows[i]["서비스 날짜"].ToString()).ToString("MM/dd/yyyy");
                dtSorted.Rows[i]["의료기관명"] = dtCloneSorted.Rows[i]["의료기관명"];
                dtSorted.Rows[i]["청구액(원금)"] = dtCloneSorted.Rows[i]["청구액(원금)"];
                dtSorted.Rows[i]["Type"] = dtCloneSorted.Rows[i]["Type"];
                dtSorted.Rows[i]["PR Type: Member Payment"] = dtCloneSorted.Rows[i]["PR Type: Member Payment"];
                dtSorted.Rows[i]["PR Type: Member Discount"] = dtCloneSorted.Rows[i]["PR Type: Member Discount"];
                dtSorted.Rows[i]["PR Type: 3rd Party Discount"] = dtCloneSorted.Rows[i]["PR Type: 3rd Party Discount"];
                dtSorted.Rows[i]["Personal Responsibility Total"] = dtCloneSorted.Rows[i]["Personal Responsibility Total"];

                //dtSorted.Rows[i]["본인 부담금"] = dtCloneSorted.Rows[i]["본인 부담금"];
            }

            DataRow drPersonalResponsibilitySumNew = dtSorted.NewRow();

            List<PersonalResponsibilityExpense> lstPersonalResponsibilityExpense = new List<PersonalResponsibilityExpense>();

            foreach (DataRow row in dtSorted.Rows)
            {
                //lstPersonalResponsibilityExpense.Add(new PersonalResponsibilityExpense { BillAmount = Double.Parse(row[3].ToString().Substring(1)),
                //                                                                         MemberPayment = Double.Parse(row[5].ToString().Substring(1)),
                //                                                                         MemberDiscount = Double.Parse(row[6].ToString().Substring(1)) });

                PersonalResponsibilityExpense expense = new PersonalResponsibilityExpense();
                expense.BillAmount = Double.Parse(row[3].ToString().Substring(1));
                if (row[5] != DBNull.Value) expense.MemberPayment = Double.Parse(row[5].ToString().Substring(1));
                if (row[6] != DBNull.Value) expense.MemberDiscount = Double.Parse(row[6].ToString().Substring(1));
                if (row[7] != DBNull.Value) expense.ThirdPartyDiscount = Double.Parse(row[7].ToString().Substring(1));
                if (row[8] != DBNull.Value) expense.PersonalResponsiblityTotal = Double.Parse(row[8].ToString().Substring(1));

                lstPersonalResponsibilityExpense.Add(expense);

            }

            drPersonalResponsibilitySumNew["MEDBILL"] = String.Empty;
            drPersonalResponsibilitySumNew["서비스 날짜"] = String.Empty;
            drPersonalResponsibilitySumNew["의료기관명"] = "합계";

            PersonalResponsibilityExpense sumPersonalResponsibility = new PersonalResponsibilityExpense();
            foreach (PersonalResponsibilityExpense expense in lstPersonalResponsibilityExpense)
            {
                sumPersonalResponsibility.BillAmount += expense.BillAmount;
                sumPersonalResponsibility.MemberPayment += expense.MemberPayment;
                sumPersonalResponsibility.MemberDiscount += expense.MemberDiscount;
                //sumPersonalResponsibility.SettlementAmount += expense.SettlementAmount;
            }

            drPersonalResponsibilitySumNew["청구액(원금)"] = sumPersonalResponsibility.BillAmount.ToString("C");
            drPersonalResponsibilitySumNew["PR Type: Member Payment"] = sumPersonalResponsibility.MemberPayment.ToString("C");
            drPersonalResponsibilitySumNew["PR Type: Member Discount"] = sumPersonalResponsibility.MemberDiscount.ToString("C");
            drPersonalResponsibilitySumNew["PR Type: 3rd Party Discount"] = sumPersonalResponsibility.ThirdPartyDiscount.ToString("C");
            drPersonalResponsibilitySumNew["Personal Responsibility Total"] = sumPersonalResponsibility.PersonalResponsiblityTotal.ToString("C");
            //drPersonalResponsibilitySumNew["본인 부담금"] = sumPersonalResponsibility.SettlementAmount.ToString("C");

            dtSorted.Rows.Add(drPersonalResponsibilitySumNew);
            gvPersonalResponsibility.DataSource = dtSorted;
            
        }

        private void rbPersonalResponsibilityOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (rbNoSharingOnly.Checked)
            {
                txtCheckNo.Text = String.Empty;
                txtCheckNo.Enabled = false;
                dtpCheckIssueDate.Text = String.Empty;
                dtpCheckIssueDate.Enabled = false;

                txtACH_No.Text = String.Empty;
                txtACH_No.Enabled = false;
                dtpACHDate.Text = String.Empty;
                dtpACHDate.Enabled = false;

                txtCreditCardNo.Text = String.Empty;
                txtCreditCardNo.Enabled = false;
                dtpCreditCardPaymentDate.Text = String.Empty;
                dtpCreditCardPaymentDate.Enabled = false;

                txtIncidentNo.Text = String.Empty;
                txtIncidentNo.Enabled = true;
                txtIndividualID.Focus();

            }
        }
    }
}
