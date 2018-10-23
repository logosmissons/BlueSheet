namespace BlueSheetApp
{
    partial class frmBlueSheet
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmBlueSheet));
            this.btnSearch = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGeneratePDF = new System.Windows.Forms.Button();
            this.grpPaymentMethod = new System.Windows.Forms.GroupBox();
            this.rbNoSharingOnly = new System.Windows.Forms.RadioButton();
            this.rbACH = new System.Windows.Forms.RadioButton();
            this.rbCreditCard = new System.Windows.Forms.RadioButton();
            this.rbCheck = new System.Windows.Forms.RadioButton();
            this.grpPaymentInfo = new System.Windows.Forms.GroupBox();
            this.dtpACHDate = new System.Windows.Forms.DateTimePicker();
            this.dtpCheckIssueDate = new System.Windows.Forms.DateTimePicker();
            this.dtpCreditCardPaymentDate = new System.Windows.Forms.DateTimePicker();
            this.txtACH_No = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtCreditCardNo = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtCheckNo = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tabMedicalExpense = new System.Windows.Forms.TabControl();
            this.tabAll = new System.Windows.Forms.TabPage();
            this.gvSummary = new System.Windows.Forms.DataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.gvIneligible = new System.Windows.Forms.DataGridView();
            this.label4 = new System.Windows.Forms.Label();
            this.gvPending = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.gvCMMPendingPayment = new System.Windows.Forms.DataGridView();
            this.label5 = new System.Windows.Forms.Label();
            this.gvBillPaid = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPaid = new System.Windows.Forms.TabPage();
            this.gvPaidInTabPaid = new System.Windows.Forms.DataGridView();
            this.label12 = new System.Windows.Forms.Label();
            this.tabCMMPendingPayment = new System.Windows.Forms.TabPage();
            this.gvCMMPendingInTab = new System.Windows.Forms.DataGridView();
            this.label14 = new System.Windows.Forms.Label();
            this.tabPending = new System.Windows.Forms.TabPage();
            this.gvPendingInTab = new System.Windows.Forms.DataGridView();
            this.label15 = new System.Windows.Forms.Label();
            this.tabIneligible = new System.Windows.Forms.TabPage();
            this.gvIneligibleInTab = new System.Windows.Forms.DataGridView();
            this.label16 = new System.Windows.Forms.Label();
            this.tabPersonalResponsibility = new System.Windows.Forms.TabPage();
            this.label18 = new System.Windows.Forms.Label();
            this.gvPersonalResponsibility = new System.Windows.Forms.DataGridView();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.btnGenerateEnPDF = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtIndividualID = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtIncidentNo = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.gvIneligibleNoSharing = new System.Windows.Forms.DataGridView();
            this.grpPaymentMethod.SuspendLayout();
            this.grpPaymentInfo.SuspendLayout();
            this.tabMedicalExpense.SuspendLayout();
            this.tabAll.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvSummary)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligible)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPending)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvCMMPendingPayment)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvBillPaid)).BeginInit();
            this.tabPaid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPaidInTabPaid)).BeginInit();
            this.tabCMMPendingPayment.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvCMMPendingInTab)).BeginInit();
            this.tabPending.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPendingInTab)).BeginInit();
            this.tabIneligible.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligibleInTab)).BeginInit();
            this.tabPersonalResponsibility.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPersonalResponsibility)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligibleNoSharing)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSearch
            // 
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(30, 130);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(129, 30);
            this.btnSearch.TabIndex = 2;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(462, 130);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(129, 30);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGeneratePDF
            // 
            this.btnGeneratePDF.Location = new System.Drawing.Point(174, 130);
            this.btnGeneratePDF.Name = "btnGeneratePDF";
            this.btnGeneratePDF.Size = new System.Drawing.Size(129, 30);
            this.btnGeneratePDF.TabIndex = 8;
            this.btnGeneratePDF.Text = "Generate Korean PDF";
            this.btnGeneratePDF.UseVisualStyleBackColor = true;
            this.btnGeneratePDF.Click += new System.EventHandler(this.btnGeneratePDF_Click);
            // 
            // grpPaymentMethod
            // 
            this.grpPaymentMethod.Controls.Add(this.rbNoSharingOnly);
            this.grpPaymentMethod.Controls.Add(this.rbACH);
            this.grpPaymentMethod.Controls.Add(this.rbCreditCard);
            this.grpPaymentMethod.Controls.Add(this.rbCheck);
            this.grpPaymentMethod.Location = new System.Drawing.Point(30, 65);
            this.grpPaymentMethod.Name = "grpPaymentMethod";
            this.grpPaymentMethod.Size = new System.Drawing.Size(561, 50);
            this.grpPaymentMethod.TabIndex = 29;
            this.grpPaymentMethod.TabStop = false;
            this.grpPaymentMethod.Text = "Payment Method";
            // 
            // rbNoSharingOnly
            // 
            this.rbNoSharingOnly.AutoSize = true;
            this.rbNoSharingOnly.Location = new System.Drawing.Point(424, 19);
            this.rbNoSharingOnly.Name = "rbNoSharingOnly";
            this.rbNoSharingOnly.Size = new System.Drawing.Size(102, 17);
            this.rbNoSharingOnly.TabIndex = 4;
            this.rbNoSharingOnly.TabStop = true;
            this.rbNoSharingOnly.Text = "No Sharing Only";
            this.rbNoSharingOnly.UseVisualStyleBackColor = true;
            this.rbNoSharingOnly.CheckedChanged += new System.EventHandler(this.rbPersonalResponsibilityOnly_CheckedChanged);
            // 
            // rbACH
            // 
            this.rbACH.AutoSize = true;
            this.rbACH.Location = new System.Drawing.Point(167, 19);
            this.rbACH.Name = "rbACH";
            this.rbACH.Size = new System.Drawing.Size(47, 17);
            this.rbACH.TabIndex = 3;
            this.rbACH.TabStop = true;
            this.rbACH.Text = "ACH";
            this.rbACH.UseVisualStyleBackColor = true;
            this.rbACH.CheckedChanged += new System.EventHandler(this.rbACH_CheckedChanged);
            // 
            // rbCreditCard
            // 
            this.rbCreditCard.AutoSize = true;
            this.rbCreditCard.Location = new System.Drawing.Point(296, 19);
            this.rbCreditCard.Name = "rbCreditCard";
            this.rbCreditCard.Size = new System.Drawing.Size(77, 17);
            this.rbCreditCard.TabIndex = 2;
            this.rbCreditCard.TabStop = true;
            this.rbCreditCard.Text = "Credit Card";
            this.rbCreditCard.UseVisualStyleBackColor = true;
            this.rbCreditCard.CheckedChanged += new System.EventHandler(this.rbCreditCard_CheckedChanged);
            // 
            // rbCheck
            // 
            this.rbCheck.AutoSize = true;
            this.rbCheck.Location = new System.Drawing.Point(24, 19);
            this.rbCheck.Name = "rbCheck";
            this.rbCheck.Size = new System.Drawing.Size(56, 17);
            this.rbCheck.TabIndex = 1;
            this.rbCheck.TabStop = true;
            this.rbCheck.Text = "Check";
            this.rbCheck.UseVisualStyleBackColor = true;
            this.rbCheck.CheckedChanged += new System.EventHandler(this.rbCheck_CheckedChanged);
            // 
            // grpPaymentInfo
            // 
            this.grpPaymentInfo.Controls.Add(this.dtpACHDate);
            this.grpPaymentInfo.Controls.Add(this.dtpCheckIssueDate);
            this.grpPaymentInfo.Controls.Add(this.dtpCreditCardPaymentDate);
            this.grpPaymentInfo.Controls.Add(this.txtACH_No);
            this.grpPaymentInfo.Controls.Add(this.label7);
            this.grpPaymentInfo.Controls.Add(this.label13);
            this.grpPaymentInfo.Controls.Add(this.label11);
            this.grpPaymentInfo.Controls.Add(this.label10);
            this.grpPaymentInfo.Controls.Add(this.txtCreditCardNo);
            this.grpPaymentInfo.Controls.Add(this.label9);
            this.grpPaymentInfo.Controls.Add(this.txtCheckNo);
            this.grpPaymentInfo.Controls.Add(this.label8);
            this.grpPaymentInfo.Location = new System.Drawing.Point(610, 27);
            this.grpPaymentInfo.Name = "grpPaymentInfo";
            this.grpPaymentInfo.Size = new System.Drawing.Size(675, 133);
            this.grpPaymentInfo.TabIndex = 30;
            this.grpPaymentInfo.TabStop = false;
            this.grpPaymentInfo.Text = "Payment Information";
            // 
            // dtpACHDate
            // 
            this.dtpACHDate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpACHDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpACHDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpACHDate.Location = new System.Drawing.Point(438, 59);
            this.dtpACHDate.Name = "dtpACHDate";
            this.dtpACHDate.Size = new System.Drawing.Size(211, 26);
            this.dtpACHDate.TabIndex = 105;
            // 
            // dtpCheckIssueDate
            // 
            this.dtpCheckIssueDate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpCheckIssueDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpCheckIssueDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCheckIssueDate.Location = new System.Drawing.Point(438, 26);
            this.dtpCheckIssueDate.Name = "dtpCheckIssueDate";
            this.dtpCheckIssueDate.Size = new System.Drawing.Size(211, 26);
            this.dtpCheckIssueDate.TabIndex = 104;
            // 
            // dtpCreditCardPaymentDate
            // 
            this.dtpCreditCardPaymentDate.Enabled = false;
            this.dtpCreditCardPaymentDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpCreditCardPaymentDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCreditCardPaymentDate.Location = new System.Drawing.Point(438, 93);
            this.dtpCreditCardPaymentDate.Name = "dtpCreditCardPaymentDate";
            this.dtpCreditCardPaymentDate.Size = new System.Drawing.Size(211, 26);
            this.dtpCreditCardPaymentDate.TabIndex = 103;
            // 
            // txtACH_No
            // 
            this.txtACH_No.Enabled = false;
            this.txtACH_No.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtACH_No.Location = new System.Drawing.Point(101, 58);
            this.txtACH_No.Name = "txtACH_No";
            this.txtACH_No.Size = new System.Drawing.Size(161, 26);
            this.txtACH_No.TabIndex = 43;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 65);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(49, 13);
            this.label7.TabIndex = 42;
            this.label7.Text = "ACH No:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(300, 65);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(92, 13);
            this.label13.TabIndex = 40;
            this.label13.Text = "Transaction Date:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(300, 98);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(132, 13);
            this.label11.TabIndex = 33;
            this.label11.Text = "Credit Card Payment Date:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(300, 31);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(95, 13);
            this.label10.TabIndex = 32;
            this.label10.Text = "Check Issue Date:";
            // 
            // txtCreditCardNo
            // 
            this.txtCreditCardNo.Enabled = false;
            this.txtCreditCardNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCreditCardNo.Location = new System.Drawing.Point(101, 93);
            this.txtCreditCardNo.Name = "txtCreditCardNo";
            this.txtCreditCardNo.Size = new System.Drawing.Size(161, 26);
            this.txtCreditCardNo.TabIndex = 31;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(17, 98);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(79, 13);
            this.label9.TabIndex = 30;
            this.label9.Text = "Credit Card No:";
            // 
            // txtCheckNo
            // 
            this.txtCheckNo.Enabled = false;
            this.txtCheckNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCheckNo.Location = new System.Drawing.Point(101, 23);
            this.txtCheckNo.Name = "txtCheckNo";
            this.txtCheckNo.Size = new System.Drawing.Size(161, 26);
            this.txtCheckNo.TabIndex = 99;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 31);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 13);
            this.label8.TabIndex = 28;
            this.label8.Text = "Check No:";
            // 
            // tabMedicalExpense
            // 
            this.tabMedicalExpense.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabMedicalExpense.Controls.Add(this.tabAll);
            this.tabMedicalExpense.Controls.Add(this.tabPaid);
            this.tabMedicalExpense.Controls.Add(this.tabCMMPendingPayment);
            this.tabMedicalExpense.Controls.Add(this.tabPending);
            this.tabMedicalExpense.Controls.Add(this.tabIneligible);
            this.tabMedicalExpense.Controls.Add(this.tabPersonalResponsibility);
            this.tabMedicalExpense.Location = new System.Drawing.Point(30, 174);
            this.tabMedicalExpense.Name = "tabMedicalExpense";
            this.tabMedicalExpense.SelectedIndex = 0;
            this.tabMedicalExpense.Size = new System.Drawing.Size(1810, 744);
            this.tabMedicalExpense.TabIndex = 31;
            // 
            // tabAll
            // 
            this.tabAll.Controls.Add(this.gvSummary);
            this.tabAll.Controls.Add(this.label6);
            this.tabAll.Controls.Add(this.gvIneligible);
            this.tabAll.Controls.Add(this.label4);
            this.tabAll.Controls.Add(this.gvPending);
            this.tabAll.Controls.Add(this.label3);
            this.tabAll.Controls.Add(this.gvCMMPendingPayment);
            this.tabAll.Controls.Add(this.label5);
            this.tabAll.Controls.Add(this.gvBillPaid);
            this.tabAll.Controls.Add(this.label2);
            this.tabAll.Location = new System.Drawing.Point(4, 22);
            this.tabAll.Name = "tabAll";
            this.tabAll.Padding = new System.Windows.Forms.Padding(3);
            this.tabAll.Size = new System.Drawing.Size(1802, 718);
            this.tabAll.TabIndex = 0;
            this.tabAll.Text = "All";
            this.tabAll.UseVisualStyleBackColor = true;
            // 
            // gvSummary
            // 
            this.gvSummary.AllowUserToAddRows = false;
            this.gvSummary.AllowUserToDeleteRows = false;
            this.gvSummary.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvSummary.Location = new System.Drawing.Point(17, 597);
            this.gvSummary.Name = "gvSummary";
            this.gvSummary.ReadOnly = true;
            this.gvSummary.Size = new System.Drawing.Size(1757, 96);
            this.gvSummary.TabIndex = 19;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(13, 582);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(127, 16);
            this.label6.TabIndex = 18;
            this.label6.Text = "연도별 의료비 지원 내역";
            // 
            // gvIneligible
            // 
            this.gvIneligible.AllowUserToAddRows = false;
            this.gvIneligible.AllowUserToDeleteRows = false;
            this.gvIneligible.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvIneligible.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvIneligible.Location = new System.Drawing.Point(16, 470);
            this.gvIneligible.Name = "gvIneligible";
            this.gvIneligible.ReadOnly = true;
            this.gvIneligible.Size = new System.Drawing.Size(1758, 96);
            this.gvIneligible.TabIndex = 17;
            this.gvIneligible.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvIneligible_ColumnHeaderMouseClick);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(12, 452);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 16);
            this.label4.TabIndex = 16;
            this.label4.Text = "지원 불가한 의료비";
            // 
            // gvPending
            // 
            this.gvPending.AllowUserToAddRows = false;
            this.gvPending.AllowUserToDeleteRows = false;
            this.gvPending.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvPending.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvPending.Location = new System.Drawing.Point(16, 343);
            this.gvPending.Name = "gvPending";
            this.gvPending.ReadOnly = true;
            this.gvPending.Size = new System.Drawing.Size(1758, 96);
            this.gvPending.TabIndex = 15;
            this.gvPending.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvPending_ColumnHeaderMouseClick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(13, 325);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 16);
            this.label3.TabIndex = 14;
            this.label3.Text = "현재 보류 중인 의료비";
            // 
            // gvCMMPendingPayment
            // 
            this.gvCMMPendingPayment.AllowUserToAddRows = false;
            this.gvCMMPendingPayment.AllowUserToDeleteRows = false;
            this.gvCMMPendingPayment.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvCMMPendingPayment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvCMMPendingPayment.Location = new System.Drawing.Point(15, 216);
            this.gvCMMPendingPayment.Name = "gvCMMPendingPayment";
            this.gvCMMPendingPayment.ReadOnly = true;
            this.gvCMMPendingPayment.Size = new System.Drawing.Size(1759, 96);
            this.gvCMMPendingPayment.TabIndex = 13;
            this.gvCMMPendingPayment.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvCMMPendingPayment_ColumnHeaderMouseClick);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(13, 200);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(102, 16);
            this.label5.TabIndex = 12;
            this.label5.Text = "지불 예정인 의료비";
            // 
            // gvBillPaid
            // 
            this.gvBillPaid.AllowUserToAddRows = false;
            this.gvBillPaid.AllowUserToDeleteRows = false;
            this.gvBillPaid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvBillPaid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvBillPaid.Location = new System.Drawing.Point(16, 35);
            this.gvBillPaid.Name = "gvBillPaid";
            this.gvBillPaid.ReadOnly = true;
            this.gvBillPaid.Size = new System.Drawing.Size(1758, 150);
            this.gvBillPaid.TabIndex = 7;
            this.gvBillPaid.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvBillPaid_ColumnHeaderMouseClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label2.Location = new System.Drawing.Point(13, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(185, 16);
            this.label2.TabIndex = 6;
            this.label2.Text = "회원 및 의료기관으로 지불된 의료비";
            // 
            // tabPaid
            // 
            this.tabPaid.Controls.Add(this.gvPaidInTabPaid);
            this.tabPaid.Controls.Add(this.label12);
            this.tabPaid.Location = new System.Drawing.Point(4, 22);
            this.tabPaid.Name = "tabPaid";
            this.tabPaid.Padding = new System.Windows.Forms.Padding(3);
            this.tabPaid.Size = new System.Drawing.Size(1802, 718);
            this.tabPaid.TabIndex = 1;
            this.tabPaid.Text = "Paid";
            this.tabPaid.UseVisualStyleBackColor = true;
            // 
            // gvPaidInTabPaid
            // 
            this.gvPaidInTabPaid.AllowUserToAddRows = false;
            this.gvPaidInTabPaid.AllowUserToDeleteRows = false;
            this.gvPaidInTabPaid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvPaidInTabPaid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvPaidInTabPaid.Location = new System.Drawing.Point(26, 41);
            this.gvPaidInTabPaid.Name = "gvPaidInTabPaid";
            this.gvPaidInTabPaid.Size = new System.Drawing.Size(1747, 656);
            this.gvPaidInTabPaid.TabIndex = 1;
            this.gvPaidInTabPaid.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvPaidInTabPaid_ColumnHeaderMouseClick);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.label12.Location = new System.Drawing.Point(21, 14);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(185, 16);
            this.label12.TabIndex = 0;
            this.label12.Text = "회원 및 의료기관으로 지불된 의료비";
            // 
            // tabCMMPendingPayment
            // 
            this.tabCMMPendingPayment.Controls.Add(this.gvCMMPendingInTab);
            this.tabCMMPendingPayment.Controls.Add(this.label14);
            this.tabCMMPendingPayment.Location = new System.Drawing.Point(4, 22);
            this.tabCMMPendingPayment.Name = "tabCMMPendingPayment";
            this.tabCMMPendingPayment.Size = new System.Drawing.Size(1802, 718);
            this.tabCMMPendingPayment.TabIndex = 2;
            this.tabCMMPendingPayment.Text = "CMM Pending Payment";
            this.tabCMMPendingPayment.UseVisualStyleBackColor = true;
            // 
            // gvCMMPendingInTab
            // 
            this.gvCMMPendingInTab.AllowUserToAddRows = false;
            this.gvCMMPendingInTab.AllowUserToDeleteRows = false;
            this.gvCMMPendingInTab.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvCMMPendingInTab.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvCMMPendingInTab.Location = new System.Drawing.Point(24, 45);
            this.gvCMMPendingInTab.Name = "gvCMMPendingInTab";
            this.gvCMMPendingInTab.Size = new System.Drawing.Size(1747, 646);
            this.gvCMMPendingInTab.TabIndex = 14;
            this.gvCMMPendingInTab.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvCMMPendingInTab_ColumnHeaderMouseClick);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(21, 17);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(101, 13);
            this.label14.TabIndex = 13;
            this.label14.Text = "지불 예정인 의료비";
            // 
            // tabPending
            // 
            this.tabPending.Controls.Add(this.gvPendingInTab);
            this.tabPending.Controls.Add(this.label15);
            this.tabPending.Location = new System.Drawing.Point(4, 22);
            this.tabPending.Name = "tabPending";
            this.tabPending.Size = new System.Drawing.Size(1802, 718);
            this.tabPending.TabIndex = 3;
            this.tabPending.Text = "Pending";
            this.tabPending.UseVisualStyleBackColor = true;
            // 
            // gvPendingInTab
            // 
            this.gvPendingInTab.AllowUserToAddRows = false;
            this.gvPendingInTab.AllowUserToDeleteRows = false;
            this.gvPendingInTab.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvPendingInTab.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvPendingInTab.Location = new System.Drawing.Point(26, 49);
            this.gvPendingInTab.Name = "gvPendingInTab";
            this.gvPendingInTab.Size = new System.Drawing.Size(1742, 645);
            this.gvPendingInTab.TabIndex = 16;
            this.gvPendingInTab.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvPendingInTab_ColumnHeaderMouseClick);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.Red;
            this.label15.Location = new System.Drawing.Point(21, 17);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(113, 16);
            this.label15.TabIndex = 15;
            this.label15.Text = "현재 보류중인 의료비";
            // 
            // tabIneligible
            // 
            this.tabIneligible.Controls.Add(this.gvIneligibleInTab);
            this.tabIneligible.Controls.Add(this.label16);
            this.tabIneligible.Location = new System.Drawing.Point(4, 22);
            this.tabIneligible.Name = "tabIneligible";
            this.tabIneligible.Size = new System.Drawing.Size(1802, 718);
            this.tabIneligible.TabIndex = 4;
            this.tabIneligible.Text = "Ineligible";
            this.tabIneligible.UseVisualStyleBackColor = true;
            // 
            // gvIneligibleInTab
            // 
            this.gvIneligibleInTab.AllowUserToAddRows = false;
            this.gvIneligibleInTab.AllowUserToDeleteRows = false;
            this.gvIneligibleInTab.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvIneligibleInTab.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvIneligibleInTab.Location = new System.Drawing.Point(24, 51);
            this.gvIneligibleInTab.Name = "gvIneligibleInTab";
            this.gvIneligibleInTab.Size = new System.Drawing.Size(1749, 638);
            this.gvIneligibleInTab.TabIndex = 18;
            this.gvIneligibleInTab.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvIneligibleInTab_ColumnHeaderMouseClick);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.ForeColor = System.Drawing.Color.Red;
            this.label16.Location = new System.Drawing.Point(21, 16);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(102, 16);
            this.label16.TabIndex = 17;
            this.label16.Text = "지원 불가한 의료비";
            // 
            // tabPersonalResponsibility
            // 
            this.tabPersonalResponsibility.Controls.Add(this.gvIneligibleNoSharing);
            this.tabPersonalResponsibility.Controls.Add(this.label20);
            this.tabPersonalResponsibility.Controls.Add(this.label18);
            this.tabPersonalResponsibility.Controls.Add(this.gvPersonalResponsibility);
            this.tabPersonalResponsibility.Location = new System.Drawing.Point(4, 22);
            this.tabPersonalResponsibility.Name = "tabPersonalResponsibility";
            this.tabPersonalResponsibility.Size = new System.Drawing.Size(1802, 718);
            this.tabPersonalResponsibility.TabIndex = 5;
            this.tabPersonalResponsibility.Text = "Personal Responsibility";
            this.tabPersonalResponsibility.UseVisualStyleBackColor = true;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(23, 17);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(66, 16);
            this.label18.TabIndex = 13;
            this.label18.Text = "본인 부담금";
            // 
            // gvPersonalResponsibility
            // 
            this.gvPersonalResponsibility.AllowUserToAddRows = false;
            this.gvPersonalResponsibility.AllowUserToDeleteRows = false;
            this.gvPersonalResponsibility.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvPersonalResponsibility.Location = new System.Drawing.Point(20, 45);
            this.gvPersonalResponsibility.Name = "gvPersonalResponsibility";
            this.gvPersonalResponsibility.ReadOnly = true;
            this.gvPersonalResponsibility.Size = new System.Drawing.Size(1748, 253);
            this.gvPersonalResponsibility.TabIndex = 0;
            this.gvPersonalResponsibility.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvPersonalResponsibility_ColumnHeaderMouseClick);
            // 
            // btnGenerateEnPDF
            // 
            this.btnGenerateEnPDF.Location = new System.Drawing.Point(318, 130);
            this.btnGenerateEnPDF.Name = "btnGenerateEnPDF";
            this.btnGenerateEnPDF.Size = new System.Drawing.Size(129, 30);
            this.btnGenerateEnPDF.TabIndex = 32;
            this.btnGenerateEnPDF.Text = "Generate English PDF";
            this.btnGenerateEnPDF.UseVisualStyleBackColor = true;
            this.btnGenerateEnPDF.Click += new System.EventHandler(this.btnGenerateEnPDF_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(31, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 16);
            this.label1.TabIndex = 33;
            this.label1.Text = "Individual ID:";
            // 
            // txtIndividualID
            // 
            this.txtIndividualID.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIndividualID.Location = new System.Drawing.Point(121, 27);
            this.txtIndividualID.Name = "txtIndividualID";
            this.txtIndividualID.Size = new System.Drawing.Size(182, 26);
            this.txtIndividualID.TabIndex = 34;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(324, 32);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(78, 16);
            this.label17.TabIndex = 35;
            this.label17.Text = "Incident No:";
            // 
            // txtIncidentNo
            // 
            this.txtIncidentNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIncidentNo.Location = new System.Drawing.Point(409, 27);
            this.txtIncidentNo.Name = "txtIncidentNo";
            this.txtIncidentNo.Size = new System.Drawing.Size(182, 26);
            this.txtIncidentNo.TabIndex = 36;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(1725, 32);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(111, 13);
            this.label19.TabIndex = 37;
            this.label19.Text = "BlueSheet version 1.5\r\n";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label20.Location = new System.Drawing.Point(23, 328);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(91, 16);
            this.label20.TabIndex = 14;
            this.label20.Text = "지원 불가 의료비";
            // 
            // gvIneligibleNoSharing
            // 
            this.gvIneligibleNoSharing.AllowUserToAddRows = false;
            this.gvIneligibleNoSharing.AllowUserToDeleteRows = false;
            this.gvIneligibleNoSharing.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvIneligibleNoSharing.Location = new System.Drawing.Point(20, 358);
            this.gvIneligibleNoSharing.Name = "gvIneligibleNoSharing";
            this.gvIneligibleNoSharing.ReadOnly = true;
            this.gvIneligibleNoSharing.Size = new System.Drawing.Size(1748, 251);
            this.gvIneligibleNoSharing.TabIndex = 15;
            // 
            // frmBlueSheet
            // 
            this.AcceptButton = this.btnSearch;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1864, 930);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.txtIncidentNo);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.txtIndividualID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenerateEnPDF);
            this.Controls.Add(this.tabMedicalExpense);
            this.Controls.Add(this.grpPaymentInfo);
            this.Controls.Add(this.grpPaymentMethod);
            this.Controls.Add(this.btnGeneratePDF);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSearch);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmBlueSheet";
            this.Text = "Blue Sheet App";
            this.Shown += new System.EventHandler(this.frmBlueSheet_Shown);
            this.grpPaymentMethod.ResumeLayout(false);
            this.grpPaymentMethod.PerformLayout();
            this.grpPaymentInfo.ResumeLayout(false);
            this.grpPaymentInfo.PerformLayout();
            this.tabMedicalExpense.ResumeLayout(false);
            this.tabAll.ResumeLayout(false);
            this.tabAll.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvSummary)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligible)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvPending)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvCMMPendingPayment)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvBillPaid)).EndInit();
            this.tabPaid.ResumeLayout(false);
            this.tabPaid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPaidInTabPaid)).EndInit();
            this.tabCMMPendingPayment.ResumeLayout(false);
            this.tabCMMPendingPayment.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvCMMPendingInTab)).EndInit();
            this.tabPending.ResumeLayout(false);
            this.tabPending.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPendingInTab)).EndInit();
            this.tabIneligible.ResumeLayout(false);
            this.tabIneligible.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligibleInTab)).EndInit();
            this.tabPersonalResponsibility.ResumeLayout(false);
            this.tabPersonalResponsibility.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvPersonalResponsibility)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gvIneligibleNoSharing)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGeneratePDF;
        private System.Windows.Forms.GroupBox grpPaymentMethod;
        private System.Windows.Forms.GroupBox grpPaymentInfo;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtCreditCardNo;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtCheckNo;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.RadioButton rbACH;
        private System.Windows.Forms.RadioButton rbCreditCard;
        private System.Windows.Forms.RadioButton rbCheck;
        private System.Windows.Forms.TextBox txtACH_No;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TabControl tabMedicalExpense;
        private System.Windows.Forms.TabPage tabAll;
        private System.Windows.Forms.TabPage tabPaid;
        private System.Windows.Forms.TabPage tabCMMPendingPayment;
        private System.Windows.Forms.DataGridView gvSummary;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView gvIneligible;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView gvPending;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView gvCMMPendingPayment;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView gvBillPaid;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage tabPending;
        private System.Windows.Forms.TabPage tabIneligible;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView gvPaidInTabPaid;
        private System.Windows.Forms.DataGridView gvCMMPendingInTab;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.DataGridView gvPendingInTab;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.DataGridView gvIneligibleInTab;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnGenerateEnPDF;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtIndividualID;
        private System.Windows.Forms.DateTimePicker dtpCreditCardPaymentDate;
        private System.Windows.Forms.DateTimePicker dtpACHDate;
        private System.Windows.Forms.DateTimePicker dtpCheckIssueDate;
        private System.Windows.Forms.TabPage tabPersonalResponsibility;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.DataGridView gvPersonalResponsibility;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtIncidentNo;
        private System.Windows.Forms.RadioButton rbNoSharingOnly;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.DataGridView gvIneligibleNoSharing;
        private System.Windows.Forms.Label label20;
    }
}

