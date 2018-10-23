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
using BlueSheetSForce = BlueSheetApp.Salesforce;

namespace BlueSheetApp
{
    public partial class frmLogin : Form
    {
        public String strUserName { get; set; }
        public String strPassword { get; set; }

        private BlueSheetSForce.SforceService Sfdcbinding;
        private BlueSheetSForce.LoginResult CurrentLoginResult;

        public BlueSheetSForce.SforceService SalesforceBinding
        {
            get { return Sfdcbinding; }
        }

        public BlueSheetSForce.LoginResult SalesforceLoginResult
        {
            get { return CurrentLoginResult; }
        }

        public void InitializeSfdcbinding(String strUserName, String strPassword)
        {
            Sfdcbinding = new BlueSheetSForce.SforceService();
            try
            {
                CurrentLoginResult = Sfdcbinding.login(strUserName, strPassword);
            }
            catch (System.Web.Services.Protocols.SoapException e)
            {
                Sfdcbinding = null; // bad username of password
                throw (e);
            }
            catch (Exception e)
            {
                // some other error
                Sfdcbinding = null;
                throw (e);
            }
            Sfdcbinding.Url = CurrentLoginResult.serverUrl;
            Sfdcbinding.SessionHeaderValue = new BlueSheetSForce.SessionHeader();
            Sfdcbinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
        }

        public frmLogin()
        {
            InitializeComponent();

            String strUserIDPasswdPath = @"C:\BlueSheetAppUserIDPasswd\UserIDPasswd.txt";

            if (File.Exists(strUserIDPasswdPath))
            {
                FileStream fs = File.OpenRead(strUserIDPasswdPath);

                StreamReader sr = new StreamReader(fs);

                txtUserId.Text = sr.ReadLine();
                txtPassword.Text = sr.ReadLine();
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            strUserName = txtUserId.Text;
            strPassword = txtPassword.Text;
         
            try
            {
                InitializeSfdcbinding(strUserName, strPassword);
                //Sfdcbinding.Url = CurrentLoginResult.serverUrl;
                //Sfdcbinding.SessionHeaderValue = new BlueSheetSForce.SessionHeader();
                //Sfdcbinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Salesforce Connection Error");
                this.DialogResult = DialogResult.No;
                return;
            }

            String strUserIDPasswdFolderHidden = @"C:\BlueSheetAppUserIDPasswd";

            if (!Directory.Exists(strUserIDPasswdFolderHidden))
            {
                DirectoryInfo di = Directory.CreateDirectory(strUserIDPasswdFolderHidden);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }

            String strUserIDPasswdPath = @"C:\BlueSheetAppUserIDPasswd\UserIDPasswd.txt";

            FileStream fs = File.Open(strUserIDPasswdPath, FileMode.Create, FileAccess.Write, FileShare.None);
            StreamWriter sw = new StreamWriter(fs);

            sw.WriteLine(strUserName);
            sw.WriteLine(strPassword);

            sw.Close();

            this.DialogResult = DialogResult.OK;
            Cursor.Current = Cursors.Default;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
            return;
        }
    }
}
