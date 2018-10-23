using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BlueSheetApp
{
    public partial class frmDocReceivedDate : Form
    {

        private DateTime dtReceivedDate;

        public DateTime ReceivedDate
        {
            get { return dtReceivedDate; }
        }

        public frmDocReceivedDate()
        {
            InitializeComponent();

            dtpDocRcvdDate.Value = DateTime.Today.AddMonths(-2);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            dtReceivedDate = DateTime.Parse(dtpDocRcvdDate.Value.ToString("MM/dd/yyyy"));
            this.DialogResult = DialogResult.OK;
            return;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            return;
        }
    }
}
