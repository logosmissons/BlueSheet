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
    public partial class frmStartEndDate : Form
    {

        public DateTime? StartDate;
        public DateTime? EndDate;


        public frmStartEndDate()
        {
            InitializeComponent();

            StartDate = null;
            EndDate = null;

            //int nYear = DateTime.Today.Year;

            //for (int year = nYear; year > nYear - 5; year--)
            //{
            //    comboYear.Items.Add(year);
            //}

            //comboYear.SelectedItem = nYear;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            StartDate = dtpStartDate.Value;
            EndDate = dtpEndDate.Value;

            DialogResult = DialogResult.OK;

            Close();
            //YearSelected = Int32.Parse(comboYear.SelectedValue.ToString());
            //YearSelected = Int32.Parse(comboYear.SelectedItem.ToString());
            //DialogResult = DialogResult.OK;
            //Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
