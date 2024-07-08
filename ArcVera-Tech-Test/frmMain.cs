namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        private Report reportCls;
        public frmMain()
        {
            InitializeComponent();
            if(reportCls == null)
                reportCls = new Report();
        }

        private void btnImportEra5_Click(object sender, EventArgs e)
        {

        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {

        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {

        }
    }
}
