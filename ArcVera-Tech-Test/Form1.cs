namespace ArcVera_Tech_Test
{
    public partial class Form1 : Form
    {
        private Report reportCls;
        public Form1()
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
