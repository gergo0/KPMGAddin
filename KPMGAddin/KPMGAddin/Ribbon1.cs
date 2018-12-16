using Microsoft.Office.Tools.Ribbon;



namespace KPMGAddin
{
   
    public partial class Ribbon1
    {
        const string path = "C:\\teszt.accdb";     
        const string TableName = "KPMG";


        ThisAddIn.AccessHandler access = new ThisAddIn.AccessHandler();
        ThisAddIn.MNB mNB = new ThisAddIn.MNB();
        ThisAddIn.Viewer viewer = new ThisAddIn.Viewer();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            access.Path_ = path;
            access.CreateTableName = TableName;
            viewer.Path = path;
        }

        private void buttonMNB_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {                   
            mNB.GetData();
            access.PutInfoToAccess();                                         
        }

        private void buttonLog_Click(object sender, RibbonControlEventArgs e)
        {
            if (access.GetConnection().State.ToString() == "Open") viewer.ShowWindow(access.GetTableData());         
        }
        private void ActiveWindow()
        {
           
        }
    }
}
