using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using KPMGAddin.MNBServiceReference;
using System.Data;
using System.IO;
using ADOX;
using ADODB;
using System.Data.OleDb;
using System.Windows.Forms;


namespace KPMGAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public class Viewer
        {
            Form form = new Form();
            public string Path { get; set; }
            public Viewer() { }
           
            public Viewer(string path_)
            {             
                Path = path_;
            }

            public void ShowWindow(DataTable table)  //Check window state, if there is an active control to window than close it
            {
                bool check = false;                                     
                foreach (Control con in form.Controls)
                {
                    form.Close();
                    form = new Form();                          
                    check = true;
                }
                if(check==false) form = new Form();
                WindowHandler(table);
                form.Show();
            } 
            public void WindowHandler(DataTable table)   //Create Datagridview to Mainwindow
            {
                DataGridView view = new DataGridView();
                #region GridViewProperties    
                view.Anchor = AnchorStyles.Left;
                view.Anchor = AnchorStyles.Right;
                view.Dock = DockStyle.Fill;
                view.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                view.RowHeadersVisible = false;
                view.CellValueChanged += new DataGridViewCellEventHandler(view_CellValueChanged);
                view.DataSource = new BindingSource(table, null);              
                #endregion
                form.Text =table.TableName + " " + Path;             
                form.Controls.Add(view);
               
            }

            private void view_CellValueChanged(object sender, DataGridViewCellEventArgs e)
            {
                DataGridView view = sender as DataGridView;   
                if(e.ColumnIndex.ToString()==(view.Columns.Count-1).ToString()) UpdateValue((string)view[e.ColumnIndex, e.RowIndex].Value, e.ColumnIndex, e.RowIndex);
            }

            public void UpdateValue(string newValue, int ColI, int RowI)
            {
                AccessHandler handler = new AccessHandler(Path);
                handler.AddNewValue(newValue, ColI, RowI);              
            }

        }

        public class AccessHandler
        {
            OleDbConnection conn = new OleDbConnection();
            public string Path_ { get; set; }
            public string CreateTableName { get; set; }
            public string SqlDef { get; set; } = @"Provider = Microsoft.ACE.OLEDB.12.0; Data source= ";          
            public AccessHandler(string path) { Path_ = path; }
           
            public AccessHandler()
            {

            }          
            public void PutInfoToAccess() //Check and create or open database
            {
                try
                {                 
                        if (CheckExistDb() == false) CreateDatabase(CreateTableName);
                        AddDataTable(CreateRow(), GetTableData());                    
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }

            public void AddNewValue(string newValue, int ColI, int RowI) // Add new value to database
            {              
                DataTable table = GetTableData();
                table.Rows[RowI].SetField<string>(table.Columns[ColI], newValue);

                OleDbDataAdapter adapter = new OleDbDataAdapter("", GetConnection());
                OleDbCommandBuilder oleDb = new OleDbCommandBuilder(adapter);

                string query = "UPDATE "+table.TableName+" SET ["+table.Columns[ColI].ColumnName+"] = ? WHERE "+ table.Columns[0].ColumnName + " = ?";
                var accessUpdateCommand = new OleDbCommand(query, conn);
                accessUpdateCommand.Parameters.AddWithValue(table.Columns[ColI].ColumnName, newValue);
                accessUpdateCommand.Parameters.AddWithValue(table.Columns[0].ColumnName, RowI+1);

                adapter.UpdateCommand = accessUpdateCommand;
                adapter.UpdateCommand.ExecuteNonQuery();

                GetConnection().Close();
            }
            public OleDbConnection GetConnection()
            {
                try
                {
                    if (conn.State.ToString() == "Closed")
                    {
                        conn.ConnectionString = SqlDef + Path_;
                        conn.Open();
                    }
                }
                catch(Exception ex) { MessageBox.Show(ex.ToString()); }
                return conn;
            }
            
            private void AddDataTable(DataRow row, DataTable table) //Include the new record in the database 
            {
                               
                OleDbDataAdapter adapter = new OleDbDataAdapter("", GetConnection());
                List<string> columnName = new List<string>();

                string strSt = @"INSERT INTO " + table.TableName + "(";
                string strEn = ")VALUES(";
                foreach(var item in table.Columns)
                {
                    columnName.Add(item.ToString());
                    strSt += "[" + item.ToString() + "],";
                    strEn += "?,";                 
                }
                string strFi = strSt.Remove(strSt.Length - 1, 1) + strEn.Remove(strEn.Length - 1, 1)+ ")"; //Final query string
              
                int i = 0;

                using (var command = conn.CreateCommand())
                {
                    command.CommandText = strFi;
                    foreach (var item in columnName)
                    {                    
                        command.Parameters.AddWithValue(item, row[i]);
                        i++;
                    }
                    adapter.InsertCommand = command;
                    adapter.InsertCommand.ExecuteNonQuery();
                }
                GetConnection().Close();
            }
            private DataRow CreateRow() // Include items in the new row
            {
                int i = 0;
                List<object> items = new List<object>();
                string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                DateTime dateTime = DateTime.Now;
                DataRow row;
                DataTable table = GetTableData();
               
                row = table.NewRow();
                               
                items.Add((table.Rows.Count+1).ToString()); //Id
                items.Add(userName);
                items.Add(dateTime);
                items.Add(String.Empty);

                while (items.Count > i)
                {
                    table.Columns.Add();
                    i++;
                }
                row.ItemArray = items.ToArray();

                return row;   
            }

      
            public DataTable GetTableData()// Get chosen tables from database 
            {
                
                List<string> list = new List<string>();   //Database names           
                DataTable schema = GetConnection().GetSchema("Tables");
               
                foreach (DataRow row in schema.Rows)        //Get all tables
                {
                    if(row[3].ToString()=="TABLE") list.Add(row[2].ToString());   //TABLE TYPE == TABLE  add TABLE NAME to list             
                }

                OleDbCommand cmd = new OleDbCommand("SELECT * FROM " + list[0], conn); 
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);             
                DataTable table = new DataTable();
                table.TableName = list[0];
                adapter.Fill(table);
                GetConnection().Close();
                return table;
            }
            private bool CheckExistDb()
            {
                bool check = false;
                if (System.IO.File.Exists(Path_)) check = true;
                return check;             
            }
            private bool CreateDatabase(string TableName) 
            {
                Connection connection = new Connection();
                bool check = false;
                Catalog catalog = new Catalog();
                Table table = new Table
                {
                    Name = TableName
                };

                table.Columns.Append("ID");
                table.Columns.Append("Felh_nev");
                table.Columns.Append("Timestamp");
                table.Columns.Append("Indoklas");
               

                try
                {
                    catalog.Create(SqlDef + Path_ + "; Jet OLEDB:Engine Type=5");
                    catalog.Tables.Append(table);

                    connection = catalog.ActiveConnection as Connection;
                    check = true;
                }
                catch (Exception ex) { check = false; }

                catalog = null;
                connection.Close();
                return check;
            }
        }
        public class MNB
        {
            
            public MNB()
            {

            }        
            public void GetData()
            {             
                DisplayMNBServiceData();
            }
     
            private void FillDataToExcel(Excel.Worksheet sheet, MNBArfolyamServiceSoapClient soapClient) //Fill excel worksheet
            {
                
                GetCurrenciesRequestBody getCurrenciesBody = new GetCurrenciesRequestBody();
                GetExchangeRatesRequestBody getRatesBody = new GetExchangeRatesRequestBody();

                var range = sheet.get_Range("A1", "A1");
                var currencies = soapClient.GetCurrencies(getCurrenciesBody);

                int DayID = 0, countRates = 0;
                string xmlResult = currencies.GetCurrenciesResult;
                DataTable table = XmlToDataTable(xmlResult, 1);
                DataTable dataRates = new DataTable();
                DataTable dataDays = new DataTable();
                List<string> currenciesList = new List<string>();

                ((Excel.Range)range.Cells[1, 1]).Value = "Dátum/ISO";
                ((Excel.Range)range.Cells[2, 1]).Value = "Egység";

                //Fill and get all currencies//
                int columnIndex = 2;
                foreach (DataRow items in table.Rows) 
                {                 
                    ((Excel.Range)range.Cells[1, columnIndex]).Value = items[0];                 
                    currenciesList.Add(items[0].ToString());
                    columnIndex++;
                }
                //////////////////////////////

                //Fill days//
                var exchRates = soapClient.GetExchangeRates(getRatesBody);
                xmlResult = exchRates.GetExchangeRatesResult;
                table = XmlToDataTable(xmlResult, 0);

                columnIndex = 3;
                foreach (DataRow item in table.Rows)
                    {
                    ((Excel.Range)range.Cells[columnIndex, 1]).Value2 = item[0];
                    columnIndex++;
                    }
                /////////////

                //Fill Rates//             
                getRatesBody = new GetExchangeRatesRequestBody();
                dataRates = new DataTable();
                dataDays = new DataTable();
                
                for (int k = 0; k < currenciesList.Count; k++)//Loop all currencies and fill rates value
                {
                    //if (k > 2) break;  Set how many item will appaer
                    range.Cells[1, k+2].Select();
                    DayID = 0;
                    countRates = 0;
                    columnIndex = 3;

                    getRatesBody.currencyNames = currenciesList[k];
                    exchRates = soapClient.GetExchangeRates(getRatesBody);
                    xmlResult = exchRates.GetExchangeRatesResult;
                    dataRates = XmlToDataTable(xmlResult, 1);
                    dataDays = XmlToDataTable(xmlResult, 0);
                
                    if (dataRates != null)
                    {
                        ((Excel.Range)range.Cells[2, k + 2]).Value = dataRates.Rows[countRates][0]; //Fill unit
                        while (countRates < dataRates.Rows.Count)
                        {
                            if (dataRates.Rows[countRates][3].ToString() == dataDays.Rows[DayID][0].ToString())//Date ID check
                            { 
                                ((Excel.Range)range.Cells[columnIndex, k + 2]).Value = dataRates.Rows[countRates][2].ToString().Replace(",", "."); //Change string format 
                                columnIndex++;
                                DayID++;
                                countRates++;
                            }
                            else { columnIndex++; DayID++; }
                        }
                    }
                }
            }         
            private DataTable XmlToDataTable(string xmlResult, int tableindex) //Get datatable from MNB serviceresult xml string
            {
               
                    StringReader reader = new StringReader(xmlResult);
                    DataSet ds = new DataSet();
                    DataTable table = new DataTable();
                    ds.ReadXml(reader);
                    if (ds.Tables.Count > tableindex) table = ds.Tables[tableindex];
                    else table = null;
                    reader.Close();

                    return table;               
            }
            private void DisplayMNBServiceData() //Create client, sheet and call Fill method
            {
                try
                {
                    using (MNBArfolyamServiceSoapClient soapClient = new MNBArfolyamServiceSoapClient())
                    {
                        Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
                        FillDataToExcel(worksheet, soapClient);                        
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }                        
            }
        }
        
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
