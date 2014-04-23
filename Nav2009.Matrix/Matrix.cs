using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Text;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Dynamics.Framework.UI.Extensibility;
using Microsoft.Dynamics.Framework.UI.Extensibility.WinForms;


namespace Nav2009.Matrix
{
    [ControlAddInExport("Nav2009.Matrix")]
    // public token: 86ce8986b7a3f4a7
    public class Matrix : StringControlAddInBase, IValueControlAddInDefinition<string>
    {

        private Timer Timer1 = new Timer();
        bool RefreshNeeded;
        string xValue;
        private DataGridView dataGridView1 = new DataGridView();
        protected override Control CreateControl()
        {
            //dataGridView1.Scroll += dataGridView1_Scroll;
            dataGridView1.CellDoubleClick +=dataGridView1_CellDoubleClick;
            Timer1.Interval = 500;
            Timer1.Enabled = true;
            Timer1.Tick += new EventHandler(Timer1Tick);

            addScrollListener(dataGridView1);
            
            InitializeDataGridView();
            return dataGridView1;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                this.RaiseControlAddInEvent(e.ColumnIndex, e.RowIndex.ToString());
                //ControlAddIn(e.ColumnIndex, e.RowIndex.ToString());
            }
            catch(Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
            
        }

        private void Timer1Tick(object sender, EventArgs e)
        {
            // todo: consider redesign
            if (Value == "")
            {
                Timer1.Dispose();
            }
            else
            {
                RefreshNeeded = !Value.Equals(xValue, StringComparison.Ordinal);
                if (RefreshNeeded)
                {
                    RefreshNeeded = false;
                    xValue = Value;
                 
                    try
                    {                        
                        string tmpFilePath = Path.Combine(Path.GetTempPath().ToString(), Path.GetRandomFileName().ToString() + ".xml");
                        System.IO.File.WriteAllText(tmpFilePath, Value);
                        SetXMLDoc(tmpFilePath);                     
                        File.Delete(tmpFilePath);
                        
                        // todo: replace with direct loading of Value to XML:
                        //XmlDocument xmlDoc1 = new XmlDocument();
                        //xmlDoc1.LoadXml(Value);
                    }
                    catch(Exception e1)
                    {
                        Timer1.Enabled = false;
                        Timer1.Dispose();
                        MessageBox.Show(e1.Message);
                    }                    
                }
            }
        }

        //[ApplicationVisible]
        public void SetXMLDoc(string xmlDocPath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlDocPath);
            LoadXMLDoc(xmlDoc);
        }
        
        private void InitializeDataGridView()
        {
   
            dataGridView1.BorderStyle = BorderStyle.None;
            dataGridView1.BackgroundColor = Color.GhostWhite;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.AdvancedCellBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            dataGridView1.AdvancedCellBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
            dataGridView1.AdvancedCellBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
            dataGridView1.AdvancedCellBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.Single;

            for (int i = 0; i <= 10; i++)
            {
                dataGridView1.Columns.Add("", "");
            }
        }

        //[ApplicationVisible]
        public void ClearDataGridView()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
        }

        private void LoadXMLDoc(XmlDocument xmlDoc)
        {
            ClearDataGridView();
            int rowID = 0;
            int totalColNum = 0;

            foreach (XmlNode recNode in xmlDoc.DocumentElement.ChildNodes)
            {

                int columnID = 0;
                if (rowID == 0)
                {
                    foreach (XmlNode columnNode in recNode.ChildNodes)
                    {
                        AddColumn(columnNode.Name, columnNode.Attributes["Caption"].Value);                                          
                    }
                }

                AddRow((recNode.Attributes["Caption"] != null) && (recNode.Attributes["Caption"].Value != "") ? recNode.Attributes["Caption"].Value  : "");

                foreach (XmlNode columnNode in recNode.ChildNodes)
                {
                    SetCellValue(rowID, columnID, columnNode.InnerText);
                    columnID++;
                }
                rowID++;
                totalColNum = columnID;
                columnID = 0;
            }
            
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            
            AddColumn("","");
            dataGridView1.Columns[totalColNum].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            bool changeColor = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (changeColor)
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(242,247,255);
                }
                changeColor = !changeColor;    
            }
        }
        
        //[ApplicationVisible]
        public void SetCellValue(int rowID, int columnID, string value)
        {
            dataGridView1.Rows[rowID].Cells[columnID].Value = value;
        }

        //[ApplicationVisible]
        public void AddRow(string RowCaption)
        {
            int rowID = dataGridView1.Rows.Add();
            if (RowCaption != "")
            {
                dataGridView1.Rows[rowID].HeaderCell.Value = RowCaption;
                dataGridView1.RowHeadersVisible = true;
            }
        }

        //[ApplicationVisible]
        public void AddColumn(string columnName, string columnCaption)
        {
            dataGridView1.Columns.Add(columnName, columnCaption);
        }
        
        public void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            MessageBox.Show(HasValueChanged.ToString());
            s_Scroll(sender, e);
        }

        bool addScrollListener(DataGridView dgv)
        {
            bool ret = false;

            Type t = dgv.GetType();
            PropertyInfo pi = t.GetProperty("HorizontalScrollBar", BindingFlags.Instance | BindingFlags.NonPublic);
            ScrollBar s = null;

            if (pi != null)
                s = pi.GetValue(dgv, null) as ScrollBar;

            if (s != null)
            {
                s.Scroll += new ScrollEventHandler(s_Scroll);
                ret = true;
            }

            return ret;
        }

        void s_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll && e.Type == ScrollEventType.EndScroll)
            {               
                //ControlAddIn(1, "");
            }
        }

        #region IObjectControlAddInDefinition Member

        //public event ControlAddInEventHandler ControlAddIn;

        #endregion    
    }
}
