using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace Markt
{
    public partial class ViewMarkt : Form

    {
        private string file = @"C:\Desktop\Gewürzliste";
        private ComboBox list = new ComboBox();

        public ViewMarkt()
        {
            InitializeComponent();
            ReadExcel(file, iGrid, comboBox1);      
         }

        static private void ReadExcel(string sFile, DataGridView iGrid, ComboBox comboBox)
        {
            DataTable dataTable = new DataTable();
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                int rCnt = 0;
                xlWorkBook = xlApp.Workbooks.Open(sFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                range = xlWorkSheet.UsedRange;

                //Gehe das ganze Zabellenblatt durch
                for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {

                    //Hier haben wir Zugriff auf jede Zeile
                    if ((range.Cells[rCnt, 1] as Excel.Range).Value2 != null)
                    {
                        try
                        {
                            int rowid = iGrid.Rows.Add();

                            DataGridViewRow row = iGrid.Rows[rowid];

                            row.Cells["GNAME"].Value = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                            row.Cells["NAMENZUSATZ"].Value = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                            row.Cells["ZUTATEN"].Value = (range.Cells[rCnt, 3] as Excel.Range).Value2;
                            row.Cells["GEWICHT"].Value = (range.Cells[rCnt, 4] as Excel.Range).Value2;
                            row.Cells["PREIS"].Value = (range.Cells[rCnt, 5] as Excel.Range).Value2;

                            //Combobox füllen
                            comboBox.Items.Add(row.Cells["GNAME"].Value);
                                                       
                        }
                        catch { }
                    }
                }


                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler in ReadExcel: " + ex.Message);
            }

        }

        private void datenUbertragen_Click(object sender, EventArgs e)
        {

        }

        private void btnExceltable_Click(object sender, EventArgs e)
        {
           
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();

            try
            {
                if (openFileDialog.FileName != "")
                {
                    iGrid.Rows.Clear();
                    ReadExcel(openFileDialog.FileName, iGrid, comboBox1);
                    iGrid.Refresh();
                }
                   

            }
            catch (Exception)
            {
                MessageBox.Show("Es ist ein fehlher beim änderen der Exceltabelle aufgetretten");
                throw;
            }

            openFileDialog.Dispose();
        }

        private void comboBox1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox1.SelectedIndex;
            Object selectedItem = comboBox1.SelectedItem;

            if (selectedItem != null)
                selectedItem.ToString();


        }
  
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            int selectedIndex = comboBox1.SelectedIndex;

            if (selectedIndex != -1)
            {
                textBox1.Text = comboBox1.Text;
                textBox2.Text = iGrid.Rows[selectedIndex].Cells["NAMENZUSATZ"].Value.ToString();
                textBox5.Text = iGrid.Rows[selectedIndex].Cells["ZUTATEN"].Value.ToString();
                textBox4.Text = iGrid.Rows[selectedIndex].Cells["GEWICHT"].Value.ToString();
                textBox3.Text = iGrid.Rows[selectedIndex].Cells["PREIS"].Value.ToString();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

            if (monthCalendar1.Visible)
                monthCalendar1.Visible = false;
            else
            {
                monthCalendar1.Left = pictureBox1.Left + 30;
                monthCalendar1.Top = pictureBox1.Top;
                monthCalendar1.Visible = false;
            }

           
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            textBox6.Text = e.Start.ToShortDateString();
        }
    }
}
