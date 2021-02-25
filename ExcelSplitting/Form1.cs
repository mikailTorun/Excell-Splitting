using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel= Microsoft.Office.Interop.Excel;


namespace ExcelSplitting
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        } 

        public static String path;
        DataTableCollection tableCollection;
        private void Button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog1.FileName.ToString();
                textBox2.Text = path;

                using (var Stream = File.Open(openFileDialog1.FileName,FileMode.Open, FileAccess.Read))
                {
                    using (ExcelDataReader.IExcelDataReader reader = ExcelDataReader.ExcelReaderFactory.CreateReader(Stream))
                    {
                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow =true}
                        });
                        tableCollection = result.Tables;
                        comboBox1.Items.Clear();
                        foreach(DataTable table in tableCollection)
                        {
                            comboBox1.Items.Add(table.TableName);
                        }
                        reader.Close();
                        Stream.Close();
                    }
                } 
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Interval = 1000;
            //timer1.Enabled = true;
            button3.Enabled = false;
            button1.Visible = false;
            // Create the ToolTip and associate with the Form container.
            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 100;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(this.button2, "Açık veya arka planda çalışmakta olan TÜM EXCELLERİ KAPATIR");
            
        }
        Boolean lastTerm = false;
        private void split()
        {
            //MessageBox.Show((Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1).ToString());
            //Application.Exit();
            label2.Text = dataGridView1.Rows.Count.ToString();
            int startRowPoint = 0;
            int endRowPoint = 900;


            //int firstIndexOfTable = 0;
            //int lastIndexOfTable = dataGridView1.Rows.Count - 2;


            int lastTermCount = dataGridView1.Rows.Count % 900;

            double ExcelTerm = Math.Ceiling(dataGridView1.Rows.Count / 900.0);

            //MessageBox.Show("iŞLEME ALINACAK SATIR SAYISI =>"+dataGridView1.Rows.Count.ToString() +
            //    "\n KAYIT YOLU=>" + Path.GetDirectoryName(path));
            MessageBox.Show("iŞLEME ALINACAK SATIR SAYISI =>" + dataGridView1.Rows.Count.ToString() +
                "\n KAYIT YOLU=>" + Path.GetDirectoryName(path), "İşlem Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Information);

            progressBar1.Maximum = dataGridView1.Rows.Count - 2;

            for (int i = 0; i < Math.Ceiling(dataGridView1.Rows.Count / 900.0); i++)
            {
                if (dataGridView1.Rows.Count < 900)
                {
                    //createExcel(startRowPoint, lastIndexOfTable);
                    MessageBox.Show("Excel Uzunluğu Ayırma İşlemi İçin Yetersizdir. 900 Satıra Eşit ya da Büyük Olmalıdır. \n Mevcut Uzunluk => " + dataGridView1.Rows.Count.ToString(), "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else
                {
                    //MessageBox.Show("girdi \n i=" + i + "\n i=" + (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1) +
                    //    "\n startRowPoint=" + startRowPoint +
                    //    "\n lastTermCount=" + lastTermCount +
                    //    "\n endRowPoint=" + endRowPoint +
                    //    "\n dataGridView1.Rows.Count=" + dataGridView1.Rows.Count.ToString() 
                    //    );

                    if (i != (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1))
                        createExcel(startRowPoint, endRowPoint - 2);

                    if (i == (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1))
                    {
                        lastTerm = true;
                        createExcel(startRowPoint, (startRowPoint + lastTermCount));

                    }

                    //createExcel(startRowPoint, startRowPoint+ lastTermCount);

                    startRowPoint = endRowPoint - 2;

                    if (i == (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1))
                    {

                        endRowPoint += lastTermCount;
                    }
                    else
                    {
                        endRowPoint += 900;
                    }
                }
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {

            //MessageBox.Show((Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1).ToString());
            //Application.Exit();
            label2.Text = dataGridView1.Rows.Count.ToString();
            int startRowPoint = 0;
            int endRowPoint   = 900;
            

            //int firstIndexOfTable = 0;
           // int lastIndexOfTable = dataGridView1.Rows.Count-2;


            int lastTermCount = dataGridView1.Rows.Count % 900;

            double ExcelTerm = Math.Ceiling(dataGridView1.Rows.Count / 900.0);

            //MessageBox.Show("iŞLEME ALINACAK SATIR SAYISI =>"+dataGridView1.Rows.Count.ToString() +
            //    "\n KAYIT YOLU=>" + Path.GetDirectoryName(path));
            MessageBox.Show("iŞLEME ALINACAK SATIR SAYISI =>" + dataGridView1.Rows.Count.ToString() +
                "\n KAYIT YOLU=>" + Path.GetDirectoryName(path), "İşlem Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Information);

            progressBar1.Maximum = dataGridView1.Rows.Count-2;

            for (int i=0; i < Math.Ceiling(dataGridView1.Rows.Count/900.0) ; i++)
            {
                if (dataGridView1.Rows.Count < 900)
                {
                    //createExcel(startRowPoint, lastIndexOfTable);
                    MessageBox.Show("Excel Uzunluğu Ayırma İşlemi İçin Yetersizdir. 900 Satıra Eşit ya da Büyük Olmalıdır. \n Mevcut Uzunluk => " + dataGridView1.Rows.Count.ToString(), "UYARI",MessageBoxButtons.OK,MessageBoxIcon.Warning);

                }
                else
                {
                    //MessageBox.Show("girdi \n i=" + i + "\n i=" + (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1) +
                    //    "\n startRowPoint=" + startRowPoint +
                    //    "\n lastTermCount=" + lastTermCount +
                    //    "\n endRowPoint=" + endRowPoint +
                    //    "\n dataGridView1.Rows.Count=" + dataGridView1.Rows.Count.ToString() 
                    //    );

                    if (i != (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1))
                        createExcel(startRowPoint, endRowPoint-2);

                    if (i == (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0) - 1))
                    {
                        lastTerm = true;
                        createExcel(startRowPoint, (startRowPoint + lastTermCount));

                    }
                        
                        //createExcel(startRowPoint, startRowPoint+ lastTermCount);

                    startRowPoint = endRowPoint-2;

                    if (i == (int)(Math.Ceiling(dataGridView1.Rows.Count / 900.0)-1))
                    {

                        endRowPoint += lastTermCount;
                    }
                    else
                    {
                        endRowPoint += 900;
                    }
                }
            }
        }
        int term = 1;

        int iter = 0;
       // Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
       
        private void createExcel(int rowNumberStart, int rowNumberEnd)
        {
            //MessageBox.Show("path => " + path + "\n getfullpath =>" + Path.GetDirectoryName(path) );  Application.Exit();
            
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            
            Microsoft.Office.Interop.Excel.Workbook wrkbk = app.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sayfa1 = (Microsoft.Office.Interop.Excel.Worksheet)wrkbk.Sheets[1];
            


            Microsoft.Office.Interop.Excel.Range alan1 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 1];
            alan1.Value2 = "BelgeTipi";
            Microsoft.Office.Interop.Excel.Range alan2 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 2];
            alan2.Value2 = "Tarih";
            Microsoft.Office.Interop.Excel.Range alan3 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 3];
            alan3.Value2 = "SeriNo";
            Microsoft.Office.Interop.Excel.Range alan4 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 4];
            alan4.Value2 = "SiraNo";
            Microsoft.Office.Interop.Excel.Range alan5 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 5];
            alan5.Value2 = "CariKod";
            Microsoft.Office.Interop.Excel.Range alan6 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 6];
            alan6.Value2 = "CariAd";
            Microsoft.Office.Interop.Excel.Range alan7 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 7];
            alan7.Value2 = "StokKod";
            Microsoft.Office.Interop.Excel.Range alan8 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 8];
            alan8.Value2 = "StokAd";
            Microsoft.Office.Interop.Excel.Range alan9 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 9];
            alan9.Value2 = "MiktarN";
            Microsoft.Office.Interop.Excel.Range alan10 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 10];
            alan10.Value2 = "BFiyat";
            Microsoft.Office.Interop.Excel.Range alan11 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 11];
            alan11.Value2 = "Tutar";
            Microsoft.Office.Interop.Excel.Range alan12 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 12];
            alan12.Value2 = "KdvOran";
            Microsoft.Office.Interop.Excel.Range alan13 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 13];
            alan13.Value2 = "KdvTutar";
            Microsoft.Office.Interop.Excel.Range alan14 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 14];
            alan14.Value2 = "TutarN";
            Microsoft.Office.Interop.Excel.Range alan15 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 15];
            alan15.Value2 = "Nakit";
            Microsoft.Office.Interop.Excel.Range alan16 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1, 16];
            alan16.Value2 = "Kredi";

      
            if (iter > 0)
                rowNumberStart -= iter;

            int x = 1;
            try
            {
                Microsoft.Office.Interop.Excel.Range alan ;
                iter = 0;

                for (int a = rowNumberEnd - 30; a < rowNumberEnd; a++)
                {
                    //MessageBox.Show("hatamız=>"+dataGridView1.Rows[rowNumberEnd - 2].Cells[3].Value.ToString());
                    if (dataGridView1.Rows[rowNumberEnd].Cells[3].Value.ToString().Equals(dataGridView1.Rows[rowNumberEnd - x].Cells[3].Value.ToString()))//rowNumberEnd - 2
                    {
                        iter++;
                    }
                    x++;
                }
  
                if (lastTerm)
                    iter = -1;

                int startCell = 2;
                for (int rows = rowNumberStart; rows <= rowNumberEnd-(iter+1); rows++)
                {
                    for (int col = 0; col < dataGridView1.Rows[rows].Cells.Count; col++)
                    {
                        sayfa1.Range["D"+ startCell.ToString()].NumberFormat = "0";
                        alan = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[startCell, col + 1];
                        alan.Value2 = dataGridView1.Rows[rows].Cells[col].Value.ToString();
                        //progressBar1.Value = rows;
                        //label5.Text =  (rows+2).ToString();
                    }
                    startCell++;
                }
                wrkbk.SaveAs(Path.GetDirectoryName(path) + "\\" + tableCollection[comboBox1.SelectedItem.ToString()] + "_term" + term.ToString() + ".xlsx",
                Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show("UYARI \n" + ex.Message);
            }


            label2.Text = term.ToString();
            term++;

            wrkbk.Close();
            app.Quit();


            //MessageBox.Show("UYARI \n" + "progresbar geldği yer" + progressBar1.Value.ToString());
            //if (progressBar1.Value > dataGridView1.Rows.Count - 30)
            //////if (progressBar1.Value == dataGridView1.Rows.Count - 2)
            //////{
            //////    progressBar1.Value = dataGridView1.Rows.Count-2;
                

            //////}
            //////else
            //////{
            //////    //MessageBox.Show("Bir sorun var .", "Progressbar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //////}
        }
        
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
            if (dt.Rows.Count > 0)
                button3.Enabled = true;
          

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string processName = "EXCEL"; // Kapatmak İstediğimiz Program
            Process[] processes = Process.GetProcesses();// Tüm Çalışan Programlar
            foreach (Process process in processes)
            {
                if (process.ProcessName == processName)
                {
                    process.Kill();
                }

            }

            Application.Exit();
        }
      
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //button1.PerformClick();
           // timer1.Start();
            split();
        }

        private void Button3_Click(object sender, EventArgs e)
        {


            //timer1.Start();
            backgroundWorker1.RunWorkerAsync();
            button3.Enabled = false;
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Dosya başarılı bir şekilde kaydedildi.", "İşlem Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            timer1.Stop();
            button3.Enabled = true;

        }
        int sec = 0, min = 0, hour = 0;
        private void Timer1_Tick(object sender, EventArgs e)
        {

                sec++;
                label7.Text = sec.ToString();
                if (sec == 60)
                {
                    min++;
                    label8.Text = min.ToString();
                    sec = 0;
                    if (min == 60)
                    {
                        hour++;
                        label9.Text = hour.ToString();
                        min = 0;
                    }
                }

            

        }
    }
}
