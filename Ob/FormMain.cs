using System;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Ob
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
             this.CenterToScreen();

        }

        private bool Ffind(ref Excel.Worksheet ws,  string str, out int row, out int col )
        {
            var found = false;
            row = -1;
            col = -1;

            for (int i = 1; i < 20; i++)
                for (int k = 1; k < 20; k++)
                {
                    string s = ws.Cells[i, k].Text.ToUpper();
                    if (s.Equals(str, StringComparison.Ordinal))
                    {
                        found = true;
                        row = i + 1 ;
                        col = k;
                        break;
                    }

                }
            return found;
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                 textBox1.Text = openFileDialog1.FileName;
                                 
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                textBox2.Text = openFileDialog1.FileName;

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application
            {
                Visible = true
            };


            Excel.Workbook wb2 = app.Workbooks.Open(textBox2.Text);
            Excel.Worksheet ws2 = wb2.Sheets[1];
            

            var j1 = -1;
            var l1 = -1;

  
               
            if (!Ffind(ref ws2, @"АРТИКУЛ", out j1, out l1))
            {
                MessageBox.Show(@"Не найдена колонка <АРТИКУЛ> в листе цены");
                wb2.Close();
                app.Quit();
                return;
            }
             

            var j2 = -1;
            var l2 = -1;

             if (!Ffind(ref ws2, @"ЦЕНА, Р.", out j2, out l2))
            {
                MessageBox.Show(@"Не найдена колонка <ЦЕНА, Р.> в листе цены");
                wb2.Close();
                app.Quit();
                return;
            }
             

            Dictionary<string, string> dic = new Dictionary<string, string>();
            int j = j1;
            while (ws2.Cells[j,l1].Text.Length > 0)
            {
                dic.Add(ws2.Cells[j, l1].Text, ws2.Cells[j, l2].Text);
                j += 1; 
            }
            wb2.Close();

            Excel.Workbook wb = app.Workbooks.Open(textBox1.Text);
            Excel.Worksheet ws = wb.Sheets[1];

        
            var i1 = -1;
            var k1 = -1;
            

            if (!Ffind(ref ws, @"ЦЕНА", out i1, out k1))
            {
                MessageBox.Show(@"Не найдена колонка <Цена> в листе оборотки");
                wb.Close();
                app.Quit();
                return;
            }


            var i2 = -1;
            var k2 = -1;

            if (!Ffind(ref ws, @"КОД", out i2, out k2))
            {
                MessageBox.Show(@"Не найдена колонка <КОД> в листе оборотки");
                wb.Close();
                app.Quit();
                return;
            }

// заполнение цен в оборотке 

            j = i2;
            while (ws.Cells[j, k2].Text.Length > 0)
            {
                var c = ws.Cells[j, k2].Text;
                if (dic.ContainsKey(c))
                    ws.Cells[j, k1] = dic[c];
                else
                    ws.Cells[j, k1] = @"нет данных";
                j++;

                
            }
            wb.Save();
            wb.Close();
            app.Quit();
            Activate();
            MessageBox.Show(@"Завершено проставление цен в оборотную ведомость");
        }

    }
}
