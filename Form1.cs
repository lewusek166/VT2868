using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace VT2868
{
    public partial class Form1 : Form
    {
        
        string[,] program;
        Excel.Application ap;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.Range range;
        int ileOperacji;
        public Form1()
        {
            InitializeComponent();
            openFileDialog1.InitialDirectory = "@c:\\";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            ap = new Excel.Application();
           
        }
        void SprawdzPoprawnosc(string[,] tab)
        {
            
            ///jak zdjęcie to nie film i na odwrót
            for(int q = 0; ileOperacji >= q; q++)
            {
                if (tab[q,0]!=null&&tab[q,0]!="")
                {
                    if(tab[q, 2] != null && tab[q, 2] != "")
                    {
                        label3.Text = "Zastanów się czy chcesz zdjęcie poglądowe czy film może być tylko jedno pozycja:" + (q + 1).ToString();
                        break;
                    }
                }
                    
            }
            
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filestream = openFileDialog1.FileName;
                TranslatorPobranie(filestream);
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
            this.Close();
        }
        void TranslatorPobranie(string path)
        {
            string[,] tab;
            int i = 2;
            wb = ap.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);

            while (ws.Cells[i, 1].Value2 != "" && ws.Cells[i, 1].Value2 != null)
            {
                i++;
            }
            tab = new string[i-2, 5];
            for (int z = 0; z < i-2; z++)
            {
                //ws.Cells[z + 2, 4].ToString();
                tab[z, 0] = Convert.ToString(ws.Cells[z + 2, 2].Value2);//zdjęcie poglądeowe 
                tab[z, 1] = Convert.ToString(ws.Cells[z + 2, 3].Value2);//opis zdjecie
                tab[z, 2] = Convert.ToString(ws.Cells[z + 2, 4].Value2);//wideo
                tab[z, 3] = Convert.ToString(ws.Cells[z + 2, 5].Value2);//minimalny czas
                tab[z, 4] = Convert.ToString(ws.Cells[z + 2, 6].Value2);//opis do raportu 
            }
            
            program = tab;
            ileOperacji = i - 2;
            wb.Close(false, null, null);
            ap.Quit();
            SprawdzPoprawnosc(program);
        }
    }
}
