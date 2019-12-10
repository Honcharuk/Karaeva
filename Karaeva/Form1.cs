using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;
using System.IO;

namespace Karaeva
{
    public partial class Form1 : Form
    {
        string[,] arrayData = new string[26, 7];
        public Form1()
        {
            InitializeComponent();
        }

        private void openExcel_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();


            openFileDialog1.InitialDirectory = "C:\\Users\\manager\\source\\repos\\Karaeva\\Karaeva\\bin\\Debug";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filename = openFileDialog1.FileName;
                LadTable(filename);
            }
            
        }
        private void LadTable(string filePath)
        {
            //
            WorkBook workbook = WorkBook.Load(filePath);
            WorkSheet sheet = workbook.WorkSheets.First();

            int cellValue = sheet["A1"].IntValue;
            int i = 0;
            int j = 0;

            foreach (var cell in sheet["A1:G26"]){
                arrayData[i, j] = cell.Text;
                
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);


                Console.WriteLine("Cell {0}:{1} has value '{2}'", i,j, cell.Text);


                
                j++;
                if (j > 6){
                    j = 0;
                    i++;
                }
            }
            for (int ii = 0; ii < 25; ii++)
            {
                FirstdataGridView.Rows.Add(new object[] { arrayData[ii, 0], arrayData[ii, 1], arrayData[ii, 2], arrayData[ii, 3], arrayData[ii, 4], arrayData[ii, 5]});
            }
        }
    }
}
