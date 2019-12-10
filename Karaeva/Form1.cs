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

        public double ComputeCoeff(double[] values1, double[] values2)
        {
            if (values1.Length != values2.Length)
                throw new ArgumentException("Values must be the same length");

            var avg1 = values1.Average();
            var avg2 = values2.Average();

            var sum1 = values1.Zip(values2, (x1, y1) => (x1 - avg1) * (y1 - avg2)).Sum();

            var sumSqr1 = values1.Sum(x => Math.Pow((x - avg1), 2.0));
            var sumSqr2 = values2.Sum(y => Math.Pow((y - avg2), 2.0));

            var result = sum1 / Math.Sqrt(sumSqr1 * sumSqr2);

            return result;
        }

        private void LadTable(string filePath)
        {
            //
            WorkBook workbook = WorkBook.Load(filePath);
            WorkSheet sheet = workbook.WorkSheets.First();

            //int cellValue = sheet["A1"].IntValue;
            int i = 0;
            int j = 0;

            foreach (var cell in sheet["A1:G26"]){
                arrayData[i, j] = cell.Text;
                
                //Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);

                //Console.WriteLine("Cell {0}:{1} has value '{2}'", i,j, cell.Text);
                
                j++;
                if (j > 6){
                    j = 0;
                    i++;
                }
            }

            for (int ii = 0; ii < 26; ii++) {
                FirstdataGridView.Rows.Add(new object[] { arrayData[ii, 0], arrayData[ii, 1], arrayData[ii, 2], arrayData[ii, 3], arrayData[ii, 4], arrayData[ii, 5], arrayData[ii, 6] });
            }
            //заголовки коррел матрицы
            string subText1;
            string subText2;
            string subText3;
            string subText4;
            string subText5;
            string subText6;
            int pos;
            string text;

            text = arrayData[0, 1];
            pos = text.IndexOf('\n');
            subText1 = text.Substring(0, pos);
            CorreldataGridView.Columns[1].HeaderText = subText1;

            text = arrayData[0, 2];
            pos = text.IndexOf('\n');
            subText2 = text.Substring(0, pos);
            CorreldataGridView.Columns[2].HeaderText = subText2;

            text = arrayData[0, 3];
            pos = text.IndexOf('\n');
            subText3 = text.Substring(0, pos);
            CorreldataGridView.Columns[3].HeaderText = subText3;

            text = arrayData[0, 4];
            pos = text.IndexOf('\n');
            subText4 = text.Substring(0, pos);
            CorreldataGridView.Columns[4].HeaderText = subText4;

            text = arrayData[0, 5];
            pos = text.IndexOf('\n');
            subText5 = text.Substring(0, pos);
            CorreldataGridView.Columns[5].HeaderText = subText5;

            text = arrayData[0, 6];
            pos = text.IndexOf('\n');
            subText6 = text.Substring(0, pos);
            CorreldataGridView.Columns[6].HeaderText = subText6;

            double[,] correlMatrix = new double[6, 7];

            //for(int iii = 0; iii < 6; iii++) {
            //    for (int jjj = 0; jjj < 7; jjj++) {
            //        correlMatrix[iii, jjj] = iii * jjj;
            //        Console.WriteLine(correlMatrix[iii, jjj]);
            //    }
            //}

            double[] dataColumn1 = new double[25];
            double[] dataColumn2 = new double[25];
            double[] dataColumn3 = new double[25];
            double[] dataColumn4 = new double[25];
            double[] dataColumn5 = new double[25];
            double[] dataColumn6 = new double[25];

            for (int rowi = 0; rowi < 25; rowi++) {
                dataColumn1[rowi] = Convert.ToDouble(arrayData[rowi + 1, 1]);
                dataColumn2[rowi] = Convert.ToDouble(arrayData[rowi + 1, 2]);
                dataColumn3[rowi] = Convert.ToDouble(arrayData[rowi + 1, 3]);
                dataColumn4[rowi] = Convert.ToDouble(arrayData[rowi + 1, 4]);
                dataColumn5[rowi] = Convert.ToDouble(arrayData[rowi + 1, 5]);
                dataColumn6[rowi] = Convert.ToDouble(arrayData[rowi + 1, 6]);
            }

            correlMatrix[0, 0] = ComputeCoeff(dataColumn1, dataColumn1);
            correlMatrix[0, 1] = ComputeCoeff(dataColumn1, dataColumn2);
            correlMatrix[0, 2] = ComputeCoeff(dataColumn1, dataColumn3);
            correlMatrix[0, 3] = ComputeCoeff(dataColumn1, dataColumn4);
            correlMatrix[0, 4] = ComputeCoeff(dataColumn1, dataColumn5);
            correlMatrix[0, 5] = ComputeCoeff(dataColumn1, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText1, Math.Round(correlMatrix[0, 0], 2), Math.Round(correlMatrix[0, 1], 2), Math.Round(correlMatrix[0, 2], 2), Math.Round(correlMatrix[0, 3], 2), Math.Round(correlMatrix[0, 4], 2), Math.Round(correlMatrix[0, 5], 2), Math.Round(correlMatrix[0, 6], 2) });

            correlMatrix[1, 0] = correlMatrix[0, 1];
            correlMatrix[1, 1] = ComputeCoeff(dataColumn2, dataColumn2);
            correlMatrix[1, 2] = ComputeCoeff(dataColumn2, dataColumn3);
            correlMatrix[1, 3] = ComputeCoeff(dataColumn2, dataColumn4);
            correlMatrix[1, 4] = ComputeCoeff(dataColumn2, dataColumn5);
            correlMatrix[1, 5] = ComputeCoeff(dataColumn2, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText2, Math.Round(correlMatrix[1, 0], 2), Math.Round(correlMatrix[1, 1], 2), Math.Round(correlMatrix[1, 2], 2), Math.Round(correlMatrix[1, 3], 2), Math.Round(correlMatrix[1, 4], 2), Math.Round(correlMatrix[1, 5], 2), Math.Round(correlMatrix[1, 6], 2) });

            correlMatrix[2, 0] = correlMatrix[0, 2];
            correlMatrix[2, 1] = correlMatrix[1, 2];
            correlMatrix[2, 2] = ComputeCoeff(dataColumn3, dataColumn3);
            correlMatrix[2, 3] = ComputeCoeff(dataColumn3, dataColumn4);
            correlMatrix[2, 4] = ComputeCoeff(dataColumn3, dataColumn5);
            correlMatrix[2, 5] = ComputeCoeff(dataColumn3, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText3, Math.Round(correlMatrix[2, 0], 2), Math.Round(correlMatrix[2, 1], 2), Math.Round(correlMatrix[2, 2], 2), Math.Round(correlMatrix[2, 3], 2), Math.Round(correlMatrix[2, 4], 2), Math.Round(correlMatrix[2, 5], 2), Math.Round(correlMatrix[2, 6], 2) });

            correlMatrix[3, 0] = correlMatrix[0, 3];
            correlMatrix[3, 1] = correlMatrix[1, 3];
            correlMatrix[3, 2] = correlMatrix[2, 3];
            correlMatrix[3, 3] = ComputeCoeff(dataColumn4, dataColumn4);
            correlMatrix[3, 4] = ComputeCoeff(dataColumn4, dataColumn5);
            correlMatrix[3, 5] = ComputeCoeff(dataColumn4, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText4, Math.Round(correlMatrix[3, 0], 2), Math.Round(correlMatrix[3, 1], 2), Math.Round(correlMatrix[3, 2], 2), Math.Round(correlMatrix[3, 3], 2), Math.Round(correlMatrix[3, 4], 2), Math.Round(correlMatrix[3, 5], 2), Math.Round(correlMatrix[3, 6], 2) });

            correlMatrix[4, 0] = correlMatrix[0, 4];
            correlMatrix[4, 1] = correlMatrix[1, 4];
            correlMatrix[4, 2] = correlMatrix[2, 4];
            correlMatrix[4, 3] = correlMatrix[3, 4];
            correlMatrix[4, 4] = ComputeCoeff(dataColumn5, dataColumn5);
            correlMatrix[4, 5] = ComputeCoeff(dataColumn5, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText5, Math.Round(correlMatrix[4, 0], 2), Math.Round(correlMatrix[4, 1], 2), Math.Round(correlMatrix[4, 2], 2), Math.Round(correlMatrix[4, 3], 2), Math.Round(correlMatrix[4, 4], 2), Math.Round(correlMatrix[4, 5], 2), Math.Round(correlMatrix[4, 6], 2) });

            correlMatrix[5, 0] = correlMatrix[0, 5];
            correlMatrix[5, 1] = correlMatrix[1, 5];
            correlMatrix[5, 2] = correlMatrix[2, 5];
            correlMatrix[5, 3] = correlMatrix[3, 5];
            correlMatrix[5, 4] = correlMatrix[4, 5];
            correlMatrix[5, 5] = ComputeCoeff(dataColumn6, dataColumn6);

            CorreldataGridView.Rows.Add(new object[] { subText5, Math.Round(correlMatrix[5, 0], 2), Math.Round(correlMatrix[5, 1], 2), Math.Round(correlMatrix[5, 2], 2), Math.Round(correlMatrix[5, 3], 2), Math.Round(correlMatrix[5, 4], 2), Math.Round(correlMatrix[5, 5], 2), Math.Round(correlMatrix[5, 6], 2) });


        }
    }
}
