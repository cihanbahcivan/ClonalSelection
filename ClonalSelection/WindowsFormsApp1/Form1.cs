using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using GemBox.Spreadsheet;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        List<ExcelRow> DataSet = new List<ExcelRow>();
        static Random rnd = new Random();
        int ColumnCount = 0;
        int K = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void Run()
        {
            LoadExcel();
            List<ExcelRow> testDataSet = GetTestDataSet(DataSet);
            if (K > testDataSet.Count) MessageBox.Show($"K must not be higher than { testDataSet.Count }", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            List<ExcelRow> ExcludedDataSet = ExcludeTestFromDataset(testDataSet, DataSet);
            List<ExcelRow> Pn = ApplyFormula(testDataSet, ExcludedDataSet);
            List<ExcelRow> ClonnedPn = CloneRows(Pn);
            List<ExcelRow> Shuffled = ShuffleColumn(ClonnedPn);
            List<ExcelRow> Pn2 = ApplyFormula(testDataSet, Shuffled);
            ExcludedDataSet.AddRange(Pn2);
            ApplyFormula2(testDataSet, ExcludedDataSet);

            ShowGrid(dataGridView1, DataSet);
            ShowGrid(dataGridView2, testDataSet);
            ShowGrid(dataGridView3, ClonnedPn);
            ShowGrid(dataGridView4, Shuffled);
            ShowGrid(dataGridView5, ExcludedDataSet);
            ShowGrid(dataGridView6, testDataSet);

            CalculateRate(testDataSet);
        }
        private List<ExcelRow> GetTestDataSet(List<ExcelRow> dataSet)
        {
            double count = Math.Floor(dataSet.Count * 0.30);
            List<ExcelRow> _selectedDataSet = new List<ExcelRow>();
            for (int i = 0; i < count; i++)
            {
                int rndIndex = rnd.Next(1, dataSet.Count);
                if (_selectedDataSet.Any(n => n.RowNumber == dataSet[rndIndex].RowNumber))
                {
                    i -= 1; continue;
                }
                _selectedDataSet.Add(dataSet[rndIndex]);
            }
            return _selectedDataSet;
        }
        private void LoadExcel()
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile workbook = ExcelFile.Load("data.xlsx");
            ExcelWorksheet worksheet = workbook.Worksheets[0];
            ColumnCount = worksheet.Rows[0].AllocatedCells.Count - 1;
            for (int r = 1; r < worksheet.Rows.Count; r++)
            {
                List<int> columnList = new List<int>();
                for (int c = 0; c < ColumnCount; c++)
                    columnList.Add(Convert.ToInt32(worksheet.Cells[r, c].Value));
                DataSet.Add(new ExcelRow
                {
                    RowNumber = r,
                    Data = columnList,
                    y = Convert.ToInt32(worksheet.Cells[r, ColumnCount].Value) == 1
                });
            }
        }
        private List<ExcelRow> ExcludeTestFromDataset(List<ExcelRow> testDataSet, List<ExcelRow> dataSet)
        {
            return dataSet.Where(n => !testDataSet.Any(m => m.RowNumber == n.RowNumber)).ToList();
        }
        private List<ExcelRow> ApplyFormula(List<ExcelRow> testDataSet, List<ExcelRow> dataSet)
        {
            List<ExcelRow> Pn = new List<ExcelRow>();
            foreach (var i in testDataSet)
            {
                List<Tuple<ExcelRow, double>> results = new List<Tuple<ExcelRow, double>>();
                foreach (var x in dataSet)
                {
                    double result = EuclidCalculate(i.Data, x.Data);
                    results.Add(Tuple.Create<ExcelRow, double>(x, result));
                }
                var smallest = results.OrderBy(n => n.Item2).FirstOrDefault();
                Pn.Add(smallest.Item1);
            }
            return Pn;
        }
        private void ApplyFormula2(List<ExcelRow> testDataSet, List<ExcelRow> dataSet)
        {
            foreach (var i in testDataSet)
            {
                List<Tuple<ExcelRow, double>> results = new List<Tuple<ExcelRow, double>>();
                foreach (var x in dataSet)
                {
                    double result = EuclidCalculate(i.Data, x.Data);
                    results.Add(Tuple.Create<ExcelRow, double>(x, result));
                }
                var smallest = results.OrderBy(n => n.Item2).Take(K).ToList();
                bool _Result = smallest.Count(n => n.Item1.y == true) > smallest.Count(n => n.Item1.y == false);
                bool _Result2 = i.y == _Result;
                i.y2 = _Result2;
            }
        }

        private double EuclidCalculate(List<int> x1, List<int> x2)
        {
            double total = 0;
            for (int k = 0; k < x1.Count; k++)
                total += Math.Pow(x1[k] - x2[k], 2);
            return Math.Sqrt(total);
        }
        private List<ExcelRow> CloneRows(List<ExcelRow> Pn)
        {
            List<ExcelRow> _pn = new List<ExcelRow>();
            foreach (var i in Pn)
            {
                _pn.Add(Clone(i));
                _pn.Add(Clone(i));
            }
            return _pn;
        }
        private List<ExcelRow> ShuffleColumn(List<ExcelRow> Pn)
        {
            int rndIndex = rnd.Next(0, ColumnCount - 1);
            List<ExcelRow> _Pn = Clone(Pn);
            List<ExcelRow> _Pn2 = Clone(Pn);
            Shuffle(_Pn);

            for (int i = 0; i < _Pn.Count; i++)
            {
                int _rndIndex = rnd.Next(0, _Pn2.Count - 1);
                _Pn[i].Data[rndIndex] = _Pn2[_rndIndex].Data[rndIndex];
                _Pn2.RemoveAt(_rndIndex);
            }
            return _Pn;
        }
        private void CalculateRate(List<ExcelRow> testDataSet)
        {
            int trueCount = testDataSet.Count(n => n.y2 == true);
            int falseCount = testDataSet.Count(n => n.y2 == false);
            double rate = (trueCount * 100) / testDataSet.Count;
            label3.Text = "Accuracy Rate = %" + rate.ToString();
        }
        private static T Clone<T>(T source)
        {
            if (!typeof(T).IsSerializable)
            {
                throw new ArgumentException("The type must be serializable.", "source");
            }

            if (Object.ReferenceEquals(source, null))
            {
                return default(T);
            }

            System.Runtime.Serialization.IFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            Stream stream = new MemoryStream();
            using (stream)
            {
                formatter.Serialize(stream, source);
                stream.Seek(0, SeekOrigin.Begin);
                return (T)formatter.Deserialize(stream);
            }
        }
        public static void Shuffle<T>(IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rnd.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }

        private void ShowGrid(DataGridView gridView, List<ExcelRow> list)
        {
            List<int> data = list.FirstOrDefault().Data;
            //gridView.Columns.Add("RowNumber", "RowNumber");
            for (int i = 0; i < data.Count; i++)
                gridView.Columns.Add("x" + (i + 1).ToString(), "x" + (i + 1).ToString());
            gridView.Columns.Add("y", "y");
            gridView.Columns.Add("ComparisonResult", "Comparison Result");

            foreach (DataGridViewColumn c in gridView.Columns)
                c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            foreach (ExcelRow r in list.OrderBy(n => n.RowNumber)) {
                DataGridViewRow row = new DataGridViewRow();
                //DataGridViewCell cell4 = new DataGridViewTextBoxCell();
                //cell4.Value = r.RowNumber;
                //row.Cells.Add(cell4);

                foreach (int i in r.Data)
                {
                    DataGridViewCell cell = new DataGridViewTextBoxCell();
                    cell.Value = i;
                    row.Cells.Add(cell);
                }
                DataGridViewCell cell2 = new DataGridViewTextBoxCell();
                cell2.Value = r.y;
                row.Cells.Add(cell2);
                if(gridView.Name == "dataGridView6")
                {
                    DataGridViewCell cell3 = new DataGridViewTextBoxCell();
                    cell3.Value = r.y2;
                    row.Cells.Add(cell3);
                }
                gridView.Rows.Add(row);
            }
            gridView.Refresh();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            K = Convert.ToInt32(textBox2.Text);
            if (K % 2 == 0)
            {
                MessageBox.Show("K must be only odd number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataSet.Clear();
            dataGridView1.Columns.Clear();
            dataGridView2.Columns.Clear();
            dataGridView3.Columns.Clear();
            dataGridView4.Columns.Clear();
            dataGridView5.Columns.Clear();
            dataGridView6.Columns.Clear();

            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
            Run();
        }
    }
}
