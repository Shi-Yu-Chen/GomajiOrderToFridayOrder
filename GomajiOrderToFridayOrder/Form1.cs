using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace GomajiOrderToFridayOrder
{
    public partial class Form1 : Form
    {
        private string appPath_;
        private static Excel.Application _Excel = null;

        public Form1()
        {
            InitializeComponent();

            this.appPath_ = Directory.GetCurrentDirectory();
            textBox1.Text = this.appPath_ + "\\價格設定.xlsx";
        }

        private void initailExcel()
        {
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                _Excel = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                _Excel = obj as Excel.Application;
            }

            _Excel.Visible = true;//設false效能會比較好
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = this.appPath_;
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                            FileStream fs = myStream as FileStream;
                            if (fs != null)
                            {
                                textBox1.Text = fs.Name.ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = this.appPath_;
            openFileDialog1.Filter = "Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                            FileStream fs = myStream as FileStream;
                            if (fs != null)
                            {
                                textBox2.Text = fs.Name.ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {            
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = this.appPath_;
            saveFileDialog1.FileName = "";
            saveFileDialog1.DefaultExt = ".xls";
            saveFileDialog1.Filter = "Excel files 2003~2007 (*.xls)|*.xls|All files (*.*)|*.*";

            saveFileDialog1.ShowDialog();
            this.textBox3.Text = saveFileDialog1.FileName;
    
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string configPath = this.textBox1.Text;
            string inputPath = this.textBox2.Text;
            string outputPath = this.textBox3.Text;

            this.label1.Text = "";
            this.label2.Text = "";
            this.label3.Text = "";

            if (!File.Exists(configPath))
            {
                label1.Text = "設定檔不存在!";
                label1.ForeColor = Color.Red;
                return;
            }

            if (!File.Exists(inputPath))
            {
                label2.Text = "輸入檔案不存在!";
                label2.ForeColor = Color.Red;
                return;
            }

            this.initailExcel();

            Dictionary<string, double> PIDPriceMap = new Dictionary<string, double>();
            if (!this.readPIDPriceConfig(configPath, ref PIDPriceMap))
            {
                label1.Text = "設定檔讀檔失敗";
                return;
            }

            List<Dictionary<string, string>> OrderList = new List<Dictionary<string, string>>();
            if (!this.ReadGomajiOrder(inputPath, ref OrderList))
            {
                label2.Text = "Gomoji 訂單讀取失敗";
                return;
            }

            if (this.WriteFridayOrder(outputPath, OrderList, PIDPriceMap))
            {
                return;
            }
        }

        private bool readPIDPriceConfig(string path, ref Dictionary<string, double> map)
        {
            Excel.Workbook book = null;
            Excel.Range range = null;

            try
            {
                book = _Excel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "工作表1";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
             
                range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    var ProductId = (excelWorksheet.Cells[r, 1] as Excel.Range).Value.ToString();
                    
                    if (map.ContainsKey(ProductId))
                    {
                        continue;
                    }
                    else
                    {                        
                        map.Add(ProductId, Convert.ToDouble((excelWorksheet.Cells[r, 2] as Excel.Range).Value.ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }

        private bool ReadGomajiOrder(string inpath, ref List<Dictionary<string,string>> OrderList)
        {
            Excel.Workbook book = null;
            Excel.Range range = null;

            try
            {
                book = _Excel.Workbooks.Open(inpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "Worksheet";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count;
                int lastUsedCol = range.Column;

                string OrderDate = "";
                string OrderID = "";
                string OrderName = "";
                string ReceveName = "";
                string RecevePhone = "";
                string ReceveAddress = "";
                string OrderNote = "";
                string ProductName = "";
                string PID = "";
                string ProductCount = "";

                int OrderCount = 0;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    if (r == 1)
                    {
                        continue;
                    }

                    //var store_id = Convert.ToString((excelWorksheet.Cells[r, 1] as Excel.Range).Value);
                    Dictionary<string, string> OrderDetail = new Dictionary<string, string>();
                    var tmp  = Convert.ToString((excelWorksheet.Cells[r, 3] as Excel.Range).Value);
                    if (tmp == null)
                    {

                    }
                    else
                    {
                        OrderDate = tmp;
                        OrderCount++;
                        OrderID = String.Format("{0}-{1:0000}", OrderDate, OrderCount);
                        OrderName = Convert.ToString((excelWorksheet.Cells[r, 4] as Excel.Range).Value);
                        ReceveName = Convert.ToString((excelWorksheet.Cells[r, 5] as Excel.Range).Value);
                        RecevePhone = Convert.ToString((excelWorksheet.Cells[r, 7] as Excel.Range).Value);
                        ReceveAddress = Convert.ToString((excelWorksheet.Cells[r, 10] as Excel.Range).Value);
                        OrderNote = Convert.ToString((excelWorksheet.Cells[r, 15] as Excel.Range).Value);
                        OrderNote += Convert.ToString((excelWorksheet.Cells[r, 22] as Excel.Range).Value);
                    }

                    ProductName = Convert.ToString((excelWorksheet.Cells[r, 11] as Excel.Range).Value);
                    PID = Convert.ToString((excelWorksheet.Cells[r, 13] as Excel.Range).Value);
                    ProductCount = Convert.ToString((excelWorksheet.Cells[r, 14] as Excel.Range).Value);

                    OrderDetail.Add("OrderDate", OrderDate);
                    OrderDetail.Add("OrderID", OrderID);
                    OrderDetail.Add("OrderName", OrderName);
                    OrderDetail.Add("ReceveName", ReceveName);
                    OrderDetail.Add("RecevePhone", RecevePhone);
                    OrderDetail.Add("ReceveAddress", ReceveAddress);
                    OrderDetail.Add("OrderNote", OrderNote);
                    OrderDetail.Add("PID", PID);
                    OrderDetail.Add("ProductName", ProductName);
                    OrderDetail.Add("ProductCount", ProductCount);

                    OrderList.Add(OrderDetail);
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }

        private bool WriteFridayOrder(string outpath, List<Dictionary<string, string>> OrderList, Dictionary<string, double> PIDPriceMap)
        {
            Excel.Workbook book = null;

            try
            {
                book = _Excel.Workbooks.Open(this.appPath_ + "\\FridayTemp.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "Worksheet";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);

                //cleanup
                Excel.Range range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count - 1;
                excelWorksheet.get_Range(string.Format("3:{0}", lastUsedRow), Type.Missing).Delete();

                int r = 1;
                foreach (var Orderdetail in OrderList)
                {
                    r++;
                    excelWorksheet.Cells[r, 1].Value = (r - 1).ToString();
                    excelWorksheet.Cells[r, 2].Value = Orderdetail["OrderID"];
                    excelWorksheet.Cells[r, 3].Value = "一般";
                    excelWorksheet.Cells[r, 4].Value = Orderdetail["OrderDate"];
                    excelWorksheet.Cells[r, 5].Value = Orderdetail["OrderDate"];
                    excelWorksheet.Cells[r, 6].Value = Orderdetail["OrderDate"];
                    excelWorksheet.Cells[r, 7].Value = Orderdetail["OrderName"];
                    excelWorksheet.Cells[r, 8].Value = Orderdetail["ReceveName"];
                    excelWorksheet.Cells[r, 9].Value = Orderdetail["RecevePhone"];
                    excelWorksheet.Cells[r, 10].Value = Orderdetail["ReceveAddress"];
                    excelWorksheet.Cells[r, 11].Value = Orderdetail["OrderNote"];
                    excelWorksheet.Cells[r, 12].Value = Orderdetail["ProductName"];
                    excelWorksheet.Cells[r, 13].Value = Orderdetail["PID"];
                    excelWorksheet.Cells[r, 14].Value = "無";
                    excelWorksheet.Cells[r, 15].Value = Orderdetail["ProductCount"];
                    if (!PIDPriceMap.ContainsKey(Orderdetail["PID"]))
                    {
                        label3.Text = String.Format("找不到 {0} 成本，請編輯 價格設定.xlsx 後再試一次", Orderdetail["PID"]);
                        return false;
                    }
                    excelWorksheet.Cells[r, 16].Value = PIDPriceMap[Orderdetail["PID"]].ToString();
                    excelWorksheet.Cells[r, 17].Value = "無";
                }

                //excelWorksheet.get_Range(string.Format("A1:Q{0}", r), Type.Missing).NumberFormatLocal = "@";
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                //book.SaveAs(outpath, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


                /*book.SaveAs(outpath, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution, true, Type.Missing, Type.Missing, Type.Missing);*/
                book.SaveAs(outpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }

            return true;
        }
    }
}
