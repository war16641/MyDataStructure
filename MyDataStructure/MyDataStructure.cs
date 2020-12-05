using System;
using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace MyDataStructure
{
    public class DataUnit
    {
        public Dictionary<string, object> data = new Dictionary<string, object>();
        public FlatDataModel fdm;

        public DataUnit(FlatDataModel f)
        {
            this.fdm = f;
        }
        public string detail()
        {
            string s = string.Format("共{0:D}个值。\n", this.data.Count);
            foreach (string item in this.data.Keys)
            {
                s += item + ":" + this.data[item].ToString() + "\n";
            }
            return s;
        }
    }

    public class FlatDataModel : IEnumerable
    {
        public List<DataUnit> units = new List<DataUnit>();
        public List<string> vn = new List<string>();//字段名
        public FlatDataModel()
        {

        }


        public IEnumerator GetEnumerator()
        {
            for (int i = 0; i < this.units.Count; i++)
            {
                yield return this.units[i];
            }
        }
        /// <summary>
        /// 从excel文件中载入
        /// 第一行作为vn 后续行作为数据
        /// </summary>
        /// <param name="pathname"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public static FlatDataModel load_from_excel(string pathname, string sheetname = null)
        {
            //create the Application object we can use in the member functions.
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = false;

            string fileName = pathname;

            //open the workbook
            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //select the first sheet       
            Worksheet worksheet;
            if (null == sheetname)
            {
                worksheet = (Worksheet)workbook.Worksheets[1];
            }
            else
            {
                worksheet = (Worksheet)workbook.Worksheets[sheetname];
            }


            //find the used range in worksheet
            Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

            //get an object array of all of the cells in the worksheet (their values)
            object[,] valueArray = (object[,])excelRange.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);

            //access the cells 创建flatdatamodel
            FlatDataModel fdm = new FlatDataModel();

            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
            {
                DataUnit du = new DataUnit(fdm);
                for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                {
                    //access each cell
                    object thiscell = valueArray[row, col];
                    if (row == 1)//第一行 设定vn
                    {
                        fdm.vn.Add(thiscell.ToString());
                    }
                    else
                    {//其他行 赋值du
                        du.data.Add(fdm.vn[col - 1], thiscell);
                    }
                    //System.Diagnostics.Debug.Print(valueArray[row, col].ToString());



                }
                if (du.data.Count != 0)
                {
                    fdm.units.Add(du);
                }

            }

            //clean up stuffs
            workbook.Close(false, Type.Missing, Type.Missing);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

            _excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelApp);

            return fdm;
        }

        public int Count
        {
            get
            {
                return this.units.Count;
            }
        }


        /// <summary>
        /// 在excel中显示自己
        /// 不会block
        /// </summary>
        public void show_in_excel()
        {
            //create the Application object we can use in the member functions.
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;



            //open the workbook
            Workbook workbook = _excelApp.Workbooks.Add();


            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
            //显示VN
            for (int i = 0; i < this.vn.Count; i++)
            {
                worksheet.Cells[1, i + 1] = this.vn[i];
            }
            //显示数据
            for (int i = 0; i < this.Count; i++)
            {
                for (int j = 0; j < this.vn.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = this.units[i].data[this.vn[j]];
                }
            }

        }
    }
}
