using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using Execl = Microsoft.Office.Interop.Excel;
namespace 累進稅率
{
    class DataProcess
    {
        static public List<exl> _list = new List<exl>();
        static public void ReadFile()
        {
            Execl.Application xlApp = new Execl.Application();
            Execl.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\user\Desktop\練習\Homework\累進稅率\bin\Debug\IcomeTaxTable.xlsx");
            Execl._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Execl.Range xlRange = xlWorksheet.UsedRange;
            for(int i = 2; i <= xlRange.Rows.Count; i++)
            {
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    if (j == 1)
                        Console.Write("\r\n");
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        var item = new exl
                        {
                            Number = Convert.ToInt32(xlRange.Cells[i, 1].Value2.ToString()),
                            Range = xlRange.Cells[i, 2].Value2.ToString(),
                            TaxRate = Convert.ToDouble(xlRange.Cells[i, 3].Value2.ToString())
                        };
                        _list.Add(item);
                        break;
                    }
                }
            }
            _list[0].RangeGap = 0;
            xlWorkbook.Close();
            xlApp.Quit();
            ComputRange();
        }
        static public void ComputRange()
        {
            int i = 0;
            for(i = 1; i < _list.Count-1; i++)
            {
                _list[i].RangeGap = _list[i - 1].RangeGap + Max(_list[i - 1].Range) * Convert.ToDecimal((_list[i].TaxRate - _list[i - 1].TaxRate) * 0.01);
            }
            _list[_list.Count - 1].RangeGap = _list[i - 1].RangeGap + Max(_list[i - 1].Range) * Convert.ToDecimal((_list[i].TaxRate - _list[i - 1].TaxRate) * 0.01);
        }
        static public decimal Max(string num)
        {
            return Convert.ToDecimal(num.Split('~')[1]);
        }
        static public string ComputeValue(string s)
        {
            int i = 0;
            var num = Convert.ToDecimal(s);
            for(i = 0; i < _list.Count-1; i++)
            {
                if(num <= Max(_list[i].Range))
                {
                    return (num * Convert.ToDecimal(_list[i].TaxRate * 0.01) - _list[i].RangeGap).ToString();
                }
            }
            return (num * Convert.ToDecimal(_list[i].TaxRate * 0.01) - _list[i].RangeGap).ToString();
        }
    }
    class exl
    {
        public int Number { get; set; }
        public string Range { get; set; }
        public double TaxRate { get; set; }
        public decimal RangeGap { get; set; }
    }
}
