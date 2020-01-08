using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace ExcelTool_Xset
{

    public partial class Form1 : Form
    {
        Excel.Application ExcelApp;
        Excel.Workbook WorkBook;
        Excel.Worksheet WorkSheet;

        public Form1()
        {
            InitializeComponent();
            ExcelApp = Globals.ThisAddIn.Application;//ExcelApp获取全局控制权
        }

        private void Form1_Load(object sender, EventArgs e) //窗体启动
        {
            //获取当前运行的Excel
            ExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            richTextBox1.Multiline = true; //实现显示多行
            richTextBox2.Multiline = true;

            richTextBox1.ScrollBars = RichTextBoxScrollBars.Vertical;//设置ScrollBar属性实现只显示垂直滚动
            richTextBox2.ScrollBars = RichTextBoxScrollBars.Vertical;
        }

        private void Button1_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Excel.Worksheet wbk;
            Excel.Range rg;//新建一个Range对象

            wbk = (Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[1];

            rg = wbk.Range["A1:O6"];//设定需要选取的范围
            //把range中的数值输出到richTextBox1
            richTextBox1.Text = rg.ToString();
        }
    }
}
