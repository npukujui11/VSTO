using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    
    public partial class Form1 : Form
    {
        public Excel.Application ExcelApp;
        public Form1()
        {
            InitializeComponent();
            ExcelApp = Globals.ThisAddIn.Application;//ExcelApp获取全局控制权
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //单击按钮，会把textbox1中的文本，发送给当前单元格
            //ExcelApp.ActiveCell.Value = this.textBox1.Text;//改动range value的一个范例
            this.textBox1.Text = ExcelApp.ActiveCell.get_Address();//文本框获取当前单元格的地址
            Excel.Workbook wbk = ExcelApp.Workbooks.Add();//增加一个工作簿
            Excel.Range rg;
            rg = ExcelApp.Range["B6:D8"];
            int i = ExcelApp.ActiveCell.Count;//获取当前Excel的工作簿个数；
            this.Text = i.ToString();//窗体的标签等于数目
        }
    }
}
