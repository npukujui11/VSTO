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

namespace TryToGetExcel
{
    public partial class Form1 : Form
    {
        Excel.Application application;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) //窗体启动
        {
            //获取当前运行的Excel
            application = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(application.UserName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Range rg;
            rg = application.Range["B2:C5"];//设定需要选取的范围
            rg.Select();//对设定的范围进行选取
            rg.Interior.Color = Color.Blue;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Range rg;
            rg = application.Range["B2:C5"];
            //获取数据类型为double的活动单元格的值
            double d = application.ActiveCell.Value;
            this.Text = d.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //遍历Excel工作簿的名称
            /*foreach (Excel.Workbook wbk in application.Workbooks) {
             *   MessageBox.Show(wbk.Name);
             *}
             */
            Excel.Workbook wbk = application.Workbooks[1];
            MessageBox.Show(wbk.Name);
        }

        private void button5_Click(object sender, EventArgs e)
        {   //创建一个Excel工作表事件
            Excel.Workbook wbk = application.Workbooks[1];
            Excel.Worksheet wst = wbk.Worksheets["Sheet1"];
            //事件的对象名需要额外的声明
            wst.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(MyEvent);//创建事件
            //wst.SelectionChange -= new Excel.DocEvents_SelectionChangeEventHandler(MyEvent);//取消事件
        }

        private void MyEvent(Excel.Range Target) {
            Target.Merge();
        }
    }
}
 