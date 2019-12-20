using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSharpBasic
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = "Hello World!";
            button1.Text = s;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s = "This is my first Csharp programme!";
            button2.Text = s;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int i = 25;
            if (i > 0)
            {
                MessageBox.Show(i + "是正数");
            }
            else 
            {
                MessageBox.Show(i + "不是正数");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int i = 0;
            while (i < 10) {
                i++;
            }
            this.button4.Text = i.ToString();
        }
    }
}
