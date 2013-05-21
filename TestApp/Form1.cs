using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LixinFastReadExcel07;
using System.IO;
using System.Diagnostics;

namespace TestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stopwatch myTimer = new Stopwatch();
            myTimer.Start();
            LixinFastReadExcel myExcel = new LixinFastReadExcel(@"d:\website\test2.xlsx");

            richTextBox1.AppendText("初始化耗时:" + (myTimer.ElapsedMilliseconds / 1000.0) + "\n");
            myTimer.Reset();
            myTimer.Start();
            myExcel.loadShareString();
            richTextBox1.AppendText("加载shareString耗时:" + (myTimer.ElapsedMilliseconds / 1000.0) + "\n");
            myTimer.Reset();
            myTimer.Start();
            string[] mySheetNames = myExcel.getSheetName();
            richTextBox1.AppendText("获取所有sheet名称耗时：:" + (myTimer.ElapsedMilliseconds / 1000.0) + "\n");
            for (int i = 0; i < mySheetNames.Length; i++)
            {

                richTextBox1.AppendText(mySheetNames[i] + "   ");
                myTimer.Reset();
                myTimer.Start();
                myExcel.OpenSheet(i);
                richTextBox1.AppendText("打开" + mySheetNames[i] + " 耗时：" + (myTimer.ElapsedMilliseconds / 1000.0) + "\n");
                int total = 0;
                myTimer.Reset();
                myTimer.Start();
                while (myExcel.SheetRead())
                {
                    total++;
                }

                richTextBox1.AppendText("读取行数：" + total + "\n");
                richTextBox1.AppendText("读完" + mySheetNames[i] + "数据 耗时：" + (myTimer.ElapsedMilliseconds / 1000.0)+"\n");
            }
            myExcel.close();
            MessageBox.Show("ok");
        }
    }
}
