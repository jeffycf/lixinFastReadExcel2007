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
            myExcel.loadSheetsInfo();
            SheetData mySheet = myExcel.GetSheetData(0);
            myTimer.Stop();
            richTextBox1.AppendText("耗时:"+(myTimer.ElapsedMilliseconds/1000).ToString ()+"\n");
            StringBuilder ss = new StringBuilder();
            for (int i = 0; i < 10; i++)
            {
                Row oneRow = mySheet.ReadOneRow();
                ss.AppendLine(string.Join("  |  ", oneRow.cells.Select(a => a.value).ToArray()));
            }
            richTextBox1.AppendText(ss.ToString());
            
            MessageBox.Show("ok");
        }
    }
}
