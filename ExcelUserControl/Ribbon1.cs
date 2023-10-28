using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelUserControl
{
    public partial class Ribbon1
    {
        public Excel.Application Excelapp;
        Common common = new Common();  // 实例化类
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Excelapp = Globals.ThisAddIn.Application;
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //输入单元格值按钮
            string str = Excelapp.Range["A1"].Value;
            common.WriteText(str);
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // 显示/隐藏按钮事件
            common.ShowCustomTask();
        }
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string str = Excelapp.Range["A1"].Value;
            common.WriteText(str);
        }
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            /*string str = common.HttpPost("请告诉我一年有多少天");
            common.WriteText(str);*/
            common.testformu("计算A1列的第三行和第四行的公式为：=SUM（A3:A4) ");

        }
    }
}
