using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
/*using Microsoft.Office.Tools.Excel;*/


namespace ExcelUserControl
{
    internal class Common
    {
        private static Microsoft.Office.Tools.CustomTaskPane CustomTask;
        // 创建显示用户窗体的方法，并对外暴露 关键字 public
        public void ShowCustomTask()
        {
            if (CustomTask == null)
            {
                UserControl1 userControl1 = new UserControl1();
                CustomTask = Globals.ThisAddIn.CustomTaskPanes.Add(userControl1, "任务窗格");
                /*CustomTask.Visible = true;*/
                CustomTask.Width = 720;
                CustomTask.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight; // 可选：将任务窗格固定在右侧
            }
            /*else
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(CustomTask);
                CustomTask = null;
            }*/
            CustomTask.Visible = !CustomTask.Visible;
        }
        // 创建写入用户窗体的文本，将这个文本写到设计的richTextBox中，可以设计
        public void WriteText(string str)
        {
            if (CustomTask == null)
            {
                ShowCustomTask();
            }                                 
            CustomTask.Control.Controls["richTextBox2"].Text = str;
            
        }
        public string HttpPost(string str)
        {
            string url = "https://dev.iflyrpa.com/api/gpt/spark_model/data_assistant";
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            string cookieStr = "JSESSIONID=OTA4ZjZiMjAtMzM5Yi00YTc2LWJhMmItOGU3OWZiNjE4NzNk";

            JObject xinhuoData = new JObject();
            JArray array = new JArray();
            JObject queryData = new JObject();

            queryData.Add(new JProperty("role", "user"));
            queryData.Add(new JProperty("content", str));
            array.Add(queryData);
            xinhuoData.Add(new JProperty("query", array));
            Console.WriteLine(xinhuoData.ToString());

            string xinhuoStr = JsonConvert.SerializeObject(xinhuoData);
            //字符串转换为字节码
            byte[] bs = Encoding.UTF8.GetBytes(xinhuoStr);
            //参数类型，这里是json类型
            //还有别的类型如"application/x-www-form-urlencoded"，不过我没用过(逃
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) iFlyRPAStudio/3.0.1 Chrome/108.0.5359.215 Electron/22.3.7 Safari/537.36";
            httpWebRequest.Headers.Add("Cookie", cookieStr);
            //参数数据长度
            httpWebRequest.ContentLength = bs.Length;
            //设置请求类型
            httpWebRequest.Method = "POST";
            //设置超时时间
            httpWebRequest.Timeout = 20000;
            //将参数写入请求对象中
            httpWebRequest.GetRequestStream().Write(bs, 0, bs.Length);
            //发送请求
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            //读取返回数据
            StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.UTF8);
            string responseContent = streamReader.ReadToEnd();
            streamReader.Close();
            httpWebResponse.Close();
            httpWebRequest.Abort();
            return responseContent;
        }

        public string Jiexi(string str)//string str,先写死后面接收参数
        {
            //解析大模型返回的json字符串
            JObject jsonObject = JObject.Parse(str);
            string textValue = (string)jsonObject["data"]["text"];
            return textValue;  
        }
        public void DisplayText1(string message,int id) // 在这个版本的展示信息中，只显示一条message
        {
            //x,y设计label显示的位置
            int x_label = 10;int y = 0;
            int x_text = 70;
            CustomTask.Control.Controls["panel1"].AutoSize = false;
            //创建一个label，表示是用户或大模型
            System.Windows.Forms.Label label = new  System.Windows.Forms.Label();
            if (0 == id % 2)
            {
                label.Text = "LLM：";
            }
            else
            {
                label.Text = "用户：";
            }
            label.Size = new Size(62,24);
            label.TextAlign = ContentAlignment.MiddleCenter;
            // 将 Label 控件添加到容器中
            CustomTask.Control.Controls["panel1"].Controls.Add(label);
            
            //为了实现可以复制，用TextBox控件试一下
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            textBox.Text = message;
            textBox.AutoSize = true;
            textBox.ReadOnly = true;
            textBox.Multiline = true;
            textBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;//水平和垂直滚动条都有
            textBox.Size = new Size(600, 65);//set box size
            /*textBox.Width = 600;*/
            textBox.BorderStyle = BorderStyle.None;
            // 将 TextBox 控件添加到容器中
            CustomTask.Control.Controls["panel1"].Controls.Add(textBox);
     
            y = 80 * id - 10;
            label.Location = new System.Drawing.Point(x_label, y);
            textBox.Location = new System.Drawing.Point(x_text, y);
            return;

        }       
        public static string JiwxiFormular(string str)
        {
            //从解析好的大模型返回的字符串中二次解析，目的是解析出公式
            string pattern = @"`=(.*?)`";
            Match match = Regex.Match(str, pattern);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "未找到公式";
            }
        }
        public void testformu(string str)
        {
            //测试直接执行公式的方案
            string new_str = JiwxiFormular(str);//二次解析大模型的返回，目的是只要公式部分           
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;// 获取当前活动的工作表           
            if ("未找到公式" == new_str)// 在指定单元格中设置公式
            {
                activeSheet.Range["A1"].Value = "公式有误";
            }
            else
            {
                activeSheet.Range["A1"].Formula = new_str;//"=SUM(A3:A4)"
                /*double result = (double)activeSheet.Range["A1"].Value;*/
            }
        }
    }
}
