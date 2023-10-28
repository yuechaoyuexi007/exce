using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelUserControl
{
   
    public partial class UserControl1 : UserControl
    {
        Common common = new Common();  // 实例化类
        int id = 1;//多线乘得时候肯定会出问题
        public UserControl1()
        {
            InitializeComponent();
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            //第二个富文本的内容，用于展示用户的提问（暂时，后续再深入）

        }
        private void button1_Click(object sender, EventArgs e)
        {
            //公式解释按钮
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //公式修正按钮
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //发送按钮
            string propmt = richTextBox2.Text;//在excel中，给出求A列的第三行第四行的和的公式，仅返回公式
            common.DisplayText1(propmt, id);
            id++;
            string str = common.HttpPost(propmt);
            string new_str = common.Jiexi(str);
            common.DisplayText1(new_str, id);
            id++;
            common.testformu(new_str);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //公式生成按钮
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel1.Controls.Clear();
            id = 1;
        }
    }
}
