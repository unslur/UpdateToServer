using MSFramework.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UpdateToServer
{
    public partial class Setting : Form
    {
        IniFile myIni = new IniFile(Environment.CurrentDirectory + "\\config.ini");
        public Setting()
        {

            InitializeComponent();
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            
            textBox1.Text = myIni.ReadString("db", "ip", "localhost");
            textBox2.Text = myIni.ReadString("db", "username", "sa");
            textBox3.Text = myIni.ReadString("db", "password", "cyychenQQ");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            myIni.WriteString("db", "ip", textBox1.Text);
            myIni.WriteString("db", "username", textBox2.Text);
            myIni.WriteString("db", "password", textBox3.Text);
            MessageBox.Show("更新成功");
            this.DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection())
            {
                try
                {
                    String connsql = "server=" + textBox1.Text + ";database=StmsCq;uid=" + textBox2.Text + ";pwd=" + textBox3.Text;// 数据库连接字符串,database设置为自己的数据库名，以Windows身份验证
                    conn.ConnectionString = connsql;
                    conn.Open(); // 打开数据库连接
                    SqlDataAdapter sqlDa = new SqlDataAdapter("select GETDATE()", conn);
                    DataTable dt = new DataTable();
                    sqlDa.Fill(dt);
                    MessageBox.Show("连接成功");
                }
                catch (Exception f)
                {

                    MessageBox.Show(f.Message);
                }
            }
        }
    }
}
