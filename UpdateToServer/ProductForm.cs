using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UpdateToServer
{
    public partial class 商品编码 : Form
    {
        public 商品编码()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            product one = new product();
            one.Spdm = textBox1.Text;
            one.SpdmGB = textBox2.Text;
            one.Spmc = textBox3.Text;
            using (var db = SugarDao.GetInstance())
            {
                db.Insert<product>(one);
                MessageBox.Show("添加成功");
            }
            ProductForm_Load(sender, e);
        }

        private void ProductForm_Load(object sender, EventArgs e)
        {
            List<product> productList;
            using (var db = SugarDao.GetInstance())
            {
                string sql = string.Format("SELECT spdmgb,Spmc,Spdm from product order by spdm desc ");
                productList = db.SqlQuery<product>(sql);
            }
            dataGridView1.DataSource = productList;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            List<product> productList;
            using (var db = SugarDao.GetInstance())
            {
                int x;
                x = dataGridView1.CurrentRow.Index;
                string spdmgb = dataGridView1.Rows[x].Cells[0].Value.ToString();
                string Spmc = dataGridView1.Rows[x].Cells[1].Value.ToString();
                string Spdm = dataGridView1.Rows[x].Cells[2].Value.ToString();
                string sql = string.Format("delete from product where spdmgb='{0}'and Spmc='{1}' and Spdm='{2}'",spdmgb,Spmc,Spdm);

                product delone = new product();
                delone.SpdmGB = spdmgb;
                delone.Spmc = Spmc;
                delone.Spdm = Spdm;

                if (!db.Delete<product>(string.Format("spdmgb='{1}' and Spmc='{2}' and Spdm='{0}'", spdmgb, Spmc, Spdm))){
                    MessageBox.Show("删除失败");
                }
               
                sql = string.Format("SELECT spdmgb,Spmc,Spdm from product order by spdm desc ");
                productList = db.SqlQuery<product>(sql);
            }
            dataGridView1.DataSource = productList;
        }
    }
}
