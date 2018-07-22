using MSFramework.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UpdateToServer
{
    public partial class Form1 : Form
    {
        String ip = "localhost";
        String username = "sa";
        String userpassword = "cyychenQQ";
        private JhInfo newInitInfo()
        {
            JhInfo jhInfo = new JhInfo();
            jhInfo.KhdmGB = "5001120040046";
            jhInfo.Khdm = "0100003";
            jhInfo.Sl = 1.0;
            jhInfo.Cph = "渝A5Y965";
            jhInfo.UserGuid = "643c91d0-003b-4c2b-aa68-3ef39bafceee";
            jhInfo.Remark = "";
            jhInfo.TotalChargeMoneyValue = 0.00;
            jhInfo.CountChargeMoneyValue = 0;
            jhInfo.SbyId = 11;
            jhInfo.IsEnabled = 1;
            jhInfo.Status = 1;
            jhInfo.Kh_mc = "观农贸批发市场";
            jhInfo.SbyXm = "梅巧洪";
            jhInfo.SbyCode = "001";
            jhInfo.IsJz = 2;
            jhInfo.CountJybs = 0;
            jhInfo.JhXzCode = "01";
            jhInfo.JhXzName = "代销";
            jhInfo.KhJfFsCode = "01";
            jhInfo.KhJfFsName = "本地菜";
            jhInfo.SourceType = 1;
            jhInfo.SourceName = "手工录入";
            jhInfo.JcChargeGuid = "";

            return jhInfo;

        }
        private JhInfoList newInitInfoList()
        {
            JhInfoList jhInfoList = new JhInfoList();
            jhInfoList.SpdmGB = "";
            jhInfoList.Cddm = "500000";
            jhInfoList.CddmGB = "500000";
            jhInfoList.Status = 1;
            jhInfoList.JhUnitprice = "0.00";
            jhInfoList.Cdmc = "重庆市";
            jhInfoList.Khdm = "0100003";
            jhInfoList.KhdmGB = "5001120040046";
            jhInfoList.Khmc = "观农贸批发市场";
            jhInfoList.Scz = "";
            jhInfoList.Scjd = "";
            jhInfoList.Lxdh = "";
            jhInfoList.Jczh = "";
            jhInfoList.Cdzh = "";
            jhInfoList.Yjscdm = "01";
            jhInfoList.Yjscmc = "重庆市得立农业发展有限公司";
            jhInfoList.YjscdmGB = "500112004";
            jhInfoList.Ejscdm = "";
            jhInfoList.EjscdmGB = "";
            jhInfoList.Ejscmc = "";
            jhInfoList.IsJz = 2;
            jhInfoList.IsEnabled = 1;
            jhInfoList.Cph = "渝A5Y965";
            jhInfoList.Sppz = "蔬菜";
            jhInfoList.Yjzsm = "";
            jhInfoList.JhlxCode = "1";
            jhInfoList.Jhlxmc = "一级批";
            jhInfoList.JhXzCode = "01";
            jhInfoList.JhXzName = "代销";
            jhInfoList.KhJfFsCode = "01";
            jhInfoList.KhJfFsName = "本地菜";
            jhInfoList.SourceType = 1;
            jhInfoList.SourceName = "手工录入";
            return jhInfoList;

        }
        DataSet ExcelToDS(string Path)
        {
            string strConn = "";
            if (Path.ToLower().IndexOf(".xlsx") > 0)
            {
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Path + "';Extended Properties='Excel 12.0;HDR=YES'";
            }
            else
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + Path + "';Extended Properties='Excel 8.0;HDR=YES;'";
            }
            //strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Path + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";
            DataSet ds = new DataSet();

            string sql = string.Format("select 配送日期 as  send_time ,商品 as  product_name ,配货数 as  counts ,金额  as total_price from [Worksheet$] where 品类='蔬菜'");
            //string sql = string.Format("select {0} as id,送货单编号 as delivery_id,配送日期 as  send_time,客户名称  as customer_name, 品类  as categoty,子品类  as sub_category, 商品 as  product_name ,规格  as specification ,单位  as unit,配货数 as  counts,单价 as  price,金额  as total_price,备注 as  remark from [Worksheet$]", batch_code);
            // string sql = "select * from [Worksheet$]";
           
                OleDbDataAdapter oada = new OleDbDataAdapter(sql, strConn);
                oada.Fill(ds);
                
           
            return ds;

        }
        public Form1()
        {
            InitializeComponent();
        }
       // List<string> invaildProduct=new List<string>();
        Dictionary<string, string> invaildProduct = new Dictionary<string, string>();
        private void 打开文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            invaildProduct.Clear();
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //设置打开对话框的标题
            openFileDialog1.Title = "请选择要打开的文件";
            //设置打开对话框可以多选
            openFileDialog1.Multiselect = true;
            openFileDialog1.FileName = "";
            //设置对话框打开的文件类型
            openFileDialog1.Filter = "Excel文件|*.xls";
            //设置文件对话框当前选定的筛选器的索引
            openFileDialog1.FilterIndex = 2;
            //设置对话框是否记忆之前打开的目录
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择的文件完整路径
                string filePath = openFileDialog1.FileName;
                //获取对话框中所选文件的文件名和扩展名，文件名不包括路径
                string fileName = openFileDialog1.SafeFileName;
                System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1, 0, 0, 0, 0));
                long t = (DateTime.Now.Ticks - startTime.Ticks) / 10000;   //除10000调整为13位    
                string batch_code = "S101" + t.ToString();
                DataSet dataSet;
                try
                {
                     dataSet = ExcelToDS(filePath);
                }
                catch (Exception f)
                {

                    MessageBox.Show("打开Excel文件失败" + f.Message);
                    return;
                }
              

                JhInfo order = newInitInfo();
                List < JhInfoList > orderList = new List<JhInfoList>();
                DataTable table = dataSet.Tables[0];

                dataGridView1.DataSource = table;
                this.dataGridView1.Columns[0].FillWeight = 25;
                this.dataGridView1.Columns[1].FillWeight = 20;
                this.dataGridView1.Columns[2].FillWeight = 25;
                this.dataGridView1.Columns[3].FillWeight = 25;
                this.dataGridView1.Columns[0].HeaderText = "日期";
                this.dataGridView1.Columns[1].HeaderText = "蔬菜名称";
                this.dataGridView1.Columns[2].HeaderText = "重量（斤）";
                this.dataGridView1.Columns[3].HeaderText = "金额";
                String connsql = "server="+ip+";database=StmsCq;uid="+username+";pwd="+userpassword;// 数据库连接字符串,database设置为自己的数据库名，以Windows身份验证
                string year = DateTime.Now.ToString("yy");
                string data = DateTime.Now.ToString(" HH:mm:ss.fff");
                try
                {
                    List<product> productList;
                    using (var db = SugarDao.GetInstance())
                    {
                        string sql = string.Format("SELECT id,spdmgb,Spmc,Spdm from product");
                        productList = db.SqlQuery<product>(sql);
                        using (SqlConnection conn = new SqlConnection())
                        {
                            conn.ConnectionString = connsql;
                            conn.Open(); // 打开数据库连接
                            SqlDataAdapter sqlDa = new SqlDataAdapter("select top 1 JhCode,JhCodeSimple from JhInfo order by JhId desc", conn);
                            DataTable dt = new DataTable();
                            sqlDa.Fill(dt);
                            long JhCode = long.Parse(dt.Rows[0][0].ToString()) + 1;
                            string JhCodeSimple = year + (long.Parse(dt.Rows[0][1].ToString().Substring(3, 3))+1).ToString() + "001";
                            order.JhCodeSimple = JhCodeSimple;
                            order.JhCode = JhCode.ToString();
                            order.CountSplist = dataSet.Tables[0].Rows.Count;
                            order.WriteDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                            order.EditDate = order.WriteDate;
                           
                            DataTable dts = GetTableSchema();
                            SqlBulkCopy bulkCopy = new SqlBulkCopy(conn);
                            bulkCopy.DestinationTableName = "JhInfoList";
                            bulkCopy.BatchSize = dt.Rows.Count;
                            sqlDa = new SqlDataAdapter("select top 1 JhId from JhInfo order by JhId desc", conn);
                            DataTable dt1 = new DataTable();
                            sqlDa.Fill(dt1);
                            long orderListStart = long.Parse(dt1.Rows[0][0].ToString()) ;
                            foreach (DataRow dr in dataSet.Tables[0].Rows)
                            {
                                JhInfoList one = newInitInfoList();
                                DataRow drs = dts.NewRow();
                                foreach (var tmp in productList)
                                {
                                    string product_name = "";
                                    if (dr[1].ToString().Length < 3) {
                                        product_name = dr[1].ToString();
                                    }
                                    else
                                    {
                                        product_name = dr[1].ToString().Substring(2);
                                    }
                                    
                                    if (tmp.Spmc == product_name)
                                    {
                                        one.SpdmGB = tmp.SpdmGB;
                                        one.Spdm = tmp.Spdm;
                                        one.Spmc = tmp.Spmc;
                                        break;
                                    }
                                }
                                if (one.SpdmGB == "") {
                                    string info =  dr[1].ToString() + " 没有该商品代码请添加，或者修改Excel源数据";
                                    try
                                    {
                                        invaildProduct.Add(info, info);
                                    }
                                    catch (Exception)
                                    {

                                        
                                    }
                                    continue;
                                    MessageBox.Show(info);
                                    return;
                                    one.SpdmGB = "000";
                                    one.Spdm = "0000";
                                    one.Spmc = dr[1].ToString();
                                }
                                one.JhId = orderListStart+1;// + dataSet.Tables[0].Rows.IndexOf(dr);
                                one.IndexOfOrder = dataSet.Tables[0].Rows.IndexOf(dr)+1;
                                one.Spzl = dr[2].ToString();
                                one.Spsl = "1.00";
                                
                                one.JhRq = dr[0].ToString().Substring(0,10) + data;
                                order.Jhrq = one.JhRq;
                                one.JhCode = JhCode.ToString();
                                one.JhCodeSimple = order.JhCodeSimple;
                                order.Zl += System.Convert.ToDouble(one.Spzl);
                                orderList.Add(one);
                                drs[0] = one.JhId;
                                drs[1] = one.SpdmGB;
                                drs[2] = one.Spdm;
                                drs[3] = one.CddmGB;
                                drs[4] = one.Cddm;
                                drs[5] = one.IndexOfOrder;
                                drs[6] = one.Spzl;
                                drs[7] = one.Spsl;
                                drs[8] = one.Status;
                                drs[9] = one.JhUnitprice;
                                drs[10] = one.Spmc;
                                drs[11] = one.Cdmc;
                                drs[12] = one.JhRq;
                                drs[13] = one.JhCode;
                                drs[14] = one.JhCodeSimple;
                                drs[15] = one.Khdm;
                                drs[16] = one.KhdmGB;
                                drs[17] = one.Khmc;
                                drs[18] = one.Scz;
                                drs[19] = one.Scjd;
                                drs[20] = one.Lxdh;
                                drs[21] = one.Jczh;
                                drs[22] = one.Cdzh;
                                drs[23] = one.Yjscdm;
                                drs[24] = one.Yjscmc;
                                drs[25] = one.YjscdmGB;
                                drs[26] = one.Ejscdm;
                                drs[27] = one.EjscdmGB;
                                drs[28] = one.Ejscmc;
                                drs[29] = one.IsJz;
                                drs[30] = one.IsEnabled;
                                drs[31] = one.Cph;
                                drs[32] = one.Sppz;
                                drs[33] = one.Yjzsm;
                                drs[34] = one.JhlxCode;
                                drs[35] = one.Jhlxmc;
                                drs[36] = one.JhXzCode;
                                drs[37] = one.JhXzName;
                                drs[38] = one.KhJfFsCode;
                                drs[39] = one.KhJfFsName;
                                drs[40] = one.SourceType;
                                drs[41] = one.SourceName;
                                dts.Rows.Add(drs);
                            }
                            if (invaildProduct.Count > 0) {
                                string errstring = "";
                                foreach (var info in invaildProduct) {
                                    errstring += info.Value + "\n";
                                }
                                MessageBox.Show(errstring);
                                return;
                            }
                            bulkCopy.WriteToServer(dts);

                            sql = "INSERT INTO JhInfo(Jhrq,JhCode,JhCodeSimple,KhdmGB,Khdm,Sl,Zl,Cph,UserGuid,Remark,WriteDate,EditDate,TotalChargeMoneyValue,CountChargeMoneyValue,CountSplist,SbyId,IsEnabled,Status,Kh_mc,SbyXm,SbyCode,IsJz,CountJybs,JhXzCode,JhXzName,KhJfFsCode,KhJfFsName,SourceType,SourceName,JcChargeGuid) VALUES(@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14,@15,@16,@17,@18,@19,@20,@21,@22,@23,@24,@25,@26,@27,@28,@29,@30)";
                            using (SqlCommand cmd = new SqlCommand(sql, conn))
                            {
                               
                                cmd.Parameters.AddWithValue("@1", order.Jhrq);
                                cmd.Parameters.AddWithValue("@2", order.JhCode);
                                cmd.Parameters.AddWithValue("@3", order.JhCodeSimple);
                                cmd.Parameters.AddWithValue("@4", order.KhdmGB);
                                cmd.Parameters.AddWithValue("@5", order.Khdm);
                                cmd.Parameters.AddWithValue("@6", order.Sl);
                                cmd.Parameters.AddWithValue("@7", order.Zl);
                                cmd.Parameters.AddWithValue("@8", order.Cph);
                                cmd.Parameters.AddWithValue("@9", order.UserGuid);
                                cmd.Parameters.AddWithValue("@10", order.Remark);
                                cmd.Parameters.AddWithValue("@11", order.WriteDate);
                                cmd.Parameters.AddWithValue("@12", order.EditDate);
                                cmd.Parameters.AddWithValue("@13", 0.00);
                                cmd.Parameters.AddWithValue("@14", 0);
                                cmd.Parameters.AddWithValue("@15", order.CountSplist);
                                cmd.Parameters.AddWithValue("@16", order.SbyId);
                                cmd.Parameters.AddWithValue("@17", order.IsEnabled);
                                cmd.Parameters.AddWithValue("@18", order.Status);
                                cmd.Parameters.AddWithValue("@19", order.Kh_mc);
                                cmd.Parameters.AddWithValue("@20", order.SbyXm);
                                cmd.Parameters.AddWithValue("@21", order.SbyCode);
                                cmd.Parameters.AddWithValue("@22", order.IsJz);
                                cmd.Parameters.AddWithValue("@23", order.CountJybs);
                                cmd.Parameters.AddWithValue("@24", order.JhXzCode);
                                cmd.Parameters.AddWithValue("@25", order.JhXzName);
                                cmd.Parameters.AddWithValue("@26", order.KhJfFsCode);
                                cmd.Parameters.AddWithValue("@27", order.KhJfFsName);
                                cmd.Parameters.AddWithValue("@28", order.SourceType);
                                cmd.Parameters.AddWithValue("@29", order.SourceName);
                                cmd.Parameters.AddWithValue("@30", order.JcChargeGuid);
                                int count=cmd.ExecuteNonQuery();
                                //return;

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("错误信息：" + ex.Message, "出现错误");
                    return;
                }
                MessageBox.Show("添加完成");
            }
          
        }
        static DataTable GetTableSchema()
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[] {
                new DataColumn("JhId",typeof(long)),
      new DataColumn("SpdmGB",typeof(string)),
new DataColumn("Spdm",typeof(string)),
new DataColumn("CddmGB",typeof(string)),
new DataColumn("Cddm",typeof(string)),
new DataColumn("IndexOfOrder",typeof(int)),
new DataColumn("Spzl",typeof(decimal)),
new DataColumn("Spsl",typeof(decimal)),
new DataColumn("Status",typeof(int)),
new DataColumn("JhUnitprice",typeof(decimal)),
new DataColumn("Spmc",typeof(string)),
new DataColumn("Cdmc",typeof(string)),
new DataColumn("JhRq",typeof(string)),
new DataColumn("JhCode",typeof(string)),
new DataColumn("JhCodeSimple",typeof(string)),
new DataColumn("Khdm",typeof(string)),
new DataColumn("KhdmGB",typeof(string)),
new DataColumn("Khmc",typeof(string)),
new DataColumn("Scz",typeof(string)),
new DataColumn("Scjd",typeof(string)),
new DataColumn("Lxdh",typeof(string)),
new DataColumn("Jczh",typeof(string)),
new DataColumn("Cdzh",typeof(string)),
new DataColumn("Yjscdm",typeof(string)),
new DataColumn("Yjscmc",typeof(string)),
new DataColumn("YjscdmGB",typeof(string)),
new DataColumn("Ejscdm",typeof(string)),
new DataColumn("EjscdmGB",typeof(string)),
new DataColumn("Ejscmc",typeof(string)),
new DataColumn("IsJz",typeof(int)),
new DataColumn("IsEnabled",typeof(int)),
new DataColumn("Cph",typeof(string)),
new DataColumn("Sppz",typeof(string)),
new DataColumn("Yjzsm",typeof(string)),
new DataColumn("JhlxCode",typeof(string)),
new DataColumn("Jhlxmc",typeof(string)),
new DataColumn("JhXzCode",typeof(string)),
new DataColumn("JhXzName",typeof(string)),
new DataColumn("KhJfFsCode",typeof(string)),
new DataColumn("KhJfFsName",typeof(string)),
new DataColumn("SourceType",typeof(int)),
new DataColumn("SourceName",typeof(string))});
            return dt;
        }

        private void 添加ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            商品编码 form = new 商品编码();
            form.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            IniFile myIni = new IniFile(Environment.CurrentDirectory + "\\config.ini");
            ip = myIni.ReadString("db", "ip", "localhost");
            username = myIni.ReadString("db", "username", "sa");
            userpassword = myIni.ReadString("db", "password", "cyychenQQ");
        }

        private void 数据库配置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Setting form = new Setting();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                IniFile myIni = new IniFile(Environment.CurrentDirectory + "\\config.ini");
                ip = myIni.ReadString("db", "ip", "localhost");
                username = myIni.ReadString("db", "username", "sa");
                userpassword = myIni.ReadString("db", "password", "cyychenQQ");
            }
        }
    }
}