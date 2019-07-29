using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public static DataSet ExcelToDataSet(string filename)  //函数用来读取一个excel文件到DataSet集中  
        {
            string strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                            "Extended Properties=Excel 8.0;" +
                            "data source=" + filename;
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [Sheet1$]";     //"Sheet1"为表单标签页名  
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            DataSet ds;
            ds = new DataSet();
            myCommand.Fill(ds);
            myConn.Close();
            return ds;
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            String input_name = this.comboBox1.Text;
            String input_id = this.comboBox2.Text;
            String s = "0";
            this.label5.Text = s;

            DataSet ds = ExcelToDataSet("D:\\dataset.xls");            //括号中为表格地址  
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string name = ds.Tables[0].Rows[i]["姓名"].ToString();  //Rows[i]["col1"]表示i行"col1"字段  
                string id = ds.Tables[0].Rows[i]["身份证"].ToString();
                if ((input_name == name) && (input_id == id))
                {
                    this.label5.Text = ds.Tables[0].Rows[i]["分数"].ToString();
                    break;
                }

            }

        }

        private static void NewMethod1(OleDbConnection conn)
        {
            conn.Open();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            using (Form dlg = new Form()) //caozuo是窗口类名，确保访问；后面的是构造函数
            {
                dlg.StartPosition = FormStartPosition.Manual; //窗体的位置由Location属性决定
                dlg.Location = (Point)new Size(this.Location.X, this.Location.Y);
                dlg.ClientSize = new System.Drawing.Size(this.Size.Width, this.Size.Height);


                Label lab = new Label();
                lab.Text = "积分政策";
                lab.Font = new Font("宋体", 10, FontStyle.Bold);
                lab.Size = new Size(100, 30);
                lab.Location = new Point(300, 180);
                dlg.Controls.Add(lab);


                dlg.ShowDialog();
            }



        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            using (Form dlg = new Form())
            {
                dlg.StartPosition = FormStartPosition.Manual;
                dlg.Location = (Point)new Size(this.Location.X, this.Location.Y);
                dlg.ClientSize = new System.Drawing.Size(this.Size.Width, this.Size.Height);


                Label lab = new Label();
                lab.Text = "使用方法";
                lab.Font = new Font("宋体", 10, FontStyle.Bold);
                lab.Size = new Size(100, 30);
                lab.Location = new Point(300, 180);
                dlg.Controls.Add(lab);


                dlg.ShowDialog();
            }
        }

    }
}
