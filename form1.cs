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
        public static DataSet ExcelToDataSet(string filename)  //����������ȡһ��excel�ļ���DataSet����  
        {
            string strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                            "Extended Properties=Excel 8.0;" +
                            "data source=" + filename;
            OleDbConnection myConn = new OleDbConnection(strCon);
            string strCom = " SELECT * FROM [Sheet1$]";     //"Sheet1"Ϊ����ǩҳ��  
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

            DataSet ds = ExcelToDataSet("D:\\dataset.xls");            //������Ϊ����ַ  
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string name = ds.Tables[0].Rows[i]["����"].ToString();  //Rows[i]["col1"]��ʾi��"col1"�ֶ�  
                string id = ds.Tables[0].Rows[i]["���֤"].ToString();
                if ((input_name == name) && (input_id == id))
                {
                    this.label5.Text = ds.Tables[0].Rows[i]["����"].ToString();
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
            using (Form dlg = new Form()) //caozuo�Ǵ���������ȷ�����ʣ�������ǹ��캯��
            {
                dlg.StartPosition = FormStartPosition.Manual; //�����λ����Location���Ծ���
                dlg.Location = (Point)new Size(this.Location.X, this.Location.Y);
                dlg.ClientSize = new System.Drawing.Size(this.Size.Width, this.Size.Height);


                Label lab = new Label();
                lab.Text = "��������";
                lab.Font = new Font("����", 10, FontStyle.Bold);
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
                lab.Text = "ʹ�÷���";
                lab.Font = new Font("����", 10, FontStyle.Bold);
                lab.Size = new Size(100, 30);
                lab.Location = new Point(300, 180);
                dlg.Controls.Add(lab);


                dlg.ShowDialog();
            }
        }

    }
}
