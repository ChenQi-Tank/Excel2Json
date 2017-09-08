using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel2Json
{
    public partial class Form1 : Form
    {
        string path;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件|*.xls;*.xlsx";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = ofd.FileName;
                DataTable dt = Excel.GetXlsSheetName(ofd.FileName);
                if (dt != null)
                {
                    foreach(DataRow dr in dt.Rows)
                    {
                        comboBox1.Items.Add(dr[2]);
                    }
                }
                else
                {
                    MessageBox.Show("Excel文件格式异常");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(path))
            {
               DataSet ds = Excel.SelectFromXLS(path, comboBox1.Text);
               if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
               {
                   Transform(ds.Tables[0]);
               }
               else
                   MessageBox.Show("Excel文件中指定工作表格式异常");

            }
            else
                MessageBox.Show("Excel文件不存在");
        }

        void Transform(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("[");
            foreach (DataRow dr in dt.Rows)
            {
                sb.Append("{");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append("\"");
                    sb.Append(dt.Columns[i].ColumnName);
                    sb.Append("\":");
                    sb.Append("\"");
                    sb.Append(dr[i].ToString().Replace("\n",""));
                    sb.Append("\",");
                }
                sb.Remove(sb.Length - 1, 1);
                sb.Append("},");
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append("]");
            richTextBox1.Text = sb.ToString();
        }
    }
}
