using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TEST
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataTable dt = new DataTable();
            DataColumn column1 = new DataColumn("A", typeof(System.Int32));
            DataColumn column2 = new DataColumn("B", typeof(System.String));
            DataColumn column3 = new DataColumn("C", typeof(System.DateTime));
            DataColumn column4 = new DataColumn("D", typeof(System.Int32));
            DataColumn column5 = new DataColumn("E", typeof(System.Int32));
            dt.Columns.Add(column1);
            dt.Columns.Add(column2);
            dt.Columns.Add(column3);
            dt.Columns.Add(column4);
            dt.Columns.Add(column5);
            DataRow dr;
            for (int i = 0; i < 10; i++)
            {
                dr = dt.NewRow();
                dr["A"] = i + 1;
                dr["B"] = i + 2;
                dr["C"] = DateTime.Now;
                dr["D"] = i + 4;
                dr["E"] = i + 5;
                dt.Rows.Add(dr);
            }

            this.exDataGridView1.AutoGenerateColumns = true;
            this.exDataGridView1.DataSource = dt.DefaultView;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            DataTable dt = GetDgvToTable(this.exDataGridView1);
            Helper.Excel.ExcelHelper a = new Helper.Excel.ExcelHelper(@"d:\1.xls");
            a.DataTableToExcel(dt, "sheet1", true);
            
        }



        /// <summary>
        /// 绑定DataGridView数据到DataTable
        /// </summary>
        /// <param name="dgv">复制数据的DataGridView</param>
        /// <returns>返回的绑定数据后的DataTable</returns>
        public DataTable GetDgvToTable(DataGridView dgv)

        {

            DataTable dt = new DataTable();

            // 列强制转换
            
            for (int count = 0; count < dgv.Columns.Count; count++)

            {

                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());

                dt.Columns.Add(dc);

               

            }

            // 循环行

            for (int count = 0; count < dgv.Rows.Count; count++)

            {

                DataRow dr = dt.NewRow();

                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)

                {

                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);

                }

                dt.Rows.Add(dr);

            }

            return dt;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            // this.Enabled = false;

            MyContrals.WaitFormService.Show(this);

            MyContrals.WaitFormService.SetLeftText("lefttext");
            MyContrals.WaitFormService.SetProgressBarMax(100, "啥意思?");
            MyContrals.WaitFormService.SetRightText("righttext");
            MyContrals.WaitFormService.SetTopText("toptext");

            for (int i = 1; i <= 100; i++)
            {

                MyContrals.WaitFormService.ProgressBarGrow();
                System.Threading.Thread.Sleep(100);
                // Application.DoEvents();
            }

            //System.Threading.Thread.Sleep(10000);
            MyContrals.WaitFormService.Close();

            MessageBox.Show("aac");
            // this.Activate();
            // this.Enabled = true;

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }
    }
}
