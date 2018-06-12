using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace MyContrals
{

     public partial class BusyForm : Form
    {

        [DllImport("user32")]
        public static extern long SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


        private Form TheParentForm;


        private DateTime Begin_time = DateTime.Now;

        private Label label2;
        private Label label1;
        private ProgressBar progressBar1;
        private PictureBox pictureBox1;
        private System.Windows.Forms.Timer timer1;
        private System.ComponentModel.IContainer components;
        private SplitContainer splitContainer1;
        private Label label3;



        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label3 = new System.Windows.Forms.Label();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(8, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 17);
            this.label2.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "总用时:0天0小时0分0秒";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar1.Location = new System.Drawing.Point(0, 0);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(389, 29);
            this.progressBar1.TabIndex = 6;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::MyContrals.Properties.Resources.index;
            this.pictureBox1.Location = new System.Drawing.Point(138, 55);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(192, 187);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(9, 255);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 17);
            this.label3.TabIndex = 9;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(6, 275);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.progressBar1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.label2);
            this.splitContainer1.Size = new System.Drawing.Size(457, 29);
            this.splitContainer1.SplitterDistance = 389;
            this.splitContainer1.TabIndex = 10;
            // 
            // BusyForm
            // 
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(469, 310);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label3);
            this.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "BusyForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.BusyForm_FormClosing);
            this.Load += new System.EventHandler(this.BusyForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void BusyForm_Load(object sender, EventArgs e)
        {


            // SetParent(this.Handle, TheParentForm.Handle);
            //this.label1.Text = "";
            //this.label2.Text = "";
            //this.label3.Text = "";



            //int OldLeft = this.Left;

            this.Top = TheParentForm.Top + (TheParentForm.Height - this.Height) / 2;// - this.Top * 2;
            this.Left = TheParentForm.Left + (TheParentForm.Width - this.Width) / 2;// - OldLeft;


            //this.Top = 0;// (TheParentForm.Height - this.Height) / 2;
            //this.Left = 0; //(TheParentForm.Width - this.Width) / 2;;


            //Point dp = new Point(); //定义一个坐标结构
            ////PointToScreen
            //dp.X = TheParentForm.Top;
            //dp.Y = TheParentForm.Left;

            //dp = PointToScreen(dp);
            //dp = PointToClient(dp);
            //this.Top = dp.X;
            //this.Left = dp.Y;

            //dp.X = MousePosition.X; //用mousepostion获取鼠标在屏幕中的X轴位置

            //dp.Y = MousePosition.Y;//用mousepostion获取鼠标在屏幕中的Y轴位置





            //dp = this.PointToClient(dp);  //this 屏幕点的位置转化成工作区坐标


            //progressBar1.Maximum = 0;
            //progressBar1.Minimum = 0;
            //progressBar1.Value = 0;
            //progressBar1.Visible = false;
            //this.TopMost = true;
            //this.FormBorderStyle = FormBorderStyle.None;

            //this.timer1.Enabled = true;
            this.timer1.Start();
            TheParentForm.Enabled = false;

            //foreach (Control cc in TheParentForm.Controls)
            //{
            //    cc.Enabled = false;
            //}
        }

        public BusyForm(Form ParentForm)
        //public BusyForm()
        {



            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;



            TheParentForm = ParentForm;


            //progressBar1.Maximum = 0;
            //progressBar1.Minimum = 0;
            //progressBar1.Value = 0;
            //progressBar1.Visible = false;
            //this.TopMost = true;
            //this.FormBorderStyle = FormBorderStyle.None;

        }


        private delegate void SetTopTextHandler(string text);
        private delegate void SetRightTextHandler(string text);
        private delegate void SetLeftTextHandler(string text);
        private delegate void SetProgressBarMaxHandler(int ProgressBarMax, string StepMemo);
        private delegate void ProgressBarGrowHandler();



        /// <summary>
        /// 置头顶上的label(label1)
        /// </summary>
        /// <param name="text">想要填入的文字</param>
        public void SetTopText(string text)
        {
            if (this.label1.InvokeRequired)
            {
                this.Invoke(new SetTopTextHandler(SetTopText), text);
            }
            else
            {
                this.label1.Text = text;
            }
        }


        /// <summary>
        /// 置右下角的label(label2)
        /// </summary>
        /// <param name="text">想要填入的文字</param>
        public void SetRightText(string text)
        {
            if (this.label2.InvokeRequired)
            {
                this.Invoke(new SetRightTextHandler(SetRightText), text);
            }
            else
            {
                this.label2.Text = text;
            }
        }

        /// <summary>
        /// 置左下角的label(label3)
        /// </summary>
        /// <param name="text">想要填入的文字</param>
        public void SetLeftText(string text)
        {
            if (this.label3.InvokeRequired)
            {
                this.Invoke(new SetLeftTextHandler(SetLeftText), text);
            }
            else
            {
                this.label3.Text = text;
            }
        }



        public void SetProgressBarMax(int ProgressBarMax, string StepMemo)
        {

            if (this.progressBar1.InvokeRequired)
            {

                this.Invoke(new SetProgressBarMaxHandler(SetProgressBarMax), ProgressBarMax, StepMemo);


            }
            else
            {
                if (ProgressBarMax == 0)
                {
                    progressBar1.Visible = false;
                }
                else
                {
                    progressBar1.Maximum = ProgressBarMax;
                    progressBar1.Value = 0;
                    progressBar1.Visible = true;

                }
                label3.Text = StepMemo.ToString().Trim();

            }


        }

        public void ProgressBarGrow()
        {
            if (this.progressBar1.InvokeRequired)
            {
                this.Invoke(new ProgressBarGrowHandler(ProgressBarGrow));
            }
            else
            {
                if (this.progressBar1.Value < this.progressBar1.Maximum)
                {
                    progressBar1.Value += 1;
                    //this.label2.Text = Convert.ToString(Math.Round(Convert.ToDecimal(progressBar1.Value / progressBar1.Maximum) * 10000) / 100) + "%";
                    SetRightText(Convert.ToString(Math.Round(Convert.ToDecimal(progressBar1.Value * 1.0 / progressBar1.Maximum) * 10000) / 100) + "%");

                }
            }


        }


        private void timer1_Tick(object sender, EventArgs e)
        {

            DateTime Date1 = Begin_time;
            DateTime Date2 = DateTime.Now;

            TimeSpan d3 = Date2.Subtract(Date1);

            if (this.label1.InvokeRequired)
            {
                this.Invoke(new SetTopTextHandler(SetTopText), "总用时：" + d3.Days.ToString() + "天" + d3.Hours.ToString() + "小时" + d3.Minutes.ToString() + "分钟" + d3.Seconds.ToString() + "秒");
            }
            else
            {
                this.label1.Text = "总用时：" + d3.Days.ToString() + "天" + d3.Hours.ToString() + "小时" + d3.Minutes.ToString() + "分钟" + d3.Seconds.ToString() + "秒";
            }


        }

        private void BusyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //foreach (Control cc in TheParentForm.Controls)
            //{
            //    cc.Enabled = true;
            //}
            TheParentForm.Enabled = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }



    /// <summary>  
    /// Using Singleton Design Pattern  
    /// </summary>  
    public class WaitFormService
    {

        /// <summary>  
        /// 单例模式  
        /// </summary>  
        public static WaitFormService Instance
        {
            get
            {
                if (WaitFormService._instance == null)
                {
                    lock (syncLock)
                    {
                        if (WaitFormService._instance == null)
                        {
                            WaitFormService._instance = new WaitFormService();
                        }
                    }
                }
                return WaitFormService._instance;
            }
        }

        /// <summary>  
        /// 为了单例模式防止new 实例化..  
        /// </summary>  
        private WaitFormService()
        {
        }


        private static WaitFormService _instance;
        private static readonly Object syncLock = new Object();
        //private Thread waitThread;
        private static BusyForm waitForm;
        /// <summary>  
        /// 显示等待窗体  
        /// </summary>  
        //public static void Show()
        public static void Show(Form ParentForm)
        {

            //try  
            //{              
            WaitFormService.Instance._CreateForm(ParentForm);
            Thread.Sleep(100);
            //isclose = false;  
            //}  
            //catch (Exception ex)  
            //{   
            //}  
        }

        /// <summary>  
        /// 关闭等待窗体  
        /// </summary>  
        public static void Close()
        {

            //if (isclose == true)  
            //{  
            //    return;  
            //}  
            //try  
            //{  
            Thread.Sleep(100);
            WaitFormService.Instance._CloseForm();
            //isclose = true;  
            //}  
            //catch (Exception ex)  
            //{   
            //}  
        }

        /// <summary>  
        /// 设置等待窗体标题  
        /// </summary>  
        /// <param name="text"></param>  
        public static void SetTopText(string text)
        {
            //if (isclose == true)  
            //{  
            //    return;  
            //}  
            //try  
            //{  
            WaitFormService.Instance.SetWaiteTopText(text);
            //}  
            //catch (Exception ex)  
            //{   
            //}  
        }


        public static void SetLeftText(string text)
        {
            //if (isclose == true)  
            //{  
            //    return;  
            //}  
            //try  
            //{  
            WaitFormService.Instance.SetWaiteLeftText(text);
            //}  
            //catch (Exception ex)  
            //{   
            //}  
        }


        public static void SetRightText(string text)
        {
            //if (isclose == true)  
            //{  
            //    return;  
            //}  
            //try  
            //{  
            WaitFormService.Instance.SetWaiteRightText(text);
            //}  
            //catch (Exception ex)  
            //{   
            //}  
        }

        public static void SetProgressBarMax(int ProgressBarMax, string StepMemo)
        {
            WaitFormService.Instance.SetWaitProgressBarMax(ProgressBarMax, StepMemo);
        }

        public static void ProgressBarGrow()
        {
            WaitFormService.Instance.WatiProgressBarGrow();
        }

        /// <summary>  
        /// 创建等待窗体  
        /// </summary>  
        //public void _CreateForm()
        public void _CreateForm(Form ParentForm)
        {
            waitForm = null;
            //waitThread = new Thread(new ThreadStart(this._ShowWaitForm));
            //waitThread = new Thread(new ThreadStart(this._ShowWaitForm(ParentForm)));
            //waitThread.Start();



            Thread waitThread = new Thread(new ParameterizedThreadStart(this._ShowWaitForm));

            waitThread.Start(ParentForm);  //启动异步线程


            Thread.Sleep(100);
        }


        //private void _ShowWaitForm()
        private void _ShowWaitForm(object ParentForm)
        {
            try
            {
                //waitForm = new BusyForm();
                waitForm = new BusyForm((Form)ParentForm);
                waitForm.ShowDialog();

            }
            catch (ThreadAbortException )
            {
                waitForm.Close();
                Thread.ResetAbort();
            }
        }

        /// <summary>  
        /// 关闭窗体  
        /// </summary>  
        private void _CloseForm()
        {

            //waitForm.Close();  
            //waitForm = null;  
            if (waitForm != null)
            {
                Application.DoEvents();
                waitForm.Close();
                //waitThread.Abort();
            }
        }


        /// <summary>  
        /// 设置窗体标题  
        /// </summary>  
        /// <param name="text"></param>  
        public void SetWaiteTopText(string text)
        {
            if (waitForm != null)
            {
                //try  
                //{  

                waitForm.Show();

                waitForm.SetTopText(text);

                //}  
                //catch (Exception ex)  
                //{   
                //}  
            }
        }

        public void SetWaiteLeftText(string text)
        {
            if (waitForm != null)
            {
                //try  
                //{  

                waitForm.Show();

                waitForm.SetLeftText(text);

                //}  
                //catch (Exception ex)  
                //{   
                //}  
            }
        }


        public void SetWaiteRightText(string text)
        {
            if (waitForm != null)
            {
                //try  
                //{  

                waitForm.Show();

                waitForm.SetRightText(text);

                //}  
                //catch (Exception ex)  
                //{   
                //}  
            }
        }

        public void SetWaitProgressBarMax(int ProgressBarMax, string StepMemo)
        {
            if (waitForm != null)
            {
                //try  
                //{  

                waitForm.Show();

                waitForm.SetProgressBarMax(ProgressBarMax, StepMemo);

                //}  
                //catch (Exception ex)  
                //{   
                //}  
            }
        }

        public void WatiProgressBarGrow()
        {
            if (waitForm != null)
            {
                //try  
                //{  

                waitForm.Show();

                waitForm.ProgressBarGrow();

                //}  
                //catch (Exception ex)  
                //{   
                //}  
            }
        }

    }
}
