namespace GetALotOfMoneyForYears
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lab_NowTime = new System.Windows.Forms.Label();
            this.timer_NowTime = new System.Windows.Forms.Timer(this.components);
            this.txt_NowTime = new System.Windows.Forms.TextBox();
            this.btu_SelfGet = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lab_NowTime
            // 
            this.lab_NowTime.AutoSize = true;
            this.lab_NowTime.Location = new System.Drawing.Point(39, 23);
            this.lab_NowTime.Name = "lab_NowTime";
            this.lab_NowTime.Size = new System.Drawing.Size(65, 12);
            this.lab_NowTime.TabIndex = 0;
            this.lab_NowTime.Text = "当前时间：";
            // 
            // timer_NowTime
            // 
            this.timer_NowTime.Enabled = true;
            this.timer_NowTime.Tick += new System.EventHandler(this.Timer_NowTime_Tick);
            // 
            // txt_NowTime
            // 
            this.txt_NowTime.Location = new System.Drawing.Point(101, 20);
            this.txt_NowTime.Name = "txt_NowTime";
            this.txt_NowTime.ReadOnly = true;
            this.txt_NowTime.Size = new System.Drawing.Size(153, 21);
            this.txt_NowTime.TabIndex = 1;
            this.txt_NowTime.TextChanged += new System.EventHandler(this.Txt_NowTime_TextChanged);
            // 
            // btu_SelfGet
            // 
            this.btu_SelfGet.Location = new System.Drawing.Point(330, 18);
            this.btu_SelfGet.Name = "btu_SelfGet";
            this.btu_SelfGet.Size = new System.Drawing.Size(75, 23);
            this.btu_SelfGet.TabIndex = 2;
            this.btu_SelfGet.Text = "手动刷新";
            this.btu_SelfGet.UseVisualStyleBackColor = true;
            this.btu_SelfGet.Click += new System.EventHandler(this.Btu_SelfGet_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btu_SelfGet);
            this.Controls.Add(this.txt_NowTime);
            this.Controls.Add(this.lab_NowTime);
            this.Name = "Main";
            this.Text = "主页";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lab_NowTime;
        private System.Windows.Forms.Timer timer_NowTime;
        private System.Windows.Forms.TextBox txt_NowTime;
        private System.Windows.Forms.Button btu_SelfGet;
    }
}