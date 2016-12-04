namespace EM_Client
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.serstatus = new System.Windows.Forms.Label();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.svip = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.svport = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.svstatus = new System.Windows.Forms.ToolStripLabel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.站别绑定ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Modelcombo = new System.Windows.Forms.ComboBox();
            this.StationLab = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.GridView = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.bs = new System.Windows.Forms.DataGridViewImageColumn();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Step = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EntBut = new System.Windows.Forms.DataGridViewButtonColumn();
            this.toolStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridView)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(619, 447);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(64, 34);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // serstatus
            // 
            this.serstatus.AutoSize = true;
            this.serstatus.Location = new System.Drawing.Point(454, 458);
            this.serstatus.Name = "serstatus";
            this.serstatus.Size = new System.Drawing.Size(35, 13);
            this.serstatus.TabIndex = 1;
            this.serstatus.Text = "label1";
            // 
            // toolStrip1
            // 
            this.toolStrip1.AutoSize = false;
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.svip,
            this.toolStripSeparator1,
            this.svport,
            this.toolStripSeparator2,
            this.svstatus});
            this.toolStrip1.Location = new System.Drawing.Point(0, 493);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(804, 43);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // svip
            // 
            this.svip.Name = "svip";
            this.svip.Size = new System.Drawing.Size(86, 40);
            this.svip.Text = "toolStripLabel1";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 43);
            // 
            // svport
            // 
            this.svport.Name = "svport";
            this.svport.Size = new System.Drawing.Size(86, 40);
            this.svport.Text = "toolStripLabel1";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 43);
            // 
            // svstatus
            // 
            this.svstatus.Name = "svstatus";
            this.svstatus.Size = new System.Drawing.Size(86, 40);
            this.svstatus.Text = "toolStripLabel1";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(804, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.站别绑定ToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(43, 20);
            this.toolStripMenuItem1.Text = "设置";
            // 
            // 站别绑定ToolStripMenuItem
            // 
            this.站别绑定ToolStripMenuItem.Name = "站别绑定ToolStripMenuItem";
            this.站别绑定ToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
            this.站别绑定ToolStripMenuItem.Text = "站别绑定";
            this.站别绑定ToolStripMenuItem.Click += new System.EventHandler(this.站别绑定ToolStripMenuItem_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.Modelcombo);
            this.panel1.Location = new System.Drawing.Point(700, 72);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(83, 40);
            this.panel1.TabIndex = 6;
            this.panel1.Visible = false;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(232, 102);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(108, 42);
            this.button2.TabIndex = 2;
            this.button2.Text = "Start Assembly";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "请选择要做的Model:";
            // 
            // Modelcombo
            // 
            this.Modelcombo.FormattingEnabled = true;
            this.Modelcombo.Location = new System.Drawing.Point(93, 43);
            this.Modelcombo.Name = "Modelcombo";
            this.Modelcombo.Size = new System.Drawing.Size(181, 21);
            this.Modelcombo.TabIndex = 0;
            // 
            // StationLab
            // 
            this.StationLab.AutoSize = true;
            this.StationLab.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.StationLab.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.StationLab.Location = new System.Drawing.Point(678, 33);
            this.StationLab.Name = "StationLab";
            this.StationLab.Size = new System.Drawing.Size(86, 25);
            this.StationLab.TabIndex = 7;
            this.StationLab.Text = "Station";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button4);
            this.panel2.Controls.Add(this.button3);
            this.panel2.Controls.Add(this.GridView);
            this.panel2.Location = new System.Drawing.Point(33, 33);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(537, 401);
            this.panel2.TabIndex = 8;
            this.panel2.Visible = false;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(378, 179);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(111, 54);
            this.button4.TabIndex = 2;
            this.button4.Text = "组装完成";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(378, 87);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(111, 54);
            this.button3.TabIndex = 1;
            this.button3.Text = "组装异常";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // GridView
            // 
            this.GridView.AllowUserToAddRows = false;
            this.GridView.AllowUserToDeleteRows = false;
            this.GridView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.GridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.GridView.ColumnHeadersHeight = 35;
            this.GridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.bs,
            this.ID,
            this.Step,
            this.EntBut});
            this.GridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.GridView.Location = new System.Drawing.Point(3, 3);
            this.GridView.Name = "GridView";
            this.GridView.ReadOnly = true;
            this.GridView.RowHeadersVisible = false;
            this.GridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.GridView.RowTemplate.Height = 25;
            this.GridView.Size = new System.Drawing.Size(336, 395);
            this.GridView.TabIndex = 0;
            this.GridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GridView_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.label2.Location = new System.Drawing.Point(590, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 25);
            this.label2.TabIndex = 9;
            this.label2.Text = "Station:";
            // 
            // bs
            // 
            this.bs.HeaderText = "标识";
            this.bs.Name = "bs";
            this.bs.ReadOnly = true;
            // 
            // ID
            // 
            this.ID.DataPropertyName = "ID";
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            // 
            // Step
            // 
            this.Step.DataPropertyName = "AssyStep";
            this.Step.FillWeight = 152.2843F;
            this.Step.HeaderText = "组装步骤";
            this.Step.Name = "Step";
            this.Step.ReadOnly = true;
            // 
            // EntBut
            // 
            this.EntBut.DataPropertyName = "But";
            this.EntBut.FillWeight = 47.71573F;
            this.EntBut.HeaderText = "Column1";
            this.EntBut.Name = "EntBut";
            this.EntBut.ReadOnly = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(804, 536);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.StationLab);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.serstatus);
            this.Controls.Add(this.button1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "FCT LED KanBan Client";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label serstatus;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem 站别绑定ToolStripMenuItem;
        private System.Windows.Forms.ToolStripLabel svip;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel svport;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripLabel svstatus;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox Modelcombo;
        private System.Windows.Forms.Label StationLab;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView GridView;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewImageColumn bs;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Step;
        private System.Windows.Forms.DataGridViewButtonColumn EntBut;
    }
}

