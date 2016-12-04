namespace EM_Client
{
    partial class StationSetting
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
            this.label1 = new System.Windows.Forms.Label();
            this.StationSetCombo = new System.Windows.Forms.ComboBox();
            this.ModelSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(40, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "选择现在的站别：";
            // 
            // StationSetCombo
            // 
            this.StationSetCombo.FormattingEnabled = true;
            this.StationSetCombo.Location = new System.Drawing.Point(118, 79);
            this.StationSetCombo.Name = "StationSetCombo";
            this.StationSetCombo.Size = new System.Drawing.Size(169, 21);
            this.StationSetCombo.TabIndex = 1;
            // 
            // ModelSave
            // 
            this.ModelSave.Location = new System.Drawing.Point(265, 139);
            this.ModelSave.Name = "ModelSave";
            this.ModelSave.Size = new System.Drawing.Size(93, 40);
            this.ModelSave.TabIndex = 2;
            this.ModelSave.Text = "保 存";
            this.ModelSave.UseVisualStyleBackColor = true;
            this.ModelSave.Click += new System.EventHandler(this.ModelSave_Click);
            // 
            // StationSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 217);
            this.Controls.Add(this.ModelSave);
            this.Controls.Add(this.StationSetCombo);
            this.Controls.Add(this.label1);
            this.Name = "StationSetting";
            this.Text = "StationSetting";
            this.Load += new System.EventHandler(this.StationSetting_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox StationSetCombo;
        private System.Windows.Forms.Button ModelSave;
    }
}