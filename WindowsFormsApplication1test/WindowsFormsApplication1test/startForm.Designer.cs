namespace WindowsFormsApplication1test
{
    partial class startForm
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
            this.NewSummary = new System.Windows.Forms.Button();
            this.ProcessSummary = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please choose:";
            // 
            // NewSummary
            // 
            this.NewSummary.Location = new System.Drawing.Point(12, 109);
            this.NewSummary.Name = "NewSummary";
            this.NewSummary.Size = new System.Drawing.Size(103, 23);
            this.NewSummary.TabIndex = 1;
            this.NewSummary.Text = "New Summary";
            this.NewSummary.UseVisualStyleBackColor = true;
            this.NewSummary.Click += new System.EventHandler(this.NewSummary_Click);
            // 
            // ProcessSummary
            // 
            this.ProcessSummary.Location = new System.Drawing.Point(130, 109);
            this.ProcessSummary.Name = "ProcessSummary";
            this.ProcessSummary.Size = new System.Drawing.Size(103, 23);
            this.ProcessSummary.TabIndex = 2;
            this.ProcessSummary.Text = "Process Summary";
            this.ProcessSummary.UseVisualStyleBackColor = true;
            this.ProcessSummary.Click += new System.EventHandler(this.ProcessSummary_Click);
            // 
            // startForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(245, 187);
            this.Controls.Add(this.ProcessSummary);
            this.Controls.Add(this.NewSummary);
            this.Controls.Add(this.label1);
            this.Name = "startForm";
            this.Text = "Summary Tool";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button NewSummary;
        private System.Windows.Forms.Button ProcessSummary;
    }
}