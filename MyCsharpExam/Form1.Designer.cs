namespace MyCsharpExam
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.bXlsx = new System.Windows.Forms.Button();
            this.bXml = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.button3 = new System.Windows.Forms.Button();
            this.progressLabel = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.lTime = new System.Windows.Forms.Label();
            this.helpProvider1 = new System.Windows.Forms.HelpProvider();
            this.button1 = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(55, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "XLSX документ";
            // 
            // bXlsx
            // 
            this.bXlsx.Location = new System.Drawing.Point(165, 64);
            this.bXlsx.Name = "bXlsx";
            this.bXlsx.Size = new System.Drawing.Size(75, 23);
            this.bXlsx.TabIndex = 2;
            this.bXlsx.Text = "Открыть";
            this.bXlsx.UseVisualStyleBackColor = true;
            this.bXlsx.Click += new System.EventHandler(this.bXlsx_Click);
            // 
            // bXml
            // 
            this.bXml.Location = new System.Drawing.Point(165, 106);
            this.bXml.Name = "bXml";
            this.bXml.Size = new System.Drawing.Size(75, 23);
            this.bXml.TabIndex = 3;
            this.bXml.Text = "Открыть";
            this.bXml.UseVisualStyleBackColor = true;
            this.bXml.Click += new System.EventHandler(this.bXml_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(58, 243);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(182, 23);
            this.progressBar.TabIndex = 4;
            this.progressBar.Visible = false;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(58, 313);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 5;
            this.button3.Text = "Start";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Start_Click);
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(135, 227);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(0, 13);
            this.progressLabel.TabIndex = 6;
            this.progressLabel.Visible = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.оПрограммеToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(294, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // оПрограммеToolStripMenuItem
            // 
            this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
            this.оПрограммеToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
            this.оПрограммеToolStripMenuItem.Text = "О программе";
            this.оПрограммеToolStripMenuItem.Click += new System.EventHandler(this.About_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(55, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "XML шаблон";
            // 
            // linkLabel1
            // 
            this.linkLabel1.LinkArea = new System.Windows.Forms.LinkArea(0, 5);
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel1.Location = new System.Drawing.Point(135, 63);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(24, 25);
            this.linkLabel1.TabIndex = 9;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "?     ";
            this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel1.UseCompatibleTextRendering = true;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.Help_Click);
            // 
            // linkLabel2
            // 
            this.linkLabel2.LinkArea = new System.Windows.Forms.LinkArea(0, 5);
            this.linkLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline;
            this.linkLabel2.Location = new System.Drawing.Point(123, 103);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(25, 28);
            this.linkLabel2.TabIndex = 10;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "?    ";
            this.linkLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.Help_Click);
            // 
            // lTime
            // 
            this.lTime.AutoSize = true;
            this.lTime.ForeColor = System.Drawing.Color.Lime;
            this.lTime.Location = new System.Drawing.Point(135, 269);
            this.lTime.Name = "lTime";
            this.lTime.Size = new System.Drawing.Size(0, 13);
            this.lTime.TabIndex = 11;
            this.lTime.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(165, 313);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Cancel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(294, 361);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lTime);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.bXml);
            this.Controls.Add(this.bXlsx);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Exam psv30-18 SalikovVasiliy";
            this.Load += new System.EventHandler(this.Form_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bXlsx;
        private System.Windows.Forms.Button bXml;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.Label lTime;
        private System.Windows.Forms.HelpProvider helpProvider1;
        private System.Windows.Forms.Button button1;
    }
}

