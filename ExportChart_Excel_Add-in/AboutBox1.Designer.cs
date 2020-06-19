namespace ExportChart_Excel_Add_in
{
    partial class AboutBox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.labelProductName = new System.Windows.Forms.Label();
            this.labelVersion = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.linkLabelAuthor = new System.Windows.Forms.LinkLabel();
            this.labelAuthor = new System.Windows.Forms.Label();
            this.labelCopyright = new System.Windows.Forms.Label();
            this.okButton = new System.Windows.Forms.Button();
            this.richTextBoxDescription = new System.Windows.Forms.RichTextBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.labelProductName, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.labelVersion, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.labelCopyright, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.okButton, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.richTextBoxDescription, 0, 4);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(9, 9);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 6;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(519, 272);
            this.tableLayoutPanel1.TabIndex = 0;
            this.tableLayoutPanel1.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel1_Paint);
            // 
            // labelProductName
            // 
            this.labelProductName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelProductName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelProductName.Location = new System.Drawing.Point(6, 0);
            this.labelProductName.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.labelProductName.MaximumSize = new System.Drawing.Size(0, 17);
            this.labelProductName.Name = "labelProductName";
            this.labelProductName.Size = new System.Drawing.Size(510, 17);
            this.labelProductName.TabIndex = 0;
            this.labelProductName.Text = "Product Name";
            this.labelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelVersion
            // 
            this.labelVersion.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelVersion.Location = new System.Drawing.Point(6, 27);
            this.labelVersion.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.labelVersion.MaximumSize = new System.Drawing.Size(0, 17);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(510, 17);
            this.labelVersion.TabIndex = 1;
            this.labelVersion.Text = "Version";
            this.labelVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.linkLabelAuthor);
            this.panel1.Controls.Add(this.labelAuthor);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(6, 54);
            this.panel1.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.panel1.MaximumSize = new System.Drawing.Size(0, 17);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(510, 17);
            this.panel1.TabIndex = 2;
            // 
            // linkLabelAuthor
            // 
            this.linkLabelAuthor.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.linkLabelAuthor.AutoSize = true;
            this.linkLabelAuthor.Location = new System.Drawing.Point(106, 0);
            this.linkLabelAuthor.Name = "linkLabelAuthor";
            this.linkLabelAuthor.Size = new System.Drawing.Size(105, 13);
            this.linkLabelAuthor.TabIndex = 1;
            this.linkLabelAuthor.TabStop = true;
            this.linkLabelAuthor.Text = "http://omidjafari.com";
            this.linkLabelAuthor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabelAuthor.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelAuthor_LinkClicked);
            // 
            // labelAuthor
            // 
            this.labelAuthor.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.labelAuthor.AutoSize = true;
            this.labelAuthor.Location = new System.Drawing.Point(0, 0);
            this.labelAuthor.Name = "labelAuthor";
            this.labelAuthor.Size = new System.Drawing.Size(108, 13);
            this.labelAuthor.TabIndex = 0;
            this.labelAuthor.Text = "Author: Omid Jafari -  ";
            this.labelAuthor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelCopyright
            // 
            this.labelCopyright.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelCopyright.Location = new System.Drawing.Point(6, 81);
            this.labelCopyright.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            this.labelCopyright.MaximumSize = new System.Drawing.Size(0, 17);
            this.labelCopyright.Name = "labelCopyright";
            this.labelCopyright.Size = new System.Drawing.Size(510, 17);
            this.labelCopyright.TabIndex = 3;
            this.labelCopyright.Text = "Copyright";
            this.labelCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.okButton.Location = new System.Drawing.Point(441, 247);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 22);
            this.okButton.TabIndex = 5;
            this.okButton.Text = "&OK";
            // 
            // richTextBoxDescription
            // 
            this.richTextBoxDescription.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxDescription.Location = new System.Drawing.Point(6, 111);
            this.richTextBoxDescription.Margin = new System.Windows.Forms.Padding(6, 3, 3, 3);
            this.richTextBoxDescription.Name = "richTextBoxDescription";
            this.richTextBoxDescription.ReadOnly = true;
            this.richTextBoxDescription.Size = new System.Drawing.Size(510, 130);
            this.richTextBoxDescription.TabIndex = 4;
            this.richTextBoxDescription.TabStop = false;
            this.richTextBoxDescription.Text = "Description";
            this.richTextBoxDescription.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.richTextBoxDescription_LinkClicked);
            // 
            // AboutBox
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 291);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutBox";
            this.Padding = new System.Windows.Forms.Padding(9);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "AboutBox1";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label labelProductName;
        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.LinkLabel linkLabelAuthor;
        private System.Windows.Forms.Label labelAuthor;
        private System.Windows.Forms.Label labelCopyright;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.RichTextBox richTextBoxDescription;
    }
}
