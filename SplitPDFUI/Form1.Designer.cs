namespace SplitPDFUI
{
    partial class Form1
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
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblPDFFolder = new System.Windows.Forms.Label();
            this.txtPDFFolder = new System.Windows.Forms.TextBox();
            this.lblFolder = new System.Windows.Forms.Label();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.chkPDFs = new System.Windows.Forms.CheckBox();
            this.chkThumbs = new System.Windows.Forms.CheckBox();
            this.chkConsolidate = new System.Windows.Forms.CheckBox();
            this.chkText = new System.Windows.Forms.CheckBox();
            this.chkNav = new System.Windows.Forms.CheckBox();
            this.cmdDefault = new System.Windows.Forms.Button();
            this.btnCurlTests = new System.Windows.Forms.Button();
            this.txtProject = new System.Windows.Forms.TextBox();
            this.lblProject = new System.Windows.Forms.Label();
            this.txtOutputFolder = new System.Windows.Forms.TextBox();
            this.lblOutputFolder = new System.Windows.Forms.Label();
            this.cmbGit = new System.Windows.Forms.ComboBox();
            this.txtExcelSource = new System.Windows.Forms.TextBox();
            this.Excel = new System.Windows.Forms.Label();
            this.btnBookmark = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(101, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.Text = "browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lblPDFFolder
            // 
            this.lblPDFFolder.AutoSize = true;
            this.lblPDFFolder.Location = new System.Drawing.Point(12, 9);
            this.lblPDFFolder.Name = "lblPDFFolder";
            this.lblPDFFolder.Size = new System.Drawing.Size(28, 13);
            this.lblPDFFolder.TabIndex = 6;
            this.lblPDFFolder.Text = "PDF";
            // 
            // txtPDFFolder
            // 
            this.txtPDFFolder.Location = new System.Drawing.Point(298, 7);
            this.txtPDFFolder.Name = "txtPDFFolder";
            this.txtPDFFolder.Size = new System.Drawing.Size(636, 20);
            this.txtPDFFolder.TabIndex = 30;
            this.txtPDFFolder.TextChanged += new System.EventHandler(this.txtPDFFolder_TextChanged);
            // 
            // lblFolder
            // 
            this.lblFolder.AutoSize = true;
            this.lblFolder.Location = new System.Drawing.Point(217, 10);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(60, 13);
            this.lblFolder.TabIndex = 29;
            this.lblFolder.Text = "PDF Folder";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(15, 44);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(89, 25);
            this.btnGenerate.TabIndex = 31;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // chkPDFs
            // 
            this.chkPDFs.AutoSize = true;
            this.chkPDFs.Location = new System.Drawing.Point(12, 161);
            this.chkPDFs.Name = "chkPDFs";
            this.chkPDFs.Size = new System.Drawing.Size(82, 17);
            this.chkPDFs.TabIndex = 32;
            this.chkPDFs.Text = "createPDFs";
            this.chkPDFs.UseVisualStyleBackColor = true;
            // 
            // chkThumbs
            // 
            this.chkThumbs.AutoSize = true;
            this.chkThumbs.Checked = true;
            this.chkThumbs.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkThumbs.Location = new System.Drawing.Point(12, 184);
            this.chkThumbs.Name = "chkThumbs";
            this.chkThumbs.Size = new System.Drawing.Size(94, 17);
            this.chkThumbs.TabIndex = 33;
            this.chkThumbs.Text = "createThumbs";
            this.chkThumbs.UseVisualStyleBackColor = true;
            // 
            // chkConsolidate
            // 
            this.chkConsolidate.AutoSize = true;
            this.chkConsolidate.Location = new System.Drawing.Point(12, 207);
            this.chkConsolidate.Name = "chkConsolidate";
            this.chkConsolidate.Size = new System.Drawing.Size(80, 17);
            this.chkConsolidate.TabIndex = 34;
            this.chkConsolidate.Text = "consolidate";
            this.chkConsolidate.UseVisualStyleBackColor = true;
            // 
            // chkText
            // 
            this.chkText.AutoSize = true;
            this.chkText.Location = new System.Drawing.Point(12, 230);
            this.chkText.Name = "chkText";
            this.chkText.Size = new System.Drawing.Size(79, 17);
            this.chkText.TabIndex = 35;
            this.chkText.Text = "extractText";
            this.chkText.UseVisualStyleBackColor = true;
            // 
            // chkNav
            // 
            this.chkNav.AutoSize = true;
            this.chkNav.Location = new System.Drawing.Point(11, 253);
            this.chkNav.Name = "chkNav";
            this.chkNav.Size = new System.Drawing.Size(76, 17);
            this.chkNav.TabIndex = 36;
            this.chkNav.Text = "export nav";
            this.chkNav.UseVisualStyleBackColor = true;
            // 
            // cmdDefault
            // 
            this.cmdDefault.Location = new System.Drawing.Point(111, 44);
            this.cmdDefault.Name = "cmdDefault";
            this.cmdDefault.Size = new System.Drawing.Size(75, 23);
            this.cmdDefault.TabIndex = 37;
            this.cmdDefault.Text = "Default";
            this.cmdDefault.UseVisualStyleBackColor = true;
            this.cmdDefault.Click += new System.EventHandler(this.cmdDefault_Click);
            // 
            // btnCurlTests
            // 
            this.btnCurlTests.Location = new System.Drawing.Point(15, 85);
            this.btnCurlTests.Name = "btnCurlTests";
            this.btnCurlTests.Size = new System.Drawing.Size(89, 23);
            this.btnCurlTests.TabIndex = 38;
            this.btnCurlTests.Text = "Curl Tests";
            this.btnCurlTests.UseVisualStyleBackColor = true;
            this.btnCurlTests.Click += new System.EventHandler(this.btnCurlTests_Click);
            // 
            // txtProject
            // 
            this.txtProject.Location = new System.Drawing.Point(298, 33);
            this.txtProject.MaxLength = 3;
            this.txtProject.Name = "txtProject";
            this.txtProject.Size = new System.Drawing.Size(46, 20);
            this.txtProject.TabIndex = 40;
            this.txtProject.TextChanged += new System.EventHandler(this.txtProject_TextChanged);
            // 
            // lblProject
            // 
            this.lblProject.AutoSize = true;
            this.lblProject.Location = new System.Drawing.Point(217, 36);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(54, 13);
            this.lblProject.TabIndex = 41;
            this.lblProject.Text = "Project ID";
            this.lblProject.Click += new System.EventHandler(this.label1_Click);
            // 
            // txtOutputFolder
            // 
            this.txtOutputFolder.Location = new System.Drawing.Point(298, 62);
            this.txtOutputFolder.Name = "txtOutputFolder";
            this.txtOutputFolder.Size = new System.Drawing.Size(636, 20);
            this.txtOutputFolder.TabIndex = 43;
            this.txtOutputFolder.TextChanged += new System.EventHandler(this.txtOutputFolder_TextChanged);
            // 
            // lblOutputFolder
            // 
            this.lblOutputFolder.AutoSize = true;
            this.lblOutputFolder.Location = new System.Drawing.Point(217, 65);
            this.lblOutputFolder.Name = "lblOutputFolder";
            this.lblOutputFolder.Size = new System.Drawing.Size(71, 13);
            this.lblOutputFolder.TabIndex = 42;
            this.lblOutputFolder.Text = "Output Folder";
            this.lblOutputFolder.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // cmbGit
            // 
            this.cmbGit.FormattingEnabled = true;
            this.cmbGit.Location = new System.Drawing.Point(11, 274);
            this.cmbGit.Name = "cmbGit";
            this.cmbGit.Size = new System.Drawing.Size(121, 21);
            this.cmbGit.TabIndex = 44;
            this.cmbGit.SelectedIndexChanged += new System.EventHandler(this.cmbGit_SelectedIndexChanged);
            // 
            // txtExcelSource
            // 
            this.txtExcelSource.Location = new System.Drawing.Point(298, 88);
            this.txtExcelSource.Name = "txtExcelSource";
            this.txtExcelSource.Size = new System.Drawing.Size(636, 20);
            this.txtExcelSource.TabIndex = 46;
            // 
            // Excel
            // 
            this.Excel.AutoSize = true;
            this.Excel.Location = new System.Drawing.Point(217, 91);
            this.Excel.Name = "Excel";
            this.Excel.Size = new System.Drawing.Size(33, 13);
            this.Excel.TabIndex = 45;
            this.Excel.Text = "Excel";
            // 
            // btnBookmark
            // 
            this.btnBookmark.Location = new System.Drawing.Point(15, 114);
            this.btnBookmark.Name = "btnBookmark";
            this.btnBookmark.Size = new System.Drawing.Size(89, 23);
            this.btnBookmark.TabIndex = 47;
            this.btnBookmark.Text = "Bookmark PDF";
            this.btnBookmark.UseVisualStyleBackColor = true;
            this.btnBookmark.Click += new System.EventHandler(this.btnBookmark_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(946, 307);
            this.Controls.Add(this.btnBookmark);
            this.Controls.Add(this.txtExcelSource);
            this.Controls.Add(this.Excel);
            this.Controls.Add(this.cmbGit);
            this.Controls.Add(this.txtOutputFolder);
            this.Controls.Add(this.lblOutputFolder);
            this.Controls.Add(this.lblProject);
            this.Controls.Add(this.txtProject);
            this.Controls.Add(this.btnCurlTests);
            this.Controls.Add(this.cmdDefault);
            this.Controls.Add(this.chkNav);
            this.Controls.Add(this.chkText);
            this.Controls.Add(this.chkConsolidate);
            this.Controls.Add(this.chkThumbs);
            this.Controls.Add(this.chkPDFs);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.txtPDFFolder);
            this.Controls.Add(this.lblFolder);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.lblPDFFolder);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label lblPDFFolder;
        private System.Windows.Forms.TextBox txtPDFFolder;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.CheckBox chkPDFs;
        private System.Windows.Forms.CheckBox chkThumbs;
        private System.Windows.Forms.CheckBox chkConsolidate;
        private System.Windows.Forms.CheckBox chkText;
        private System.Windows.Forms.CheckBox chkNav;
        private System.Windows.Forms.Button cmdDefault;
        private System.Windows.Forms.Button btnCurlTests;
        private System.Windows.Forms.TextBox txtProject;
        private System.Windows.Forms.Label lblProject;
        private System.Windows.Forms.TextBox txtOutputFolder;
        private System.Windows.Forms.Label lblOutputFolder;
        private System.Windows.Forms.ComboBox cmbGit;
        private System.Windows.Forms.TextBox txtExcelSource;
        private System.Windows.Forms.Label Excel;
        private System.Windows.Forms.Button btnBookmark;
    }
}

