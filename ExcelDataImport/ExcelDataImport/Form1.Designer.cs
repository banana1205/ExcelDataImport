namespace ExcelDataImport
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_Import = new System.Windows.Forms.Button();
            this.cmb_modelList = new System.Windows.Forms.ComboBox();
            this.cmb_sheetName = new System.Windows.Forms.ComboBox();
            this.txt_filePath = new System.Windows.Forms.TextBox();
            this.btn_openFile = new System.Windows.Forms.Button();
            this.dgv_pvw = new System.Windows.Forms.DataGridView();
            this.gruopBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_pvw)).BeginInit();
            this.gruopBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_Import);
            this.groupBox1.Controls.Add(this.cmb_modelList);
            this.groupBox1.Controls.Add(this.cmb_sheetName);
            this.groupBox1.Controls.Add(this.txt_filePath);
            this.groupBox1.Controls.Add(this.btn_openFile);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(984, 76);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel操作";
            // 
            // btn_Import
            // 
            this.btn_Import.Location = new System.Drawing.Point(792, 32);
            this.btn_Import.Name = "btn_Import";
            this.btn_Import.Size = new System.Drawing.Size(75, 23);
            this.btn_Import.TabIndex = 4;
            this.btn_Import.Text = "导入";
            this.btn_Import.UseVisualStyleBackColor = true;
            this.btn_Import.Click += new System.EventHandler(this.btn_Import_Click);
            // 
            // cmb_modelList
            // 
            this.cmb_modelList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_modelList.FormattingEnabled = true;
            this.cmb_modelList.Location = new System.Drawing.Point(664, 35);
            this.cmb_modelList.Name = "cmb_modelList";
            this.cmb_modelList.Size = new System.Drawing.Size(121, 20);
            this.cmb_modelList.TabIndex = 3;
            // 
            // cmb_sheetName
            // 
            this.cmb_sheetName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_sheetName.FormattingEnabled = true;
            this.cmb_sheetName.Location = new System.Drawing.Point(574, 36);
            this.cmb_sheetName.Name = "cmb_sheetName";
            this.cmb_sheetName.Size = new System.Drawing.Size(83, 20);
            this.cmb_sheetName.TabIndex = 2;
            this.cmb_sheetName.SelectedValueChanged += new System.EventHandler(this.cmb_sheetName_SelectedValueChanged);
            // 
            // txt_filePath
            // 
            this.txt_filePath.Location = new System.Drawing.Point(155, 35);
            this.txt_filePath.Name = "txt_filePath";
            this.txt_filePath.ReadOnly = true;
            this.txt_filePath.Size = new System.Drawing.Size(413, 21);
            this.txt_filePath.TabIndex = 1;
            // 
            // btn_openFile
            // 
            this.btn_openFile.Location = new System.Drawing.Point(44, 34);
            this.btn_openFile.Name = "btn_openFile";
            this.btn_openFile.Size = new System.Drawing.Size(105, 21);
            this.btn_openFile.TabIndex = 0;
            this.btn_openFile.Text = "打开Excel文件";
            this.btn_openFile.UseVisualStyleBackColor = true;
            this.btn_openFile.Click += new System.EventHandler(this.btn_openFile_Click);
            // 
            // dgv_pvw
            // 
            this.dgv_pvw.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_pvw.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_pvw.Location = new System.Drawing.Point(3, 17);
            this.dgv_pvw.Name = "dgv_pvw";
            this.dgv_pvw.RowTemplate.Height = 23;
            this.dgv_pvw.Size = new System.Drawing.Size(978, 593);
            this.dgv_pvw.TabIndex = 1;
            // 
            // gruopBox2
            // 
            this.gruopBox2.Controls.Add(this.dgv_pvw);
            this.gruopBox2.Location = new System.Drawing.Point(12, 94);
            this.gruopBox2.Name = "gruopBox2";
            this.gruopBox2.Size = new System.Drawing.Size(984, 613);
            this.gruopBox2.TabIndex = 2;
            this.gruopBox2.TabStop = false;
            this.gruopBox2.Text = "预览";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1008, 707);
            this.Controls.Add(this.gruopBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_pvw)).EndInit();
            this.gruopBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_openFile;
        private System.Windows.Forms.TextBox txt_filePath;
        private System.Windows.Forms.ComboBox cmb_sheetName;
        private System.Windows.Forms.DataGridView dgv_pvw;
        private System.Windows.Forms.GroupBox gruopBox2;
        private System.Windows.Forms.ComboBox cmb_modelList;
        private System.Windows.Forms.Button btn_Import;
    }
}

