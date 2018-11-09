namespace ClassLibraryNavisworksROOMDFS
{
    partial class DFSControl
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.comboBoxTo = new System.Windows.Forms.ComboBox();
            this.comboBoxFrom = new System.Windows.Forms.ComboBox();
            this.btnFilter = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataGridViewPaths = new System.Windows.Forms.DataGridView();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnShow = new System.Windows.Forms.Button();
            this.btnR = new System.Windows.Forms.Button();
            this.dataGridViewDFS = new System.Windows.Forms.DataGridView();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabelS = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPaths)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDFS)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.groupBox2);
            this.splitContainer1.Panel1.Controls.Add(this.groupBox1);
            this.splitContainer1.Panel1.Controls.Add(this.btnClear);
            this.splitContainer1.Panel1.Controls.Add(this.btnShow);
            this.splitContainer1.Panel1.Controls.Add(this.btnR);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewDFS);
            this.splitContainer1.Panel2.Controls.Add(this.statusStrip1);
            this.splitContainer1.Size = new System.Drawing.Size(600, 450);
            this.splitContainer1.SplitterDistance = 190;
            this.splitContainer1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.comboBoxTo);
            this.groupBox2.Controls.Add(this.comboBoxFrom);
            this.groupBox2.Controls.Add(this.btnFilter);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(11, 84);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(172, 99);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "空间过滤";
            // 
            // comboBoxTo
            // 
            this.comboBoxTo.FormattingEnabled = true;
            this.comboBoxTo.Location = new System.Drawing.Point(55, 44);
            this.comboBoxTo.Name = "comboBoxTo";
            this.comboBoxTo.Size = new System.Drawing.Size(111, 20);
            this.comboBoxTo.TabIndex = 4;
            // 
            // comboBoxFrom
            // 
            this.comboBoxFrom.FormattingEnabled = true;
            this.comboBoxFrom.Location = new System.Drawing.Point(55, 18);
            this.comboBoxFrom.Name = "comboBoxFrom";
            this.comboBoxFrom.Size = new System.Drawing.Size(111, 20);
            this.comboBoxFrom.TabIndex = 3;
            // 
            // btnFilter
            // 
            this.btnFilter.Location = new System.Drawing.Point(91, 70);
            this.btnFilter.Name = "btnFilter";
            this.btnFilter.Size = new System.Drawing.Size(75, 23);
            this.btnFilter.TabIndex = 2;
            this.btnFilter.Text = "过滤";
            this.btnFilter.UseVisualStyleBackColor = true;
            this.btnFilter.Click += new System.EventHandler(this.btnFilter_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "保护:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "外部:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataGridViewPaths);
            this.groupBox1.Location = new System.Drawing.Point(15, 189);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(168, 258);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "路径描述";
            // 
            // dataGridViewPaths
            // 
            this.dataGridViewPaths.AllowUserToAddRows = false;
            this.dataGridViewPaths.AllowUserToDeleteRows = false;
            this.dataGridViewPaths.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPaths.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridViewPaths.Location = new System.Drawing.Point(3, 17);
            this.dataGridViewPaths.Name = "dataGridViewPaths";
            this.dataGridViewPaths.ReadOnly = true;
            this.dataGridViewPaths.RowHeadersVisible = false;
            this.dataGridViewPaths.RowTemplate.Height = 23;
            this.dataGridViewPaths.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewPaths.Size = new System.Drawing.Size(162, 213);
            this.dataGridViewPaths.TabIndex = 1;
            this.dataGridViewPaths.DoubleClick += new System.EventHandler(this.dataGridViewPaths_DoubleClick);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(134, 53);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(49, 23);
            this.btnClear.TabIndex = 3;
            this.btnClear.Text = "清除";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnShow
            // 
            this.btnShow.Location = new System.Drawing.Point(11, 53);
            this.btnShow.Name = "btnShow";
            this.btnShow.Size = new System.Drawing.Size(117, 23);
            this.btnShow.TabIndex = 1;
            this.btnShow.Text = "显示";
            this.btnShow.UseVisualStyleBackColor = true;
            this.btnShow.Click += new System.EventHandler(this.btnShow_Click);
            // 
            // btnR
            // 
            this.btnR.Location = new System.Drawing.Point(11, 10);
            this.btnR.Name = "btnR";
            this.btnR.Size = new System.Drawing.Size(172, 38);
            this.btnR.TabIndex = 0;
            this.btnR.Text = "读取";
            this.btnR.UseVisualStyleBackColor = true;
            this.btnR.Click += new System.EventHandler(this.btnR_Click);
            // 
            // dataGridViewDFS
            // 
            this.dataGridViewDFS.AllowUserToAddRows = false;
            this.dataGridViewDFS.AllowUserToDeleteRows = false;
            this.dataGridViewDFS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDFS.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewDFS.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewDFS.Name = "dataGridViewDFS";
            this.dataGridViewDFS.ReadOnly = true;
            this.dataGridViewDFS.RowTemplate.Height = 23;
            this.dataGridViewDFS.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewDFS.Size = new System.Drawing.Size(406, 428);
            this.dataGridViewDFS.TabIndex = 2;
            this.dataGridViewDFS.SelectionChanged += new System.EventHandler(this.dataGridViewDFS_SelectionChanged);
            this.dataGridViewDFS.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridViewDFS_KeyDown);
            this.dataGridViewDFS.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dataGridViewDFS_KeyPress);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelS});
            this.statusStrip1.Location = new System.Drawing.Point(0, 428);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(406, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabelS
            // 
            this.toolStripStatusLabelS.Name = "toolStripStatusLabelS";
            this.toolStripStatusLabelS.Size = new System.Drawing.Size(0, 17);
            // 
            // DFSControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "DFSControl";
            this.Size = new System.Drawing.Size(600, 450);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPaths)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDFS)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button btnShow;
        private System.Windows.Forms.Button btnR;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.DataGridView dataGridViewDFS;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelS;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox comboBoxTo;
        private System.Windows.Forms.ComboBox comboBoxFrom;
        private System.Windows.Forms.Button btnFilter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridViewPaths;
    }
}
