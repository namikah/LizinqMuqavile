namespace Lizinq_Muqavile
{
    partial class Sekiller
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
            this.t1 = new System.Windows.Forms.RichTextBox();
            this.label47 = new System.Windows.Forms.Label();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Melumat = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Images = new System.Windows.Forms.DataGridViewImageColumn();
            this.button15 = new System.Windows.Forms.Button();
            this.pb1 = new System.Windows.Forms.PictureBox();
            this.button16 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pb1)).BeginInit();
            this.SuspendLayout();
            // 
            // t1
            // 
            this.t1.Location = new System.Drawing.Point(106, 25);
            this.t1.Name = "t1";
            this.t1.Size = new System.Drawing.Size(177, 45);
            this.t1.TabIndex = 20;
            this.t1.Text = "";
            // 
            // label47
            // 
            this.label47.AutoSize = true;
            this.label47.Location = new System.Drawing.Point(108, 9);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(128, 13);
            this.label47.TabIndex = 19;
            this.label47.Text = "Avtomobil haqda melumat";
            // 
            // dataGridView4
            // 
            this.dataGridView4.AllowUserToAddRows = false;
            this.dataGridView4.AllowUserToDeleteRows = false;
            this.dataGridView4.AllowUserToResizeColumns = false;
            this.dataGridView4.AllowUserToResizeRows = false;
            this.dataGridView4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.id,
            this.Melumat,
            this.Images});
            this.dataGridView4.Location = new System.Drawing.Point(3, 110);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.ReadOnly = true;
            this.dataGridView4.RowHeadersVisible = false;
            this.dataGridView4.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView4.Size = new System.Drawing.Size(1045, 486);
            this.dataGridView4.StandardTab = true;
            this.dataGridView4.TabIndex = 18;
            // 
            // id
            // 
            this.id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.id.DataPropertyName = "id";
            this.id.FillWeight = 50F;
            this.id.HeaderText = "id";
            this.id.MinimumWidth = 50;
            this.id.Name = "id";
            this.id.ReadOnly = true;
            this.id.Width = 50;
            // 
            // Melumat
            // 
            this.Melumat.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Melumat.DataPropertyName = "c1";
            this.Melumat.FillWeight = 50F;
            this.Melumat.HeaderText = "Melumat";
            this.Melumat.MinimumWidth = 50;
            this.Melumat.Name = "Melumat";
            this.Melumat.ReadOnly = true;
            // 
            // Images
            // 
            this.Images.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Images.DataPropertyName = "c2";
            this.Images.FillWeight = 200F;
            this.Images.HeaderText = "Images";
            this.Images.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Stretch;
            this.Images.MinimumWidth = 200;
            this.Images.Name = "Images";
            this.Images.ReadOnly = true;
            this.Images.Width = 200;
            // 
            // button15
            // 
            this.button15.Location = new System.Drawing.Point(106, 76);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(177, 28);
            this.button15.TabIndex = 17;
            this.button15.Text = "Bazada saxla";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Click += new System.EventHandler(this.button15_Click);
            // 
            // pb1
            // 
            this.pb1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pb1.Location = new System.Drawing.Point(3, 3);
            this.pb1.Name = "pb1";
            this.pb1.Size = new System.Drawing.Size(93, 67);
            this.pb1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pb1.TabIndex = 16;
            this.pb1.TabStop = false;
            // 
            // button16
            // 
            this.button16.Location = new System.Drawing.Point(3, 76);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(93, 28);
            this.button16.TabIndex = 15;
            this.button16.Text = "Şəkil yüklə";
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Click += new System.EventHandler(this.button16_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Sekiller
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1049, 598);
            this.Controls.Add(this.t1);
            this.Controls.Add(this.label47);
            this.Controls.Add(this.dataGridView4);
            this.Controls.Add(this.button15);
            this.Controls.Add(this.pb1);
            this.Controls.Add(this.button16);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Sekiller";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sekiller";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Sekiller_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pb1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox t1;
        private System.Windows.Forms.Label label47;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.DataGridViewTextBoxColumn id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Melumat;
        private System.Windows.Forms.DataGridViewImageColumn Images;
        private System.Windows.Forms.Button button15;
        private System.Windows.Forms.PictureBox pb1;
        private System.Windows.Forms.Button button16;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}