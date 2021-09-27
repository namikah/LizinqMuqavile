namespace DollarKurs
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txtAZN = new System.Windows.Forms.TextBox();
            this.txtUSD = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.button6 = new System.Windows.Forms.Button();
            this.txtEUR = new System.Windows.Forms.TextBox();
            this.txtAZN2 = new System.Windows.Forms.TextBox();
            this.button9 = new System.Windows.Forms.Button();
            this.txtRUB = new System.Windows.Forms.TextBox();
            this.txtAZN3 = new System.Windows.Forms.TextBox();
            this.button12 = new System.Windows.Forms.Button();
            this.txtTRY = new System.Windows.Forms.TextBox();
            this.txtAZN4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtAZN
            // 
            this.txtAZN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtAZN.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtAZN.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtAZN.ForeColor = System.Drawing.Color.Black;
            this.txtAZN.Location = new System.Drawing.Point(154, 52);
            this.txtAZN.Name = "txtAZN";
            this.txtAZN.Size = new System.Drawing.Size(97, 19);
            this.txtAZN.TabIndex = 0;
            // 
            // txtUSD
            // 
            this.txtUSD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtUSD.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtUSD.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtUSD.ForeColor = System.Drawing.Color.Black;
            this.txtUSD.Location = new System.Drawing.Point(8, 52);
            this.txtUSD.Name = "txtUSD";
            this.txtUSD.Size = new System.Drawing.Size(97, 19);
            this.txtUSD.TabIndex = 2;
            this.txtUSD.Text = "1.0000";
            this.txtUSD.TextChanged += new System.EventHandler(this.txtUSD_TextChanged);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(8, 243);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(243, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Convert";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btConvert_Click);
            // 
            // button3
            // 
            this.button3.FlatAppearance.BorderSize = 0;
            this.button3.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button3.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.ForeColor = System.Drawing.Color.Black;
            this.button3.Location = new System.Drawing.Point(3, 27);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(41, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "USD";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.button5.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button5.ForeColor = System.Drawing.Color.Yellow;
            this.button5.Location = new System.Drawing.Point(-6, 0);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(275, 23);
            this.button5.TabIndex = 12;
            this.button5.Text = "VALYUTA MƏZƏNNƏSİ";
            this.button5.UseVisualStyleBackColor = false;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            // 
            // button6
            // 
            this.button6.FlatAppearance.BorderSize = 0;
            this.button6.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button6.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button6.ForeColor = System.Drawing.Color.Black;
            this.button6.Location = new System.Drawing.Point(3, 79);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(41, 23);
            this.button6.TabIndex = 16;
            this.button6.Text = "EUR";
            this.button6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button6.UseVisualStyleBackColor = true;
            // 
            // txtEUR
            // 
            this.txtEUR.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtEUR.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtEUR.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtEUR.ForeColor = System.Drawing.Color.Black;
            this.txtEUR.Location = new System.Drawing.Point(8, 104);
            this.txtEUR.Name = "txtEUR";
            this.txtEUR.Size = new System.Drawing.Size(97, 19);
            this.txtEUR.TabIndex = 14;
            this.txtEUR.Text = "1.0000";
            this.txtEUR.TextChanged += new System.EventHandler(this.txtEUR_TextChanged);
            // 
            // txtAZN2
            // 
            this.txtAZN2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtAZN2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtAZN2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtAZN2.ForeColor = System.Drawing.Color.Black;
            this.txtAZN2.Location = new System.Drawing.Point(154, 104);
            this.txtAZN2.Name = "txtAZN2";
            this.txtAZN2.Size = new System.Drawing.Size(97, 19);
            this.txtAZN2.TabIndex = 13;
            // 
            // button9
            // 
            this.button9.FlatAppearance.BorderSize = 0;
            this.button9.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button9.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.button9.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button9.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button9.ForeColor = System.Drawing.Color.Black;
            this.button9.Location = new System.Drawing.Point(3, 131);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(41, 23);
            this.button9.TabIndex = 21;
            this.button9.Text = "RUB";
            this.button9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button9.UseVisualStyleBackColor = true;
            // 
            // txtRUB
            // 
            this.txtRUB.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtRUB.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtRUB.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtRUB.ForeColor = System.Drawing.Color.Black;
            this.txtRUB.Location = new System.Drawing.Point(8, 156);
            this.txtRUB.Name = "txtRUB";
            this.txtRUB.Size = new System.Drawing.Size(97, 19);
            this.txtRUB.TabIndex = 19;
            this.txtRUB.Text = "1.0000";
            this.txtRUB.TextChanged += new System.EventHandler(this.txtRUB_TextChanged);
            // 
            // txtAZN3
            // 
            this.txtAZN3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtAZN3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtAZN3.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtAZN3.ForeColor = System.Drawing.Color.Black;
            this.txtAZN3.Location = new System.Drawing.Point(154, 156);
            this.txtAZN3.Name = "txtAZN3";
            this.txtAZN3.Size = new System.Drawing.Size(97, 19);
            this.txtAZN3.TabIndex = 18;
            // 
            // button12
            // 
            this.button12.FlatAppearance.BorderSize = 0;
            this.button12.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button12.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.button12.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button12.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button12.ForeColor = System.Drawing.Color.Black;
            this.button12.Location = new System.Drawing.Point(3, 183);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(41, 23);
            this.button12.TabIndex = 26;
            this.button12.Text = "TRY";
            this.button12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button12.UseVisualStyleBackColor = true;
            // 
            // txtTRY
            // 
            this.txtTRY.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtTRY.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtTRY.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtTRY.ForeColor = System.Drawing.Color.Black;
            this.txtTRY.Location = new System.Drawing.Point(8, 208);
            this.txtTRY.Name = "txtTRY";
            this.txtTRY.Size = new System.Drawing.Size(97, 19);
            this.txtTRY.TabIndex = 24;
            this.txtTRY.Text = "1.0000";
            this.txtTRY.TextChanged += new System.EventHandler(this.txtTRY_TextChanged);
            // 
            // txtAZN4
            // 
            this.txtAZN4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.txtAZN4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtAZN4.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtAZN4.ForeColor = System.Drawing.Color.Black;
            this.txtAZN4.Location = new System.Drawing.Point(154, 208);
            this.txtAZN4.Name = "txtAZN4";
            this.txtAZN4.Size = new System.Drawing.Size(97, 19);
            this.txtAZN4.TabIndex = 23;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(263, 274);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.txtTRY);
            this.Controls.Add(this.txtAZN4);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.txtRUB);
            this.Controls.Add(this.txtAZN3);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.txtEUR);
            this.Controls.Add(this.txtAZN2);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtUSD);
            this.Controls.Add(this.txtAZN);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MƏRKƏZİ BANK";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtAZN;
        private System.Windows.Forms.TextBox txtUSD;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TextBox txtEUR;
        private System.Windows.Forms.TextBox txtAZN2;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.TextBox txtRUB;
        private System.Windows.Forms.TextBox txtAZN3;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.TextBox txtTRY;
        private System.Windows.Forms.TextBox txtAZN4;
    }
}

