namespace ExcelAyirmaV2
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            fileselectbtn = new Button();
            adresstxt = new TextBox();
            cellvaluetxt = new TextBox();
            dataGridView1 = new DataGridView();
            progressBar1 = new ProgressBar();
            saveadressfoldertxt = new TextBox();
            saveselectedfolderbtn = new Button();
            saveexcelbtn = new Button();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            label4 = new Label();
            rdybtn = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // fileselectbtn
            // 
            fileselectbtn.Location = new Point(566, 9);
            fileselectbtn.Name = "fileselectbtn";
            fileselectbtn.Size = new Size(149, 32);
            fileselectbtn.TabIndex = 0;
            fileselectbtn.Tag = "";
            fileselectbtn.Text = "DOSYA SEÇ";
            fileselectbtn.UseVisualStyleBackColor = true;
            fileselectbtn.Click += fileselectbtn_Click;
            // 
            // adresstxt
            // 
            adresstxt.BorderStyle = BorderStyle.None;
            adresstxt.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            adresstxt.Location = new Point(148, 9);
            adresstxt.Multiline = true;
            adresstxt.Name = "adresstxt";
            adresstxt.ReadOnly = true;
            adresstxt.ScrollBars = ScrollBars.Vertical;
            adresstxt.Size = new Size(389, 23);
            adresstxt.TabIndex = 1;
            adresstxt.WordWrap = false;
            // 
            // cellvaluetxt
            // 
            cellvaluetxt.BorderStyle = BorderStyle.None;
            cellvaluetxt.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            cellvaluetxt.Location = new Point(148, 67);
            cellvaluetxt.Name = "cellvaluetxt";
            cellvaluetxt.Size = new Size(389, 22);
            cellvaluetxt.TabIndex = 3;
            cellvaluetxt.Text = "01-15";
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(11, 125);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(704, 240);
            dataGridView1.TabIndex = 4;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(12, 96);
            progressBar1.Maximum = 10000;
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(423, 23);
            progressBar1.TabIndex = 6;
            // 
            // saveadressfoldertxt
            // 
            saveadressfoldertxt.BorderStyle = BorderStyle.None;
            saveadressfoldertxt.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            saveadressfoldertxt.Location = new Point(148, 38);
            saveadressfoldertxt.Name = "saveadressfoldertxt";
            saveadressfoldertxt.ReadOnly = true;
            saveadressfoldertxt.Size = new Size(389, 22);
            saveadressfoldertxt.TabIndex = 9;
            // 
            // saveselectedfolderbtn
            // 
            saveselectedfolderbtn.Location = new Point(566, 47);
            saveselectedfolderbtn.Name = "saveselectedfolderbtn";
            saveselectedfolderbtn.Size = new Size(149, 32);
            saveselectedfolderbtn.TabIndex = 10;
            saveselectedfolderbtn.Text = "KAYDEDİLECEK DİZİN";
            saveselectedfolderbtn.UseVisualStyleBackColor = true;
            saveselectedfolderbtn.Click += saveselectedfolderbtn_Click;
            // 
            // saveexcelbtn
            // 
            saveexcelbtn.Location = new Point(638, 87);
            saveexcelbtn.Name = "saveexcelbtn";
            saveexcelbtn.Size = new Size(77, 32);
            saveexcelbtn.TabIndex = 11;
            saveexcelbtn.Text = "KAYDET";
            saveexcelbtn.UseVisualStyleBackColor = true;
            saveexcelbtn.Click += saveexcelbtn_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(54, 12);
            label1.Name = "label1";
            label1.Size = new Size(88, 15);
            label1.TabIndex = 12;
            label1.Text = "DOSYA ADRESİ:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(15, 41);
            label2.Name = "label2";
            label2.Size = new Size(127, 15);
            label2.TabIndex = 12;
            label2.Text = "KAYDEDİLECEK ADRES:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(73, 70);
            label3.Name = "label3";
            label3.Size = new Size(69, 15);
            label3.TabIndex = 12;
            label3.Text = "DOSYA ADI:";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            label4.Location = new Point(441, 96);
            label4.Name = "label4";
            label4.Size = new Size(96, 21);
            label4.TabIndex = 13;
            label4.Text = "1000 / 1000";
            // 
            // rdybtn
            // 
            rdybtn.Location = new Point(566, 87);
            rdybtn.Name = "rdybtn";
            rdybtn.Size = new Size(75, 32);
            rdybtn.TabIndex = 14;
            rdybtn.Text = "HAZIRLA";
            rdybtn.UseVisualStyleBackColor = true;
            rdybtn.Click += rdybtn_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(727, 377);
            Controls.Add(rdybtn);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(saveexcelbtn);
            Controls.Add(saveselectedfolderbtn);
            Controls.Add(saveadressfoldertxt);
            Controls.Add(progressBar1);
            Controls.Add(dataGridView1);
            Controls.Add(cellvaluetxt);
            Controls.Add(adresstxt);
            Controls.Add(fileselectbtn);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button fileselectbtn;
        private TextBox adresstxt;
        private TextBox cellvaluetxt;
        private DataGridView dataGridView1;
        private ProgressBar progressBar1;
        private TextBox saveadressfoldertxt;
        private Button saveselectedfolderbtn;
        private Button saveexcelbtn;
        private Label label1;
        private Label label2;
        private Label label3;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private Label label4;
        private Button rdybtn;
    }
}
