
namespace InformeMensusalV4
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dateTimePickerDesde = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerHasta = new System.Windows.Forms.DateTimePicker();
            this.btnCargarExcel = new System.Windows.Forms.Button();
            this.btnGenerarPDF = new System.Windows.Forms.Button();
            this.dataGridViewTickets = new System.Windows.Forms.DataGridView();
            this.webBrowserPDF = new System.Windows.Forms.WebBrowser();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTickets)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dateTimePickerDesde
            // 
            this.dateTimePickerDesde.Location = new System.Drawing.Point(12, 38);
            this.dateTimePickerDesde.Name = "dateTimePickerDesde";
            this.dateTimePickerDesde.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerDesde.TabIndex = 0;
            // 
            // dateTimePickerHasta
            // 
            this.dateTimePickerHasta.Location = new System.Drawing.Point(232, 38);
            this.dateTimePickerHasta.Name = "dateTimePickerHasta";
            this.dateTimePickerHasta.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerHasta.TabIndex = 1;
            // 
            // btnCargarExcel
            // 
            this.btnCargarExcel.Location = new System.Drawing.Point(508, 35);
            this.btnCargarExcel.Name = "btnCargarExcel";
            this.btnCargarExcel.Size = new System.Drawing.Size(75, 23);
            this.btnCargarExcel.TabIndex = 2;
            this.btnCargarExcel.Text = "Cargar Excel";
            this.btnCargarExcel.UseVisualStyleBackColor = true;
            this.btnCargarExcel.Click += new System.EventHandler(this.btnCargarExcel_Click);
            // 
            // btnGenerarPDF
            // 
            this.btnGenerarPDF.Location = new System.Drawing.Point(661, 35);
            this.btnGenerarPDF.Name = "btnGenerarPDF";
            this.btnGenerarPDF.Size = new System.Drawing.Size(95, 23);
            this.btnGenerarPDF.TabIndex = 3;
            this.btnGenerarPDF.Text = "Generar Informe";
            this.btnGenerarPDF.UseVisualStyleBackColor = true;
            this.btnGenerarPDF.Click += new System.EventHandler(this.btnGenerarPDF_Click);
            // 
            // dataGridViewTickets
            // 
            this.dataGridViewTickets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewTickets.Location = new System.Drawing.Point(55, 95);
            this.dataGridViewTickets.Name = "dataGridViewTickets";
            this.dataGridViewTickets.Size = new System.Drawing.Size(690, 250);
            this.dataGridViewTickets.TabIndex = 4;
            // 
            // webBrowserPDF
            // 
            this.webBrowserPDF.Location = new System.Drawing.Point(55, 388);
            this.webBrowserPDF.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowserPDF.Name = "webBrowserPDF";
            this.webBrowserPDF.Size = new System.Drawing.Size(690, 306);
            this.webBrowserPDF.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(88, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Desde ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(324, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Hasta";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 742);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.webBrowserPDF);
            this.Controls.Add(this.dataGridViewTickets);
            this.Controls.Add(this.btnGenerarPDF);
            this.Controls.Add(this.btnCargarExcel);
            this.Controls.Add(this.dateTimePickerHasta);
            this.Controls.Add(this.dateTimePickerDesde);
            this.Name = "Form1";
            this.Text = "Informe Mensual";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTickets)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DateTimePicker dateTimePickerDesde;
        private System.Windows.Forms.DateTimePicker dateTimePickerHasta;
        private System.Windows.Forms.Button btnCargarExcel;
        private System.Windows.Forms.Button btnGenerarPDF;
        private System.Windows.Forms.DataGridView dataGridViewTickets;
        private System.Windows.Forms.WebBrowser webBrowserPDF;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

