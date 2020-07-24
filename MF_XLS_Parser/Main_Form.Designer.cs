namespace MF_XLS_Parser
{
    partial class Main_Form
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
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.MainTextBox = new System.Windows.Forms.TextBox();
            this.DataTextBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.CleanupButton = new System.Windows.Forms.Button();
            this.CellSearchTextBox = new System.Windows.Forms.TextBox();
            this.cellBox1 = new System.Windows.Forms.TextBox();
            this.cellBox2 = new System.Windows.Forms.TextBox();
            this.LoadingImage = new System.Windows.Forms.PictureBox();
            this.RowConfirmButton = new System.Windows.Forms.Button();
            this.DataRowBox = new System.Windows.Forms.TextBox();
            this.AppLoadingImage = new System.Windows.Forms.PictureBox();
            this.FilterStartButton = new System.Windows.Forms.Button();
            this.TypeRowBox = new System.Windows.Forms.TextBox();
            this.ListBox = new System.Windows.Forms.TextBox();
            this.MissingBox = new System.Windows.Forms.TextBox();
            this.YearTextBox = new System.Windows.Forms.TextBox();
            this.MonthTextBox = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.CancelButton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.AppLoadingImage)).BeginInit();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(63, 119);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(142, 52);
            this.OpenFileButton.TabIndex = 0;
            this.OpenFileButton.Text = "Abrir archivo";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // MainTextBox
            // 
            this.MainTextBox.Location = new System.Drawing.Point(211, 119);
            this.MainTextBox.Multiline = true;
            this.MainTextBox.Name = "MainTextBox";
            this.MainTextBox.ReadOnly = true;
            this.MainTextBox.Size = new System.Drawing.Size(142, 73);
            this.MainTextBox.TabIndex = 1;
            // 
            // DataTextBox
            // 
            this.DataTextBox.Location = new System.Drawing.Point(63, 22);
            this.DataTextBox.Multiline = true;
            this.DataTextBox.Name = "DataTextBox";
            this.DataTextBox.ReadOnly = true;
            this.DataTextBox.Size = new System.Drawing.Size(202, 77);
            this.DataTextBox.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(406, 131);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(142, 52);
            this.button2.TabIndex = 3;
            this.button2.Text = "Mostrar contenidos de celda";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.ShowCellButton_Click);
            // 
            // CleanupButton
            // 
            this.CleanupButton.Location = new System.Drawing.Point(63, 177);
            this.CleanupButton.Name = "CleanupButton";
            this.CleanupButton.Size = new System.Drawing.Size(142, 52);
            this.CleanupButton.TabIndex = 4;
            this.CleanupButton.Text = "Cerrar Excel Cargado";
            this.CleanupButton.UseVisualStyleBackColor = true;
            this.CleanupButton.Click += new System.EventHandler(this.CleanupButton_Click);
            // 
            // CellSearchTextBox
            // 
            this.CellSearchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CellSearchTextBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CellSearchTextBox.Location = new System.Drawing.Point(406, 84);
            this.CellSearchTextBox.Name = "CellSearchTextBox";
            this.CellSearchTextBox.ReadOnly = true;
            this.CellSearchTextBox.Size = new System.Drawing.Size(138, 15);
            this.CellSearchTextBox.TabIndex = 6;
            this.CellSearchTextBox.Text = "Ubicacion de celda:";
            // 
            // cellBox1
            // 
            this.cellBox1.Location = new System.Drawing.Point(406, 105);
            this.cellBox1.Name = "cellBox1";
            this.cellBox1.Size = new System.Drawing.Size(66, 20);
            this.cellBox1.TabIndex = 7;
            // 
            // cellBox2
            // 
            this.cellBox2.Location = new System.Drawing.Point(477, 105);
            this.cellBox2.Name = "cellBox2";
            this.cellBox2.Size = new System.Drawing.Size(72, 20);
            this.cellBox2.TabIndex = 8;
            // 
            // LoadingImage
            // 
            this.LoadingImage.Image = global::MF_XLS_Parser.Properties.Resources.loading;
            this.LoadingImage.InitialImage = global::MF_XLS_Parser.Properties.Resources.loading;
            this.LoadingImage.Location = new System.Drawing.Point(12, 280);
            this.LoadingImage.Name = "LoadingImage";
            this.LoadingImage.Size = new System.Drawing.Size(50, 52);
            this.LoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.LoadingImage.TabIndex = 11;
            this.LoadingImage.TabStop = false;
            this.LoadingImage.UseWaitCursor = true;
            this.LoadingImage.Visible = false;
            // 
            // RowConfirmButton
            // 
            this.RowConfirmButton.Enabled = false;
            this.RowConfirmButton.Location = new System.Drawing.Point(140, 362);
            this.RowConfirmButton.Name = "RowConfirmButton";
            this.RowConfirmButton.Size = new System.Drawing.Size(124, 44);
            this.RowConfirmButton.TabIndex = 12;
            this.RowConfirmButton.Text = "Confirmar filas (categoria,datos)";
            this.RowConfirmButton.UseVisualStyleBackColor = true;
            this.RowConfirmButton.Click += new System.EventHandler(this.RowConfirmButton_Click);
            // 
            // DataRowBox
            // 
            this.DataRowBox.Location = new System.Drawing.Point(270, 386);
            this.DataRowBox.Name = "DataRowBox";
            this.DataRowBox.Size = new System.Drawing.Size(66, 20);
            this.DataRowBox.TabIndex = 13;
            this.DataRowBox.Text = "9";
            // 
            // AppLoadingImage
            // 
            this.AppLoadingImage.Image = global::MF_XLS_Parser.Properties.Resources.loading;
            this.AppLoadingImage.InitialImage = global::MF_XLS_Parser.Properties.Resources.loading;
            this.AppLoadingImage.Location = new System.Drawing.Point(12, 119);
            this.AppLoadingImage.Name = "AppLoadingImage";
            this.AppLoadingImage.Size = new System.Drawing.Size(50, 52);
            this.AppLoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.AppLoadingImage.TabIndex = 14;
            this.AppLoadingImage.TabStop = false;
            this.AppLoadingImage.UseWaitCursor = true;
            this.AppLoadingImage.Visible = false;
            // 
            // FilterStartButton
            // 
            this.FilterStartButton.Enabled = false;
            this.FilterStartButton.Location = new System.Drawing.Point(211, 280);
            this.FilterStartButton.Name = "FilterStartButton";
            this.FilterStartButton.Size = new System.Drawing.Size(142, 52);
            this.FilterStartButton.TabIndex = 15;
            this.FilterStartButton.Text = "Empezar proceso (Filtro)";
            this.FilterStartButton.UseVisualStyleBackColor = true;
            this.FilterStartButton.Click += new System.EventHandler(this.FilterStartButton_Click);
            // 
            // TypeRowBox
            // 
            this.TypeRowBox.Location = new System.Drawing.Point(270, 362);
            this.TypeRowBox.Name = "TypeRowBox";
            this.TypeRowBox.Size = new System.Drawing.Size(66, 20);
            this.TypeRowBox.TabIndex = 19;
            this.TypeRowBox.Text = "5";
            // 
            // ListBox
            // 
            this.ListBox.Location = new System.Drawing.Point(565, 105);
            this.ListBox.Multiline = true;
            this.ListBox.Name = "ListBox";
            this.ListBox.ReadOnly = true;
            this.ListBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.ListBox.Size = new System.Drawing.Size(197, 288);
            this.ListBox.TabIndex = 20;
            this.ListBox.WordWrap = false;
            // 
            // MissingBox
            // 
            this.MissingBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.MissingBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MissingBox.Location = new System.Drawing.Point(565, 84);
            this.MissingBox.Name = "MissingBox";
            this.MissingBox.ReadOnly = true;
            this.MissingBox.Size = new System.Drawing.Size(138, 15);
            this.MissingBox.TabIndex = 21;
            this.MissingBox.Text = "Productos no encontrados:";
            // 
            // YearTextBox
            // 
            this.YearTextBox.Location = new System.Drawing.Point(441, 312);
            this.YearTextBox.Name = "YearTextBox";
            this.YearTextBox.Size = new System.Drawing.Size(66, 20);
            this.YearTextBox.TabIndex = 22;
            // 
            // MonthTextBox
            // 
            this.MonthTextBox.Location = new System.Drawing.Point(441, 280);
            this.MonthTextBox.Name = "MonthTextBox";
            this.MonthTextBox.Size = new System.Drawing.Size(66, 20);
            this.MonthTextBox.TabIndex = 23;
            // 
            // textBox3
            // 
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(382, 280);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(53, 15);
            this.textBox3.TabIndex = 24;
            this.textBox3.Text = "Mes:";
            // 
            // textBox4
            // 
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(382, 312);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(53, 15);
            this.textBox4.TabIndex = 25;
            this.textBox4.Text = "Año:";
            // 
            // CancelButton
            // 
            this.CancelButton.Enabled = false;
            this.CancelButton.Location = new System.Drawing.Point(63, 280);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(142, 52);
            this.CancelButton.TabIndex = 26;
            this.CancelButton.Text = "Cancelar";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(342, 365);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(184, 13);
            this.textBox1.TabIndex = 27;
            this.textBox1.Text = "Fila de primera seccion o en blanco";
            // 
            // textBox2
            // 
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(342, 389);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(130, 13);
            this.textBox2.TabIndex = 28;
            this.textBox2.Text = "Fila de primer elemento";
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.MonthTextBox);
            this.Controls.Add(this.YearTextBox);
            this.Controls.Add(this.MissingBox);
            this.Controls.Add(this.ListBox);
            this.Controls.Add(this.TypeRowBox);
            this.Controls.Add(this.FilterStartButton);
            this.Controls.Add(this.AppLoadingImage);
            this.Controls.Add(this.DataRowBox);
            this.Controls.Add(this.RowConfirmButton);
            this.Controls.Add(this.LoadingImage);
            this.Controls.Add(this.cellBox2);
            this.Controls.Add(this.cellBox1);
            this.Controls.Add(this.CellSearchTextBox);
            this.Controls.Add(this.CleanupButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.DataTextBox);
            this.Controls.Add(this.MainTextBox);
            this.Controls.Add(this.OpenFileButton);
            this.MaximizeBox = false;
            this.Name = "Main_Form";
            this.Text = " Procesador Excel Supermercado \"La Bomba\"";
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.AppLoadingImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.TextBox MainTextBox;
        private System.Windows.Forms.TextBox DataTextBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button CleanupButton;
        private System.Windows.Forms.TextBox CellSearchTextBox;
        private System.Windows.Forms.TextBox cellBox1;
        private System.Windows.Forms.TextBox cellBox2;
        private System.Windows.Forms.PictureBox LoadingImage;
        private System.Windows.Forms.Button RowConfirmButton;
        private System.Windows.Forms.TextBox DataRowBox;
        private System.Windows.Forms.PictureBox AppLoadingImage;
        private System.Windows.Forms.Button FilterStartButton;
        private System.Windows.Forms.TextBox TypeRowBox;
        private System.Windows.Forms.TextBox ListBox;
        private System.Windows.Forms.TextBox MissingBox;
        private System.Windows.Forms.TextBox YearTextBox;
        private System.Windows.Forms.TextBox MonthTextBox;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
    }
}

