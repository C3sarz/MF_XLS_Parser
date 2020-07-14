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
            this.StartButton = new System.Windows.Forms.Button();
            this.LoadingImage = new System.Windows.Forms.PictureBox();
            this.RowConfirmButton = new System.Windows.Forms.Button();
            this.RowBox1 = new System.Windows.Forms.TextBox();
            this.AppLoadingImage = new System.Windows.Forms.PictureBox();
            this.FilterStartButton = new System.Windows.Forms.Button();
            this.TestButton = new System.Windows.Forms.Button();
            this.RowBox2 = new System.Windows.Forms.TextBox();
            this.ListBox = new System.Windows.Forms.TextBox();
            this.MissingBox = new System.Windows.Forms.TextBox();
            this.YearTextBox = new System.Windows.Forms.TextBox();
            this.MonthTextBox = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
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
            this.MainTextBox.Size = new System.Drawing.Size(142, 41);
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
            this.button2.Location = new System.Drawing.Point(406, 140);
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
            // StartButton
            // 
            this.StartButton.Enabled = false;
            this.StartButton.Location = new System.Drawing.Point(63, 280);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(142, 52);
            this.StartButton.TabIndex = 10;
            this.StartButton.Text = "Empezar proceso";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
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
            this.RowConfirmButton.Location = new System.Drawing.Point(140, 352);
            this.RowConfirmButton.Name = "RowConfirmButton";
            this.RowConfirmButton.Size = new System.Drawing.Size(124, 52);
            this.RowConfirmButton.TabIndex = 12;
            this.RowConfirmButton.Text = "Confirmar primera fila de datos";
            this.RowConfirmButton.UseVisualStyleBackColor = true;
            this.RowConfirmButton.Click += new System.EventHandler(this.RowConfirmButton_Click);
            // 
            // RowBox1
            // 
            this.RowBox1.Location = new System.Drawing.Point(270, 386);
            this.RowBox1.Name = "RowBox1";
            this.RowBox1.Size = new System.Drawing.Size(66, 20);
            this.RowBox1.TabIndex = 13;
            this.RowBox1.Text = "9";
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
            // TestButton
            // 
            this.TestButton.Enabled = false;
            this.TestButton.Location = new System.Drawing.Point(211, 177);
            this.TestButton.Name = "TestButton";
            this.TestButton.Size = new System.Drawing.Size(142, 52);
            this.TestButton.TabIndex = 17;
            this.TestButton.Text = "Prueba";
            this.TestButton.UseVisualStyleBackColor = true;
            this.TestButton.Click += new System.EventHandler(this.TestButton_Click);
            // 
            // RowBox2
            // 
            this.RowBox2.Location = new System.Drawing.Point(270, 362);
            this.RowBox2.Name = "RowBox2";
            this.RowBox2.Size = new System.Drawing.Size(66, 20);
            this.RowBox2.TabIndex = 19;
            this.RowBox2.Text = "5";
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
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.MonthTextBox);
            this.Controls.Add(this.YearTextBox);
            this.Controls.Add(this.MissingBox);
            this.Controls.Add(this.ListBox);
            this.Controls.Add(this.RowBox2);
            this.Controls.Add(this.TestButton);
            this.Controls.Add(this.FilterStartButton);
            this.Controls.Add(this.AppLoadingImage);
            this.Controls.Add(this.RowBox1);
            this.Controls.Add(this.RowConfirmButton);
            this.Controls.Add(this.LoadingImage);
            this.Controls.Add(this.StartButton);
            this.Controls.Add(this.cellBox2);
            this.Controls.Add(this.cellBox1);
            this.Controls.Add(this.CellSearchTextBox);
            this.Controls.Add(this.CleanupButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.DataTextBox);
            this.Controls.Add(this.MainTextBox);
            this.Controls.Add(this.OpenFileButton);
            this.Name = "Main_Form";
            this.Text = " ";
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
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.PictureBox LoadingImage;
        private System.Windows.Forms.Button RowConfirmButton;
        private System.Windows.Forms.TextBox RowBox1;
        private System.Windows.Forms.PictureBox AppLoadingImage;
        private System.Windows.Forms.Button FilterStartButton;
        private System.Windows.Forms.Button TestButton;
        private System.Windows.Forms.TextBox RowBox2;
        private System.Windows.Forms.TextBox ListBox;
        private System.Windows.Forms.TextBox MissingBox;
        private System.Windows.Forms.TextBox YearTextBox;
        private System.Windows.Forms.TextBox MonthTextBox;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
    }
}

