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
            this.InfoTextBox = new System.Windows.Forms.TextBox();
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
            this.FilterBox = new System.Windows.Forms.TextBox();
            this.TestButton = new System.Windows.Forms.Button();
            this.RowBox2 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.AppLoadingImage)).BeginInit();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(84, 146);
            this.OpenFileButton.Margin = new System.Windows.Forms.Padding(4);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(189, 64);
            this.OpenFileButton.TabIndex = 0;
            this.OpenFileButton.Text = "Abrir archivo";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // InfoTextBox
            // 
            this.InfoTextBox.Location = new System.Drawing.Point(84, 26);
            this.InfoTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.InfoTextBox.Multiline = true;
            this.InfoTextBox.Name = "InfoTextBox";
            this.InfoTextBox.ReadOnly = true;
            this.InfoTextBox.Size = new System.Drawing.Size(583, 94);
            this.InfoTextBox.TabIndex = 1;
            // 
            // DataTextBox
            // 
            this.DataTextBox.Location = new System.Drawing.Point(728, 26);
            this.DataTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.DataTextBox.Multiline = true;
            this.DataTextBox.Name = "DataTextBox";
            this.DataTextBox.Size = new System.Drawing.Size(268, 94);
            this.DataTextBox.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(673, 146);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(189, 64);
            this.button2.TabIndex = 3;
            this.button2.Text = "Mostrar contenidos de celda";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.ShowCellButton_Click);
            // 
            // CleanupButton
            // 
            this.CleanupButton.Location = new System.Drawing.Point(84, 218);
            this.CleanupButton.Margin = new System.Windows.Forms.Padding(4);
            this.CleanupButton.Name = "CleanupButton";
            this.CleanupButton.Size = new System.Drawing.Size(189, 64);
            this.CleanupButton.TabIndex = 4;
            this.CleanupButton.Text = "Cerrar Excel Cargado";
            this.CleanupButton.UseVisualStyleBackColor = true;
            this.CleanupButton.Click += new System.EventHandler(this.CleanupButton_Click);
            // 
            // CellSearchTextBox
            // 
            this.CellSearchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CellSearchTextBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CellSearchTextBox.Location = new System.Drawing.Point(870, 146);
            this.CellSearchTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.CellSearchTextBox.Name = "CellSearchTextBox";
            this.CellSearchTextBox.ReadOnly = true;
            this.CellSearchTextBox.Size = new System.Drawing.Size(184, 18);
            this.CellSearchTextBox.TabIndex = 6;
            this.CellSearchTextBox.Text = "Ubicacion de celda:";
            // 
            // cellBox1
            // 
            this.cellBox1.Location = new System.Drawing.Point(870, 172);
            this.cellBox1.Margin = new System.Windows.Forms.Padding(4);
            this.cellBox1.Name = "cellBox1";
            this.cellBox1.Size = new System.Drawing.Size(87, 22);
            this.cellBox1.TabIndex = 7;
            // 
            // cellBox2
            // 
            this.cellBox2.Location = new System.Drawing.Point(966, 172);
            this.cellBox2.Margin = new System.Windows.Forms.Padding(4);
            this.cellBox2.Name = "cellBox2";
            this.cellBox2.Size = new System.Drawing.Size(87, 22);
            this.cellBox2.TabIndex = 8;
            // 
            // StartButton
            // 
            this.StartButton.Enabled = false;
            this.StartButton.Location = new System.Drawing.Point(84, 345);
            this.StartButton.Margin = new System.Windows.Forms.Padding(4);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(189, 64);
            this.StartButton.TabIndex = 10;
            this.StartButton.Text = "Empezar proceso";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // LoadingImage
            // 
            this.LoadingImage.Image = global::MF_XLS_Parser.Properties.Resources.loading;
            this.LoadingImage.InitialImage = global::MF_XLS_Parser.Properties.Resources.loading;
            this.LoadingImage.Location = new System.Drawing.Point(16, 345);
            this.LoadingImage.Margin = new System.Windows.Forms.Padding(4);
            this.LoadingImage.Name = "LoadingImage";
            this.LoadingImage.Size = new System.Drawing.Size(67, 64);
            this.LoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.LoadingImage.TabIndex = 11;
            this.LoadingImage.TabStop = false;
            this.LoadingImage.UseWaitCursor = true;
            this.LoadingImage.Visible = false;
            // 
            // RowConfirmButton
            // 
            this.RowConfirmButton.Enabled = false;
            this.RowConfirmButton.Location = new System.Drawing.Point(673, 345);
            this.RowConfirmButton.Margin = new System.Windows.Forms.Padding(4);
            this.RowConfirmButton.Name = "RowConfirmButton";
            this.RowConfirmButton.Size = new System.Drawing.Size(165, 64);
            this.RowConfirmButton.TabIndex = 12;
            this.RowConfirmButton.Text = "Confirmar primera fila de datos";
            this.RowConfirmButton.UseVisualStyleBackColor = true;
            this.RowConfirmButton.Click += new System.EventHandler(this.RowConfirmButton_Click);
            // 
            // RowBox1
            // 
            this.RowBox1.Location = new System.Drawing.Point(846, 366);
            this.RowBox1.Margin = new System.Windows.Forms.Padding(4);
            this.RowBox1.Name = "RowBox1";
            this.RowBox1.Size = new System.Drawing.Size(87, 22);
            this.RowBox1.TabIndex = 13;
            this.RowBox1.Text = "9";
            // 
            // AppLoadingImage
            // 
            this.AppLoadingImage.Image = global::MF_XLS_Parser.Properties.Resources.loading;
            this.AppLoadingImage.InitialImage = global::MF_XLS_Parser.Properties.Resources.loading;
            this.AppLoadingImage.Location = new System.Drawing.Point(16, 146);
            this.AppLoadingImage.Margin = new System.Windows.Forms.Padding(4);
            this.AppLoadingImage.Name = "AppLoadingImage";
            this.AppLoadingImage.Size = new System.Drawing.Size(67, 64);
            this.AppLoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.AppLoadingImage.TabIndex = 14;
            this.AppLoadingImage.TabStop = false;
            this.AppLoadingImage.UseWaitCursor = true;
            this.AppLoadingImage.Visible = false;
            // 
            // FilterStartButton
            // 
            this.FilterStartButton.Enabled = false;
            this.FilterStartButton.Location = new System.Drawing.Point(281, 345);
            this.FilterStartButton.Margin = new System.Windows.Forms.Padding(4);
            this.FilterStartButton.Name = "FilterStartButton";
            this.FilterStartButton.Size = new System.Drawing.Size(189, 64);
            this.FilterStartButton.TabIndex = 15;
            this.FilterStartButton.Text = "Empezar proceso (Filtro)";
            this.FilterStartButton.UseVisualStyleBackColor = true;
            this.FilterStartButton.Click += new System.EventHandler(this.FilterStartButton_Click);
            // 
            // FilterBox
            // 
            this.FilterBox.Location = new System.Drawing.Point(281, 417);
            this.FilterBox.Margin = new System.Windows.Forms.Padding(4);
            this.FilterBox.Name = "FilterBox";
            this.FilterBox.Size = new System.Drawing.Size(188, 22);
            this.FilterBox.TabIndex = 16;
            // 
            // TestButton
            // 
            this.TestButton.Enabled = false;
            this.TestButton.Location = new System.Drawing.Point(281, 218);
            this.TestButton.Margin = new System.Windows.Forms.Padding(4);
            this.TestButton.Name = "TestButton";
            this.TestButton.Size = new System.Drawing.Size(189, 64);
            this.TestButton.TabIndex = 17;
            this.TestButton.Text = "Prueba";
            this.TestButton.UseVisualStyleBackColor = true;
            this.TestButton.Click += new System.EventHandler(this.TestButton_Click);
            // 
            // RowBox2
            // 
            this.RowBox2.Location = new System.Drawing.Point(846, 438);
            this.RowBox2.Margin = new System.Windows.Forms.Padding(4);
            this.RowBox2.Name = "RowBox2";
            this.RowBox2.Size = new System.Drawing.Size(87, 22);
            this.RowBox2.TabIndex = 19;
            this.RowBox2.Text = "5";
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.RowBox2);
            this.Controls.Add(this.TestButton);
            this.Controls.Add(this.FilterBox);
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
            this.Controls.Add(this.InfoTextBox);
            this.Controls.Add(this.OpenFileButton);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Main_Form";
            this.Text = " ";
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.AppLoadingImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.TextBox InfoTextBox;
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
        private System.Windows.Forms.TextBox FilterBox;
        private System.Windows.Forms.Button TestButton;
        private System.Windows.Forms.TextBox RowBox2;
    }
}

