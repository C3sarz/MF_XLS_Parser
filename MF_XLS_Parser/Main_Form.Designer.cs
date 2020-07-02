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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main_Form));
            this.Button1 = new System.Windows.Forms.Button();
            this.InfoTextBox = new System.Windows.Forms.TextBox();
            this.DataTextBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.CleanupButton = new System.Windows.Forms.Button();
            this.CellSearchTextBox = new System.Windows.Forms.TextBox();
            this.cellBox1 = new System.Windows.Forms.TextBox();
            this.cellBox2 = new System.Windows.Forms.TextBox();
            this.WriteButton = new System.Windows.Forms.Button();
            this.ParsingButton = new System.Windows.Forms.Button();
            this.LoadingImage = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).BeginInit();
            this.SuspendLayout();
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(84, 146);
            this.Button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(189, 64);
            this.Button1.TabIndex = 0;
            this.Button1.Text = "Abrir archivo";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // InfoTextBox
            // 
            this.InfoTextBox.Location = new System.Drawing.Point(84, 26);
            this.InfoTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.InfoTextBox.Multiline = true;
            this.InfoTextBox.Name = "InfoTextBox";
            this.InfoTextBox.ReadOnly = true;
            this.InfoTextBox.Size = new System.Drawing.Size(583, 94);
            this.InfoTextBox.TabIndex = 1;
            // 
            // DataTextBox
            // 
            this.DataTextBox.Location = new System.Drawing.Point(728, 26);
            this.DataTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.DataTextBox.Multiline = true;
            this.DataTextBox.Name = "DataTextBox";
            this.DataTextBox.Size = new System.Drawing.Size(268, 94);
            this.DataTextBox.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(84, 246);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(189, 64);
            this.button2.TabIndex = 3;
            this.button2.Text = "Show Cell Contents";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // CleanupButton
            // 
            this.CleanupButton.Location = new System.Drawing.Point(84, 345);
            this.CleanupButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CleanupButton.Name = "CleanupButton";
            this.CleanupButton.Size = new System.Drawing.Size(189, 64);
            this.CleanupButton.TabIndex = 4;
            this.CleanupButton.Text = "Close all";
            this.CleanupButton.UseVisualStyleBackColor = true;
            this.CleanupButton.Click += new System.EventHandler(this.CleanupButton_Click);
            // 
            // CellSearchTextBox
            // 
            this.CellSearchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CellSearchTextBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CellSearchTextBox.Location = new System.Drawing.Point(728, 183);
            this.CellSearchTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CellSearchTextBox.Name = "CellSearchTextBox";
            this.CellSearchTextBox.ReadOnly = true;
            this.CellSearchTextBox.Size = new System.Drawing.Size(184, 18);
            this.CellSearchTextBox.TabIndex = 6;
            this.CellSearchTextBox.Text = "Cell Selection:";
            // 
            // cellBox1
            // 
            this.cellBox1.Location = new System.Drawing.Point(728, 209);
            this.cellBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cellBox1.Name = "cellBox1";
            this.cellBox1.Size = new System.Drawing.Size(87, 22);
            this.cellBox1.TabIndex = 7;
            // 
            // cellBox2
            // 
            this.cellBox2.Location = new System.Drawing.Point(824, 209);
            this.cellBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cellBox2.Name = "cellBox2";
            this.cellBox2.Size = new System.Drawing.Size(87, 22);
            this.cellBox2.TabIndex = 8;
            // 
            // WriteButton
            // 
            this.WriteButton.Location = new System.Drawing.Point(297, 246);
            this.WriteButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.WriteButton.Name = "WriteButton";
            this.WriteButton.Size = new System.Drawing.Size(189, 64);
            this.WriteButton.TabIndex = 9;
            this.WriteButton.Text = "Write to Cell";
            this.WriteButton.UseVisualStyleBackColor = true;
            this.WriteButton.Click += new System.EventHandler(this.WriteButton_Click);
            // 
            // ParsingButton
            // 
            this.ParsingButton.Location = new System.Drawing.Point(297, 345);
            this.ParsingButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ParsingButton.Name = "ParsingButton";
            this.ParsingButton.Size = new System.Drawing.Size(189, 64);
            this.ParsingButton.TabIndex = 10;
            this.ParsingButton.Text = "Start parsing";
            this.ParsingButton.UseVisualStyleBackColor = true;
            this.ParsingButton.Click += new System.EventHandler(this.ParsingButton_Click);
            // 
            // LoadingImage
            // 
            this.LoadingImage.Image = ((System.Drawing.Image)(resources.GetObject("LoadingImage.Image")));
            this.LoadingImage.InitialImage = ((System.Drawing.Image)(resources.GetObject("LoadingImage.InitialImage")));
            this.LoadingImage.Location = new System.Drawing.Point(495, 345);
            this.LoadingImage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.LoadingImage.Name = "LoadingImage";
            this.LoadingImage.Size = new System.Drawing.Size(67, 60);
            this.LoadingImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.LoadingImage.TabIndex = 11;
            this.LoadingImage.TabStop = false;
            this.LoadingImage.UseWaitCursor = true;
            this.LoadingImage.Visible = false;
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.LoadingImage);
            this.Controls.Add(this.ParsingButton);
            this.Controls.Add(this.WriteButton);
            this.Controls.Add(this.cellBox2);
            this.Controls.Add(this.cellBox1);
            this.Controls.Add(this.CellSearchTextBox);
            this.Controls.Add(this.CleanupButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.DataTextBox);
            this.Controls.Add(this.InfoTextBox);
            this.Controls.Add(this.Button1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Main_Form";
            this.Text = " ";
            ((System.ComponentModel.ISupportInitialize)(this.LoadingImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Button1;
        private System.Windows.Forms.TextBox InfoTextBox;
        private System.Windows.Forms.TextBox DataTextBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button CleanupButton;
        private System.Windows.Forms.TextBox CellSearchTextBox;
        private System.Windows.Forms.TextBox cellBox1;
        private System.Windows.Forms.TextBox cellBox2;
        private System.Windows.Forms.Button WriteButton;
        private System.Windows.Forms.Button ParsingButton;
        private System.Windows.Forms.PictureBox LoadingImage;
    }
}

