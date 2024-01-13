namespace WinFormsApp1
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
            browseInput = new Button();
            label1 = new Label();
            textBox1 = new TextBox();
            run = new Button();
            SuspendLayout();
            // 
            // browseInput
            // 
            browseInput.Location = new Point(184, 27);
            browseInput.Name = "browseInput";
            browseInput.Size = new Size(75, 23);
            browseInput.TabIndex = 0;
            browseInput.Text = "Browse";
            browseInput.UseVisualStyleBackColor = true;
            browseInput.Click += browseInput_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(55, 15);
            label1.TabIndex = 2;
            label1.Text = "Input xls:";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(12, 27);
            textBox1.Name = "textBox1";
            textBox1.ReadOnly = true;
            textBox1.Size = new Size(166, 23);
            textBox1.TabIndex = 4;
            // 
            // run
            // 
            run.Location = new Point(12, 65);
            run.Name = "run";
            run.Size = new Size(247, 23);
            run.TabIndex = 6;
            run.Text = "Run";
            run.UseVisualStyleBackColor = true;
            run.Click += run_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(273, 95);
            Controls.Add(run);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Controls.Add(browseInput);
            Name = "Form1";
            Text = "NN renamer";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button browseInput;

        private Label label1;
        private TextBox textBox1;
        private Button run;

    }
}