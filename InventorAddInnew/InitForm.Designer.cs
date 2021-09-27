
namespace InvAddIn
{
    partial class InitForm
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
            this.ShellButton = new System.Windows.Forms.Button();
            this.BottomPButton = new System.Windows.Forms.Button();
            this.NozzleButton = new System.Windows.Forms.Button();
            this.AssembleButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ShellButton
            // 
            this.ShellButton.Location = new System.Drawing.Point(79, 37);
            this.ShellButton.Name = "ShellButton";
            this.ShellButton.Size = new System.Drawing.Size(130, 23);
            this.ShellButton.TabIndex = 0;
            this.ShellButton.Text = "Create Shell Plates";
            this.ShellButton.UseVisualStyleBackColor = true;
            this.ShellButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // BottomPButton
            // 
            this.BottomPButton.Location = new System.Drawing.Point(79, 92);
            this.BottomPButton.Name = "BottomPButton";
            this.BottomPButton.Size = new System.Drawing.Size(130, 23);
            this.BottomPButton.TabIndex = 1;
            this.BottomPButton.Text = "Create Bottom Plates";
            this.BottomPButton.UseVisualStyleBackColor = true;
            this.BottomPButton.Click += new System.EventHandler(this.button2_Click);
            // 
            // NozzleButton
            // 
            this.NozzleButton.Location = new System.Drawing.Point(79, 143);
            this.NozzleButton.Name = "NozzleButton";
            this.NozzleButton.Size = new System.Drawing.Size(130, 23);
            this.NozzleButton.TabIndex = 2;
            this.NozzleButton.Text = "Create Nozzle";
            this.NozzleButton.UseVisualStyleBackColor = true;
            this.NozzleButton.Click += new System.EventHandler(this.button3_Click);
            // 
            // AssembleButton
            // 
            this.AssembleButton.Location = new System.Drawing.Point(79, 201);
            this.AssembleButton.Name = "AssembleButton";
            this.AssembleButton.Size = new System.Drawing.Size(130, 23);
            this.AssembleButton.TabIndex = 3;
            this.AssembleButton.Text = "Assemble Model";
            this.AssembleButton.UseVisualStyleBackColor = true;
            this.AssembleButton.Click += new System.EventHandler(this.AssembleButton_Click);
            // 
            // InitForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(297, 257);
            this.Controls.Add(this.AssembleButton);
            this.Controls.Add(this.NozzleButton);
            this.Controls.Add(this.BottomPButton);
            this.Controls.Add(this.ShellButton);
            this.Name = "InitForm";
            this.Text = "InitForm";
            this.Load += new System.EventHandler(this.InitForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ShellButton;
        private System.Windows.Forms.Button BottomPButton;
        private System.Windows.Forms.Button NozzleButton;
        private System.Windows.Forms.Button AssembleButton;
    }
}