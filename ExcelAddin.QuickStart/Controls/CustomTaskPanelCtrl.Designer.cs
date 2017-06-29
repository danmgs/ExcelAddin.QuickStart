namespace ExcelAddin.QuickStart.Controls
{
    partial class CustomTaskPanelCtrl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonLaunchAction = new System.Windows.Forms.Button();
            this.textBoxLogin = new System.Windows.Forms.TextBox();
            this.labelLogin = new System.Windows.Forms.Label();
            this.labelPassword = new System.Windows.Forms.Label();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.tabMainControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.labelTitleTab1 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.labelTitleTab2 = new System.Windows.Forms.Label();
            this.labelVersionTab1 = new System.Windows.Forms.Label();
            this.tabMainControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonLaunchAction
            // 
            this.buttonLaunchAction.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonLaunchAction.Location = new System.Drawing.Point(60, 207);
            this.buttonLaunchAction.Name = "buttonLaunchAction";
            this.buttonLaunchAction.Size = new System.Drawing.Size(159, 44);
            this.buttonLaunchAction.TabIndex = 2;
            this.buttonLaunchAction.Text = "Launch Action";
            this.buttonLaunchAction.UseVisualStyleBackColor = true;
            this.buttonLaunchAction.Click += new System.EventHandler(this.buttonLaunchAction_Click);
            // 
            // textBoxLogin
            // 
            this.textBoxLogin.Location = new System.Drawing.Point(119, 105);
            this.textBoxLogin.Name = "textBoxLogin";
            this.textBoxLogin.Size = new System.Drawing.Size(100, 20);
            this.textBoxLogin.TabIndex = 3;
            // 
            // labelLogin
            // 
            this.labelLogin.AutoSize = true;
            this.labelLogin.Location = new System.Drawing.Point(56, 112);
            this.labelLogin.Name = "labelLogin";
            this.labelLogin.Size = new System.Drawing.Size(33, 13);
            this.labelLogin.TabIndex = 4;
            this.labelLogin.Text = "Login";
            // 
            // labelPassword
            // 
            this.labelPassword.AutoSize = true;
            this.labelPassword.Location = new System.Drawing.Point(56, 154);
            this.labelPassword.Name = "labelPassword";
            this.labelPassword.Size = new System.Drawing.Size(53, 13);
            this.labelPassword.TabIndex = 5;
            this.labelPassword.Text = "Password";
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(119, 147);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.Size = new System.Drawing.Size(100, 20);
            this.textBoxPassword.TabIndex = 6;
            this.textBoxPassword.UseSystemPasswordChar = true;
            // 
            // tabMainControl
            // 
            this.tabMainControl.Controls.Add(this.tabPage1);
            this.tabMainControl.Controls.Add(this.tabPage2);
            this.tabMainControl.Location = new System.Drawing.Point(34, 27);
            this.tabMainControl.Name = "tabMainControl";
            this.tabMainControl.SelectedIndex = 0;
            this.tabMainControl.Size = new System.Drawing.Size(301, 404);
            this.tabMainControl.TabIndex = 7;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.labelVersionTab1);
            this.tabPage1.Controls.Add(this.labelTitleTab1);
            this.tabPage1.Controls.Add(this.textBoxLogin);
            this.tabPage1.Controls.Add(this.textBoxPassword);
            this.tabPage1.Controls.Add(this.labelPassword);
            this.tabPage1.Controls.Add(this.labelLogin);
            this.tabPage1.Controls.Add(this.buttonLaunchAction);
            this.tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(293, 378);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Tab 1";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // labelTitleTab1
            // 
            this.labelTitleTab1.AutoSize = true;
            this.labelTitleTab1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitleTab1.Location = new System.Drawing.Point(112, 35);
            this.labelTitleTab1.Name = "labelTitleTab1";
            this.labelTitleTab1.Size = new System.Drawing.Size(71, 13);
            this.labelTitleTab1.TabIndex = 7;
            this.labelTitleTab1.Text = "Connection";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.labelTitleTab2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(293, 378);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Tab 2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // labelTitleTab2
            // 
            this.labelTitleTab2.AutoSize = true;
            this.labelTitleTab2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitleTab2.Location = new System.Drawing.Point(112, 35);
            this.labelTitleTab2.Name = "labelTitleTab2";
            this.labelTitleTab2.Size = new System.Drawing.Size(70, 13);
            this.labelTitleTab2.TabIndex = 0;
            this.labelTitleTab2.Text = "Parameters";
            // 
            // labelVersionTab1
            // 
            this.labelVersionTab1.AutoSize = true;
            this.labelVersionTab1.Location = new System.Drawing.Point(237, 359);
            this.labelVersionTab1.Name = "labelVersionTab1";
            this.labelVersionTab1.Size = new System.Drawing.Size(10, 13);
            this.labelVersionTab1.TabIndex = 8;
            this.labelVersionTab1.Text = "-";
            // 
            // CustomTaskPanelCtrl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tabMainControl);
            this.Name = "CustomTaskPanelCtrl";
            this.Size = new System.Drawing.Size(376, 510);
            this.tabMainControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonLaunchAction;
        private System.Windows.Forms.TextBox textBoxLogin;
        private System.Windows.Forms.Label labelLogin;
        private System.Windows.Forms.Label labelPassword;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.TabControl tabMainControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label labelTitleTab2;
        private System.Windows.Forms.Label labelTitleTab1;
        private System.Windows.Forms.Label labelVersionTab1;
    }
}
