using System.Runtime.InteropServices;

namespace OutlookDSD
{
    partial class OptionsPage: System.Windows.Forms.UserControl, Microsoft.Office.Interop.Outlook.PropertyPage
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBox_ribbon_showOnExplorer = new System.Windows.Forms.CheckBox();
            this.checkBox_ribbon_showOnEmail = new System.Windows.Forms.CheckBox();
            this.checkBox_bar_show = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lbl_VersionInstalled = new System.Windows.Forms.Label();
            this.lbl_VersionAvaliable = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_Update = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btn_Help = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox1.Image = global::OutlookDSD.Properties.Resources.logo;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(215, 103);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 116);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Validation Display";
            // 
            // checkBox_ribbon_showOnExplorer
            // 
            this.checkBox_ribbon_showOnExplorer.AutoSize = true;
            this.checkBox_ribbon_showOnExplorer.Location = new System.Drawing.Point(9, 134);
            this.checkBox_ribbon_showOnExplorer.Name = "checkBox_ribbon_showOnExplorer";
            this.checkBox_ribbon_showOnExplorer.Size = new System.Drawing.Size(206, 17);
            this.checkBox_ribbon_showOnExplorer.TabIndex = 3;
            this.checkBox_ribbon_showOnExplorer.Text = "Show Validation on Ribbon in Explorer";
            this.checkBox_ribbon_showOnExplorer.UseVisualStyleBackColor = true;
            this.checkBox_ribbon_showOnExplorer.CheckedChanged += new System.EventHandler(this.SettingChanged);
            // 
            // checkBox_ribbon_showOnEmail
            // 
            this.checkBox_ribbon_showOnEmail.AutoSize = true;
            this.checkBox_ribbon_showOnEmail.Location = new System.Drawing.Point(9, 157);
            this.checkBox_ribbon_showOnEmail.Name = "checkBox_ribbon_showOnEmail";
            this.checkBox_ribbon_showOnEmail.Size = new System.Drawing.Size(267, 17);
            this.checkBox_ribbon_showOnEmail.TabIndex = 4;
            this.checkBox_ribbon_showOnEmail.Text = "Show Validation on Ribbon in Message/Email View";
            this.checkBox_ribbon_showOnEmail.UseVisualStyleBackColor = true;
            this.checkBox_ribbon_showOnEmail.CheckedChanged += new System.EventHandler(this.SettingChanged);
            // 
            // checkBox_bar_show
            // 
            this.checkBox_bar_show.AutoSize = true;
            this.checkBox_bar_show.Location = new System.Drawing.Point(9, 180);
            this.checkBox_bar_show.Name = "checkBox_bar_show";
            this.checkBox_bar_show.Size = new System.Drawing.Size(121, 17);
            this.checkBox_bar_show.TabIndex = 5;
            this.checkBox_bar_show.Text = "Show Validation Bar";
            this.checkBox_bar_show.UseVisualStyleBackColor = true;
            this.checkBox_bar_show.CheckedChanged += new System.EventHandler(this.SettingChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(221, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Add-in Information";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(224, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Installed Version:";
            // 
            // lbl_VersionInstalled
            // 
            this.lbl_VersionInstalled.AutoSize = true;
            this.lbl_VersionInstalled.Location = new System.Drawing.Point(321, 20);
            this.lbl_VersionInstalled.Name = "lbl_VersionInstalled";
            this.lbl_VersionInstalled.Size = new System.Drawing.Size(40, 13);
            this.lbl_VersionInstalled.TabIndex = 8;
            this.lbl_VersionInstalled.Text = "1.1.1.1";
            // 
            // lbl_VersionAvaliable
            // 
            this.lbl_VersionAvaliable.AutoSize = true;
            this.lbl_VersionAvaliable.Location = new System.Drawing.Point(321, 36);
            this.lbl_VersionAvaliable.Name = "lbl_VersionAvaliable";
            this.lbl_VersionAvaliable.Size = new System.Drawing.Size(40, 13);
            this.lbl_VersionAvaliable.TabIndex = 10;
            this.lbl_VersionAvaliable.Text = "1.1.1.1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(224, 36);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Latest Version:";
            // 
            // btn_Update
            // 
            this.btn_Update.Location = new System.Drawing.Point(224, 57);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(90, 29);
            this.btn_Update.TabIndex = 11;
            this.btn_Update.Text = "Update Add-in";
            this.btn_Update.UseVisualStyleBackColor = true;
            this.btn_Update.Visible = false;
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label4.Location = new System.Drawing.Point(247, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(170, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Copyright (C) 2023 Matthew Hana.";
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label6.Location = new System.Drawing.Point(0, 110);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(450, 2);
            this.label6.TabIndex = 13;
            this.label6.Text = "XX";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_Help
            // 
            this.btn_Help.Location = new System.Drawing.Point(320, 57);
            this.btn_Help.Name = "btn_Help";
            this.btn_Help.Size = new System.Drawing.Size(90, 29);
            this.btn_Help.TabIndex = 14;
            this.btn_Help.Text = "Help";
            this.btn_Help.UseVisualStyleBackColor = true;
            this.btn_Help.Click += new System.EventHandler(this.btn_Help_Click);
            // 
            // OptionsPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btn_Help);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btn_Update);
            this.Controls.Add(this.lbl_VersionAvaliable);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lbl_VersionInstalled);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkBox_bar_show);
            this.Controls.Add(this.checkBox_ribbon_showOnEmail);
            this.Controls.Add(this.checkBox_ribbon_showOnExplorer);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "OptionsPage";
            this.Size = new System.Drawing.Size(420, 420);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox_ribbon_showOnExplorer;
        private System.Windows.Forms.CheckBox checkBox_ribbon_showOnEmail;
        private System.Windows.Forms.CheckBox checkBox_bar_show;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lbl_VersionInstalled;
        private System.Windows.Forms.Label lbl_VersionAvaliable;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btn_Update;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btn_Help;
    }
}