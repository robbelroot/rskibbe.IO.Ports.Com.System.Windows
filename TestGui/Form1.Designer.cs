﻿namespace TestGui;

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
            this.lbConnectedDevices = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // lbConnectedDevices
            // 
            this.lbConnectedDevices.FormattingEnabled = true;
            this.lbConnectedDevices.ItemHeight = 15;
            this.lbConnectedDevices.Location = new System.Drawing.Point(172, 78);
            this.lbConnectedDevices.Name = "lbConnectedDevices";
            this.lbConnectedDevices.Size = new System.Drawing.Size(235, 244);
            this.lbConnectedDevices.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lbConnectedDevices);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

    }

    #endregion

    private ListBox lbConnectedDevices;
}
