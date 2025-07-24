using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;


namespace WindowsFormsApp1
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
            this.textbox = new System.Windows.Forms.TextBox();
            this.directorylabel = new System.Windows.Forms.Label();
            this.browsebutton = new System.Windows.Forms.Button();
            this.modellabel = new System.Windows.Forms.Label();
            this.mixerlabel = new System.Windows.Forms.Label();
            this.doublerlabel = new System.Windows.Forms.Label();
            this.mixercheckbox = new System.Windows.Forms.CheckBox();
            this.doublercheckbox = new System.Windows.Forms.CheckBox();
            this.generatebutton = new System.Windows.Forms.Button();
            this.cancelbutton = new System.Windows.Forms.Button();
            this.resultlabel = new System.Windows.Forms.Label();
<<<<<<< HEAD
            this.outputlabel = new System.Windows.Forms.Label();
            this.textbox2 = new System.Windows.Forms.TextBox();
            this.outputbutton = new System.Windows.Forms.Button();
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.SuspendLayout();
            // 
            // textbox
            // 
            this.textbox.Location = new System.Drawing.Point(195, 37);
            this.textbox.Name = "textbox";
<<<<<<< HEAD
            this.textbox.Size = new System.Drawing.Size(368, 22);
=======
            this.textbox.Size = new System.Drawing.Size(332, 22);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.textbox.TabIndex = 0;
            // 
            // directorylabel
            // 
            this.directorylabel.AutoSize = true;
            this.directorylabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.directorylabel.Location = new System.Drawing.Point(12, 37);
            this.directorylabel.Name = "directorylabel";
            this.directorylabel.Size = new System.Drawing.Size(177, 20);
            this.directorylabel.TabIndex = 1;
            this.directorylabel.Text = "Directory file Excel:";
            this.directorylabel.Click += new System.EventHandler(this.label1_Click);
            // 
            // browsebutton
            // 
<<<<<<< HEAD
            this.browsebutton.Location = new System.Drawing.Point(569, 31);
            this.browsebutton.Name = "browsebutton";
            this.browsebutton.Size = new System.Drawing.Size(110, 35);
=======
            this.browsebutton.Location = new System.Drawing.Point(559, 34);
            this.browsebutton.Name = "browsebutton";
            this.browsebutton.Size = new System.Drawing.Size(130, 49);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.browsebutton.TabIndex = 2;
            this.browsebutton.Text = "Browse";
            this.browsebutton.UseVisualStyleBackColor = true;
            // 
            // modellabel
            // 
            this.modellabel.AutoSize = true;
            this.modellabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
<<<<<<< HEAD
            this.modellabel.Location = new System.Drawing.Point(124, 136);
=======
            this.modellabel.Location = new System.Drawing.Point(124, 78);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.modellabel.Name = "modellabel";
            this.modellabel.Size = new System.Drawing.Size(65, 20);
            this.modellabel.TabIndex = 3;
            this.modellabel.Text = "Model:";
            // 
            // mixerlabel
            // 
            this.mixerlabel.AutoSize = true;
            this.mixerlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
<<<<<<< HEAD
            this.mixerlabel.Location = new System.Drawing.Point(195, 136);
=======
            this.mixerlabel.Location = new System.Drawing.Point(195, 78);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.mixerlabel.Name = "mixerlabel";
            this.mixerlabel.Size = new System.Drawing.Size(130, 20);
            this.mixerlabel.TabIndex = 4;
            this.mixerlabel.Text = "Mixer CSM2-13 ";
            this.mixerlabel.Click += new System.EventHandler(this.label3_Click);
            // 
            // doublerlabel
            // 
            this.doublerlabel.AutoSize = true;
            this.doublerlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
<<<<<<< HEAD
            this.doublerlabel.Location = new System.Drawing.Point(195, 171);
=======
            this.doublerlabel.Location = new System.Drawing.Point(195, 109);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.doublerlabel.Name = "doublerlabel";
            this.doublerlabel.Size = new System.Drawing.Size(150, 20);
            this.doublerlabel.TabIndex = 5;
            this.doublerlabel.Text = "Doubler CSFD25H";
            // 
            // mixercheckbox
            // 
            this.mixercheckbox.AutoSize = true;
<<<<<<< HEAD
            this.mixercheckbox.Location = new System.Drawing.Point(377, 140);
=======
            this.mixercheckbox.Location = new System.Drawing.Point(377, 82);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.mixercheckbox.Name = "mixercheckbox";
            this.mixercheckbox.Size = new System.Drawing.Size(18, 17);
            this.mixercheckbox.TabIndex = 6;
            this.mixercheckbox.UseVisualStyleBackColor = true;
            // 
            // doublercheckbox
            // 
            this.doublercheckbox.AutoSize = true;
<<<<<<< HEAD
            this.doublercheckbox.Location = new System.Drawing.Point(377, 175);
=======
            this.doublercheckbox.Location = new System.Drawing.Point(377, 113);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.doublercheckbox.Name = "doublercheckbox";
            this.doublercheckbox.Size = new System.Drawing.Size(18, 17);
            this.doublercheckbox.TabIndex = 7;
            this.doublercheckbox.UseVisualStyleBackColor = true;
            // 
            // generatebutton
            // 
<<<<<<< HEAD
            this.generatebutton.Location = new System.Drawing.Point(559, 144);
=======
            this.generatebutton.Location = new System.Drawing.Point(559, 96);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.generatebutton.Name = "generatebutton";
            this.generatebutton.Size = new System.Drawing.Size(130, 48);
            this.generatebutton.TabIndex = 8;
            this.generatebutton.Text = "Generate";
            this.generatebutton.UseVisualStyleBackColor = true;
            this.generatebutton.Click += new System.EventHandler(this.button2_Click);
            // 
            // cancelbutton
            // 
<<<<<<< HEAD
            this.cancelbutton.Location = new System.Drawing.Point(559, 198);
=======
            this.cancelbutton.Location = new System.Drawing.Point(559, 158);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.cancelbutton.Name = "cancelbutton";
            this.cancelbutton.Size = new System.Drawing.Size(130, 48);
            this.cancelbutton.TabIndex = 9;
            this.cancelbutton.Text = "Cancel";
            this.cancelbutton.UseVisualStyleBackColor = true;
            // 
            // resultlabel
            // 
            this.resultlabel.AutoSize = true;
            this.resultlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
<<<<<<< HEAD
            this.resultlabel.Location = new System.Drawing.Point(12, 265);
=======
            this.resultlabel.Location = new System.Drawing.Point(12, 171);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.resultlabel.Name = "resultlabel";
            this.resultlabel.Size = new System.Drawing.Size(75, 20);
            this.resultlabel.TabIndex = 10;
            this.resultlabel.Text = "Result: ";
            // 
<<<<<<< HEAD
            // outputlabel
            // 
            this.outputlabel.AutoSize = true;
            this.outputlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.outputlabel.Location = new System.Drawing.Point(57, 87);
            this.outputlabel.Name = "outputlabel";
            this.outputlabel.Size = new System.Drawing.Size(132, 20);
            this.outputlabel.TabIndex = 11;
            this.outputlabel.Text = "Output file INI:";
            this.outputlabel.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // textbox2
            // 
            this.textbox2.Location = new System.Drawing.Point(195, 85);
            this.textbox2.Name = "textbox2";
            this.textbox2.Size = new System.Drawing.Size(368, 22);
            this.textbox2.TabIndex = 12;
            // 
            // outputbutton
            // 
            this.outputbutton.Location = new System.Drawing.Point(569, 80);
            this.outputbutton.Name = "outputbutton";
            this.outputbutton.Size = new System.Drawing.Size(109, 36);
            this.outputbutton.TabIndex = 13;
            this.outputbutton.Text = "Browse";
            this.outputbutton.UseVisualStyleBackColor = true;
            // 
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
<<<<<<< HEAD
            this.ClientSize = new System.Drawing.Size(714, 294);
            this.Controls.Add(this.outputbutton);
            this.Controls.Add(this.textbox2);
            this.Controls.Add(this.outputlabel);
=======
            this.ClientSize = new System.Drawing.Size(714, 218);
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
            this.Controls.Add(this.resultlabel);
            this.Controls.Add(this.cancelbutton);
            this.Controls.Add(this.generatebutton);
            this.Controls.Add(this.doublercheckbox);
            this.Controls.Add(this.mixercheckbox);
            this.Controls.Add(this.doublerlabel);
            this.Controls.Add(this.mixerlabel);
            this.Controls.Add(this.modellabel);
            this.Controls.Add(this.browsebutton);
            this.Controls.Add(this.directorylabel);
            this.Controls.Add(this.textbox);
            this.Name = "Form1";
            this.Text = "MACOM_converter ";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TextBox textbox;            //text box for INPUT FILE EXCEL         
        private Label directorylabel;                // label Directory    
        private Button browsebutton ;       // button browse 
        private Label modellabel;          // label model 
        private Label mixerlabel ;          // lable mixer 
        private Label doublerlabel ;        // label doubler 
        private CheckBox mixercheckbox ;      // check box for mixer
        private CheckBox doublercheckbox ;    // check box for doubler
        private Button generatebutton ;      // button generate
        private Button cancelbutton ;        // button cancel
        private Label resultlabel;
<<<<<<< HEAD
        private Label outputlabel;
        private TextBox textbox2;
        private Button outputbutton;
=======
>>>>>>> 2742ae2f9991af8eb6b861e2a4bb4cf034f1f554
    }
}

