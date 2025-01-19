using Microsoft.VisualBasic.ApplicationServices;
using System.Reflection;
using System.Reflection.Emit;

namespace DefiPilot
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            folderSelect = new FolderBrowserDialog();
            folderSelectBtn = new Button();
            VarCheck = new CheckBox();
            MsgCheck = new CheckBox();
            UnitCheck = new CheckBox();
            MsgHelpCheck = new CheckBox();
            BtnExecute = new Button();
            actionBox = new ComboBox();
            saveFileDialog1 = new SaveFileDialog();
            progressBar1 = new ProgressBar();
            labelPlc = new System.Windows.Forms.Label();
            labelAction = new System.Windows.Forms.Label();
            IOCheck = new CheckBox();
            RecetteCheck = new CheckBox();
            SuspendLayout();
            // 
            // folderSelectBtn
            // 
            folderSelectBtn.Location = new Point(12, 12);
            folderSelectBtn.Name = "folderSelectBtn";
            folderSelectBtn.Size = new Size(150, 50);
            folderSelectBtn.TabIndex = 0;
            folderSelectBtn.Text = "Sélectionner Dossier Machine";
            folderSelectBtn.UseVisualStyleBackColor = true;
            folderSelectBtn.Click += FolderSelectBtn_Click;
            // 
            // VarCheck
            // 
            VarCheck.AutoSize = true;
            VarCheck.BackColor = Color.Transparent;
            VarCheck.Location = new Point(15, 100);
            VarCheck.Name = "VarCheck";
            VarCheck.Size = new Size(72, 19);
            VarCheck.TabIndex = 1;
            VarCheck.Text = "Variables";
            VarCheck.UseVisualStyleBackColor = false;
            VarCheck.CheckedChanged += VarCheck_CheckedChanged;
            // 
            // MsgCheck
            // 
            MsgCheck.AutoSize = true;
            MsgCheck.BackColor = Color.Transparent;
            MsgCheck.Location = new Point(15, 125);
            MsgCheck.Name = "MsgCheck";
            MsgCheck.Size = new Size(77, 19);
            MsgCheck.TabIndex = 2;
            MsgCheck.Text = "Messages";
            MsgCheck.UseVisualStyleBackColor = false;
            MsgCheck.Click += MsgCheck_Click;
            // 
            // UnitCheck
            // 
            UnitCheck.AutoSize = true;
            UnitCheck.BackColor = Color.Transparent;
            UnitCheck.Location = new Point(98, 100);
            UnitCheck.Name = "UnitCheck";
            UnitCheck.Size = new Size(112, 19);
            UnitCheck.TabIndex = 3;
            UnitCheck.Text = "Unités / Options";
            UnitCheck.UseVisualStyleBackColor = false;
            UnitCheck.Visible = false;
            // 
            // MsgHelpCheck
            // 
            MsgHelpCheck.AutoSize = true;
            MsgHelpCheck.BackColor = Color.Transparent;
            MsgHelpCheck.Location = new Point(98, 125);
            MsgHelpCheck.Name = "MsgHelpCheck";
            MsgHelpCheck.Size = new Size(99, 19);
            MsgHelpCheck.TabIndex = 4;
            MsgHelpCheck.Text = "Aide Message";
            MsgHelpCheck.UseVisualStyleBackColor = false;
            MsgHelpCheck.Visible = false;
            // 
            // BtnExecute
            // 
            BtnExecute.Location = new Point(12, 268);
            BtnExecute.Name = "BtnExecute";
            BtnExecute.Size = new Size(150, 45);
            BtnExecute.TabIndex = 5;
            BtnExecute.Text = "GO GO GO !";
            BtnExecute.UseVisualStyleBackColor = true;
            BtnExecute.Click += CreateFiles;
            // 
            // actionBox
            // 
            actionBox.DropDownStyle = ComboBoxStyle.DropDownList;
            actionBox.FormattingEnabled = true;
            actionBox.Items.AddRange(new object[] { "Création", "Complétion" });
            actionBox.Location = new Point(12, 227);
            actionBox.Name = "actionBox";
            actionBox.Size = new Size(150, 23);
            actionBox.TabIndex = 0;
            actionBox.SelectedIndexChanged += actionBox_SelectedIndexChanged;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(12, 366);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(488, 23);
            progressBar1.Step = 1;
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.TabIndex = 6;
            progressBar1.Tag = "toto";
            progressBar1.Visible = false;
            // 
            // labelPlc
            // 
            labelPlc.AutoSize = true;
            labelPlc.BackColor = Color.Transparent;
            labelPlc.Location = new Point(12, 348);
            labelPlc.Name = "labelPlc";
            labelPlc.Size = new Size(23, 15);
            labelPlc.TabIndex = 7;
            labelPlc.Text = "plc";
            labelPlc.Visible = false;
            // 
            // labelAction
            // 
            labelAction.AutoSize = true;
            labelAction.BackColor = Color.Transparent;
            labelAction.Location = new Point(98, 348);
            labelAction.Name = "labelAction";
            labelAction.Size = new Size(42, 15);
            labelAction.TabIndex = 8;
            labelAction.Text = "Action";
            labelAction.Visible = false;
            // 
            // IOCheck
            // 
            IOCheck.AutoSize = true;
            IOCheck.BackColor = Color.Transparent;
            IOCheck.Location = new Point(15, 150);
            IOCheck.Name = "IOCheck";
            IOCheck.Size = new Size(110, 19);
            IOCheck.TabIndex = 9;
            IOCheck.Text = "Entrées / Sorties";
            IOCheck.UseVisualStyleBackColor = false;
            // 
            // RecetteCheck
            // 
            RecetteCheck.AutoSize = true;
            RecetteCheck.BackColor = Color.Transparent;
            RecetteCheck.Location = new Point(15, 175);
            RecetteCheck.Name = "RecetteCheck";
            RecetteCheck.Size = new Size(70, 19);
            RecetteCheck.TabIndex = 10;
            RecetteCheck.Text = "Recettes";
            RecetteCheck.UseVisualStyleBackColor = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            BackgroundImageLayout = ImageLayout.Center;
            ClientSize = new Size(512, 436);
            Controls.Add(RecetteCheck);
            Controls.Add(IOCheck);
            Controls.Add(labelAction);
            Controls.Add(labelPlc);
            Controls.Add(progressBar1);
            Controls.Add(actionBox);
            Controls.Add(BtnExecute);
            Controls.Add(MsgHelpCheck);
            Controls.Add(UnitCheck);
            Controls.Add(MsgCheck);
            Controls.Add(VarCheck);
            Controls.Add(folderSelectBtn);
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            HelpButton = true;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            SizeGripStyle = SizeGripStyle.Hide;
            Text = "DefiPilot";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private FolderBrowserDialog folderSelect;
        private Button folderSelectBtn;
        private CheckBox VarCheck;
        private CheckBox MsgCheck;
        private CheckBox UnitCheck;
        private CheckBox MsgHelpCheck;
        private Button BtnExecute;
        private ComboBox actionBox;
        private SaveFileDialog saveFileDialog1;
        private ProgressBar progressBar1;
        private System.Windows.Forms.Label labelPlc;
        private System.Windows.Forms.Label labelAction;
        private CheckBox IOCheck;
        private CheckBox RecetteCheck;
    }
}
