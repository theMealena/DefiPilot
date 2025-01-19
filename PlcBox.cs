using System;
using System.Windows.Forms;

namespace DefiPilot
{
    /// <summary>
    /// Dialogues pour un seul PLC / dossier.
    /// </summary>    
    public class PlcBox : UserControl //héritage de UserControl
    {
        public string DstFileName { get; set; }
        public bool isSelected;
        public FileInfo? XcelFile;

        private readonly TextBox FileNameBox     = new TextBox();
        private readonly Button FileSelectBtn    = new Button();
        private readonly CheckBox FileCheck      = new CheckBox();


        public PlcBox()
        {
            DstFileName = string.Empty;
            XcelFile    = null;
            isSelected  = false;

            InitializeComponents();
        }

        private void InitializeComponents()
        {//créé les boutons et les boîtes de texte pour un seul PLC
         //fait sur le modèle de Form1.Designer.cs

            Size = new Size(300, 50);//taille de la boîte locale

            //Nouveau groupe de boutons
            //FileCheck       = new CheckBox();   //traiter ou non
            //FileNameBox     = new TextBox();    //nom du ficher (sans chemin)
            //FileSelectBtn   = new Button();     //sélection du fichier (si complétion)
            BackColor = Color.Transparent;
            SuspendLayout();

            //Les positions sont relatives à ce groupe, donc à (0,0)
            FileCheck.Location  = new Point(0, 0);
            FileCheck.Name      = "FileCheck";
            FileCheck.Size      = new Size(450, 25);
            FileCheck.Checked   = isSelected;
            FileCheck.CheckedChanged += CheckBoxClick;            

            FileNameBox.Location    = new Point(0, FileCheck.Location.Y + FileCheck.Size.Height);
            FileNameBox.Name        = "FileNameBox";
            FileNameBox.Size        = new Size(250, 25);
            FileNameBox.TabIndex    = 1;
            FileNameBox.Enabled     = false;

            FileSelectBtn.Location = new Point(FileNameBox.Location.X + FileNameBox.Size.Width, FileNameBox.Location.Y);
            FileSelectBtn.Name = "FileSelectBtn";
            FileSelectBtn.Size = new Size(30, FileNameBox.Size.Height);
            FileSelectBtn.TabIndex = 2;
            FileSelectBtn.Text = "...";
            FileSelectBtn.Visible = false;
            FileSelectBtn.Click += DstFileSelect;
            
            //ça ajoute à cette PLCBox, il faut ensuite ajouter la PLCBox au control de Form1
            Controls.Add(FileCheck);
            Controls.Add(FileNameBox);
            Controls.Add(FileSelectBtn);

            ResumeLayout(false);
            PerformLayout();
        }

        private void CheckBoxClick(object? sender, EventArgs e)
        {//validation pour ce plc

            isSelected = FileCheck.Checked;
        }

        public void CheckBoxTextSet(string text)
        {//acces public au texte de la checkbox
            FileCheck.Text = text;
        }

        public void ButonVisibleSet(bool val)
        {//accès public à la visibilité du bouton
            FileSelectBtn.Visible = val;
        }

        public void FileNameBoxTextSet(string text)
        {//accès public au texte de la boîte de texte
            FileNameBox.Text = text;
        }

        private void DstFileSelect(object? sender, EventArgs e)
        {//sélection du fichier de destination

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter   = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title    = "Select Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                XcelFile = new FileInfo(openFileDialog.FileName);
                DstFileName = XcelFile.Name;
                FileNameBox.Text = XcelFile.Name;
            }

            openFileDialog.Dispose();
        }
        

        public string GetFileName()
        {
            if (XcelFile != null)
                return XcelFile.FullName;
            else
                return string.Empty;
        }        
    }
}