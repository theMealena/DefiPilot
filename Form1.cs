using OfficeOpenXml;
using System.CodeDom;
using System.Globalization;
using System.Text;
using System.Xml.Linq;


namespace DefiPilot
{
    public partial class Form1 : Form
    {
        public static readonly string AppName = "DefiPilot (Beta)";
        public static string WorkFolder = string.Empty;        
        List<Messages> MsgList = new List<Messages>();
        Folder FolderSts = new Folder();
        List<PlcBox> DstPlc = new List<PlcBox>();


        public Form1()
        {
            InitializeComponent();
            actionBox.SelectedIndex = 0;
            this.Text = AppName;
        }

        private void FolderSelectBtn_Click(object sender, EventArgs e)
        {//sélection d'un dossier de travail

            if (folderSelect.ShowDialog() == DialogResult.OK)
            {//ok
                WorkFolder = folderSelect.SelectedPath; //Le dossier de la machine dans lequel on va travailler
                //on nettoie les anciennes données
                FolderSts.Clear();

                //on charge les nouvelles données
                FolderSts.FolderPath = WorkFolder;

                if (FolderSts.ParamFetch())
                {
                    this.Text = $"{AppName} : {FolderSts.Name} : {FolderSts.Customer} {FolderSts.City} {FolderSts.Country}";
                    //folderSelectBtn.BackColor = Color.Green;
                }
                else
                    goto err;
            }
            else
                goto err;

            PlcBoxesClear();    //on efface les anciennes PLC box
            PlcBoxesCreate();   //on en créé de nouvelles            

            return;

         err:
            //erreur, on reset tout
            WorkFolder = string.Empty;
            this.Text = AppName;
            //on nettoie les anciennes données
            FolderSts.Clear();     

        }

        private void CreateFiles(object sender, EventArgs e)
        {                        

            for (int i = 1; i <= FolderSts.numOfPlc; i++)
            {//pour chaque PLC

                VerboseDisplay(true);

                if (DstPlc[i - 1].isSelected)
                {//s'il est sélectionné

                    labelPlc.Text = $"PLC " + i.ToString("D2"); labelPlc.Update();
                    labelAction.Text = "Lecture variables"; labelAction.Update(); Update();

                    FolderSts.FetchVariables(i);//lecture de la liste des variables (DossierPLC) et premier parse

                    progressBar1.Value += 10; progressBar1.Update();
                    labelAction.Text = "Classement variables"; labelAction.Update(); Update();

                    FolderSts.SortVarByID();//on range les variables par id croissant (pour recherche et ordre Xcel)

                    progressBar1.Value += 10; progressBar1.Update();
                    labelAction.Text = "Lecture Adresses"; labelAction.Update(); Update();

                    FolderSts.FillAdress(i);//on remplit les adresses (OPC.xxxx.0X)

                    progressBar1.Value += 10; progressBar1.Update();

                    if (UnitCheck.Checked)  //si on veut récupérer tous les détails
                    {
                        labelAction.Text = "Lecture des fichiers ani"; labelAction.Update(); Update();
                        FolderSts.FillOptions(i);//on va scanner tous les fichiers.ani
                        progressBar1.Value += 10; progressBar1.Update();
                    }

                    if (RecetteCheck.Checked)
                    {
                        labelAction.Text = "Lecture des fichiers recettes"; labelAction.Update(); Update();
                        Recipes.GetNames(WorkFolder);
                        Recipes.ScanForReceipes(FolderSts, i);
                        progressBar1.Value += 10; progressBar1.Update();
                    }

                    if (MsgCheck.Checked)
                    {//si on doit relire les messages

                        labelAction.Text = "Lecture des messages"; labelAction.Update(); Update();
                        MsgList.Clear();//on vide la liste au cas où.
                        ReadMsg(i);     //on va lire tous les messages (BamDefaut.fr), les autres langues sont  ignorées
                        Messages.SortById(MsgList);//on range les messages par id croissant.
                        progressBar1.Value += 10; progressBar1.Update();

                        if (MsgHelpCheck.Checked)//S'il faut compléter avec les fichiers d'aide
                        {
                            labelAction.Text = "Lecture de l'aide aux messages"; labelAction.Update(); Update();
                            ReadMsgHelp();//on va les lire
                            progressBar1.Value += 10; progressBar1.Update();
                        }
                    }

                    if (IOCheck.Checked)
                    {
                        labelAction.Text = "Lecture des entrées / sorties"; labelAction.Update(); Update();
                        IOCard.ExtractOPC(FolderSts.FolderPath, i);
                        IOCard.ExtractDefinition(FolderSts.FolderPath, i);
                        progressBar1.Value += 10; progressBar1.Update();
                    }

                    labelAction.Text = "Ecriture du fichier"; labelAction.Update();
                    if (actionBox.SelectedIndex == 0) //selon sélection,
                        CreateXcel(i);//on créé
                    else
                        UpdateXcel(i);//on met à jour

                    progressBar1.Value = 100; progressBar1.Update();
                }

                FolderSts.Variables.Clear();//on nettoie la liste des variables
                MsgList.Clear();            //on nettoie la liste des messages
            }

            MessageBox.Show("Opération terminée", "DefiPilot");
            VerboseDisplay(false);
            
        }

        private void VerboseDisplay(bool show)
        {    
            progressBar1.Visible =
            labelPlc.Visible     =
            labelAction.Visible  = show;
            
            progressBar1.Value = 0;            
        }

        private void PlcBoxesCreate()
        {
            for (int i = 0; i < FolderSts.numOfPlc; i++)
            {//ajout d'autant de PLC box que de PLC                

                PlcBox newPlcFile = new PlcBox();
                newPlcFile.Location = new Point(220, 10 + (i * 60));
                newPlcFile.FileNameBoxTextSet(FolderSts.PlcNames[i] + ".xlsx");
                newPlcFile.CheckBoxTextSet("PLC " + (i + 1).ToString("D2") + " : " + FolderSts.PlcNames[i]);
                DstPlc.Add(newPlcFile);
                Controls.Add(newPlcFile);
            }
        }

        private void PlcBoxesClear()
        {//supprime toutes les PLC box

            foreach (var plc in DstPlc)
            {
                Controls.Remove(plc);
                plc.Dispose();
            }

            DstPlc.Clear();
        }

        private int CreateXcel(int plc)
        {
            string PathAndFileName = dstFileNameBuild(plc - 1); //chemin complet du fichier
            var file = new FileInfo(PathAndFileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(file))
            { // Ajouter une nouvelle feuille de calcul

                if (package is null)
                {
                    MessageBox.Show("Erreur lors de l'ouverture du fichier Excel", "Erreur fichier");
                    return 0;
                }

                string[] IOnglets = { "ENTREE", "SORTIE" };

                for (int i = 0; i < IOnglets.Length; i++)
                {
                    ExcelWorksheet? WsActive = null;//onglet actif
                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, IOnglets[i]);


                    if (WsIndex < 0)
                    {
                        WsActive = package.Workbook.Worksheets.Add(IOnglets[i]);
                    }
                    else
                    {
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }

                    IOCard.XcelHeaderWrite(WsActive);

                    if (IOCheck.Checked)
                    {
                        IOCard.XcelLineColIndexGet(WsActive);
                        IOCard.XcelCreate(WsActive, i);
                    }
                }

                for (int i = 2; i < 7; i++)
                {//pour chaque type de variable

                    string WsName = Uties.IdTypeToName(i); //nom de l'onglet
                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, WsName); //existe déjà ?
                    ExcelWorksheet? WsActive = null;//onglet actif

                    if (WsIndex >= 0)
                    {//existe, on va l'utiliser
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }
                    else
                    {//sinon on créé
                        WsActive = package.Workbook.Worksheets.Add(WsName);
                        // Écriture des noms de colonnes
                        Variable.XcelHeaderWrite(WsActive);
                    }

                    Variable.XcelLineColIndexGet(WsActive);//on récupère les index des colonnes

                    // Écrire les données dans les cellules
                    int firstIndex = FirstIndexOfId(FolderSts.Variables, i); //premier index dans les variables
                    int lastIndex = LastIndexOfId(FolderSts.Variables, i); //dernier index dans les variables
                    int dstLine = 3;

                    for (int srcIndex = firstIndex; srcIndex <= lastIndex; srcIndex++)
                    {//on écrit toutes les variables de ce type
                        FolderSts.Variables[srcIndex].XcelLineWrite(WsActive, dstLine++);
                    }
                }


                if (MsgCheck.Checked)
                {//si on doit créer les messages

                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, "DEFAUT");
                    ExcelWorksheet? WsActive = null;

                    if (WsIndex >= 0)
                    {//pareil, si existe on va chercher
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }
                    else
                    {//sinon on créé
                        WsActive = package.Workbook.Worksheets.Add("DEFAUT");
                        Messages.XcelHeaderWrite(WsActive);
                    }

                    Messages.XcelLineColIndexGet(WsActive);//on récupère les index des colonnes

                    int line = 3;//début décriture ligne 3
                    foreach (var Msg in MsgList)
                    {//on dump tout sans trop réfléchir (attention : tous les PLCS) /*TODO : filtrer par PLC*/
                        Msg.XcelLineWrite(WsActive, line++);
                    }
                }

                // Enregistrer le fichier
                try
                {
                    package.Save();
                }
                catch (Exception e)
                {
                    MessageBox.Show($"Erreur lors de l'enregistrement de {file.Name} : {e.Message}", "Erreur fichier");
                    return 0;
                }
            }

            return 1;
        }

        private int UpdateXcel(int plc)
        {
            string PathAndFileName = dstFileNameBuild(plc - 1); //chemin complet du fichier
            var file = new FileInfo(PathAndFileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(file))
            { // ouvrir la feuille de calcul

                if (package is null)
                {//erreur à l'ouverture
                    MessageBox.Show($"Erreur lors de l'ouverture de {PathAndFileName}", "Erreur fichier");
                    return 0;
                }

                string[] IOnglets = { "ENTREE", "SORTIE" };

                for (int i = 0; i < IOnglets.Length; i++)
                {
                    ExcelWorksheet? WsActive = null;//onglet actif
                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, IOnglets[i]);


                    if (WsIndex < 0)
                    {
                        WsActive = package.Workbook.Worksheets.Add(IOnglets[i]);
                    }
                    else
                    {
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }

                    IOCard.XcelHeaderWrite(WsActive);

                    if (IOCheck.Checked)
                    {
                        IOCard.XcelLineColIndexGet(WsActive);
                        IOCard.XcelCreate(WsActive, i);
                    }
                }

                for (int i = 2; i < 7; i++)
                {//pour chaque type de variable

                    string WsName = Uties.IdTypeToName(i); //nom de l'onglet
                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, WsName); //existe déjà ?
                    ExcelWorksheet? WsActive = null;//onglet actif

                    if (WsIndex >= 0)
                    {//existe, on va l'utiliser
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }
                    else
                    {//on le créé
                        WsActive = package.Workbook.Worksheets.Add(WsName);
                        // Écriture des noms de colonnes
                        Variable.XcelHeaderWrite(WsActive);
                    }

                    Variable.XcelLineColIndexGet(WsActive);//on récupère les index des colonnes

                    // Écrire les données dans les cellules
                    int srcIndex = FirstIndexOfId(FolderSts.Variables, i); //prmier index dans les variables
                    int srcLastIndex = LastIndexOfId(FolderSts.Variables, i);  //dernier index dans les variables
                    int dstLine = 3; //début d'écriture dans le dossier

                    while ((srcIndex <= srcLastIndex) && (dstLine <= WsActive.Dimension.End.Row))
                    {
                        var item = FolderSts.Variables[srcIndex];//la variable de travail

                        if (item.Id == "") /*TODO : il ne devrait pas y avoir de variable sans id, à corriger à la lecture*/
                        {//id inattendu, on passe à la suivante
                            srcIndex++;
                            continue;
                        }

                        int srcId = int.Parse(item.Id);
                        int dstLineId = XcelLineId(WsActive, dstLine);

                        if (dstLineId < 0)
                        {//si on a pas trouvé d'id (ligne exotique ou mal formatée), on va à la ligne suivante
                            dstLine++;
                            continue;
                        }

                        if (dstLineId == srcId)
                        {//même id
                            item.XcelLineWrite(WsActive, dstLine++);//on met à jour avec ce qu'on trouvé dans le dossier
                            srcIndex++;//on passe à la recherche du message suivant
                        }
                        else if (dstLineId < srcId)
                        {//si l'id du fichier est plus petit, 
                            WsActive.Cells[dstLine, Variable.ColumnsIndex[(int)Variable.Column.used]].Value = "";//on décoche 
                            dstLine++; //on va à la ligne suivante
                        }
                        else
                        {//si l'id du fichier est plus grand                                    
                            WsActive.InsertRow(dstLine, 1); //on insere une ligne
                            item.XcelLineWrite(WsActive, dstLine++); //on écrit la ligne et passe à la suivante
                            srcIndex++;//variable suivante
                        }
                    }
                }

                if (MsgCheck.Checked)
                {//lecture des messages demandée

                    int WsIndex = WorkSheetExists(package.Workbook.Worksheets, "DEFAUT"); //existe déjà ?
                    ExcelWorksheet? WsActive = null;//onglet actif

                    if (WsIndex >= 0)
                    {//si existe, on va chercher
                        WsActive = package.Workbook.Worksheets[WsIndex];
                    }
                    else
                    {//sinon on créé
                        WsActive = package.Workbook.Worksheets.Add("DEFAUT");
                        Messages.XcelHeaderWrite(WsActive);
                    }

                    Messages.XcelLineColIndexGet(WsActive);//on récupère les index des colonnes

                    int dstLine = 3;
                    int srcIndex = 0;

                    while ((srcIndex < MsgList.Count) && (dstLine <= WsActive.Dimension.End.Row))
                    {
                        var Msg = MsgList[srcIndex];

                        uint thisId = Msg.id;

                        int dstLineId = XcelLineId(WsActive, dstLine);

                        if (dstLineId < 0)
                        {//si on a pas trouvé d'id, on avance à la ligne suivante
                            dstLine++;
                            continue;
                        }

                        if (dstLineId == thisId)
                        {//même id, on met à jour avec ce qu'on trouvé dans le dossier
                            Msg.XcelLineWrite(WsActive, dstLine++);
                            srcIndex++;
                        }
                        else if (dstLineId < thisId)
                        {//si l'id du fichier est plus petit, on décoche puis on va à la ligne suivante
                            WsActive.Cells[dstLine, Messages.ColumnsIndex[(int)Messages.Column.used]].Value = "";
                            dstLine++;
                        }
                        else
                        {//si l'id du fichier est plus grand                                    
                            WsActive.InsertRow(dstLine, 1); //on insere une ligne
                            Msg.XcelLineWrite(WsActive, dstLine++); //on écrit la ligne
                            srcIndex++;
                        }
                    }
                }                

                // Enregistrer le fichier
                package.Save();
            }

            return 1;
        }

        private int XcelLineId(ExcelWorksheet WsSrc, int line)
        {
            var idCell = WsSrc.Cells[line, Variable.ColumnsIndex[(int)Variable.Column.id]];

            if (idCell == null || idCell.Value == null)
            {
                return -1;
            }

            int id = GetNumericValue(idCell.Value);

            return id;
        }

        private static int GetNumericValue(object source)
        {//source doit être un numérique ou un string qui peut être parsé en int

            if (source is double dvalue)
            {
                return (int)dvalue;
            }
            else if (source is int intValue)
            {
                return intValue;
            }
            else if (source is string strValue && int.TryParse(strValue, out int result))
            {
                return result;
            }

            return -1;
        }

        private int WorkSheetExists(ExcelWorksheets WsSrc, string name)
        {
            for (int i = 0; i < WsSrc.Count; i++)
                if (WsSrc[i].Name == name)
                    return i;

            return -1;
        }

        private void MsgCheck_Click(object sender, EventArgs e)
        {
            //MsgHelpCheck.Visible = MsgCheck.Checked;

            if (!(MsgHelpCheck.Visible = MsgCheck.Checked))
                MsgHelpCheck.Checked = false;
        }

        private void ReadMsgHelp()
        {
            string FolderPath = $"{WorkFolder}\\DEFAUT\\EXPLI";
            string[] Files = Directory.GetFiles(FolderPath, $"*.rtf");

            foreach (string File in Files)
            {
                if (File.Contains("_FR"))
                {
                    int lid = Messages.ExtractIDFromRtfFileName(File);
                    Messages? newMsg = Messages.FetchByID(MsgList, lid);
                    newMsg?.ReadMsgHelp(File);

                    if (newMsg != null)
                    {
                        //newMsg.id = lid;
                        //MsgList.Add(newMsg);
                    }
                }
            }
        }

        private void ReadMsg(int plc)
        {
            string FileName = $"{WorkFolder}\\SUPER\\TEXTE\\BamDefaut.FR";
            plc--;

            StreamReader SrcFile = new StreamReader(FileName);

            SrcFile.ReadLine();
            // On cherche les espaces qui séparent chaque PLC
            while (plc > 0)
            {
                if (SrcFile.ReadLine()?.Trim() == "")
                    plc--;
            }

            while (!SrcFile.EndOfStream)
            {
                string Line = SrcFile.ReadLine() ?? "";

                if ((Line != "") && char.IsNumber(Line[0]))
                {
                    Messages Msg = new Messages();
                    Msg.Parse(Line);
                    MsgList.Add(Msg);
                }
                else
                    break;

            }

            SrcFile.Close();
        }

        private void actionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool sts = actionBox.SelectedIndex == 1;

            if (DstPlc == null)
                return;

            for (int i = 0; i < FolderSts.numOfPlc; i++)
            {
                DstPlc[i].ButonVisibleSet(sts);

                if (sts && DstPlc[i].XcelFile != null)
                    DstPlc[i].FileNameBoxTextSet(DstPlc[i]?.XcelFile?.Name ?? "");
                else
                    DstPlc[i].FileNameBoxTextSet(FolderSts.PlcNames[i] + ".xlsx");
            }
        }

        private string dstFileNameBuild(int index)
        {
            bool retry = false;

            if (actionBox.SelectedIndex == 1)
            {
                string Rslt = DstPlc[index].GetFileName();

                if (Rslt != string.Empty)
                    return Rslt;
                else
                {
                    MessageBox.Show($"PLC 0{index + 1} : Nom du fichier indéfini, création du nom par défaut");
                    retry = true;
                }
            }

            if ((actionBox.SelectedIndex == 0) || retry)
            {
                DateTime now = DateTime.Now;
                string formattedDate = now.ToString("yyyy_MMdd", CultureInfo.InvariantCulture);
                return WorkFolder
                        + "\\DOSSIER\\(auto) PLC"
                        + (index + 1).ToString("D2")
                        + $"_{FolderSts.Name}_{FolderSts.PlcNames[index] + '_' + formattedDate}.xlsx";
            }

            return string.Empty; //ne devrait jamais arriver, fait taire le compilateur
        }

        public int FirstIndexOfId(List<Variable> Variables, int seek)
        {//premier occurrence de la famille d'id

            if (seek < 2 || seek >= 7)
                return 0;

            for (int i = 0; i < Variables.Count; i++)
            {
                if (Variables[i].id / 10000 >= seek)
                    return i;
            }

            return 0;
        }

        public int LastIndexOfId(List<Variable> Variables, int seek)
        {//dernier occurence de la famille d'id

            if (seek < 2 || seek >= 7)
                return Variables.Count - 1;

            for (int i = Variables.Count - 1; i >= 0; i--)
            {
                if (Variables[i].id / 10000 <= seek)
                    return i;
            }

            return Variables.Count - 1;
        }


        private void VarCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (!VarCheck.Checked)
                UnitCheck.Checked = false;

            UnitCheck.Visible = VarCheck.Checked;
        }
    }
}
