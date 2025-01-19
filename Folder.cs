using System;
using System.Text;

namespace DefiPilot
{

    /// <summary>
    /// Description du dossier
    /// </summary>
    public class Folder
    {        
        public string   Name,       //nom (machine)
                        Customer, 
                        Country, 
                        City, 
                        FolderPath;

        public List<string> PlcNames;
        public int numOfPlc;
        public List<Variable> Variables;

        public Folder()
        {// nouvelle instance de classe

            Name =
            Customer =
            Country =
            City = 
            FolderPath = string.Empty;

            numOfPlc = 0;

            PlcNames = new List<string>();
            Variables = new List<Variable>();
        }


        public void Clear()
        {
            Name =
            Customer =
            Country =
            City =
            FolderPath = string.Empty;

            numOfPlc = 0;

            PlcNames.Clear();
            Variables.Clear();
        }

        public bool ParamFetch()
        {//lecture des paramètres du dossier

            List<string>? Parameters = new List<string>();//liste de string pour faciliter la navigation

            string FileName = FolderPath + @"\SUPER\GENERAL\Parametre.ini"; //nom du fichier de paramètres

            StreamReader Param; //fichier en stream puis ouverture
            try
            {//ok
                Param = new StreamReader(FileName);
            }
            catch (Exception)
            {//echec
                MessageBox.Show($"Impossible de trouver le fichier de paramètre.\nSélectionner le dossier machine eg : \"1234MN\"" , "Erreur fichier");
                return false;
            }
            
            Uties.FileToList(Param, Parameters);
            Param.Close();

            //la chasse aux infos :
            Name        = Uties.ListParamFetch("NumMachine", Parameters);
            Customer    = Uties.ListParamFetch("NomClient", Parameters);
            Country     = Uties.ListParamFetch("Pays", Parameters);
            City        = Uties.ListParamFetch("Ville", Parameters);
            numOfPlc    = int.Parse(Uties.ListParamFetch("NbPlc", Parameters));

            for(int i = 0; i < numOfPlc; i++)
            {//à chaque PLC son nom
                int index       = Uties.ListParamIndexOf($"[PLC{i + 1}]", Parameters);
                string PlcName  = Uties.ListParamFetch("Caption", Parameters, index);

                if(PlcName != null)
                    PlcNames.Add(PlcName);
            }

            //MessageBox.Show(@$"{this.Name};{this.Customer};{this.Country};{this.City}; " + this.numOfPlc, "debug");            
            Parameters.Clear();
            Parameters = null; //inutile, mais je ne veux pas perdre l'habitude de libérer la mémoire
            return true;
        }

        public int FetchVariables(int plc)
        {//récupère tous les Id des variables du dossier PLC + commentaires

            Variables.Clear();//on vide s'il y avait quelque chose

            string FileName = FolderPath + @$"\SUPER\TEXTE\DossierPlc.{plc.ToString("D2")}.FR"; //nom du fichier

            StreamReader Param;
            try
            {//ok
                Param = new StreamReader(FileName );
            }
            catch (Exception e)
            {//echec
                MessageBox.Show(@$"{e.Message}", "Erreur fichier");
                return 0;
            }

            while (!Param.EndOfStream)
            {//lecture ligne par ligne

                string? Line = Param.ReadLine();//lecture ligne

                if (Line != null)
                {
                    Variable thisVar = new Variable();
                    if (thisVar.FromDossierPLC(Line)) //on va chercher les infos disponibles dans cette ligne
                        Variables.Add(thisVar);//on ajoute si ça a fonctionné
                }
            }

            Param.Close();//fermeture fichier

            return Variables.Count;//nombre de vraibles lues
        }

        public void SortVarByID()
        {//tri à bulle des variables par id

            for (int i = 0; i < Variables.Count; i++)
            {
                for (int j = 0; j < Variables.Count - i - 1; j++)
                {
                    if (string.Compare(Variables[j].Id, Variables[j + 1].Id, StringComparison.OrdinalIgnoreCase) > 0)
                        (Variables[j], Variables[j + 1]) = (Variables[j + 1], Variables[j]);
                }
            }
        }

        public Variable? FetchVarByID(string Seek)
        {//recherche rapide de variable par son id (dichotomie) :/*TODO il y a une fonction native il paraît.*/

            int beg = 0;
            int end = Variables.Count - 1;

            while (beg <= end)
            {
                int mid = beg + (end - beg) / 2;

                int comp = string.Compare(Seek, Variables[mid].Id, StringComparison.OrdinalIgnoreCase);

                if (comp == 0)
                {//trouvé
                    return Variables[mid];
                }
                else if (comp > 0)
                {//dans la seconde moitié
                    beg = mid + 1;
                }
                else
                {//dans la première moitié
                    end = mid - 1;
                }
            }
            return null; //pas trouvé
        }

        public void FillAdress(int plc)
        {//remplissage des adresses des variables

            for (int i = 2; i < 7; i ++)
            {
                string FileName = $"{FolderPath}\\SUPER\\PLC\\OPC.{Uties.IdTypeToName(i).ToLower()}.{plc.ToString("D2")}";

                string Id = "",
                        Adress = "";

                StreamReader SrcFile = new StreamReader(FileName);

                while (!SrcFile.EndOfStream)
                {
                    (Id, Adress) = Uties.OpcExtract(SrcFile.ReadLine() ?? "");

                    if (Id != "")
                    {
                        var Line = FetchVarByID(Id);

                        if (Line != null)
                            Line.Adress = Adress;
                    }
                }

                SrcFile.Close();
            }

            //MessageBox.Show($"Address de 52421 : {FetchVarByID("52421").Adress}");
        }

        public void FillOptions(int plc)
        {
            string Folder = $"{FolderPath}\\SUPER\\ANI"; //dossiers des ani
            string[] Files = Directory.GetFiles(Folder, $"*.ani"); //liste des fichiers.ani dans le dossier
            List<string> AniDivList = new List<string>();//liste des anidivers appelés dans les fichiers ani

            foreach (string File in Files) //pour chaque fichier
            {
                StreamReader SrcFile = new StreamReader(File, Encoding.Latin1); //ouvrir le fichier

                while (!SrcFile.EndOfStream)
                {//on va scanner chaque ligne contenue dans le fichier

                    string Line = SrcFile.ReadLine() ?? "";
                    
                    string? AniExt = AniLineParse(Line, plc);//parse de la ligne et affectation à variable concernée si besoin

                    if (AniExt != null && !AniDivList.Contains(AniExt))//si appelle anidivers et pas déjà enregistré
                        AniDivList.Add(AniExt);//on ajoute le nom à la liste des anidivers à scanner
                }

                SrcFile.Close();//ferme le fichier
            }

            //maintenant les anidivers qui ont été appelés
            Folder = $"{FolderPath}\\SUPER\\ANIDIVERS"; //dossiers des anidivers

            for (int i=0; i<AniDivList.Count; i++)
            {
                string completeFileName = Folder + "\\" + AniDivList[i] + ".ani";

                StreamReader SrcFile;

                try
                {
                    SrcFile = new StreamReader(completeFileName, Encoding.Latin1); //ouvrir le fichier
                }
                catch (Exception) {continue;}

                while (!SrcFile.EndOfStream)
                {//on va scanner chaque ligne contenue dans le fichier

                    string Line = SrcFile.ReadLine() ?? "";

                    string? AniExt = AniLineParse(Line, plc);//parse de la ligne et affectation à la variable concernée si besoin

                    if (AniExt != null && !AniDivList.Contains(AniExt))//si appelle anidivers et n'est pas déjà enregistré
                        AniDivList.Add(AniExt);//on ajoute le nom à la liste des anidivers à scanner

                }
                SrcFile.Close();//ferme le fichier
            }

            AniDivList.Clear();
        }

        private string? AniLineParse(string Line, int plc)
        {
            //on découpe avec le séparateur
            string[] Words = Line.Split('|', StringSplitOptions.TrimEntries);

            bool isFormatedLine = Words.Length == 44;//la ligne est correctement formatée

            if (!isFormatedLine)//sinon on passe à la suivante
                return null;

            bool isUsed = Words[0].ToLower() == "a";     //la varable est utilisée
            bool isPlc = int.Parse(Words[9]) == plc;    //c'est le bon PLC

            if (isUsed && isPlc)
            {
                string Id = Words[11];  //on choppe l'id

                var LineVar = FetchVarByID(Id); //on choppe la varible

                if (LineVar != null)
                {//on attribue les valeurs à chaque membre
                 //en s'assurant de ne grader que la valeur de la première occurence

                    LineVar.TypeAction = (LineVar.TypeAction == "") ? Words[13].Trim() : LineVar.TypeAction;
                    LineVar.NumOfDecimals = (LineVar.NumOfDecimals == "") ? Words[19].Trim() : LineVar.NumOfDecimals;
                    LineVar.AccessLvl = (LineVar.AccessLvl == "") ? Words[25].Trim() : LineVar.AccessLvl;
                    LineVar.Operand = (LineVar.Operand == "") ? Words[28].Trim() : LineVar.Operand;
                    LineVar.ValueMin = (LineVar.ValueMin == "") ? Words[29].Trim() : LineVar.ValueMin;
                    LineVar.ValueMax = (LineVar.ValueMax == "") ? Words[30].Trim() : LineVar.ValueMax;
                    LineVar.Enable = (LineVar.Enable == "") ? Words[36].Trim() : LineVar.Enable;
                    LineVar.Units = (LineVar.Units == "") ? Words[41].Trim() : LineVar.Units;
                }
            }
            else if (isUsed && Words[1].Trim() == "FICHIERANI")
            {
                return  Words[43].Trim();
            }

            return null;
        }
    }
}