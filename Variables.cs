using OfficeOpenXml;
using System;

namespace DefiPilot
{

    /// <summary>
    /// Summary description for Variable
    /// </summary>
    public class Variable
    {
        //index des colonnes dans le fichier Xcel
        public enum Column
        {
            used = 0,
            id,
            zone,
            zoneRef,
            module,
            moduleRef,
            text,
            textRef,
            adress,
            recipe,
            actionType = 12,
            magic,
            accessLvl = 14,
            enable,
            numOfDecimals = 16,
            units,
            operand,
            valueMin,
            valueMax,
            commentHelp,
            File
        }

        //nom des colonnes dans le fichier Xcel
        public static readonly string[] Colonnes =
        {
            "Utilisée",
            "Iden (Efidrive)",
            "Désignation_Zone",
            "Désignation_Zone_Ref",
            "Désignation_Module",
            "Désignation_Module_Ref",
            "Désignation_Texte",
            "Désignation_Texte_Ref",
            "Adresse",
            "Recette",
            "Bit",
            "Type",
            "Type_Action",
            "Numéro_fonction_magique",
            "Niveau_Accès",
            "Variariable_Autorisation",
            "Nb_Décimales",
            "Unité",
            "Opération",
            "Borne_Min",
            "Borne_Max",
            "Borne_Seuil",
            "Aide_Commentaire",
            "Fichier"
        };

        public static int[] ColumnsIndex = new int[Colonnes.Length];

        //les champs de la variable
        public string   Id,
                        Zone,
                        ZoneRef,
                        Module,
                        ModuleRef,
                        Text,
                        TexteRef,
                        Adress,
                        Recipe,
                        Type,
                        TypeAction,
                        Magic,
                        Enable,
                        Comment,
                        AccessLvl,
                        Units,
                        NumOfDecimals,
                        Operand,
                        ValueMin,
                        ValueMax;

        public int id;//id en numérique

        public Variable()
        {//nouvelle instance de classe

            Id = Zone = ZoneRef = Module =
            ModuleRef = Text = TexteRef =
            Adress = Recipe = Type = TypeAction =
            Comment = AccessLvl = Magic = Enable =
            NumOfDecimals = Operand = Units =
            ValueMin = ValueMax = string.Empty;

            id = 0;
        }

        public bool FromDossierPLC(string Line)
        {//parse une ligne de fichier de variables PLC retourne faux sur erreur

            char[] separators = { '=', ':', '-' };

            string[] Result = { "", "", "", "", "" };
            string Test = Line;

            for (int i = 0; i < separators.Length; i++)
            {
                if (Test.IndexOf(separators[i]) > 0)
                {
                    string[] Words = Test.Split(separators[i]);

                    if (Words.Length > 1)
                    {
                        Result[i] = Words[0];
                        Result[i + 1] = Test = Words[1];
                    }
                    else
                        Result[i + 1] = string.Empty;
                }
            }

            Id = Result[0].Trim();

            if(Id == "")
                return false;
            else

            id = int.Parse(Id);
            (Zone, ZoneRef) = CheckSubRef(Result[1].Trim());
            (Module, ModuleRef) = CheckSubRef(Result[2].Trim());
            (Text, TexteRef) = CheckSubRef(Result[3].Trim());

            return true;
        }

        private static (string, string) CheckSubRef(string Src)
        {//du numérique en fin de texte ? probablement une sous-reférence

            string[] Splitter = Src.Split(' ');
            string Wrd1 = Src, Wrd2 = "";

            if (Splitter.Length > 1 && Splitter[Splitter.Length - 1].Any(char.IsDigit))
            {
                Wrd2 = Splitter[Splitter.Length - 1];
                Wrd1 = Src.Remove(Src.LastIndexOf(' '));
            }

            return (Wrd1, Wrd2);
        }

        public static void XcelHeaderWrite(ExcelWorksheet WsDst)
        {  //ecrit les en-t^tes de colonnes        

            for (int i = 1; i <= Colonnes.Length; i++)
                WsDst.Cells[2, i].Value = Colonnes[i-1];
        }

        public void XcelLineWrite(ExcelWorksheet WsDst, int line)
        {//ecrit une variable dans la ligne Xcel

            WsDst.Cells[line, ColumnsIndex[(int)Column.used]].Value          = "X";
            WsDst.Cells[line, ColumnsIndex[(int)Column.id]].Value            = this.id;
            WsDst.Cells[line, ColumnsIndex[(int)Column.zone]].Value          = this.Zone;
            WsDst.Cells[line, ColumnsIndex[(int)Column.zoneRef]].Value       = this.ZoneRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.module]].Value        = this.Module;
            WsDst.Cells[line, ColumnsIndex[(int)Column.moduleRef]].Value     = this.ModuleRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.text]].Value          = this.Text;
            WsDst.Cells[line, ColumnsIndex[(int)Column.textRef]].Value       = this.TexteRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.adress]].Value        = this.Adress;
            WsDst.Cells[line, ColumnsIndex[(int)Column.recipe]].Value        = this.Recipe;
            WsDst.Cells[line, ColumnsIndex[(int)Column.actionType]].Value    = this.TypeAction;
            WsDst.Cells[line, ColumnsIndex[(int)Column.magic]].Value         = this.Magic;
            WsDst.Cells[line, ColumnsIndex[(int)Column.enable]].Value        = this.Enable;
            WsDst.Cells[line, ColumnsIndex[(int)Column.accessLvl]].Value     = this.AccessLvl;
            WsDst.Cells[line, ColumnsIndex[(int)Column.numOfDecimals]].Value = this.NumOfDecimals;
            WsDst.Cells[line, ColumnsIndex[(int)Column.units]].Value         = this.Units;
            WsDst.Cells[line, ColumnsIndex[(int)Column.operand]].Value       = this.Operand;
            WsDst.Cells[line, ColumnsIndex[(int)Column.valueMin]].Value      = this.ValueMin;
            WsDst.Cells[line, ColumnsIndex[(int)Column.valueMax]].Value      = this.ValueMax;
        }

        public static void XcelLineColIndexGet(ExcelWorksheet WsSrc)
        {
            for (int i = 1; i <= Colonnes.Length; i++)
                for (int j = 1; j <= 50; j++)
                {
                    if (WsSrc.Cells[2, j].Value?.ToString() == Colonnes[i - 1])
                    {
                        ColumnsIndex[i - 1] = j;
                        break;
                    }

                    else
                        ColumnsIndex[i - 1] = 50;
                }
        }
    }
}
