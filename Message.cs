using OfficeOpenXml;
using System;
using System.Reflection;
using System.Security.Policy;


namespace DefiPilot
{
    public class Messages
    {
        //index des colonnes dans le fichier Xcel
        public enum Column
        {
            used = 0,
            id = 1,
            type,
            responsability,
            zone,
            zoneRef,
            module,
            moduleRef,
            module2,
            module2Ref,
            msg,
            msgRef,
            msgHelp,
            msgAction
        }

        //nom des colonnes dans le fichier Xcel
        private static readonly string[] colonnes =
        {
            "Utilisé",
            "Iden (Efidrive)",
            "Type",
            "Responsabilité",
            "Zone",
            "Zone_Ref",
            "Module",
            "Module_Ref",
            "Module2",
            "Module2_Ref",
            "Désignation",
            "Désignation_Ref",
            "Aide_Commentaire",
            "Aide_Action",
            "Aide_Action2",
            "Adresse",            
            "Aide_Capteur1_Module",
            "Aide_Capteur1_Bit",
            "Aide_Capteur2_Module",
            "Aide_Capteur2_Bit",
            "Aide_Capteur3_Module",
            "Aide_Capteur3_Bit",
            "Aide_Capteur4_Module",
            "Aide_Capteur4_Bit",
            "Complement1_Iden",
            "Complement1_Type",
            "Complement1_FichierBam",
            "Complement2_Iden",
            "Complement2_Type",
            "Complement2_FichierBam",
            "Complement3_Iden",
            "Complement3_Type",
            "Complement3_FichierBam",
            "Complement4_Iden",
            "Complement4_Type",
            "Complement4_FichierBam",
            "Shunt_Iden",
            "PDF",
            "Illustration_1",
            "Illustration_2",
            "Illustration_3",
            "Vérificateur",
            "Méthode"
        };

        public static int[] ColumnsIndex = new int[colonnes.Length];

        public uint id;

        public string   Type,
                        Responsability,
                        Zone,
                        ZoneRef,
                        Module,
                        ModuleRef,
                        Module2,
                        Module2Ref,
                        Msg,
                        MsgRef,
                        MsgHelp,
                        MsgAction,
                        Adress;


        public Messages()
        {//nouvel instance d'objet

            id = 0;
            Type =
            Responsability =
            Zone =
            ZoneRef =
            Module =
            ModuleRef =
            Module2 =
            Module2Ref =
            Msg =
            MsgRef =
            MsgHelp =
            MsgAction =
            Adress = String.Empty;
        }

        public void ReadMsgHelp(string SrcFileName)
        {//extraction du contenu du fichier d'aide RTF
         //RTF est un format de texte enrichi (Rich Text File), on utilise un RichTextBox pour le lire
         //il y a peut-être mieux, mais c'est le premier truc que j'ai trouvé en cherchant ce format de texte dans les docs

            string[] titles = { "Commentaire", "Action(s)" };

            using (RichTextBox richTextBox = new RichTextBox())
            {
                // Charger le fichier RTF
                richTextBox.LoadFile(SrcFileName, RichTextBoxStreamType.RichText);

                // Extraire le texte
                string extractedText = richTextBox.Text;
                string[] extractedLines = extractedText.Split('\n');
                int breakT = 0;

                //on ecarte les lignes vides et on se place après le premier titre
                while (breakT < extractedLines.Length
                        && (
                            (extractedLines[breakT] == titles[0])
                            || (extractedLines[breakT] == "")
                            )
                        )
                    breakT++;

                this.MsgHelp = extractedLines[breakT++];//le message d'aide

                //maintenant, on ecarte les lignes vides et on se place après le second titre
                while (breakT < extractedLines.Length 
                        && (    
                            (extractedLines[breakT] == titles[1]) 
                            || (extractedLines[breakT] == "")
                            )
                        )
                    breakT++;

                for (int i = breakT; i < extractedLines.Length; i++)
                {//on concatènes les lignes non vides pour compléter les actions
                    this.MsgAction += (extractedLines[i] != "") ? $"{extractedLines[i]}\n" : "";
                }
            }

            return;
        }

        public void Parse(string Src)
        {//extrait les membres d'un message depuis une ligne du fichier BAM

            char[] separators = { '=', ':', '-' }; //les séparateurs

            string[] Words = Src.Split(separators, StringSplitOptions.TrimEntries /*| StringSplitOptions.RemoveEmptyEntries*/);

            //au moins 4 morceaux distincts
            if (Words.Length >= 4)
            {
                id              = uint.Parse(Words[0]);
                Type            = Words[1];
                (Zone, ZoneRef) = CheckSubRef(Words[2]);

                if (Words.Length >= 5)//si on a un morceau de plus
                {
                    (Module, ModuleRef) = CheckSubRef(Words[3]);
                    (Msg, MsgRef)       = CheckSubRef(Words[4]);
                }
                else
                {
                    (Msg, MsgRef) = CheckSubRef(Words[3]);
                }
            }
            else
                return;
        }

        private static (string, string) CheckSubRef(string Src)
        {//du numérique en fin de texte ? => probablement une sous-reférence, on extrait le numérique

            string[] Splitter = Src.Split(' ');
            string Wrd1 = Src, Wrd2 = string.Empty;

            if (Splitter.Length > 1 && Splitter[Splitter.Length - 1].Any(char.IsDigit))
            {
                Wrd2 = Splitter[Splitter.Length - 1];
                Wrd1 = Src.Remove(Src.LastIndexOf(' '));
            }

            return (Wrd1, Wrd2);
        }        

        public static int ExtractIDFromRtfFileName(string SrcFileName)
        {//Extrait l'id auquel fait référence le nom du fichier RTF

            string[] titles = SrcFileName.Split('\\');
            string file = titles[titles.Length - 1];

            return int.Parse(file.Split('-')[0]);
        }

        public static void SortById(List<Messages> MsgList)
        {//range les messages par ordre croissant d'id (tri à bulle), pas trouvé Qsort() en C#

            for (int i = 0; i < MsgList.Count; i++)
            {
                for (int j = 0; j < MsgList.Count - i - 1; j++)
                {
                    if (MsgList[j].id > MsgList[j + 1].id)
                        (MsgList[j], MsgList[j + 1]) = (MsgList[j + 1], MsgList[j]);
                }
            }
        }

        public static Messages? FetchByID(List<Messages> Src, int Seek)
        {//trouve un message par son id (dichotomie), nécessite d'avoir fait Qsort pour fonctionner

            int beg = 0;
            int end = Src.Count - 1;

            while (beg <= end)
            {
                int mid = beg + (end - beg) / 2;                

                if (Seek == Src[mid].id)
                {//trouvé
                    return Src[mid];
                }
                else if (Seek > Src[mid].id)
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

        public static void XcelHeaderWrite(ExcelWorksheet WsDst)
        {//écriture des en-têtes des colonnes du fichier Xcel

            for (int i = 1; i <= colonnes.Length; i++)
                WsDst.Cells[2, i].Value = colonnes[i-1];
        }

        public void XcelLineWrite(ExcelWorksheet WsDst, int line)
        {//écriture d'un message sur sa ligne dans le fichier Xcel

            WsDst.Cells[line, ColumnsIndex[(int)Column.used]].Value             = "X";
            WsDst.Cells[line, ColumnsIndex[(int)Column.id]].Value               = this.id + 80000;
            WsDst.Cells[line, ColumnsIndex[(int)Column.type]].Value             = this.Type;
            WsDst.Cells[line, ColumnsIndex[(int)Column.responsability]].Value   = this.Responsability;
            WsDst.Cells[line, ColumnsIndex[(int)Column.zone]].Value             = this.Zone;
            WsDst.Cells[line, ColumnsIndex[(int)Column.zoneRef]].Value          = this.ZoneRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.module]].Value           = this.Module;
            WsDst.Cells[line, ColumnsIndex[(int)Column.moduleRef]].Value        = this.ModuleRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.module2]].Value          = this.Module2;
            WsDst.Cells[line, ColumnsIndex[(int)Column.module2Ref]].Value       = this.Module2Ref;
            WsDst.Cells[line, ColumnsIndex[(int)Column.msg]].Value              = this.Msg;
            WsDst.Cells[line, ColumnsIndex[(int)Column.msgRef]].Value           = this.MsgRef;
            WsDst.Cells[line, ColumnsIndex[(int)Column.msgHelp]].Value          = this.MsgHelp;
            WsDst.Cells[line, ColumnsIndex[(int)Column.msgAction]].Value        = this.MsgAction;
        }

        public static void XcelLineColIndexGet(ExcelWorksheet WsSrc)
        {
            for (int i = 1; i <= colonnes.Length; i++)
                for (int j = 1; j <= 50; j++)
                {
                    if (WsSrc.Cells[2, j].Value?.ToString() == colonnes[i - 1])
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
