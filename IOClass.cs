using DefiPilot;
using OfficeOpenXml;
using System;
using System.Drawing.Drawing2D;
using System.Text;

namespace DefiPilot
{

    public class IOBit
    {
        public int bit;

        public string[][] Texts;

        public IOBit()
        {
            bit = 0;

            Texts = new string[5][];//[Texts 1..5][Text ref 1..5] 

            for (int i = 0; i < Texts.Length; i++)
                Texts[i] = new string[2];
        }

        public bool ParseDefinitionLine(string Line)
        {
            if (string.IsNullOrEmpty(Line))
                return false;

            //récup de ce qui est derrière '|'
            int bitDefStart = Line.LastIndexOf('|');

            if (bitDefStart == -1)
                return false;

            string BitDef = Line.Substring(bitDefStart + 1).Trim();

            //entre '|' et ':'
            string[] Parts = BitDef.Split(':', StringSplitOptions.TrimEntries);

            if (Parts.Length > 1)
            {
                (this.Texts[0][0], this.Texts[0][1]) = CheckSubRef(Parts[0]);
                BitDef = Parts[1];
            }
            else
                BitDef = Parts[0];

            //de chaque côté de '-'
            Parts = BitDef.Split("-");

            if (Parts.Length > 1)
            {
                (this.Texts[1][0], this.Texts[1][1]) = CheckSubRef(Parts[0]);
                (this.Texts[2][0], this.Texts[2][1]) = CheckSubRef(Parts[1]);
            }
            else
                (this.Texts[1][0], this.Texts[1][1]) = CheckSubRef(BitDef);

            return true;
        }


        private static (string, string) CheckSubRef(string Src)
        {
            if (string.IsNullOrEmpty(Src))
                return (string.Empty, string.Empty);

            string[] Splitter = Src.Split(' ');
            string Wrd1 = Src, Wrd2 = "";

            if (Splitter.Length > 1 && Splitter[Splitter.Length - 1].Any(char.IsDigit))
            {
                Wrd2 = Splitter[Splitter.Length - 1];
                Wrd1 = Src.Remove(Src.LastIndexOf(' '));
            }

            return (Wrd1, Wrd2);
        }

        public int XcelWrite(ExcelWorksheet WsDst, int line)
        {
            WsDst.Cells[line, IOCard.ColumnsIndex[(int)IOCard.Column.bit_adr]].Value = this.bit;

            for (int i = 0; i < this.Texts.Length; i++)
            {
                WsDst.Cells[line, IOCard.ColumnsIndex[(int)IOCard.Column.text_1] + (i * 2)].Value = this.Texts[i][0];
                WsDst.Cells[line, IOCard.ColumnsIndex[(int)IOCard.Column.text_1Ref + (i * 2)]].Value = this.Texts[i][1];
            }

            return 0;
        }
    }

    public class IOCard
    {
        //index des colonnes dans le fichier Xcel
        public enum Column
        {
            id,
            mod_adr,
            bit_adr,
            text_1,
            text_1Ref,
            text_2,
            text_2Ref,
            text_3,
            text_3Ref,
            text_4,
            text_4Ref,
            text_5,
            text_5Ref,
            designation,
            adress
        };

        private static readonly string[] Colonnes =
        { "Iden(Efidrive)",
        "Adr_Module",
        "Adr_Bit",
        "Texte_1",
        "Texte_1_Ref",
        "Texte_2",
        "Texte_2_Ref",
        "Texte_3",
        "Texte_3_Ref",
        "Texte_4",
        "Texte_4_Ref",
        "Texte_5",
        "Texte_5_Ref",
        "Désignation_Elec_Ref",
        "Adresse"
    };

        public string Id,
                        Mod_adr,
                        Adress;

        public string[][] Texts;

        public int id, mod_adr;

        public static int[] ColumnsIndex = new int[Colonnes.Length];

        public IOBit[] bits;

        public static List<IOCard>[] AllIOs = new List<IOCard>[2];

        static IOCard()
        {
            for (int i = 0; i < AllIOs.Length; i++)
            {
                AllIOs[i] = new List<IOCard>();
            }
        }

        public IOCard()
        {
            Id = Mod_adr = Adress = string.Empty;
            id = mod_adr = 0;

            Texts = new string[5][];

            for (int i = 0; i < Texts.Length; i++)
                Texts[i] = new string[2];

            bits = new IOBit[8];

            for (int i = 0; i < bits.Length; i++)
            {
                bits[i] = new IOBit();
                bits[i].bit = i;
            }
        }


        public static void ExtractOPC(string Folder, int plc)
        {
            AllClear();

            string[] Types = { "entree", "sortie" };
            Folder += @"\SUPER\PLC\";

            for (int i = 0; i < 2; i++)
            {
                string FileName = Folder + $"OPC.{Types[i]}.{plc.ToString("D2")}";

                StreamReader? SrcFile = null;

                try
                {
                    SrcFile = new StreamReader(FileName, Encoding.Latin1);
                }
                catch (Exception)
                {
                    MessageBox.Show($"Impossible d'ouvrir {FileName}, les {Types[i]} ne seront pas traitées", "Erreur fichier");
                    continue;
                }

                while (!SrcFile.EndOfStream)
                {
                    string? ActLine = SrcFile.ReadLine();

                    if (ActLine != null)
                    {
                        IOCard? Card = ExtractOPCLine(ActLine);

                        if (Card != null)
                            AllIOs[i].Add(Card);
                    }
                }

                SortById(AllIOs[i]);
                SrcFile.Close();
            }
        }


        private static IOCard? ExtractOPCLine(string SrcLine)
        {
            if (string.IsNullOrEmpty(SrcLine))
                return null;

            IOCard? Card = null;

            var Words = SrcLine.Split('=', StringSplitOptions.TrimEntries);

            if (Words.Length == 2)
            {
                Card = new IOCard();
                Card.Id = Words[0];
                Card.id = int.Parse(Words[0]);
                Card.Adress = Words[1];

                int start = 0;

                while (char.IsLetter(Card.Adress[start++])) ;

                Card.Mod_adr = Card.Adress.Substring(start - 1);
                try
                {
                    Card.mod_adr = int.Parse(Card.Mod_adr);
                }
                catch (Exception)
                {
                    Card.mod_adr = 0;
                }
            }

            return Card;
        }


        public static void SortById(List<IOCard> Cards)
        {//range les cartes par ordre croissant d'id (tri à bulle), pas trouvé Qsort() en C#

            for (int i = 0; i < Cards.Count; i++)
            {
                for (int j = 0; j < Cards.Count - i - 1; j++)
                {
                    if (Cards[j].id > Cards[j + 1].id)
                        (Cards[j], Cards[j + 1]) = (Cards[j + 1], Cards[j]);
                }
            }
        }


        private static void AllClear()
        {
            for (int i = 0; i < AllIOs.Length; i++)
                AllIOs[i].Clear();
        }


        public static void ExtractDefinition(string Folder, int plc)
        {
            string[] Types = { "Entree", "Sortie" };
            Folder += @"\SUPER\TEXTE\";

            for (int i = 0; i < 2; i++)
            {
                string FileName = Folder + $"DossierPlc.{Types[i]}.{plc.ToString("D2")}.FR";

                StreamReader? SrcFile = null;

                try
                {
                    SrcFile = new StreamReader(FileName, Encoding.Latin1);
                }
                catch (Exception)
                {
                    MessageBox.Show($"Impossible d'ouvrir {FileName}, les {Types[i]} ne seront pas traitées", "Erreur fichier");
                    continue;
                }

                int index = 0;
                IOCard? Card = null;

                while (!SrcFile.EndOfStream)
                {
                    string? SrcLine = SrcFile.ReadLine();

                    if (!string.IsNullOrEmpty(SrcLine))
                    {
                        if (index == 0)
                        {
                            string cardMod = ParseCard(SrcLine);
                            Card = FetchByMod_Adr(AllIOs[i], cardMod);

                            if (Card == null)
                                continue;

                            Card?.ParseDefinitionLine(SrcLine);
                        }

                        Card?.bits[index].ParseDefinitionLine(SrcLine);

                        index = (index + 1) % 8;
                    }
                }

                SrcFile.Close();
            }
        }


        public bool ParseDefinitionLine(string Line)
        {

            if (string.IsNullOrEmpty(Line))
                return false;

            string? ModAd = Line.Substring(0, Line.IndexOf('.'));

            if (string.IsNullOrEmpty(ModAd))
                return false;

            string[] Fields = Line.Split('=', '|');

            if (Fields.Length > 1)
            {
                string[] Definition = Fields[1].Split(',', StringSplitOptions.TrimEntries);

                for (int i = 0; i < Definition.Length; i++)
                {
                    (this.Texts[i][0], this.Texts[i][1]) = CheckSubRef(Definition[i]);
                }
            }

            return true;
        }

        private static (string, string) CheckSubRef(string Src)
        {
            if (string.IsNullOrEmpty(Src))
                return (string.Empty, string.Empty);

            string[] Result = Src.Split(' ', StringSplitOptions.TrimEntries);

            if (Result.Length == 2)
            {
                return (Result[0], Result[1]);
            }
            else if (Result.Length == 1)
            {
                return (Src, string.Empty);
            }

            return (string.Empty, string.Empty);
        }


        public static IOCard? FetchByMod_Adr(List<IOCard> SrcCards, string Seek)
        {
            for (int i = 0; i < SrcCards.Count; i++)
            {
                if (SrcCards[i].Mod_adr == Seek)
                    return SrcCards[i];
            }

            return null;
        }

        public static IOCard? FetchByMod_Adr(List<IOCard> SrcCards, int Seek)
        {
            for (int i = 0; i < SrcCards.Count; i++)
            {
                if (SrcCards[i].mod_adr == Seek)
                    return SrcCards[i];
            }

            return null;
        }

        private static string ParseCard(string Line)
        {
            int ptPos = Line.IndexOf('.');

            if (ptPos > 0)
                return Line.Substring(0, ptPos);

            return string.Empty;
        }

        public static void XcelHeaderWrite(ExcelWorksheet DstWS)
        {
            for (int i = 0; i < Colonnes.Length; i++)
            {
                DstWS.Cells[2, i + 2].Value = Colonnes[i];
            }
        }

        public static void XcelLineColIndexGet(ExcelWorksheet WsSrc)
        {
            for (int i = 0; i < Colonnes.Length; i++)
                for (int j = 1; j <= 26; j++)
                {
                    if (WsSrc.Cells[2, j].Value?.ToString()?.Trim() == Colonnes[i])
                    {
                        ColumnsIndex[i] = j;
                        break;
                    }

                    else
                        ColumnsIndex[i] = 26;
                }
        }


        public int XcelWrite(ExcelWorksheet WsDst, int line)
        {
            WsDst.Cells[line, ColumnsIndex[(int)Column.id]].Value = this.id.ToString("D5");
            WsDst.Cells[line, ColumnsIndex[(int)Column.adress]].Value = this.Adress;

            for (int i = 0; i < this.Texts.Length; i++)
            {
                WsDst.Cells[line, ColumnsIndex[(int)Column.text_1] + i * 2].Value = this.Texts[i][0];
                WsDst.Cells[line, ColumnsIndex[(int)Column.text_1Ref + i * 2]].Value = this.Texts[i][1];
            }

            for (int i = 0; i < this.bits.Length; i++)
            {
                line++;
                WsDst.Cells[line, ColumnsIndex[(int)Column.mod_adr]].Value = this.mod_adr;
                this.bits[i].XcelWrite(WsDst, line);
            }


            return ++line;

        }

        public static void XcelCreate(ExcelWorksheet WsDst, int mode)
        {
            int line = 5;

            for (int i = 0; i < AllIOs[mode].Count; i++)
            {
                line = AllIOs[mode][i].XcelWrite(WsDst, line);
            }
        }
    }
}