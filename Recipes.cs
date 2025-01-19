using System;

namespace DefiPilot
{
	public class Recipes
	{
		static List<string> Names = new List<string>();


		public Recipes()
		{
		}



		public static bool GetNames(string Folder)
		{
			Names.Clear();
			Folder += @"\SUPER\RECETTE\";

			string IniFileName = Folder + @"ListeTables.ini";

			StreamReader SrcFile;

			try
			{
				SrcFile = new StreamReader(IniFileName);
			}
			catch (Exception)
			{
                MessageBox.Show($"Impossible d'ouvrir {IniFileName}, les recettes ne seront pas traitées", "Erreur fichier");
                return false;
            }

			while (!SrcFile.EndOfStream)
			{
				string? SrcLine = SrcFile.ReadLine(),
						Rslt;

				if (string.IsNullOrEmpty(SrcLine))
					continue;

				( _ , Rslt) = Uties.OpcExtract(SrcLine);

                if (!string.IsNullOrEmpty(Rslt))
                    Names.Add(Rslt.Trim());
			}

			SrcFile.Close();

			return true;
        }


		public static void ScanForReceipes(Folder folder, int plc)
		{
			foreach (string Name in Names)
			{
				string IniFileName = $"{folder.FolderPath}\\SUPER\\RECETTE\\{Name}.{(plc*100)}.ini";

				StreamReader SrcFile;

				try
				{
					SrcFile = new StreamReader(IniFileName);
				}
				catch (Exception)
				{
					continue;
				}

				while (!SrcFile.EndOfStream)
				{
					string? SrcLine	= SrcFile.ReadLine(),
							Id;

					if (string.IsNullOrEmpty(SrcLine))
						continue;

					(Id, _ ) = Uties.OpcExtract(SrcLine);

					Variable? ThisVar;

					if (!string.IsNullOrEmpty(Id))
					{
						ThisVar = folder.FetchVarByID(Id);

						if (ThisVar != null)
							ThisVar.Recipe = Name;
					}
				}

				SrcFile.Close();
			}
		}
	}
}