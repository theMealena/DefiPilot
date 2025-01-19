using System;
using System.Web;

/// <summary>
/// Summary description for Uties
/// </summary>

namespace DefiPilot
{
	public class Uties
	{
		public Uties()
		{
			//
			// TODO: Add constructor logic here
			//
		}

        /// <summary> Fichier vers List. </summary>
		/// <param name="Src"> Fichier à lire.</param>
		/// <param name="Dst"> List à renseigner.</param>
		/// <returns> N/A .</returns>
        static public void FileToList(StreamReader Src, List<string> Dst)
		{
			while (!Src.EndOfStream)
			{
				string? Tmp  = Src.ReadLine();

				if(Tmp != null)
					Dst.Add(Tmp);
			}
		}

        /// <summary> Recherche et extraction de parametre dans list. </summary>
        /// <param name="Seek"> Paramètre à trouver.</param>
        /// <param name="Src"> List source.</param>
        /// <returns> string comme valeur, vide si echec.</returns>
        /// <remarks> Sous forme de string ce qui se trouve derrière le caractère '='.</remarks>
        static public string ListParamFetch(string Seek, List<string> Src)
		{
			for (int i = 0; i < Src.Count; i++)
			{
				if (Src[i].StartsWith(Seek))
				{
					int equality = Src[i].LastIndexOf('=');

					if (equality > 0)
						return Src[i].Substring(equality + 1, Src[i].Length - equality - 1).Trim();
				}
			}

			return string.Empty;
		}

        /// <summary> Recherche et extraction de parametre dans list. </summary>
        /// <param name="Seek"> Paramètre à trouver.</param>
        /// <param name="Src"> List source.</param>
		/// <param name="startFrom"> Index du début de la recherche.</param>
        /// <returns> string comme valeur, vide si echec.</returns>
        /// <remarks> Sous forme de string ce qui se trouve derrière le caractère '='.</remarks>
        static public string ListParamFetch(string Seek, List<string> Src, int startFrom)
        {
			if (startFrom < Src.Count)
			{
				for (int i = startFrom; i < Src.Count; i++)
				{
					if (Src[i].StartsWith(Seek))
					{
						int equality = Src[i].LastIndexOf('=');

						if (equality > 0)						
							return Src[i].Substring(equality + 1, Src[i].Length - equality - 1).Trim();
					}
				}
			}

            return string.Empty;
        }

        static public int ListParamIndexOf(string seek, List<string> Src)
        {
            for (int i = 0; i < Src.Count; i++)
            {
                if (Src[i].StartsWith(seek))  return i;
            }

            return -1;
        }

		static public string IdTypeToName(int Id)
		{
			switch(Id)
			{
				case 0:
                    return "ENTREE";
					//break;
                case 1:
                    return "SORTIE";
					//break;
                case 2:
					return "DIVERS";
					//break;
				case 3:
					return "ACTIF";
					//break;
				case 4:
					return "RAPIDE";
					//break;
				case 5:
					return "PARA";
					//break;
				case 6:
					return "ANALOG";
					//break;
				case 8:
					return "DEFAUT";
					//break;
				default:					
					break;
			}

			return string.Empty;
		}

		static public (string, string) OpcExtract(string Src)
		{
			string[] Rslt = Src.Split('=');

			if (Rslt.Length == 2)
			{
				return (Rslt[0].Trim(), Rslt[1].Trim());
			}
			else
			{
				return ("", "");
			}
		}
    }	
}