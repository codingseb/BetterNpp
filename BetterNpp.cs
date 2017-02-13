/***************************************************************
 * Author      : Sébastien Geiser
 * Creation    : Avril 2014
 * Licence     : MIT
 ***************************************************************/

using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NppScripts;

namespace BetterNpp
{
    /// <summary>
    /// Classe statique qui permet d'accéder à toutes les fonctions simplifiées et Pure C#
    /// disponible pour piloter Notepad++
    /// Cette classe est le point d'entrée à toutes les sous-classes et fonctions de BetterNpp.
    /// </summary>
    public static class BNpp
    {
        /// <summary>
        /// Constante défini au niveau des APIs Windows qui défini le nombre max de caractères
        /// que peut prendre un chemin d'accès complet d'un fichier
        /// </summary>
        public const int PATH_MAX = 260;

        #region BNppBase

        /// <summary>
        /// Récupère le chemin d'accès au fichier courant ouvert dans Notepad++
        /// </summary>
        public static string CurrentPath
        {
            get
            {
                string path;
                Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETFULLCURRENTPATH, 0, out path);
                return path;
            }
        }

        /// <summary>
        /// Récupère le chemin d'accès au dossier ou se trouve l'exécutable de l'instance courante de Notepad++
        /// </summary>
        public static string NppBinDirectoryPath
        {
            get
            {
                string path;
                Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETNPPDIRECTORY, 0, out path);
                return path;
            }
        }

        /// <summary>
        /// Récupère le nombre de fichiers ouverts dans Notepad++
        /// </summary>
        public static int NbrOfOpenedFiles
        {
            get
            {
                // le - 1 enlève le "new 1" en trop, toujours présent dans la liste des fichier ouverts
                return Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETNBOPENFILES, 0, (int)NppMsg.ALL_OPEN_FILES).ToInt32() - 1;
            }
        }

        /// <summary>
        /// Récupère la liste de tous les chemins d'accès aux fichiers ouverts dans Notepad++
        /// </summary>
        public static List<string> AllOpenedDocuments
        {
            get
            {
                List<string> result = new List<string>();
                int nbr = NbrOfOpenedFiles;
                ClikeStringArray cStringArray = new ClikeStringArray(nbr, PATH_MAX);

                Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETOPENFILENAMES, cStringArray.NativePointer, nbr);

                result = cStringArray.ManagedStringsUnicode;

                cStringArray.Dispose();

                return result;
            }
        }

        /// <summary>
        /// Affiche le tab du document déjà ouvert spécifié dans Notepad++
        /// </summary>
        /// <param name="tabPath">Le chemin d'accès du fichier ouvert dans Notepad++ que l'on veut afficher</param>
        /// <see>AllOpenedFilesPaths</see>
        public static void ShowOpenedDocument(string tabPath)
        {
            Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_SWITCHTOFILE , 0, tabPath);
        }

        /// <summary>
        /// Crée un nouveau document dans Notepad++ dans un nouveau tab.
        /// </summary>
        public static void CreateNewDocument()
        {
            Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_NEW, 0);
        }

        /// <summary>
        /// Essai d'ouvrir le fichier spécifié dans Notepad++
        /// </summary>
        /// <param name="fileName">Le chemin d'accès du fichier à ouvrir dans Notepad++</param>
        /// <returns><c>True</c> Si le fichier à pu s'ouvrir correctement <c>False</c> si le fichier spécifié ne peut pas être ouvert</returns>
        public static bool OpenFile(string fileName)
        {
            bool result = false;

            if(File.Exists(fileName))
            {
                result = Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_DOOPEN, 0 , fileName).ToInt32() == 1;
            }

            return result;
        }

        /// <summary>
        /// Sauvegarde le document courant
        /// </summary>
        public static void SaveCurrentDocument()
        {
            Npp.SaveCurrentDocument();
        }

        /// <summary>
        /// Sauvegarde tous les documents actuellement ouverts dans Notepad++
        /// </summary>
        public static void SaveAllOpenedDocuments()
        {
            Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_SAVEALLFILES , 0, 0);
        }

        /// <summary>
        /// Récupère les caractères de fin de lignes courant
        /// !!! Attention pour le moment bug. !!! Enlève la coloration syntaxique du fichier courant
        /// </summary>
        public static string CurrentEOL
        {
            get
            {
                string eol = "\n";
                int value = Win32.SendMessage(Npp.NppHandle, SciMsg.SCI_GETEOLMODE, 0, 0).ToInt32();

                switch(value)
                {
                    case 0:
                        eol = "\r\n";
                    break;
                    case 1:
                        eol = "\r";
                    break;
                    default:
                    break;
                }
                
                return eol;
            }
        }

        /// <summary>
        /// Récupère ou attribue le texte complet du tab Notepad++ courant
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        public static string Text
        {
            get
            {
                return BEncoding.GetUtf8TextFromScintillaText(Npp.GetAllText());
            }

            set
            {
                Npp.SetAllText(BEncoding.GetScintillaTextFromUtf8Text(value));
            }
        }

        /// <summary>
        /// Récupère ou attribue le début de la sélection de texte
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        public static int SelectionStart
        {
            get
            {
                int curPos = (int)Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_GETSELECTIONSTART, 0, 0);
                string beginingText = Npp.GetTextBetween(0, curPos);
                string text = BEncoding.GetUtf8TextFromScintillaText(beginingText);
                return text.Length;
            }

            set
            {
                string allText = Text;
                int startToUse = value;

                if(value < 0)
                {
                    startToUse = 0;
                }
                else if(value > allText.Length)
                {
                    startToUse = allText.Length;
                }

                string beforeText = allText.Substring(0, startToUse);
                string beforeTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(beforeText);
                int defaultStart = beforeTextInDefaultEncoding.Length;

                Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_SETSELECTIONSTART, defaultStart, 0);
            }
        }

        /// <summary>
        /// Récupère ou attribue la fin de la sélection de texte
        /// <br/>si aucun texte n'est sélectionné SelectionEnd = SelectionStart
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        public static int SelectionEnd
        {
            get
            {
                int curPos = (int)Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_GETSELECTIONEND, 0, 0);
                string beginingText = Npp.GetTextBetween(0, curPos);
                string text = BEncoding.GetUtf8TextFromScintillaText(beginingText);
                return text.Length;
            }

            set
            {
                string allText = Text;
                int endToUse = value;

                if(value < 0)
                {
                    endToUse = 0;
                }
                else if(value > allText.Length)
                {
                    endToUse = allText.Length;
                }

                string afterText = allText.Substring(0, endToUse);
                string afterTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(afterText);
                int defaultEnd = afterTextInDefaultEncoding.Length;

                Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_SETSELECTIONEND, defaultEnd, 0);
            }
        }

        /// <summary>
        /// Récupère ou attribue la longueur de la sélection de texte
        /// <br/>Si aucun texte n'est sélectionné SelectionEnd = 0
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        public static int SelectionLength
        {
            get
            {
                return SelectionEnd - SelectionStart;
            }

            set
            {
                SelectionEnd = SelectionStart + (value < 0 ? 0 : value);
            }
        }

        /// <summary>
        /// Récupère ou remplace le texte actuellement sélectionné
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        public static string SelectedText
        {
            get
            {
                int start = SelectionStart;
                int end = SelectionEnd;

                return end-start == 0 ? "" : Text.Substring(start, end-start);
            }

            set
            {
                string defaultNewText = BEncoding.GetScintillaTextFromUtf8Text(value);
                Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_REPLACESEL, 0 , defaultNewText);
            }
        }

        /// <summary>
        /// Sélectionne dans le tab Notepad++ courant le texte entre start et end
        /// et positionne le scroll pour voir la sélection.
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        /// <param name="start">Position du début du texte à sélectionner dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length</param>
        /// <param name="end">Position de fin du texte à sélectionner dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length<br/> Si plus petit que start -> forcé à start</param>
        public static void SelectTextAndShow(int start, int end)
        {
            string allText = Text;
            int startToUse = start;
            int endToUse = end;

            if(start < 0)
            {
                startToUse = 0;
            }
            else if(start > allText.Length)
            {
                startToUse = allText.Length;
            }

            if(end < 0)
            {
                endToUse = 0;
            }
            else if(end > allText.Length)
            {
                endToUse = allText.Length;
            }
            else if(endToUse < startToUse)
            {
                endToUse = startToUse;
            }

            string beforeText = allText.Substring(0, startToUse);
            string beforeTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(beforeText);
            int defaultStart = beforeTextInDefaultEncoding.Length;
            string endText = allText.Substring(0, endToUse);
            string endTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(endText);
            int defaultEnd = endTextInDefaultEncoding.Length;

            Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_GOTOPOS, defaultStart, 0);
            Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_SETSELECTIONEND, defaultEnd, 0);
        }
		
        /// <summary>
        /// Si la sélection multiple est activée ajoute la sélection spécifié
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        /// <param name="start">Position du début du texte à sélectionner dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length</param>
        /// <param name="end">Position de fin du texte à sélectionner dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length<br/> Si plus petit que start -> forcé à start</param>
		
		public static void AddSelection(int start, int end)
		{
			string allText = Text;
			int startToUse = start;
            int endToUse = end;

            if(start < 0)
            {
                startToUse = 0;
            }
            else if(start > allText.Length)
            {
                startToUse = allText.Length;
            }

            if(end < 0)
            {
                endToUse = 0;
            }
            else if(end > allText.Length)
            {
                endToUse = allText.Length;
            }
            else if(endToUse < startToUse)
            {
                endToUse = startToUse;
            }

			string beforeText = allText.Substring(0, startToUse);
            string beforeTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(beforeText);
            int defaultStart = beforeTextInDefaultEncoding.Length;
            string endText = allText.Substring(0, endToUse);
            string endTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(endText);
            int defaultEnd = endTextInDefaultEncoding.Length;
			
			Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_ADDSELECTION, defaultStart, defaultEnd);
		}

        /// <summary>
        /// Remplace dans le tab Notepad++ courant le texte entre start et end
        /// par newText, sélectionne le nouveau texte
        /// et positionne le scroll pour voir la sélection.
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        /// <param name="newText">Le nouveau texte</param>
        /// <param name="start">Position du début du texte à remplacer dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length</param>
        /// <param name="end">Position de fin du texte à remplacer dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length<br/> Si plus petit que start -> forcé à start</param>
        public static void ReplaceTextAtPosition(string newText, int start, int end)
        {
            SelectTextAndShow(start, end);
            string defaultNewText = BEncoding.GetScintillaTextFromUtf8Text(newText);
            Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_REPLACESEL, 0 , defaultNewText);
            SelectTextAndShow(start, start + newText.Length);
        }

        /// <summary>
        /// Insère dans le tab Notepad++ courant le texte à la position spécifiée
        /// par newText, sélectionne le nouveau texte
        /// et positionne le scroll pour voir la sélection.
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        /// <param name="text">Le texte à insérer</param>
        /// <param name="pos">la position où insérer le texte dans le texte entier<br/> Si plus petit que insertion à la position courante<br/> Si plus grand que Text.Length -> forcé à Text.Length</param>
        public static void InsertText(string text, int pos)
        {
            string allText = Text;
            int posToUse = pos;

            if(pos < 0)
            {
                posToUse = SelectionStart;
            }
            else if(pos > allText.Length)
            {
                posToUse = allText.Length;
            }

            string beforeText = allText.Substring(0, posToUse);
            string beforeTextInDefaultEncoding = BEncoding.GetScintillaTextFromUtf8Text(beforeText);
            int defaultPos = beforeTextInDefaultEncoding.Length;
            string defaultText = BEncoding.GetScintillaTextFromUtf8Text(text);

            Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_INSERTTEXT, defaultPos , defaultText);
        }

        /// <summary>
        /// Récupère la ligne courante où se situe le curseur (indice 1)
        /// Dans le tab courant
        /// </summary>
        public static int CurrentLine
        {
            get
            {
				int result = 1;
				
				try
				{
					string[] subtextLines = BNpp.Text
						.Substring(0, BNpp.SelectionStart)
						.Split(new string[] { "\r\n", "\r", "\n"}, StringSplitOptions.None);
						
						result = subtextLines.Length;				
				}
				catch {}
                
				return result;
				
                // return Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETCURRENTLINE, 0, 0).ToInt32() + 1;
            }
        }

        /// <summary>
        /// Récupère la colonne courante où se situe le curseur (indice 0)
        /// (Position du curseur dans la ligne courante)
        /// Dans le tab courant
        /// </summary>
        public static int CurrentColumn
        {
            get
            {
				int result = 0;
				
				try
				{
					string[] subtextLines = BNpp.Text
						.Substring(0, BNpp.SelectionStart)
						.Split(new string[] { "\r\n", "\r", "\n"}, StringSplitOptions.None);
						
						result = subtextLines[subtextLines.Length - 1].Length;				
				}
				catch {}
                
				return result;
				
				// return Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_GETCURRENTCOLUMN, 0, 0).ToInt32();
            }
        }

        /// <summary>
        /// Récupère le texte entre les 2 position spécifiées
        /// <br/>(Gère la conversion d'encodage Npp/C#)
        /// </summary>
        /// <param name="start">Position du début du texte à récupérer dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length</param>
        /// <param name="end">Position de fin du texte à récupérer dans le texte entier<br/> Si plus petit que 0 -> forcé à zéro<br/> Si plus grand que Text.Length -> forcé à Text.Length<br/> Si plus petit que start -> forcé à start</param>
        /// <returns>La chaine de caractères correspondant au texte entre les 2 positions spécifiée</returns>
        public static string GetTextBetween(int start, int end)
        {
            string result = "";

            try
            {
                string allText = Text;
                int startToUse = start;
                int endToUse = end;

                if(start < 0)
                {
                    startToUse = 0;
                }
                else if(start > allText.Length)
                {
                    startToUse = allText.Length;
                }

                if(end < 0)
                {
                    endToUse = 0;
                }
                else if(end > allText.Length)
                {
                    endToUse = allText.Length;
                }
                else if(endToUse < startToUse)
                {
                    endToUse = startToUse;
                }

                result = allText.Substring(startToUse, endToUse - startToUse);
            }
            catch { }

            return result;
        }
		
        /// <summary>
        /// Récupère le texte de la ligne spécifiée
        /// </summary>
        /// <param name="lineNb">Numéro de la ligne dont on veut récupérer le texte</param>
        /// <returns>Le texte de la ligne spécifiée</returns>
		public static string GetLineText(int lineNb)
		{
			string result = "";
			
			try
			{
				result = BNpp.Text.Split(new string[] {"\r\n", "\r", "\n"}, StringSplitOptions.None)[lineNb - 1];
			}
			catch {}
			
			return result;
		}

        #endregion

        #region WinFormIntegration

        /// <summary>
        /// Offre des fonctions simples pour intégrer les WinForms C# dans Notepad++
        /// Toutes les forms d'un plugin C# devrait au moins utiliser RegisterWinFormInNpp et UnregisterWinFormInNpp
        /// </summary>
        public static class BCScharpWindowsNppIntegration
        {
            [DllImport("user32.dll")]
            private static extern int SetWindowLong(IntPtr hWnd, int windowLongFlags, IntPtr dwNewLong);

            private enum WindowLongFlags : int
            {
                GWL_EXSTYLE = -20,
                GWLP_HINSTANCE = -6,
                GWLP_HWNDPARENT = -8,
                GWL_ID = -12,
                GWL_STYLE = -16,
                GWL_USERDATA = -21,
                GWL_WNDPROC = -4,
                DWLP_USER = 0x8,
                DWLP_MSGRESULT = 0x0,
                DWLP_DLGPROC = 0x4
            }

            /// <summary>
            /// Décrit la manière dont une winform a la priorité sur Notepad++
            /// au niveau de l'ordre d'affichage lorsque Notepad++ a le focus
            /// </summary>
            public enum WinFormNppZOrder
            {
                DISCONNECTED_FROM_NPP,
                KEEP_SCRIPTFORM_IN_FRONT_OF_NPP
            }

            /// <summary>
            /// Inscrit la fenêtre winform spécifier auprès de Notepad++.
            /// Cela permet à Notepad++ de gérer la fenêtre comme une fenêtre de plugin.
            /// Ca permet en en outre de gérer correctement certaines touches comme {TAB}
            /// dans la fenêtre en question
            /// </summary>
            /// <param name="scriptForm">La fenêtre à inscrire</param>
            /// <param name="winFormNppZOrder">Spécifie la manière dont doit être gérer la priorité de l'ordre d'affichage lorsque Notepad++ a le focus</param>
            /// <remarks>Ne pas oublier de désinscrire la fenêtre à la fermeture de celle-ci (Avant de disposer)</remarks>
            /// <remarks>Toutes les forms d'un plugin C# devrait au moins utiliser RegisterWinFormInNpp et UnregisterWinFormInNpp</remarks>
            public static void RegisterWinFormInNpp(Form scriptForm, WinFormNppZOrder winFormNppZOrder)
            {
                // Oubliez cette ligne sensé faire fonctionner le tab dans la fenêtre du plugin (c'est de la merde ça marche pas.
                // Et en plus ça désactive le Enter
                // Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_MODELESSDIALOG, (int)NppMsg.MODELESSDIALOGADD, scriptForm.Handle);

                // A la place on fait ça
                scriptForm.KeyDown += new System.Windows.Forms.KeyEventHandler(TabKeyManage_KeyDown);

                if(winFormNppZOrder == WinFormNppZOrder.KEEP_SCRIPTFORM_IN_FRONT_OF_NPP)
                {
                    SetWindowLong(scriptForm.Handle, (int)WindowLongFlags.GWLP_HWNDPARENT, Npp.NppHandle);
                }
            }

            /// <summary>
            /// Inscrit la fenêtre winform spécifier auprès de Notepad++.
            /// Cela permet à Notepad++ de gérer la fenêtre comme une fenêtre de plugin.
            /// Ca permet en en outre de gérer correctement certaines touches comme {TAB}
            /// dans la fenêtre en question
            /// </summary>
            /// <param name="scriptForm">La fenêtre à inscrire</param>
            /// <remarks>Ne pas oublier de désinscrire la fenêtre à la fermeture de celle-ci (Avant de disposer)</remarks>
            /// <remarks>Toutes les forms d'un plugin C# devrait au moins utiliser RegisterWinFormInNpp et UnregisterWinFormInNpp</remarks>
            public static void RegisterWinFormInNpp(Form scriptForm)
            {
                // Oubliez cette ligne sensé faire fonctionner le tab dans la fenêtre du plugin (c'est de la merde ça marche pas.
                // Et en plus ça désactive le Enter
                //Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_MODELESSDIALOG, (int)NppMsg.MODELESSDIALOGADD, scriptForm.Handle);

                // A la place on fait ça
                scriptForm.KeyDown += new System.Windows.Forms.KeyEventHandler(TabKeyManage_KeyDown);

                SetWindowLong(scriptForm.Handle, (int)WindowLongFlags.GWLP_HWNDPARENT, Npp.NppHandle);
            }

            /// <summary>
            /// Désinscrit la fenêtre winform spécifier auprès de Notepad++.
            /// Cela permet à Notepad++ de gérer la fenêtre comme une fenêtre de plugin.
            /// Ca permet en en outre de gérer correctement certaines touches comme {TAB}
            /// dans la fenêtre en question
            /// </summary>
            /// <param name="scriptForm">La fenêtre à désinscrire</param>
            /// <remarks>Toutes les forms d'un plugin C# devrait au moins utiliser RegisterWinFormInNpp et UnregisterWinFormInNpp</remarks>
            public static void UnregisterWinFormInNpp(Form scriptForm)
            {
                // On n'utilise pas ça. Voir commentaire dans RegisterWinFormInNpp
                // Win32.SendMessage(Npp.NppHandle, NppMsg.NPPM_MODELESSDIALOG, (int)NppMsg.MODELESSDIALOGREMOVE, scriptForm.Handle);

                scriptForm.KeyDown -= new System.Windows.Forms.KeyEventHandler(TabKeyManage_KeyDown);
            }

            /// <summary>
            /// Pour gérer la navigation par tab entre les composants.
            /// </summary>
            private static void TabKeyManage_KeyDown(object sender, KeyEventArgs e)
            {
                if(e.Shift && e.KeyCode == Keys.Tab)
                {
                    ((Form)sender).SelectNextControl(((Form)sender).ActiveControl, false, true, true, true);
                }
                else if(e.KeyCode == Keys.Tab)
                {
                    ((Form)sender).SelectNextControl(((Form)sender).ActiveControl, true, true, true, true);
                }
                else
                {
                    e.Handled = false;
                    e.SuppressKeyPress = false;
                }
            }
        }

        #endregion

        #region UndoRedo

        /// <summary>
        /// Offre des fonctions simples pour gérer l'historique des actions sur le fichier courant
        /// </summary>
        public static class BUndoRedo
        {
            /// <summary>
            /// Pour savoir si une/des action(s) à annuler est/sont disponible(s)
            /// </summary>
            /// <returns><c>True</c> si des action sont disponible, <c>False</c> sinon</returns>
            public static bool CanUndo
            {
                get
                {
                    return Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_CANUNDO, 0, 0).ToInt32() != 0;
                }
            }

            /// <summary>
            /// Pour savoir si une/des action(s) à réeffectuer est/sont disponible(s)
            /// </summary>
            /// <returns><c>True</c> si des action sont disponible, <c>False</c> sinon</returns>
            public static bool CanRedo
            {
                get
                {
                    return Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_CANREDO, 0, 0).ToInt32() != 0;
                }
            }

            /// <summary>
            /// Annule la dernière action effectué sur le fichier courant
            /// </summary>
            public static void Undo()
            {
                Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_UNDO, 0, 0);
            }

            /// <summary>
            /// Réeffectue la dernière action annulée du fichier courant
            /// </summary>
            public static void Redo()
            {
                Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_REDO, 0, 0);
            }
        }

        #endregion

        #region Encoding

        /// <summary>
        /// Offre des fonctions simples pour convertir l'encodage d'un texte
        /// entre l'encodage du document courant dans Notepad++ et l'encodage en C# (UTF8)
        /// </summary>
        public static class BEncoding
        {
            private static Encoding utf8 = Encoding.UTF8;

            /// <summary>
            /// Convertit le texte spécifier de l'encodage du document Notepad++ courant à l'encodage C# (UTF8)
            /// </summary>
            public static string GetUtf8TextFromScintillaText(string scText)
            {
                string result = "";
                int iEncoding = (int)Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_GETCODEPAGE, 0, 0);

                switch(iEncoding)
                {
                    case 65001 : // UTF8
                        result = utf8.GetString(Encoding.Default.GetBytes(scText));
                    break;
                    default:
                        Encoding ANSI = Encoding.GetEncoding(1252);

                        byte[] ansiBytes = ANSI.GetBytes(scText);
                        byte[] utf8Bytes = Encoding.Convert(ANSI, Encoding.UTF8, ansiBytes);

                        result = Encoding.UTF8.GetString(utf8Bytes);
                    break;
                }

                return result;
            }

            /// <summary>
            /// Convertit le texte spécifier de l'encodage C# (UTF8) à l'encodage document Notepad++ courant
            /// </summary>
            public static string GetScintillaTextFromUtf8Text(string utf8Text)
            {
                string result = "";
                int iEncoding = (int)Win32.SendMessage(Npp.CurrentScintilla, SciMsg.SCI_GETCODEPAGE, 0, 0);

                switch(iEncoding)
                {
                    case 65001 : // UTF8
                        result = Encoding.Default.GetString(utf8.GetBytes(utf8Text));
                    break;
                    default:
                        Encoding ANSI = Encoding.GetEncoding(1252);

                    byte[] utf8Bytes = utf8.GetBytes(utf8Text);
                    byte[] ansiBytes = Encoding.Convert(Encoding.UTF8, ANSI, utf8Bytes);

                    result = ANSI.GetString(ansiBytes);
                    break;
                }

                return result;
            }
        }

        #endregion

        #region NppMenusActions

        /// <summary>
        /// Offre des fonctions simples pour accéder aux commandes des menus Notepad++ les plus courants
        /// </summary>
        public static class BMenuAction
        {
            /// <summary>
            /// Effectue l'action du menu Fichier->Nouveau
            /// <br/>Créer un nouveau fichier dans un nouveau tab NotePad++
            /// </summary>
            public static void FileNew()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_NEW, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Ouvrir
            /// <br/>Ouvre la boite de dialogue permettant d'ouvrir un fichier existant dans Notepad++
            /// </summary>
            public static void FileOpen()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_OPEN, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Recharger depuis le disque
            /// <br/>Recharge le fichier courant depuis le disque et annule toute les modifications depuis la dernière sauvegarde.
            /// </summary>
            public static void FileReloadFromDisk()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_RELOAD, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Enregistrer
            /// <br/>Enregistre les modification du fichier courant.
            /// <br/>Si le fichier n'a pas encore été sauver sur le disque,
            /// <br/>ouvre la boite de dialogue "Enregistrer sous" permettant d'enregistrer un fichier sur le disque
            /// </summary>
            public static void FileSave()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_SAVE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Enregistrer sous
            /// <br/>Ouvre la boite de dialogue "Enregistrer sous" permettant d'enregistrer le fichier courant sur le disque
            /// </summary>
            public static void FileSaveAs()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_SAVEAS, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Fermer
            /// <br/>Ferme le fichier courant dans NotePad++
            /// </summary>
            public static void FileClose()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_CLOSE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Fermer tout
            /// <br/>Ferme tous les fichiers actuellement ouvert dans NotePad++
            /// </summary>
            public static void FileCloseAll()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_CLOSEALL, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Fermer tout sauf le document actuel
            /// <br/>Ferme tous les fichiers actuellement ouvert dans NotePad++ sauf le tab courant
            /// </summary>
            public static void FileCloseAllButCurrent()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_CLOSEALL_BUT_CURRENT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Imprimer
            /// <br/>Affiche la boite de dialogue d'impression de NotePad++
            /// </summary>
            public static void FilePrint()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_PRINT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Imprimer immédiatement
            /// <br/>Lance l'impression du fichier courant sur l'imprimante par défaut
            /// </summary>
            public static void FilePrintNow()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_PRINTNOW, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Fichier->Quitter
            /// <br/>Quitte NotePad++
            /// </summary>
            public static void FileExit()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_FILE_EXIT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Annuler
            /// <br/>Annule la dernière action effectué sur le fichier courant
            /// </summary>
            public static void EditUndo()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_UNDO, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Rétablir
            /// <br/>Réeffectue la dernière action annulée du fichier courant
            /// </summary>
            public static void EditRedo()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_REDO, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Copier
            /// <br/>Copie le texte actuellement sélectionné dans NotePad++ dans le presse papier
            /// </summary>
            public static void EditCopy()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_COPY, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Couper
            /// <br/>Copie le texte actuellement sélectionné dans NotePad++ dans le presse papier
            /// <br/>et le retire du texte courant.
            /// </summary>
            public static void EditCut()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_CUT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Coller
            /// <br/>Colle le texte actuellement dans le presse-papier à l'endroit de la sélection dans NotePad++
            /// </summary>
            public static void EditPaste()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_PASTE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Supprimer
            /// <br/>Supprime le texte actuellement sélectionné du fichier courant.
            /// </summary>
            public static void EditDelete()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_DELETE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Sélectionner tout
            /// <br/>Sélectionne tout le texte du fichier courant dans NotePad++
            /// </summary>
            public static void EditSelectAll()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_SELECTALL, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->MAJUSCULE/minuscule->EN MAJUSCULE
            /// <br/>Passe le texte actuellement sélectionné en majuscule
            /// </summary>
            public static void EditUpperCase()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_UPPERCASE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->MAJUSCULE/minuscule->en minuscule
            /// <br/>Passe le texte actuellement sélectionné en minuscule
            /// </summary>
            public static void EditLowerCase()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_LOWERCASE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Commentaire->Commenter le bloc sélectionné
            /// <br/>Commente le texte actuellement sélectionné en mode ligne (ex : /* block */ en C#)
            /// </summary>
            public static void EditStreamComment()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_STREAM_COMMENT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Edition->Commentaire->Commenter/Décommenter (mode ligne)
            /// <br/>Commente ou décommente le texte des lignes actuellement sélectionnées en mode ligne (ex : // ligne de code en C#)
            /// </summary>
            public static void EditBlockComment()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EDIT_BLOCK_COMMENT, 0);
            }

            /// <summary>
            /// Effectue l'action du menu Exécution->Exécuter...
            /// <br/>Lance l'exécution du fichier courant suivant son type (Si celui-ci est exécutable)
            /// </summary>
            public static void Execute()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_EXECUTE, 0);
            }

            /// <summary>
            /// Effectue l'action du menu ?->À propos de Notepad++...
            /// <br/>Affiche la boite de dialogue à propos de
            /// </summary>
            public static void About()
            {
                Win32.SendMenuCmd(Npp.NppHandle, NppMenuCmd.IDM_ABOUT, 0);
            }
        }

        #endregion
    }

    /// <summary>
    /// Simple Mappage par héritage de NppScript
    /// pour ne pas devoir utiliser using NppScript.
    /// Remplacez
    /// <br/>public class Script : NppScript
    /// <br/>par
    /// <br/>public class Script : BNppScript
    /// </summary>
    public abstract class BNppScript : NppScript
    {
        public abstract override void Run();
    }
    
}




