﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace DiplomR_PasportEko
{
    class WordHelper
    {
        private FileInfo _fileInfo;
        public WordHelper(string fileName)
        {
            if(File.Exists(fileName)) 
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not file");
            }
        }
        internal bool Process(Dictionary<string, string> items)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;
                    Object wrap = Word.WdFindWrap.wdFindContinue; 
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Object newFileName = Path.Combine(documentsPath, DateTime.Now.ToString("ddMMyyyy HHmmss") + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();
                return true;
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка", "Уведомление", MessageBoxButtons.OK);
            }
            finally
            {
                if(app != null)
                {
                    app.Quit();
                } 
            }
            return false;
        }
    }
}
