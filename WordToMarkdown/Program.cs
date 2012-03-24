using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Interop.Word;
using NDesk.Options;

// Install-Package ManyConsole

namespace WordToMarkdown
{
    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            bool help = false;
            int verbose = 0;
            List<string> names = new List<string>();

            OptionSet p = new OptionSet()
              .Add("v|verbose", delegate(string v) { if (v != null) ++verbose; })
              .Add("h|?|help", delegate(string v) { help = v != null; })
            ;

            List<string> inputPaths = p.Parse(args);

            if (inputPaths.Count == 0)
            {
                Console.WriteLine("Input files required.");
                Console.WriteLine();
                Console.WriteLine("General usage options:");
                p.WriteOptionDescriptions(Console.Out);
            }

            foreach (string path in inputPaths)
            {
                string inputPath = Path.GetFullPath(path);
                string outputName = Path.GetFileNameWithoutExtension(inputPath) + ".md";
                string outputPath = Path.Combine(Path.GetDirectoryName(inputPath), outputName);

                new Program(inputPath, outputPath);
            }
        }
        
        
        private static string pathToSublimeText = @"C:\Program Files\Sublime Text 2\sublime_text.exe";

        public Program(string inputPath, string outputPath)
        {
            Console.WriteLine("Processing: " + inputPath);

            Application word = LoadWordDocument(inputPath);

            // convert tables to text
            for (int i = word.Selection.Document.Tables.Count; i > 0; i--)
            {
                word.Selection.Document.Tables[i].ConvertToText();
            }

            Console.WriteLine("Processing: Hyperlinks");
            ReplaceHyperlinks(word);

            Console.WriteLine("Processing: Headings");
            ReplaceHeadings(word);

            Console.WriteLine("Processing: number lists");
            ReplaceListNoNumber(word);

            // bold seems to be processed by word itself, leaving here for reference.
            //Console.WriteLine("Processing: bold");
            //bool replaceOneBold = true;
            //while (replaceOneBold)
            //{
            //    replaceOneBold = ReplaceOneBold(word);
            //}

            Console.WriteLine("Processing: italic");
            bool replaceOneItalic = true;
            while (replaceOneItalic)
            {
                replaceOneItalic = ReplaceOneItalic(word);
            }

            Console.WriteLine("Processing: lists");
            ReplaceLists(word);

            Console.WriteLine("Processing: images");
            string prefix = Path.GetFileNameWithoutExtension(outputPath) + "_";

            ReplaceImages(word, prefix);

            Console.WriteLine("Processing: save");
            word.ActiveDocument.SaveAs(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDOSText);

            // OpenSublimeText(outputPath);

            KillWinWord();
        }

        private void ReplaceImages(Application word, string prefix)
        {
            InlineShapes shapes = word.ActiveDocument.InlineShapes;

            int count = 1;
            foreach (InlineShape shape in shapes)
            {
                // (shape.Type == WdInlineShapeType.wdInlineShapePicture)

                shape.Select();
                word.Selection.CopyAsPicture();

                IDataObject ido = (IDataObject)System.Windows.Clipboard.GetDataObject();
                if (null == ido)
                {
                    Console.WriteLine("Image/Picture #" + count + " could not be processed.");
                }

                // can convert to bitmap?
                if (ido.GetDataPresent(DataFormats.Bitmap))
                {
                    // cast the data into a bitmap object
                    string[] s = ido.GetFormats();
                    Bitmap bmp = (Bitmap)ido.GetData("System.Drawing.Bitmap");
                    // validate that puppy
                    if (null == bmp)
                    {
                        Console.WriteLine("Intermediate Image/Picture #" + count + " could not retreated from clipboard.");
                    }

                    string name = count.ToString();

                    string filename = name + ".png";

                    if (prefix.Length != 0)
                    {
                        filename = prefix + filename;
                    }

                    bmp.Save(filename, System.Drawing.Imaging.ImageFormat.Png);
                    word.Selection.Text = "![]" + "(" + filename + ")";
                }
                count++;
            }
        }

        static void OpenSublimeText(string f)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = pathToSublimeText;
            startInfo.Arguments = f;

            try
            {
                Process.Start(startInfo);
            }
            catch
            {
                startInfo.FileName = "Notepad.exe";
                startInfo.Arguments = f;
                Process.Start(startInfo);
            }
        }

        private static void ReplaceHyperlinks(Application word)
        {
            StoryRanges ranges = word.ActiveDocument.StoryRanges;

            foreach (Range range in word.ActiveDocument.StoryRanges)
            {
                foreach (Field field in range.Fields)
                {
                    if (field.Type == WdFieldType.wdFieldHyperlink)
                    {
                        string text = field.Result.Text;
                        string address = field.Result.Hyperlinks[1].Address;

                        field.Result.Text = "[" + text + "](" + address + ")";
                    }
                }
            }
        }

        private static void ReplaceListNoNumber(Application word)
        {
            for (int i = word.Selection.Document.Paragraphs.Count; i > 0; i--)
            {
                Paragraph para = word.Selection.Document.Paragraphs[i];

                if (para.Range.ListFormat.ListType == WdListType.wdListNoNumbering)
                {
                    if (para.LeftIndent > 0)
                    {
                        para.Range.InsertBefore(">");
                    }
                    para.Range.InsertBefore(Environment.NewLine);
                }
            }
        }

        private void ReplaceHeadings(Application word)
        {
            for (int i = 1; i < 7; i++)
            {
                word.Selection.HomeKey(WdUnits.wdStory);

                bool replaceHeading = true;
                while (replaceHeading)
                {
                    replaceHeading = ReplaceHeading(word, i);
                }
            }
        }

        private Application LoadWordDocument(object fullFilePath)
        {
            object wordObject = null;
            Application word = null;

            try
            {
                wordObject = Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                // Do nothing.
            }

            if (wordObject != null)
            {
                word = (Application)wordObject;
            }
            else
            {
                word = new Application();
            }

            //this will open the Word document
            //word.Visible = true;
            object missing = Missing.Value;

            object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            object noEncodingDialog = true;
            object f = false;
            object t = true;

            Document document = word.Documents.Open(ref fullFilePath
                , ref t
                , ref f
                , ref missing
                , ref missing
                , ref missing
                , ref missing
                , ref missing
                , ref missing
                , ref missing
                , ref encoding
                , ref missing
                , ref missing
                , ref missing
                , ref noEncodingDialog
                , ref missing
            );

            return word;
        }

        private bool ReplaceHeading(Application word, int number)
        {
            bool replaceHeading = false;
            string replacement = string.Empty;

            switch (number)
            {
                case 1:
                    replacement = "#";
                    break;
                case 2:
                    replacement = "##";
                    break;
                case 3:
                    replacement = "###";
                    break;
                case 4:
                    replacement = "####";
                    break;
                case 5:
                    replacement = "#####";
                    break;
                case 6:
                    replacement = "######";
                    break;
            }

            object heading = word.ActiveDocument.Styles["Heading " + number];
            object normal = word.ActiveDocument.Styles["Normal"];

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.set_Style(ref heading);

            while (word.Selection.Find.Execute())
            {
                replaceHeading = true;
                word.Selection.Range.InsertBefore(replacement + " ");
                word.Selection.set_Style(ref normal);
                word.Selection.Find.Execute();
            }

            return replaceHeading;
        }

        private bool ReplaceOneBold(Application word)
        {
            object findText = "";
            bool replaceOneBold = false;

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.Font.Bold = 1;
            word.Selection.HomeKey(WdUnits.wdStory);

            while (word.Selection.Find.Execute())
            {
                replaceOneBold = true;
                word.Selection.Text = "**" + word.Selection.Text + "**";
                word.Selection.Font.Bold = 0;
                word.Selection.Find.Execute();
            }

            return replaceOneBold;
        }

        private bool ReplaceOneItalic(Application word)
        {
            object findText = "";
            bool replaceOneItalic = false;

            word.Selection.Find.ClearFormatting();
            word.Selection.Find.Font.Italic = 1;
            word.Selection.HomeKey(WdUnits.wdStory);

            while (word.Selection.Find.Execute())
            {
                replaceOneItalic = true;
                word.Selection.Text = "_" + word.Selection.Text + "_";
                word.Selection.Font.Italic = 0;
                word.Selection.Find.Execute();
            }

            return replaceOneItalic;
        }

        private void ReplaceLists(Application word)
        {
            word.Selection.HomeKey(WdUnits.wdStory);

            for (int i = word.Selection.Document.Paragraphs.Count; i > 0; i--)
            {
                try
                {
                    for (int j = word.Selection.Document.Lists[i].ListParagraphs.Count; j > 0; j--)
                    {
                        Paragraph para = word.Selection.Document.Lists[i].ListParagraphs[j];

                        if (para.Range.ListFormat.ListType == WdListType.wdListBullet)
                        {
                            para.Range.InsertBefore(ListIndent(para.Range.ListFormat.ListLevelNumber, "*"));
                        }

                        if (para.Range.ListFormat.ListType == WdListType.wdListSimpleNumbering ||
                            para.Range.ListFormat.ListType == WdListType.wdListMixedNumbering ||
                            para.Range.ListFormat.ListType == WdListType.wdListListNumOnly)
                        {
                            para.Range.InsertBefore(para.Range.ListFormat.ListValue + ". ");
                        }
                    }

                    word.Selection.Document.Lists[i].Range.InsertParagraphBefore();
                    word.Selection.Document.Lists[i].Range.InsertParagraphAfter();
                    word.Selection.Document.Lists[i].RemoveNumbers();
                }
                catch
                { }
            }
        }

        private string ListIndent(int number, string text)
        {
            string returnValue = "";

            for (int i = 0; i < number - 1; i++)
            {
                returnValue = returnValue + "    ";
            }

            returnValue = returnValue + text + "    ";

            return returnValue;
        }

        private void KillWinWord()
        {
            // Get all running winword processes		 
            List<int> processIds = new List<int>();
            foreach (Process process in Process.GetProcessesByName("winword"))
            {
                process.Kill();
            }

            //foreach (Process process in Process.GetProcessesByName("winword"))
            //{
            //    if (!process.HasExited && !processIds.Contains(process.Id))
            //    {
            //        process.Kill();
            //    }
            //}

        }
    }


}
