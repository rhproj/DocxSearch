using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace DOC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            string srchWord = tbSearch.Text.ToLower();

            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK) //you only continue with the code if you click OK
            {
                var wApp = new Word.Application();
                wApp.Visible = false;

                listBox1.Items.Clear();//BELOW IS DOC! gives error, without finishing
                string[] files = Directory.GetFiles(fbd.SelectedPath, "*.doc", SearchOption.AllDirectories);
                Cursor.Current = Cursors.WaitCursor;

                foreach (string f in files)
                {
                    if (Path.GetExtension(f) == ".doc")
                    {
                        var wDoc = wApp.Documents.Open(f);
                        if (FindInDoc(srchWord, wDoc))
                        {
                            listBox1.Items.Add(f);
                        }
                        wDoc.Close();
                    }
                }
                wApp.Quit();
                if (listBox1.Items.Count == 0)
                {
                    listBox1.Items.Add("Ничего не найдено :(");
                }
                Cursor.Current = Cursors.Default;
            }
        }

        //Searh method
        bool FindInDoc(string srchWord, Word.Document wDoc)
        {
            bool result = false;
            var range = wDoc.Content;
            range.Find.ClearFormatting();

            if (range.Find.Execute(FindText: srchWord))
            {
                result = true;
            }
            return result;
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                Process.Start(listBox1.SelectedItem.ToString());
            }
        }
    }
}

//Takes an actual file and uses GetPlainText on it's XML elem:
//public static string WDRetrieveText(string filePath)
//{
//    string result = null;

//    using (var doc = WordprocessingDocument.Open(filePath, false))
//    {
//        var docPart = doc.MainDocumentPart;
//        if (docPart != null && docPart.Document != null)
//        {
//            result = GetPlainText(docPart.Document.Body); //type-> OpenXmlElement
//        }
//    }
//    return result.ToLower(); //to resolves case sensetivety problem
//}

//String building from XML element, not patricular docx file yet:
//public static string GetPlainText(OpenXmlElement elem)
//{
//    StringBuilder sb = new StringBuilder();

//    foreach (OpenXmlElement r in elem.Elements())
//    {
//        switch (r.LocalName)
//        {
//            case "cr":
//            case "br":
//                sb.Append(Environment.NewLine);
//                break;
//            case "tab":
//                sb.Append("\t");
//                break;
//            case "t":
//                sb.Append(r.InnerText);
//                break;
//            case "p":
//                sb.Append(GetPlainText(r));
//                sb.Append(Environment.NewLine);
//                break;
//            default:
//                sb.Append(GetPlainText(r));
//                break;
//        }
//    }
//    return sb.ToString();
//}

