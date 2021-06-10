using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;

namespace MSWordSearchOpenXMLSDK0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(tbSearch.Text))
            {
                string srchWord = tbSearch.Text.ToLower();

                FolderBrowserDialog fbd = new FolderBrowserDialog();

                if (fbd.ShowDialog() == DialogResult.OK) 
                {
                    listBox1.Items.Clear();
                    string[] files = Directory.GetFiles(fbd.SelectedPath, "*.doc", SearchOption.AllDirectories);
                    Cursor.Current = Cursors.WaitCursor;
                    var wApp = new Word.Application();
                    wApp.Visible = false;

                    try
                    {
                        foreach (string f in files)
                        {
                            if (Path.GetExtension(f) == ".docx")
                            {
                                if (DocXSearch(f, srchWord))
                                {
                                    listBox1.Items.Add(f);
                                }
                            }
                            else
                            {
                                var wDoc = wApp.Documents.Open(f);
                                if (DocSearch(wDoc, srchWord))
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
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Закройте все открытые документы и повторите поиск!");
                        wApp.Quit();
                    }
                    Cursor.Current = Cursors.Default;
                }
            }
            else
            {
                MessageBox.Show("Введите слово для поиска!");
            }
        }


        //Takes an actual file and uses GetPlainText on it's XML elem:
        string DocXRetrieveText(string filePath)
        {
            string result = null;

            using (var doc = WordprocessingDocument.Open(filePath, false))
            {
                var docPart = doc.MainDocumentPart;
                if (docPart != null && docPart.Document != null)
                {
                    result = GetPlainText(docPart.Document.Body); //type-> OpenXmlElement
                }
            }
            return result.ToLower();
        }

        //String building from XML element, not a particular docx file yet:
        string GetPlainText(OpenXmlElement elem)
        {
            StringBuilder sb = new StringBuilder();

            foreach (OpenXmlElement r in elem.Elements())
            {
                switch (r.LocalName)
                {
                    case "cr":
                    case "br":
                        sb.Append(Environment.NewLine);
                        break;
                    case "tab":
                        sb.Append("\t");
                        break;
                    case "t":
                        sb.Append(r.InnerText);
                        break;
                    case "p":
                        sb.Append(GetPlainText(r));
                        sb.Append(Environment.NewLine);
                        break;
                    default:
                        sb.Append(GetPlainText(r));
                        break;
                }
            }
            return sb.ToString();
        }


        /// <summary>Searh method for Docx:</summary>
        bool DocXSearch(string filePath, string srchWord)
        {
            bool result = false;
            if (DocXRetrieveText(filePath).Contains(srchWord))
            {
                result = true;
            }
            return result;
        }

        /// <summary>Searh method for Doc:</summary>
        bool DocSearch(Word.Document wDoc, string srchWord)
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

        /// <summary>Opens selected document</summary>
        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                Process.Start(listBox1.SelectedItem.ToString());
            }
        }


        private void btSearch_MouseEnter(object sender, EventArgs e)
        {
            btSearch.ForeColor = Color.Yellow;
        }

        private void btSearch_MouseLeave(object sender, EventArgs e)
        {
            btSearch.ForeColor = Color.LightYellow;
        }
    }
}
