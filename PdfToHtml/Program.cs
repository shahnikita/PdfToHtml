using Bytescout.PDF2HTML;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Xml;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;

namespace PdfToHtml
{

    class Program
    {


        static void Main(string[] args)
        {

            string appDirectory = System.IO.Path.GetDirectoryName(Environment.CurrentDirectory);
            string inFileName = appDirectory + @"\Files\demo1.pdf";
            string outFileNameViaWordApp = appDirectory + @"\Files\WordApp.html";


            byte[] inFileContent = File.ReadAllBytes(inFileName);
            string ext = System.IO.Path.GetExtension(inFileName);

            string Filename = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ext;



            var htmltest = ViaWordApp(inFileContent, Filename);
            System.IO.File.WriteAllText(outFileNameViaWordApp, htmltest);

            #region pdfbox
            string outFileNameViapdfbox = appDirectory + @"\Files\pdfbox.html";
            PDDocument doc = null;
            doc = PDDocument.load(inFileName);
            PDFTextStripper textstrip = new PDFTextStripper();
            string strPDFText = textstrip.getText(doc);

            doc.close();
            System.IO.File.WriteAllText(outFileNameViapdfbox, strPDFText);
            #endregion

            #region parser
            string outFileNameViaparser = appDirectory + @"\Files\parser.html";
            PDFParser pp = new PDFParser();
            pp.ExtractText(inFileName, outFileNameViaparser);

            #endregion parser

            #region itextsharp
            string outFileNameViaitextsharp = appDirectory + @"\Files\itextsharp.html";
            PdfReader reader = new PdfReader(inFileName);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            System.IO.File.WriteAllText(outFileNameViaitextsharp, text);
            #endregion itextsharp

            #region Bytescout
            string outFileNameViaBytescout = appDirectory + @"\Files\Bytescout.html";
            // Create Bytescout.PDF2HTML.HTMLExtractor instance
            HTMLExtractor extractor = new HTMLExtractor();
            extractor.RegistrationName = "demo";
            extractor.RegistrationKey = "demo";

            // Set HTML with CSS extraction mode
            extractor.ExtractionMode = HTMLExtractionMode.HTMLWithCSS;

            // Load sample PDF document
            extractor.LoadDocumentFromFile(inFileName);

            // Save extracted HTML to file
            extractor.SaveHtmlToFile(outFileNameViaBytescout);

            // Open output file in default associated application
            System.Diagnostics.Process.Start(outFileNameViaBytescout);
            #endregion Bytescout

            #region pdtohtml.exe
            ProcessStartInfo startInfo = new ProcessStartInfo();
            string outFileNameViapdtohtmlexe = appDirectory + @"\Files\pdtohtmlexe.html";
            
            
            //Set the PDF File Path and HTML File Path as arguments.
            startInfo.Arguments = string.Format("{0} {1}", inFileName, outFileNameViapdtohtmlexe);

            //Set the Path of the PdfToHtml exe file.
            startInfo.FileName = appDirectory + @"\Files\pdftohtml.exe";

            //Hide the Command window.
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = false;

            //Execute the PdfToHtml exe file.
            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();
            }
            #endregion pdtohtml.exe
        }


        public static string ViaWordApp(byte[] inFileContent, string inFilename)
        {
            try
            {

                string temphtmlFile = Guid.NewGuid().ToString();
                string htmlFilePath = System.IO.Path.GetTempPath() + temphtmlFile + ".html";

                File.WriteAllBytes(inFilename, inFileContent);


                //  object nullobj = System.Reflection.Missing.Value;
                object documentFormat = 8;

                Application wordApp = new Application();
                wordApp.Visible = false;

                Document document = new Document();
                document = wordApp.Documents.Open(inFilename);


                //document = wordApp.ActiveDocument;

                Range t = document.Content.FormattedText;
                string str = t.Text;
                string DocConvert = document.Content.Text;

                document.SaveAs(htmlFilePath, ref documentFormat);

                ((_Document)document).Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                ((_Application)wordApp).Quit();



                //Replace html image src to url
                string html = System.IO.File.ReadAllText(htmlFilePath);

                html = AutoCloseHtmlTags(html, temphtmlFile);


                System.IO.File.WriteAllText(htmlFilePath, html);
                return html;

            }
            catch (Exception)
            {
                foreach (var process in Process.GetProcessesByName("WINWORD"))
                {
                    process.Kill();
                }
                return null;
            }
            finally
            {
                if (File.Exists(inFilename))
                    File.Delete(inFilename);
            }

        }


        public static string AutoCloseHtmlTags(string inputHtml, string temphtmlFile)
        {
            HtmlDocument htmlDoc = new HtmlDocument();

            htmlDoc.LoadHtml(inputHtml);

            htmlDoc.OptionWriteEmptyNodes = true;
            var images = htmlDoc.DocumentNode.Descendants("img");
            byte[] imgContent = null;
            string extension = string.Empty;

            string currentDirectory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), temphtmlFile + "_Files");
            foreach (HtmlNode img in images)
            {
                var att = img.GetAttributeValue("src", null);
                if (att != null)
                {
                    imgContent = File.ReadAllBytes(System.IO.Path.GetTempPath() + att);
                    extension = System.IO.Path.GetExtension(att).Replace(".", "");
                    string base64 = Convert.ToBase64String(imgContent);
                    string imgSrc = String.Format("data:image/{0};base64,{1}", extension, base64);
                    img.SetAttributeValue("src", imgSrc);
                }
            }

            string html = RemoveTroublesomeCharacters(htmlDoc.DocumentNode.OuterHtml);

            html = Regex.Replace(html, "[^ -~]", "");
            html.Replace("&nbsp;", " ");
            return html;
        }

        public static string RemoveTroublesomeCharacters(string inString)
        {
            if (inString == null) return null;

            StringBuilder newString = new StringBuilder();
            char ch;

            for (int i = 0; i < inString.Length; i++)
            {

                ch = inString[i];
                // remove any characters outside the valid UTF-8 range as well as all control characters
                // except tabs and new lines
                //if ((ch < 0x00FD && ch > 0x001F) || ch == '\t' || ch == '\n' || ch == '\r')
                //if using .NET version prior to 4, use above logic
                if (XmlConvert.IsXmlChar(ch)) //this method is new in .NET 4
                {
                    newString.Append(ch);
                }
            }
            return newString.ToString();

        }

    }
}
