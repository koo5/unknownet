using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace eee
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            WordprocessingDocument doc =
                    WordprocessingDocument.Open("data/doc.docx", true);

            /*DocumentSettingsPart settingsPart = doc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
            // Create object to update fields on open
            UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
            updateFields.Val = new DocumentFormat.OpenXml.OnOffValue(true);
            // Insert object into settings part.
            settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
            settingsPart.Settings.Save();*/


            /*Console.WriteLine("BookmarkStart:");
            foreach (var bs in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                Console.WriteLine(">");
                Console.WriteLine(bs.OuterXml);

            }
            Console.WriteLine("BookmarkEnd:");
            foreach (var bs in doc.MainDocumentPart.RootElement.Descendants<BookmarkEnd>())
            {
                Console.WriteLine(">");
                Console.WriteLine(bs.OuterXml);

            }*/

            //Console.WriteLine("fields:");
            foreach (FieldCode field in doc.MainDocumentPart.RootElement.Descendants<FieldCode>())
            {
                //Console.WriteLine(">");
                //Console.WriteLine(field.Text);
                field.Text = "";
            }

            Console.WriteLine("paragraphs:");
            var paragraphs = doc.MainDocumentPart.Document.Body.OfType<Paragraph>().ToList();
            foreach (var p in paragraphs)
            {
                if (p.ParagraphProperties != null &&
                    p.ParagraphProperties.ParagraphStyleId != null &&
                    p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading1"))
                {

                    //Console.Out.WriteLine(p.ToString());

                    /*var fields = p.Descendants<FieldChar>();

                    foreach (var field in fields)
                    {
                        Console.WriteLine("Hi!");
                        Console.WriteLine(field.InnerXml);
                    }
                    */

                    //Console.Out.WriteLine(p.InnerText);
                    String t = "";
                    foreach (var ch in p.Descendants())
                    {
                        if (ch.InnerText != "")
                        {
                            if (ch is Text)
                                t += ch.InnerText;
                            /*foreach (var a in ch.GetAttributes())
                            {
                                Console.Out.WriteLine(a.LocalName + "=" + a.Value);
                            }
                            */
                        }
                        if (ch is TabChar)
                            t += "<TAB>";
                    }

                    Console.Out.WriteLine(t);

                }
            }
        }
    }
}
