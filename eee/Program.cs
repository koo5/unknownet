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
                    WordprocessingDocument.Open("doc.docx", true);

            /*DocumentSettingsPart settingsPart = doc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
            // Create object to update fields on open
            UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
            updateFields.Val = new DocumentFormat.OpenXml.OnOffValue(true);
            // Insert object into settings part.
            settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
            settingsPart.Settings.Save();*/


            var paragraphs = doc.MainDocumentPart.Document.Body.OfType<Paragraph>().ToList();
            foreach (var p in paragraphs)
            {
                if (p.ParagraphProperties != null &&
                    p.ParagraphProperties.ParagraphStyleId != null &&
                    p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading1"))
                {

                    Console.Out.WriteLine(p.ToString());

                    var fields = p.Descendants<FieldChar>();

                    foreach (var field in fields)
                    {
                        Console.WriteLine("Hi!");
                        Console.WriteLine(field.InnerXml);
                    }


                    Console.Out.WriteLine(p.InnerText);
                    Console.Out.WriteLine(p.InnerXml);

                }
            }
        }
    }
}
