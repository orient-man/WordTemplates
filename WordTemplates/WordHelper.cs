using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace WordTemplates
{
    public static class WordHelper
    {
        public static void FillTemplate(
            string templateFile,
            IEnumerable<TemplateField> fields,
            string outputFile)
        {
            _Application word = new Application { Visible = false, ScreenUpdating = false };

            try
            {
                var filename = (Object)new FileInfo(templateFile).FullName;
                _Document doc = word.Documents.Open(ref filename);

                try
                {
                    doc.Activate();

                    new TemplateEngine(new WordTemplateNavigator(word, doc)).SetFields(fields);

                    object outputFileName = outputFile;
                    object fileFormat = WdSaveFormat.wdFormatDocument;

                    doc.SaveAs(ref outputFileName, ref fileFormat);
                }
                finally
                {
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    doc.Close(ref saveChanges);
                }
            }
            finally
            {
                word.Quit();
            }
        }
    }
}