using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Aspose.Words;
using Zhuang.Word.AsposeWords;
using Aspose.Words.Saving;

namespace Zhuang.Word.Test
{
    [TestClass]
    public class AsposeWords
    {
        [TestMethod]
        public void TestInsertDocumentAtBookmark()
        {
            Document docA = new Document(@".\Files\a.docx");
            Document docB = new Document(@".\Files\b.docx");

            docA.InsertDocumentAtBookmark("bookmark1", docB);

            docA.Save(@".\c.docx");

        }

        [TestMethod]
        public void TestReplaceDocument()
        {
            Document docA = new Document(@".\Files\a.docx");
            Document docB = new Document(@".\Files\b.docx");

            docA.ReplaceDocument("{zwb}", docB);

            docA.Save(@".\c.docx");

        }

        [TestMethod]
        public void TestSaveAsHtml()
        {
            Document docA = new Document(@".\Files\a.docx");

            docA.Save(@".\c.html", new HtmlSaveOptions()
            {
                PrettyFormat = true,
                UseHighQualityRendering = true,
                ExportImagesAsBase64 = true,
                ExportDocumentProperties = false,
                ExportPageSetup = false,
                ExportHeadersFootersMode = ExportHeadersFootersMode.None,
                ExportDropDownFormFieldAsText = false,
            });

        }


    }
}
