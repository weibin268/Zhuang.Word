using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using System.IO;

namespace Zhuang.Word.AsposeWords
{
    public static class DocumentExtensions
    {
        public static void AppendDocument(this Document doc, Document srcDoc)
        {
            doc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
            doc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        public static void Replace(this Document doc, string oldValue, string newValue)
        {
            doc.Range.Replace(oldValue, newValue, true, false);
        }

        public static void InsertDocumentAtBookmark(this Document doc, string bookmarkName, Document srcDoc)
        {
            //ExStart:InsertDocumentAtBookmark         
            Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
            DocumentUtility.InsertDocument(bookmark.BookmarkStart.ParentNode, srcDoc);
        }
    }
}
