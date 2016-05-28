using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using System.IO;
using System.Text.RegularExpressions;

namespace Zhuang.Word.AsposeWords
{
    public static class DocumentExtensions
    {
        public static void AppendDocument(this Document doc, Document srcDoc)
        {
            doc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;
            doc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        public static void ReplaceText(this Document doc, string oldValue, string newValue)
        {
            doc.Range.Replace(oldValue, newValue, true, false);
        }

        public static void ReplaceDocument(this Document doc, string oldValue, Document newValue)
        {
            doc.Range.Replace(new Regex(oldValue), new InsertDocumentAtReplaceHandler(newValue), false);
        }

        public static void InsertDocumentAtBookmark(this Document doc, string bookmarkName, Document srcDoc)
        {
            //ExStart:InsertDocumentAtBookmark         
            Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
            DocumentUtility.InsertDocument(bookmark.BookmarkStart.ParentNode, srcDoc);
        }

        public static bool IsEmpty(this Document doc)
        {
            if (string.IsNullOrWhiteSpace(doc.GetText()))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static MemoryStream GetMemoryStream(this Document doc,SaveFormat saveFormat)
        {
            MemoryStream ms = new MemoryStream();
            doc.Save(ms, saveFormat);
            ms.Seek(0, SeekOrigin.Begin);
            return ms;
        }

        public static byte[] GetBytes(this Document doc, SaveFormat saveFormat)
        {
            var ms = doc.GetMemoryStream(saveFormat);
            byte[] buffer = new byte[ms.Length];
            ms.Read(buffer, 0, buffer.Length);
            ms.Seek(0, SeekOrigin.Begin);
            return buffer;
        }
    }
}
