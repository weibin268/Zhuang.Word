using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using System.IO;

namespace Zhuang.Word
{
    public static class DocumentExtensions
    {
        public static void AppendDocument(this Document doc, Document srcDoc)
        {
            doc.FirstSection.PageSetup.SectionStart = Aspose.Words.SectionStart.Continuous;
            doc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        public static void Replace(this Document doc, string oldValue, string newValue)
        {
            doc.Range.Replace(oldValue, newValue, true, false);
        }
    }
}
