using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Zhuang.Word.AsposeWords
{
    public class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        Document _subDoc;

        public InsertDocumentAtReplaceHandler(Document subDoc)
        {
            _subDoc = subDoc;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // Insert a document after the paragraph, containing the match text.
            Paragraph para = (Paragraph)e.MatchNode.ParentNode;
            DocumentUtility.InsertDocument(para, _subDoc);

            // Remove the paragraph with the match text.
            para.Remove();

            return ReplaceAction.Skip;
        }
    }
    //ExEnd:InsertDocumentAtReplaceHandler

}
