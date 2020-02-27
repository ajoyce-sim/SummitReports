using FastMember;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;

namespace SummitReports.Objects
{
    public static class NPoiWordExtentions
    {
        public static XWPFDocument ReplaceFieldValue(this XWPFDocument document, DataRow data, string ColumnName)
        {
            return document.ReplaceFieldValue(data, ColumnName, "");
        }
        public static XWPFDocument ReplaceFieldValue(this XWPFDocument document, DataRow data, string ColumnName, string Format)
        {
            var variableName = string.Format("%{0}%", ColumnName);
            var columnValue = data.Value(ColumnName, Format);
            foreach (var item in document.BodyElements)
            {
                if (item.ElementType == BodyElementType.PARAGRAPH)
                {
                    var p = (XWPFParagraph)item;
                    if (p.ParagraphText.Contains(variableName))
                    {
                        if (columnValue.Contains("\n"))
                        {
                            p.ReplaceText(variableName, "");
                            document.CreateParagraphs(p, columnValue.Split('\n'));
                        }
                        else
                        {
                            p.ReplaceText(variableName, columnValue);
                        }
                    }
                }
            }
            return document;
        }
        
        public static void CreateParagraphs(this XWPFDocument document, XWPFParagraph xwpfParagraph, String[] paragraphs)
        {
            if (xwpfParagraph != null)
            {
                for (int i = 0; i < paragraphs.Length; i++)
                {
                    var r = xwpfParagraph.CreateRun();
                    r.SetText(paragraphs[i]);
                    r.AddCarriageReturn();
                }
            }
        }
        
    }
}