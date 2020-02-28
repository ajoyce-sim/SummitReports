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
    public class NPoiWordDocument : ISummitDocument
    {
        private XWPFDocument document;
        public XWPFDocument Document { get => document; set => document = value; }

        public NPoiWordDocument(XWPFDocument _document)
        {
            document = _document;
        }
        public void ReplaceFieldValue(DataRow data, string ColumnName)
        {
            ReplaceFieldValue(data, ColumnName, "");
        }
        public void ReplaceFieldValue(DataRow data, string ColumnName, string Format)
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
                            CreateParagraphs(document, p, columnValue.Split('\n'));
                        }
                        else
                        {
                            p.ReplaceText(variableName, columnValue);
                        }
                    }
                }
            }

        }
        public void ReplaceFieldValue(string ColumnName, string valueToSet)
        {
            var variableName = string.Format("%{0}%", ColumnName);
            var columnValue = valueToSet;
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
                            CreateParagraphs(document, p, columnValue.Split('\n'));
                        }
                        else
                        {
                            p.ReplaceText(variableName, columnValue);
                        }
                    }
                }
            }
        }
        public void CreateParagraphs(XWPFDocument document, XWPFParagraph xwpfParagraph, String[] paragraphs)
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