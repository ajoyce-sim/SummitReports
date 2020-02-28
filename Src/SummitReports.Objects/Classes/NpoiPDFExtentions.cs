using FastMember;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;

namespace SummitReports.Objects
{


    public class NPoiPdfDocument : ISummitDocument
    {
        private HtmlDocument document;
        public HtmlDocument Document { get => document; set => document = value; }
        public NPoiPdfDocument(HtmlDocument _document)
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
            document.Text = document.Text.Replace(variableName, columnValue.Replace("\r", "").Replace("\n", "<br/>"));
        }
        public void ReplaceFieldValue(string ColumnName, string valueToSet)
        {
            var variableName = string.Format("%{0}%", ColumnName);
            var columnValue = valueToSet;
            document.Text = document.Text.Replace(variableName, columnValue.Replace("\r", "").Replace("\n", "<br/>"));
        }
    }
}