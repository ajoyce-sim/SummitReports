using FastMember;
using HtmlAgilityPack;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;

namespace SummitReports.Objects
{
    public static class NPoiPdfExtentions
    {
        public static HtmlDocument ReplaceFieldValue(this HtmlDocument document, DataRow data, string ColumnName)
        {
            return document.ReplaceFieldValue(data, ColumnName, "");
        }
        public static HtmlDocument ReplaceFieldValue(this HtmlDocument document, DataRow data, string ColumnName, string Format)
        {
            var variableName = string.Format("%{0}%", ColumnName);
            var columnValue = data.Value(ColumnName, Format);
            document.Text = document.Text.Replace(variableName, columnValue.Replace("\r", "").Replace("\n", "<br/>"));
            return document;
        }
        public static HtmlDocument ReplaceFieldValue(this HtmlDocument document, string ColumnName, string valueToSet)
        {
            var variableName = string.Format("%{0}%", ColumnName);
            var columnValue = valueToSet;
            document.Text = document.Text.Replace(variableName, columnValue.Replace("\r", "").Replace("\n", "<br/>"));
            return document;
        }
    }
}