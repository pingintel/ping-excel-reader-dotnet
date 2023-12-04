// Copyright 2023 Ping Data Intelligence
//
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System;
using System.IO;
using System.IO.Compression;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;
using ClosedXML;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;

namespace PingExcelReader
{
    internal class ExcelReader
    {
        private XLWorkbook workbook;
        private ILogger logger;

        public ExcelReader(FileInfo infile, ILogger logger = null)
        {
            this.workbook = new XLWorkbook(infile.FullName);
            this.logger = logger;
        }

        private dynamic ToDynamic(IXLCell cell)
        {
            return cell.DataType switch
            {
                XLDataType.Blank => null,
                XLDataType.Boolean => cell.CachedValue.GetBoolean(),
                XLDataType.Number => cell.CachedValue.GetNumber(),
                XLDataType.Text => cell.CachedValue.GetText(),
                XLDataType.Error => cell.CachedValue.GetError(),
                XLDataType.DateTime => cell.CachedValue.GetDateTime(),
                XLDataType.TimeSpan => cell.CachedValue.GetTimeSpan(),
                _ => throw new InvalidCastException()
            };
        }

        /*
        {
            'value': Cleaned-up data value.
            'confidence': 0.0 (lowest confidence) to 1.0 (highest confidence), indicating a reliability score for the element.
            'original': Cell value as originally read, before any scrubbing operations.
            'source': Provenance of the data.
            'units': Currency of value, if known.
            'comment': Human-readable information about the derivation of the value.
        }
        */
        public class ItemAttribute
        {
            /// <summary>
            /// Cleaned-up data value.
            /// </summary>
            public dynamic value { get; set; }
            // public double confidence { get; set; }
            // public dynamic original { get; set; }
            // public string source { get; set; }
            // public string units { get; set; }
            // public string comment { get; set; }
        }

        public static readonly string[] SIMPLE_FIELDS = new string[] {
            "item_key",
            "building_counter",
            "sheet_name",
            "sheet_row_number",
            "parsing_sheet_name",
            "parsing_sheet_row_number",
            "integration_results",
            "integration_messages",
            "extra",
            "reliability",
            "reliability_reason",
            "orig",
            "zones",
            "external_data",
            "external_data_inputs",
            "internal_data",
            "ping_viewer_url",
            "ping_pli_url"
        };

        public Dictionary<string, dynamic> CustomProperties
        {
            get
            {
                return this.workbook.CustomProperties.ToDictionary(p => p.Name, p => p.Value);
            }
        }

        public List<Dictionary<string, dynamic>> ReadItemsTable(string tableName)
        {
            var columnSpecs = this.ReadReferenceTable($"r_{tableName}_column_specification");

            var table = this.workbook.Table(tableName);
            var range = table.RangeUsed();
            var items = new List<Dictionary<string, dynamic>>();
            foreach (var row in range.Rows().Skip(1))
            {
                var dict = new Dictionary<string, dynamic>();
                foreach (var spec in columnSpecs)
                {
                    var col = spec["Col"];
                    var attribute = spec["Attribute"];
                    var props = spec["Props"].Split(new char[] { ',' });

                    // Console.WriteLine("col: {0} attr: {1}", col, attribute);
                    if (string.IsNullOrWhiteSpace(attribute)) continue;

                    var colLetter = col.Split(new char[] { '!' }).Last();
                    var cell = row.Cell(colLetter);

                    if (cell.IsEmpty()) continue;

                    try
                    {
                        SetAttr(dict, attribute, props, cell);
                    }
                    catch (Exception e)
                    {
                        this.logger?.LogWarning(e, "Error reading attribute {attribute} from cell {cell}", attribute, cell.Address);
                        throw;
                    }
                }

                // if (!dict.ContainsKey("sheet_name")) dict.Add("sheet_name", tableName);
                // if (!dict.ContainsKey("sheet_row_number")) dict.Add("sheet_row_number", row.RowNumber());
                if (!dict.ContainsKey("parsing_sheet_name")) dict.Add("parsing_sheet_name", tableName);
                if (!dict.ContainsKey("parsing_sheet_row_number")) dict.Add("parsing_sheet_row_number", row.RowNumber());

                items.Add(dict);
            }

            return items;
        }

        private void SetAttr(Dictionary<string, dynamic> dict, string attribute, string[] props, IXLCell cell)
        {
            if (SIMPLE_FIELDS.Contains(attribute))
            {
                dict.TryGetValue(attribute, out dynamic item);
                dict.Add(attribute, ToDynamic(cell));
            }
            else if (attribute.Contains('['))
            {
                // 'dictionary' attribute
                var attrsplit = attribute.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                var rootattr = attrsplit[0];
                if (attrsplit.Count() == 1)
                {
                    throw new Exception($"Invalid attribute: {attribute}");
                }
                else if (attrsplit.Count() == 2)
                {
                    dict.TryGetValue(rootattr, out dynamic item);
                    if (item == null)
                    {
                        item = new Dictionary<string, dynamic>();
                        dict.Add(rootattr, item);
                    }
                    item.Add(attrsplit[1], ToDynamic(cell));
                }
                else if (attrsplit.Count() == 3)
                {
                    var subattr = attrsplit[1];
                    var subsubattr = attrsplit[2];
                    dict.TryGetValue(rootattr, out dynamic item);
                    if (item == null)
                    {
                        item = new Dictionary<string, Dictionary<string, dynamic>>();
                        dict.Add(rootattr, item);
                    }
                    item.TryGetValue(subattr, out Dictionary<string, dynamic> subitem);
                    if (subitem == null)
                    {
                        subitem = new Dictionary<string, dynamic>();
                        item.Add(subattr, subitem);
                    }
                    subitem.Add(subsubattr, ToDynamic(cell));
                }
                else
                {
                    throw new Exception($"Invalid attribute: {attribute}");
                }
            }
            else
            {
                // Add 'regular' exploded attribute.
                dict.TryGetValue(attribute, out dynamic item);
                if (item == null)
                {
                    item = new ItemAttribute();
                    dict.Add(attribute, item);
                }

                item.value = ToDynamic(cell);
            }
        }

        public List<Dictionary<string, string>> ReadReferenceTable(string definedName)
        {
            var range = this.workbook.Range(definedName);
            var table = new List<Dictionary<string, string>>();

            if (range == null)
            {
                throw new ArgumentException($"Range {definedName} not found");
            }

            var headers = range.Row(1).RowAbove().Cells().Select(c => c.Value.ToString()).ToArray();

            foreach (var row in range.Rows())
            {
                var dict = new Dictionary<string, string>();
                for (int i = 0; i < headers.Length; i++)
                {
                    dict.Add(headers[i], row.Cell(i + 1).Value.ToString());
                }
                table.Add(dict);
            }
            return table;
        }

        internal dynamic GetCellValue(string cellref)
        {
            return ToDynamic(this.workbook.Cell(cellref));
        }

        internal string GetCustomDocumentPropertyOrDefault(string propertyName, string defaultValue)
        {
            try
            {
                var ret = this.workbook.CustomProperties.CustomProperty(propertyName);
                return ret.Value.ToString();
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        internal bool HasNamedRange(string definedName)
        {
            return this.workbook.NamedRange(definedName) != null;
        }
    }
}