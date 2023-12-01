// Copyright 2023 Ping Data Intelligence
//
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Microsoft.Extensions.Logging;
using System.Text;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PingExcelReader
{
    using Currency = System.Decimal;

    public class Layer
    {
        public string layer_name { get; set; }
        public Currency limit { get; set; }
        public Currency attachment { get; set; }
        public Currency premium { get; set; }
    }


    public class PolicyTerms
    {

        public List<Layer> layers { get; set; }
    }


    internal class PingExcelMetadata
    {
        public string ClientName;
        public string Timestamp;
        //public string Token; 
        public string SOVID;
        // public string Version;
        // public string FileFormat;
        public string Name;
        // public string CodeName;
        public string FullName;
        public string DocumentType;
        // public string UserInfo;
    }


    public class PingExcelReader
    {
        public static PingExcelReader Read(FileInfo infile, ILoggerFactory loggerFactory = null)
        {
            return new PingExcelReader()
            {
                m_infile = infile,
                m_logger = loggerFactory?.CreateLogger("PingExcelReader")
            };
        }

        private ILogger m_logger = null;

        private FileInfo m_infile;
        private ExcelReader m_reader;

        private ExcelReader reader
        {
            get
            {
                if (m_reader == null)
                {
                    m_reader = new ExcelReader(m_infile, m_logger);
                }
                return m_reader;

            }
        }

        public string id { get { return Metadata.SOVID; } }
        public string source_filename { get { return Metadata.Name; } }
        public int num_buildings { get { return buildings?.Count ?? 0; } }
        public Dictionary<string, dynamic> extra_data
        {
            get
            {
                var range = reader.ReadReferenceTable("p_extra_data_fields");
                var ret = new Dictionary<string, dynamic>();

                foreach (var row in range)
                {
                    try
                    {
                        string name = row["Label"];
                        string cellref = row["Excel Defined Name"];
                        var cellValue = reader.GetCellValue(cellref);
                        if (string.IsNullOrEmpty(cellValue)) continue;
                        ret.Add(name, cellValue);
                    }
                    catch (Exception ex)
                    {
                        m_logger?.LogWarning(ex, "Error reading extra data field");
                    }
                }

                return ret;
            }
        }

        private PingExcelMetadata m_metadata = null;
        private PingExcelMetadata Metadata
        {
            get
            {
                if (m_metadata == null)
                {
                    m_metadata = new PingExcelMetadata
                    {
                        ClientName = reader.GetCustomDocumentPropertyOrDefault("Ping Client Name", "n/a"),
                        Timestamp = DateTime.Now.ToString("s"),
                        SOVID = reader.GetCustomDocumentPropertyOrDefault("Ping Identifier", "n/a"),
                        FullName = m_infile.FullName,
                        Name = m_infile.Name,
                        DocumentType = "SOV"
                    };
                }

                return m_metadata;
            }
        }

        public PolicyTerms policy_terms
        {
            get
            {
                return new PolicyTerms()
                {
                    layers = new List<Layer>()
                    {
                        new Layer()
                        {
                            layer_name = "Layer 1",
                            limit = 1000000,
                            attachment = 0,
                            premium = 1000000
                        }
                    }
                };
            }
        }

        private List<Dictionary<string, dynamic>> m_buildings = null;

        public List<Dictionary<string, dynamic>> buildings
        {
            get
            {
                if (m_buildings == null)
                    m_buildings = reader.ReadItemsTable("Locations");
                return m_buildings;
            }
        }

        public void WritePingJson(FileInfo outfile)
        {
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(outfile.FullName, json);
        }
    }
}
