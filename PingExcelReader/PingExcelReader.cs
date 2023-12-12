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
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using System.Reflection.Metadata;

namespace PingExcelReader
{
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
                List<Dictionary<string, string>> range;
                try
                {
                    range = reader.ReadReferenceTable("p_extra_data_fields");
                }
                catch (ArgumentException)
                {
                    m_logger?.LogWarning("Cannot find p_extra_data_fields");
                    return new Dictionary<string, dynamic>();
                }

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
                    var customProperties = new Dictionary<string, dynamic>();
                    foreach (var prop in reader.CustomProperties)
                    {
                        customProperties.Add(prop.Key, prop.Value);
                    }

                    m_metadata = new PingExcelMetadata
                    {
                        ClientName = reader.GetCustomDocumentPropertyOrDefault("Ping Client Name", "n/a"),
                        Timestamp = DateTime.Now.ToString("s"),
                        SOVID = reader.GetCustomDocumentPropertyOrDefault("Ping Identifier", "n/a"),
                        FullName = m_infile.FullName,
                        Name = m_infile.Name,
                        DocumentType = "SOV",
                        PingPolicyTermsVersion = reader.GetCustomDocumentPropertyOrDefault("Ping Policy Terms Version", null),
                        PingFormatName = reader.GetCustomDocumentPropertyOrDefault("Ping Format Name", null),
                        Properties = customProperties
                    };
                }

                return m_metadata;
            }
        }

        public PolicyTerms m_policy_terms;
        public PolicyTerms policy_terms
        {
            get
            {
                if (m_policy_terms == null)
                {
                    var version = Metadata.PingPolicyTermsVersion;
                    if (string.IsNullOrEmpty(version))
                        return null;

                    if (!version.StartsWith("v"))
                        throw new Exception("Invalid Ping Policy Terms Version: " + version);

                    var parsedVersionString = version.Substring(1).Split('.').Select(s => int.Parse(s)).ToArray();
                    if (parsedVersionString[0] < 4)
                    {
                        throw new Exception("Invalid Ping Policy Terms Version, currently only support v4+: " + version);
                    }

                    if (!reader.HasNamedRange("p_L1PL"))
                    {
                        throw new Exception("Invalid Ping Policy Terms Version, no p_L1PL defined name found: " + version);
                    }

                    var pingFormatName = Metadata.PingFormatName;

                    m_policy_terms = new PolicyTerms()
                    {
                        layer_terms = ReadLayerTerms(),
                        peril_terms = ReadPerilTerms(),
                        zone_terms = new Dictionary<string, Dictionary<string, ZoneTerms>>()
                    };
                }
                return m_policy_terms;
            }
        }

        private List<LayerTerms> ReadLayerTerms()
        {
            var allLayerTerms = new List<LayerTerms>();
            int layerCounter = 0;
            while (layerCounter < 1000)
            {
                layerCounter += 1;
                m_logger?.LogDebug("Checking layer {layerCounter}...", layerCounter);
                if (!reader.HasNamedRange($"p_L{layerCounter}PL"))
                {
                    m_logger?.LogDebug("Cannot find layer {layerCounter}, stopping.", layerCounter);
                    break;
                }

                decimal? participation;
                try
                {
                    participation = reader.GetCellValue<decimal?>($"p_L{layerCounter}PL");
                    if (participation == null || participation == 0)
                    {
                        m_logger?.LogDebug("Skipping layer {layerCounter}, participation is {dsa}", layerCounter, (string)reader.GetCellValue($"p_L{layerCounter}PL").ToString());
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    m_logger?.LogWarning(ex, "Skipping layer {layerCounter}, participation invalid", layerCounter);
                    continue;
                }

                var layerTerms = new LayerTerms();
                layerTerms.participation_pct = participation ?? 0;
                layerTerms.attachment = reader.GetCellValue<decimal?>($"p_L{layerCounter}AP") ?? 0;
                var limit = reader.GetCellValue($"p_L{layerCounter}LL");
                // decide if limit (which is dynamic) is set or not
                if (limit != null && limit >= 0)
                {
                    layerTerms.limit = Convert.ToInt64(limit);
                }
                else
                {
                    var participation_pct = reader.GetCellValue($"p_L{layerCounter}PP");
                    if (participation_pct == null)
                        participation_pct = 1.0;
                    else
                        participation_pct = Convert.ToDouble(participation_pct);
                    layerTerms.limit = (decimal?)Math.Round(Convert.ToDouble(layerTerms.participation_pct) / participation_pct);
                }

                m_logger?.LogDebug("Layer {layerCounter} details: {layerTerms}", layerCounter, layerTerms.ToJson());
                allLayerTerms.Add(layerTerms);
            }

            return allLayerTerms;
        }

        private Dictionary<string, PerilTerms> ReadPerilTerms()
        {
            // provide a constant with all standard possible perils
            var perilSettings = new Dictionary<string, string>();
            perilSettings.Add("Shake", "EQ_Shake");
            perilSettings.Add("FF", "EQ_Fire");
            perilSettings.Add("SL", "EQ_Sprinkler");
            perilSettings.Add("Landslide", "EQ_Landslide");
            perilSettings.Add("Liquefaction", "EQ_Liquefaction");
            perilSettings.Add("Tsunami", "EQ_Tsunami");
            perilSettings.Add("Wind", "HU_Wind");
            perilSettings.Add("SS", "HU_Surge");
            perilSettings.Add("PF", "HU_PrecipitationFlood");
            perilSettings.Add("WS", "WinterStorm");
            perilSettings.Add("ST", "SevereConvectiveStorm");
            perilSettings.Add("STX", "SevereStorm");
            perilSettings.Add("SW", "StraightLineWind");
            perilSettings.Add("HL", "Hail");
            perilSettings.Add("TD", "Tornado");
            perilSettings.Add("IF", "InlandFlood");
            perilSettings.Add("Wildfire", "Wildfire");
            perilSettings.Add("Terrorism", "Terrorism");
            perilSettings.Add("NonCat", "NonCat");

            var perilGroups = new Dictionary<string, PerilTerms>();

            foreach (var name in perilSettings.Keys)
            {
                if (!reader.HasNamedRange($"p_{name}_Group"))
                    continue;

                string group = Convert.ToString(reader.GetCellValue($"p_{name}_Group"));
                if (string.IsNullOrWhiteSpace(group) || group == "Exclude")
                    continue;

                if (perilGroups.ContainsKey(group))
                {
                    perilGroups[group].subperil_types.Add(perilSettings[name]);
                }
                else
                {
                    var sl = reader.GetCellValue<decimal?>($"p_{name}Sublimit", null) ?? reader.GetCellValue<decimal?>($"p_{name}SubLimit", null);
                    perilGroups.Add(group, new PerilTerms()
                    {
                        subperil_types = new List<string> { perilSettings[name] },
                        sublimit = sl,
                        min_deductible = reader.GetCellValue<decimal?>($"p_{name}Ded", 0),
                        max_deductible = reader.GetCellValue<decimal?>($"p_{name}MaxDed", 0),
                        per_location_deductible = reader.GetCellValue<decimal?>($"p_{name}PerLocDed", null),
                        per_location_deductible_type = reader.GetCellValue<string>($"p_{name}PerLocDedType", null),
                        bi_days_deductible = reader.GetCellValue<decimal?>($"p_{name}BIDed", null)
                    });

                }
            }

            return perilGroups;
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
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
            File.WriteAllText(outfile.FullName, json);
        }
    }
}
