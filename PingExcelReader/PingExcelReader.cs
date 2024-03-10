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
                m_logger = loggerFactory?.CreateLogger("PingExcelReader"),
                m_TableName = "Locations"
            };
        }

        private ILogger m_logger = null;

        private FileInfo m_infile;
        private ExcelReader m_reader;

        private string m_TableName;
        public string TableName { get { return m_TableName; } }

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
                        string name = row["Attribute"];
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

        public Dictionary<string, string> subclass_mapping
        {
            get
            {
                var columnSpecs = this.reader.ReadReferenceTable($"r_{this.TableName}_column_specification");

                var dict = new Dictionary<string, string>();
                foreach (var spec in columnSpecs)
                {
                    var col = spec["Col"];
                    var attribute = spec["Attribute"];
                    var props = spec["Props"].Split(new char[] { ',' });

                    if (string.IsNullOrWhiteSpace(attribute)) continue;

                    var colLetter = col.Split(new char[] { '!' }).Last();

                    if (!reader.HasNamedRange($"r_{this.TableName}_subclass_{colLetter}")) continue;

                    var parts = attribute.Split(new char[] { '[' }, 2);
                    if (parts.Length > 1)
                    {
                        attribute = parts[1].Substring(0, parts[1].Length - 1);
                    }
                    dict[attribute] = reader.GetCellValue($"r_{this.TableName}_subclass_{colLetter}");
                }
                return dict;
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

                    var rpt = ReadPerilTerms();
                    m_policy_terms = new PolicyTerms()
                    {
                        layer_terms = ReadLayerTerms(),
                        peril_terms = rpt.Item1,
                        zone_terms = ReadZoneTerms(),
                        excluded_subperil_types = rpt.Item2,
                        notes = reader.GetCellValue<string>("p_Contract_Notes", null),
                    };
                }
                return m_policy_terms;
            }
        }

        private List<LayerTerms> ReadLayerTerms()
        {
            var allLayerTerms = new List<LayerTerms>();
            int layerCounter = 0;
            var TotalTIVLimit = reader.GetCellValue<decimal?>("TotalTIVLimit", null);
            while (layerCounter < 1000)
            {
                layerCounter += 1;
                m_logger?.LogDebug("Checking layer {layerCounter}...", layerCounter);
                if (!reader.HasNamedRange($"p_L{layerCounter}PL"))
                {
                    m_logger?.LogDebug("Cannot find layer {layerCounter}, stopping.", layerCounter);
                    break;
                }

                var limit = reader.GetCellValue<decimal?>($"p_L{layerCounter}LL");
                if (!limit.HasValue)
                    limit = TotalTIVLimit ?? 0;

                var name = reader.GetCellValue<string>($"p_L{layerCounter}Name") ?? layerCounter.ToString();
                var attachment = reader.GetCellValue<decimal?>($"p_L{layerCounter}AP") ?? 0;
                var participation_pct = reader.GetCellValue<decimal?>($"p_L{layerCounter}PP") ?? (decimal)1.0;
                var participation_amt = (decimal?)null;
                // reader.GetCellValue<decimal?>($"p_L{layerCounter}PL");
                var calculated_participation_amt = participation_amt ?? (limit.HasValue ? limit.Value : TotalTIVLimit - Math.Max(0, attachment)) * participation_pct;

                if (!calculated_participation_amt.HasValue || calculated_participation_amt == 0)
                {
                    m_logger?.LogWarning("Skipping layer {layerCounter}, participation empty", layerCounter);
                    continue;
                }

                var layerTerms = new LayerTerms()
                {
                    name = name,
                    limit = limit,
                    attachment = attachment,
                    participation = new PercentOrAmount()
                    {
                        percent = participation_pct,
                        amount = participation_amt
                    },
                    premium = reader.GetCellValue<decimal?>($"p_L{layerCounter}PR"),
                };

                m_logger?.LogDebug("Layer {layerCounter} details: {layerTerms}", layerCounter, layerTerms.ToJson());
                allLayerTerms.Add(layerTerms);
            }

            return allLayerTerms;
        }
        private Dictionary<string, string> GetLegacyVBAPerilNames()
        {
            // provide a constant with all standard possible perils
            var legacyPerilVBAAbbreviations = new Dictionary<string, string>();
            legacyPerilVBAAbbreviations.Add("Shake", "EQ_Shake");
            legacyPerilVBAAbbreviations.Add("FF", "EQ_Fire");
            legacyPerilVBAAbbreviations.Add("SL", "EQ_Sprinkler");
            legacyPerilVBAAbbreviations.Add("Landslide", "EQ_Landslide");
            legacyPerilVBAAbbreviations.Add("Liquefaction", "EQ_Liquefaction");
            legacyPerilVBAAbbreviations.Add("Tsunami", "EQ_Tsunami");
            legacyPerilVBAAbbreviations.Add("Wind", "HU_Wind");
            legacyPerilVBAAbbreviations.Add("SS", "HU_Surge");
            legacyPerilVBAAbbreviations.Add("PF", "HU_PrecipitationFlood");
            legacyPerilVBAAbbreviations.Add("WS", "WinterStorm");
            legacyPerilVBAAbbreviations.Add("ST", "SevereConvectiveStorm");
            legacyPerilVBAAbbreviations.Add("STX", "SevereStorm");
            legacyPerilVBAAbbreviations.Add("SW", "StraightLineWind");
            legacyPerilVBAAbbreviations.Add("HL", "Hail");
            legacyPerilVBAAbbreviations.Add("TD", "Tornado");
            legacyPerilVBAAbbreviations.Add("IF", "InlandFlood");
            return legacyPerilVBAAbbreviations;
        }

        private Tuple<Dictionary<string, PerilTerms>, List<string>> ReadPerilTerms()
        {
            // var perilSettings = ReadPerilSettings();
            var perilsDefinedNames = reader.GetNamedRangesMatchingRegex(@"p_(?<peril>[a-zA-Z]+)_Caption");

            var perilGroups = new Dictionary<string, PerilTerms>();
            var excludedPerils = new List<string>();
            var legacyNames = GetLegacyVBAPerilNames();

            foreach (var perilMatch in perilsDefinedNames)
            {
                var vbaName = perilMatch.Item2.Groups["peril"].Value;
                var pingConstant = legacyNames.GetValueOrDefault(vbaName, vbaName);

                if (!reader.HasNamedRange($"p_{vbaName}_Group"))
                    continue;

                string group = Convert.ToString(reader.GetCellValue($"p_{vbaName}_Group"));
                if (string.IsNullOrWhiteSpace(group) || group == "Exclude")
                {
                    excludedPerils.Add(vbaName);
                    continue;
                }

                PerilTerms pg;
                if (perilGroups.ContainsKey(group))
                {
                    pg = perilGroups[group];
                    perilGroups[group].subperil_types.Add(pingConstant);

                }
                else
                {
                    pg = new PerilTerms()
                    {
                        subperil_types = new List<string> { pingConstant },
                    };
                    perilGroups.Add(group, pg);
                }

                var sl = reader.GetCellValue<decimal?>($"p_{vbaName}Sublimit", null) ?? reader.GetCellValue<decimal?>($"p_{vbaName}SubLimit", null);
                var min_deductible = reader.GetCellValue<decimal?>($"p_{vbaName}Ded", null);
                var max_deductible = reader.GetCellValue<decimal?>($"p_{vbaName}MaxDed", null);
                var location_deductible = reader.GetCellValue<decimal?>($"p_{vbaName}PerLocDed", null);
                var location_deductible_type = reader.GetCellValue<string>($"p_{vbaName}PerLocDedType", null);
                var bi_days_deductible = reader.GetCellValue<decimal?>($"p_{vbaName}BIDed", null);

                if (sl.HasValue)
                    pg.sublimit = sl;
                if (min_deductible.HasValue)
                    pg.min_deductible = min_deductible;
                if (max_deductible.HasValue)
                    pg.max_deductible = max_deductible;
                if (location_deductible.HasValue)
                    pg.location_deductible = location_deductible;
                if (!string.IsNullOrWhiteSpace(location_deductible_type))
                    pg.location_deductible_type = location_deductible_type;
                if (bi_days_deductible.HasValue)
                    pg.bi_days_deductible = bi_days_deductible;
            };

            return new Tuple<Dictionary<string, PerilTerms>, List<string>>(perilGroups, excludedPerils);
        }


        private List<ZoneGroupSettings> ReadZoneGroupsSettings()
        {
            // provide a constant with all standard possible perils
            var zoneGroupDefinedNames = reader.GetNamedRangesMatchingRegex(@"p_(?<group>[a-zA-Z]+)_(?<zone>[a-zA-Z]+)_Caption");
            var zoneGroups = new Dictionary<string, ZoneGroupSettings>();
            foreach (var zoneGroupDefinedName in zoneGroupDefinedNames)
            {
                var zone_group_vba = zoneGroupDefinedName.Item2.Groups["group"].Value;
                var zone_vba = zoneGroupDefinedName.Item2.Groups["zone"].Value;
                var ded_prefix = $"{zone_group_vba}_{zone_vba}_";

                ZoneGroupSettings zoneGroup;
                if (!zoneGroups.ContainsKey(zone_group_vba))
                {
                    zoneGroup = new ZoneGroupSettings()
                    {
                        peril_class = zone_group_vba,
                        zones = new List<ZoneSettings>()
                    };
                    zoneGroups.Add(zone_group_vba, zoneGroup);
                }
                else
                {
                    zoneGroup = zoneGroups[zone_group_vba];
                }

                var zoneSettings = new ZoneSettings()
                {
                    zone = zone_vba,
                    caption = reader.GetCellValue<string>($"p_{ded_prefix}Caption", null)
                };

                zoneGroup.zones.Add(zoneSettings);
            }

            return zoneGroups.Values.ToList();
        }

        private Dictionary<string, Dictionary<string, PerZoneTerms>> ReadZoneTerms()
        {
            var perZoneTerms = new Dictionary<string, Dictionary<string, PerZoneTerms>>();
            var zoneGroupSettings = ReadZoneGroupsSettings();
            foreach (ZoneGroupSettings zoneGroup in zoneGroupSettings)
            {
                foreach (var zoneSettings in zoneGroup.zones)
                {
                    var dedPrefix = $"{zoneGroup.GetVbaName()}_{zoneSettings.GetVbaName()}_";
                    var perZoneTerm = new PerZoneTerms()
                    {
                        sublimit = reader.GetCellValue<decimal?>($"p_{dedPrefix}Sublimit", null),
                        min_deductible = reader.GetCellValue<decimal?>($"p_{dedPrefix}Ded", null),
                        max_deductible = reader.GetCellValue<decimal?>($"p_{dedPrefix}MaxDed", null),
                        location_deductible = reader.GetCellValue<decimal?>($"p_{dedPrefix}PerLocDed", null),
                        location_deductible_type = reader.GetCellValue<string>($"p_{dedPrefix}PerLocDedType", null),
                        is_excluded = reader.GetCellValue<string>($"p_{dedPrefix}Include", "") == "Exclude",
                    };

                    if (!perZoneTerm.IsApplicable())
                        continue;

                    if (!perZoneTerms.ContainsKey(zoneGroup.peril_class))
                        perZoneTerms[zoneGroup.peril_class] = new Dictionary<string, PerZoneTerms>();

                    var zoneName = zoneSettings.GetVbaName();
                    if (zoneSettings.zone.StartsWith("Custom"))
                        zoneName = reader.GetCellValue<string>($"p_{dedPrefix}Caption", null);

                    perZoneTerms[zoneGroup.peril_class][zoneName] = perZoneTerm;
                }
            }

            return perZoneTerms;
        }

        private List<Dictionary<string, dynamic>> m_buildings = null;

        public List<Dictionary<string, dynamic>> buildings
        {
            get
            {
                if (m_buildings == null)
                    m_buildings = reader.ReadItemsTable(this.TableName);
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
