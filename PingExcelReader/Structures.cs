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
    public class SerializableBase
    {
        public string ToJson()
        {
            return JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
        }
    }

    public class PercentOrAmount : SerializableBase
    {
        public decimal? percent { get; set; }
        public decimal? amount { get; set; }
    }


    public class LayerTerms : SerializableBase
    {
        public string name { get; set; }
        public decimal? limit { get; set; }
        public decimal? attachment { get; set; }
        public PercentOrAmount participation { get; set; }
        public decimal? premium { get; set; }
    }

    public class PerilTerms : SerializableBase
    {
        public List<string> subperil_types { get; set; }
        public decimal? sublimit { get; set; }
        public decimal? min_deductible { get; set; }
        public decimal? max_deductible { get; set; }
        public string location_deductible_type { get; set; }
        public decimal? location_deductible { get; set; }
        public decimal? bi_days_deductible { get; set; }
    }

    public class PerZoneTerms : SerializableBase
    {
        public string peril_type { get; set; }
        public string zone { get; set; }
        public decimal? sublimit { get; set; }
        public decimal? min_deductible { get; set; }
        public decimal? max_deductible { get; set; }
        public string location_deductible_type { get; set; }
        public decimal? location_deductible { get; set; }
        public bool? is_excluded { get; set; }

        private bool IsDecimalSet(decimal? value)
        {
            return value.HasValue && value.Value != 0;
        }

        internal bool IsApplicable()
        {
            if (this.sublimit.HasValue
                || IsDecimalSet(this.location_deductible)
                || IsDecimalSet(this.min_deductible)
                || IsDecimalSet(this.max_deductible)
                || (this.is_excluded.HasValue && this.is_excluded.Value))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }


    public class PolicyTerms : SerializableBase
    {
        public List<LayerTerms> layer_terms { get; set; }
        public Dictionary<string, PerilTerms> peril_terms { get; set; }
        public Dictionary<string, Dictionary<string, PerZoneTerms>> zone_terms { get; set; }
        public List<string> excluded_subperil_types;
        public string notes;

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

        public Dictionary<string, dynamic> Properties;

        public string PingPolicyTermsVersion;

        public object PingFormatName;
        // public string UserInfo;
    }

    internal class ZoneGroupSettings
    {
        public string caption { get; set; }
        public string peril_class { get; set; }
        public List<ZoneSettings> zones { get; set; }

        public string GetVbaName()
        {
            return peril_class ?? caption;
        }

    }

    internal class ZoneSettings
    {
        public string caption { get; set; }
        public string zone { get; set; }

        public string GetVbaName()
        {
            return zone;
        }
    }

}