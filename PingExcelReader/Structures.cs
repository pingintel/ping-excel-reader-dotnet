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


    public class LayerTerms : SerializableBase
    {
        public string name { get; set; }
        public decimal? limit { get; set; }
        public decimal? attachment { get; set; }
        public decimal? participation_pct { get; set; }
        public decimal? premium { get; set; }
    }

    public class PerilTerms : SerializableBase
    {
        public List<string> subperil_types { get; set; }
        public decimal? sublimit { get; set; }
        public decimal? min_deductible { get; set; }
        public decimal? max_deductible { get; set; }
        public string deductible_type { get; set; }
        public string per_location_deductible_type { get; set; }
        public decimal? per_location_deductible { get; set; }
        public decimal? bi_days_deductible { get; set; }
    }

    public class ZoneTerms : SerializableBase
    {
        public string peril_type { get; set; }
        public string zone { get; set; }
        public decimal? sublimit { get; set; }
        public decimal? min_deductible { get; set; }
        public decimal? max_deductible { get; set; }
        public string deductible_type { get; set; }
        public string per_location_deductible_type { get; set; }
        public float? per_location_deductible { get; set; }
        public bool? is_excluded { get; set; }
    }


    public class PolicyTerms : SerializableBase
    {
        public string tracking_id { get; set; }
        public string policy_number { get; set; }
        public string insured_name { get; set; }
        public DateOnly? inception_date { get; set; }
        public DateOnly? expiration_date { get; set; }
        public string underwriter { get; set; }
        public string line_of_business { get; set; }
        public string currency { get; set; }
        public bool? include_surge_as_sublimit { get; set; }
        public string air_date_format { get; set; }
        public List<LayerTerms> layer_terms { get; set; }
        public Dictionary<string, PerilTerms> peril_terms { get; set; }
        public Dictionary<string, Dictionary<string, ZoneTerms>> zone_terms { get; set; }
        public Dictionary<string, List<string>> following_perils { get; set; }
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
}
