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
    using Currency = int;

    public class LayerTerms
    {
        public string name { get; set; }
        public Currency? limit { get; set; }
        public Currency? attachment { get; set; }
        public float? participation { get; set; }
        public Currency? premium { get; set; }
    }

    public class PerilTerms
    {
        public string type { get; set; }
        public string subperil_group { get; set; }
        public int? sublimit { get; set; }
        public int? min_deductible { get; set; }
        public int? max_deductible { get; set; }
        public string deductible_type { get; set; }
        public string per_location_deductible_type { get; set; }
        public float? per_location_deductible { get; set; }
        public int? bi_days_deductible { get; set; }
    }

    public class ZoneTerms
    {
        public string peril_type { get; set; }
        public string zone { get; set; }
        public int? sublimit { get; set; }
        public int? min_deductible { get; set; }
        public int? max_deductible { get; set; }
        public string deductible_type { get; set; }
        public string per_location_deductible_type { get; set; }
        public float? per_location_deductible { get; set; }
        public bool? is_excluded { get; set; }
    }


    public class PolicyTerms
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
        public List<PerilTerms> peril_terms { get; set; }
        public List<ZoneTerms> zone_terms { get; set; }
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

        public Dictionary<string, string> Properties;

        public string PingPolicyTermsVersion;

        public object PingFormatName;
        // public string UserInfo;
    }
}
