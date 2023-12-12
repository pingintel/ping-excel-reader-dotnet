using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace pingreader
{
    class Program
    {
        static void Main(FileInfo infile, FileInfo? outfile = null)
        {
            if (outfile == null)
            {
                outfile = new FileInfo(infile.FullName + ".output.json");
            }
            using ILoggerFactory loggerFactory = LoggerFactory.Create(builder =>
            {
                builder.SetMinimumLevel(LogLevel.Debug);
                builder.AddSimpleConsole(options =>
                {
                    options.IncludeScopes = true;
                    options.SingleLine = true;
                    options.TimestampFormat = "HH:mm:ss ";
                });
            });

            ILogger logger = loggerFactory.CreateLogger("pingreader");
            logger.LogInformation("pingreader: Processing {infile}", infile.FullName);

            var pingData = PingExcelReader.PingExcelReader.Read(infile, loggerFactory);

            logger.LogInformation("SOVID: {SOVID}", pingData.id);

            try
            {
                var namedInsured = pingData.extra_data["Named Insured"];
                string namedInsuredStr = Convert.ToString(namedInsured);
                logger.LogInformation("Named Insured: {Named Insured}", namedInsuredStr);
            }
            catch (KeyNotFoundException)
            {
                logger.LogInformation("Named Insured: <not found>");
            }
            if (pingData.policy_terms != null)
            {
                foreach (var layer in pingData.policy_terms.layer_terms)
                {
                    logger.LogInformation("Layer details: {LayerDetails}", layer.ToJson());
                }
                foreach (var peril in pingData.policy_terms.peril_terms)
                {
                    logger.LogInformation("Peril {Key}: {Details}", peril.Key, peril.Value.ToJson());
                }
                foreach (var zoneGroup in pingData.policy_terms.zone_terms)
                {
                    foreach (var zone in zoneGroup.Value)
                    {
                        logger.LogInformation("Zone {GroupKey}-{Key}: {ZoneDetails}", zoneGroup.Key, zone.Key, zone.Value.ToJson());
                    }
                }

                logger.LogInformation("ExtraData Count: {Count}", pingData.extra_data.Count);

                var buildings = pingData.buildings;
                logger.LogInformation("Read Buildings Count: {Count}", buildings.Count);
                pingData.WritePingJson(outfile);
                logger.LogInformation("Wrote {0}", outfile.FullName);
            }
        }
    }
}