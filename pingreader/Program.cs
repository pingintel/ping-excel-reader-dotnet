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
            using ILoggerFactory loggerFactory =
                LoggerFactory.Create(builder =>
                    builder.AddSimpleConsole(options =>
                    {
                        options.IncludeScopes = true;
                        options.SingleLine = true;
                        options.TimestampFormat = "HH:mm:ss ";
                    }));

            ILogger logger = loggerFactory.CreateLogger("pingreader");
            logger.LogInformation("pingreader: Processing {infile}", infile.FullName);
            var pingData = PingExcelReader.PingExcelReader.Read(infile, loggerFactory);
            logger.LogInformation("SOVID: {SOVID}", pingData.id);
            var information = pingData.extra_data;
            if (information.TryGetValue("Named Insured", out dynamic? quoteCode))
                if (quoteCode.HasValue)
                    logger.LogInformation("Named Insured: {Named Insured}", (string)quoteCode.Value.ToString());
            logger.LogInformation("Info Count: {Count}", information.Count);
            // foreach (var item in information)
            //     Console.WriteLine("{0}: {1}", item.Key, item.Value);

            var buildings = pingData.buildings;
            string json = JsonSerializer.Serialize(buildings, new JsonSerializerOptions { WriteIndented = true });

            pingData.WritePingJson(outfile);
            logger.LogInformation("Wrote {0}", outfile.FullName);
        }
    }
}