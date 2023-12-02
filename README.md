# ping-excel-reader-dotnet
Ping Data Intelligence .NET Excel Reader

This tool provides a simple .NET SDK for reliably reading [Ping Data Intelligence](https://www.pingintel.com)-generated Excel SOV files.  

To use:
```
var pingData = PingExcelReader.PingExcelReader.Read(new FileInfo("scrubber.xlsx"));
pingData.WritePingJson(outfile);
```

Alternatively, data can be accessed directly from the PingExcelReader object via lazy-loaded properties such as 


## To test:
```
cd pingreader
dotnet run --infile <path to Ping-generated Excel file>
```