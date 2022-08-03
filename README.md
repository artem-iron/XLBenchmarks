# XLBenchmarks
A set of benchmarks that test IronXL.Excel library against it's previous versions and some of it's competitors.

### How to use the repository
  1. Clone the repository to your machine
  2. Add your license key to ..\XLBenchmarks\XLBenchmarks\appsettings.json file under the "LicenceKey" property, removing "PLACE YOUR KEY HERE" placeholder.
  3. Do the same for ..\XLBenchmarks\XLBenchmarks.Tests\appsettings.json file if you plan to run any tests.
  4. Check which versions of Excel-oriented nuget packages are used in the repository, update or downgrade to your taste/needs.
  5. Run the app
  6. Look for a report under ..\XLBenchmarks\XLBenchmarks\bin\Debug\net6.0\Reports (path is controlled with "ReportsFolder" property in appsettings.json). Every app run will create new report
  7. Look for saved Excel workbooks under ..\XLBenchmarks\XLBenchmarks\bin\Debug\net6.0\Results (path is controlled with "ResultsFolderName" property in appsettings.json). Files will be re-written on each app run.
  
### Notes
  * "PreviousIxlBenchmarkRunner is using the customized assembly of older version of IronXL. It was renamed to IronXLOld.dll, all of it's types' namespaces were renamed from IronXL to IronXLOld, and it is stored in an ..\XLBenchmarks\packages folder. It is added to XLBenchmarks project as an assembly, not a nuget package. To use another version of IronXL as IronXLOld perform similar renaming procedure and replace ..\XLBenchmarks\packages\IronXLOld.dll with your version.
  * Benchmarks are pretty slow, so be prepared for the app to run for several minutes, especially if you are benchmarking old IronXL library.
  * To control number of benchmark runners that are ran during benchmarking process - comment/uncomment dictionary entries in..\XLBenchmarks\XLBenchmarks\Reporting\ReportGenerator.cs file in the method GetTimeTableData().
  * To run Office Interop benchmarks you will need Office installed on your computer.
