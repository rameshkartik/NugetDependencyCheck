// See https://aka.ms/new-console-template for more information
using Config2AppSettings;
string path = @"C:\\SourceCode\\poc_net8.0_m2madapter\\LandisGyr.CellularNMS.NetworkMesssageProcessor.HostCore\Web.config";
string appSettingsJsonpath = "C:\\SourceCode\\Config2AppSettings\\appSettings.json";
Console.WriteLine("Hello, World!");

ConverterLibrary converterLibrary = new ConverterLibrary(path, appSettingsJsonpath);

converterLibrary.Convert2AppSettings();

Console.WriteLine();

