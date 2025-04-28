using Newtonsoft.Json.Linq;
using NuGet.Common;
using NuGet.Packaging.Core;
using NuGet.Protocol;
using NuGet.Protocol.Core.Types;
using NuGet.Versioning;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using ExcelReadUpdate.Library;

namespace DependencyTreeProject
{
    public partial class Form1 : Form
    {
        List<string> references = new List<string>();
        List<KeyValuePair<string, string>> currentVersions = new List<KeyValuePair<string, string>>();

        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string packageName = "Newtonsoft.Json"; // Replace with the package name
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Project Files (*.csproj)|*.csproj|All Files (*.*)|*.*",
                    Title = "Select a .csproj file"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    textBox1.Text = selectedFilePath;

                    // Do something with the selected file path
                    // For example, you can load the content of the project file.
                    // Load the .csproj file
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(selectedFilePath);

                    // Query for references
                    GetReferencesFromCsproj(xmlDocument);
                    //XmlDocument xmlPackageConfigDoc = new XmlDocument();
                    //xmlPackageConfigDoc.Load("C:\\Users\\JadhavR\\source\\repos\\DependencyTreeProject\\DependencyTreeProject\\packages.config");
                    //GetVersionsFromPackagesConfig(xmlPackageConfigDoc, packageName);

                    // Display or use the references
                    LoadGridViewData();
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void GetReferencesFromCsproj(XmlDocument xmlDocument)
        {
            // Assuming references are in the <Reference> elements within <ItemGroup>
            //XmlNodeList referenceNodes = xmlDocument.SelectNodes("//ItemGroup/Reference");
            //XmlNodeList referenceNodes = xmlDocument.SelectNodes("//*[local-name()='ItemGroup']/*[local-name()='Reference']");
            XmlNodeList referenceNodes = xmlDocument.SelectNodes("//*[local-name()='ItemGroup']/*[local-name()='PackageReference']");
            if (referenceNodes != null)
            {
                foreach (XmlNode referenceNode in referenceNodes)
                {
                    string referenceName = referenceNode.Attributes?["Include"]?.Value;
                    string currentVersion = referenceNode.Attributes?["Version"]?.Value;
                    if (!string.IsNullOrEmpty(referenceName))
                    {
                        //references.Add(referenceName);
                        //listView1.Items.Add(new ListViewItem(referenceName));
                        references.Add(referenceName);
                        currentVersions.Add(new KeyValuePair<string, string>(referenceName, currentVersion));
                    }
                }
            }

        }
        private string GetVersionsFromPackagesConfig(XmlDocument xmlDocument, string packageName)
        {
            XmlNodeList packagesNodes = xmlDocument.SelectNodes("//*[local-name()='packages']/*[local-name()='package']");

            if (packagesNodes != null)
            {
                foreach (XmlNode packageNode in packagesNodes)
                {
                    //string versionValue = packageNode.Attributes["version"].Value;
                    //string packageValue = packageNode.Attributes["id"].Value;
                    //if (!string.IsNullOrEmpty(versionValue))
                    //{
                    //    //references.Add(referenceName);
                    //    //listView1.Items.Add(new ListViewItem(referenceName));
                    //    //references.Add(versionValue);
                    //    if (packageName == packageValue)
                    //    {
                    //        var list = new List<KeyValuePair<string, string>>();
                    //        list.Add(new KeyValuePair<string, string>(packageName, versionValue));
                    //    }                        
                    //}
                    string versionValue = packageNode.Attributes["version"].Value;
                    string packageValue = packageNode.Attributes["id"].Value;

                    if (!string.IsNullOrEmpty(versionValue) && packageName == packageValue)
                    {
                        return versionValue;
                    }
                }
            }
            return null; // Return null if the package is not found
        }
        private XmlDocument GetPackagesConfigXml()
        {
            // Replace "path-to-your-packages-config" with the actual path to your packages.config file
            string packagesConfigPath = "C:\\Users\\JadhavR\\source\\repos\\DependencyTreeProject\\DependencyTreeProject\\packages.config";

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(packagesConfigPath);
                return xmlDoc;
            }
            catch (Exception ex)
            {
                // Handle exceptions, such as file not found or invalid XML format
                MessageBox.Show($"Error loading packages.config: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }


    private async Task<IEnumerable<NuGet.Packaging.PackageDependencyGroup>> GetPackageTargetFrameworks(string packageName, NuGetVersion version)
        {
            ILogger logger = NullLogger.Instance;
            CancellationToken cancellationToken = CancellationToken.None;

            SourceCacheContext cache = new SourceCacheContext();
            //SourceRepository repository = Repository.Factory.GetCoreV3("https://api.nuget.org/v3/index.json");
            SourceRepository repository = Repository.Factory.GetCoreV3("https://usnugetproxy.am.bm.net/v3/index.json");
            PackageMetadataResource resourceMeta = await repository.GetResourceAsync<PackageMetadataResource>();

            //IEnumerable<IPackageSearchMetadata> packages = await resourceMeta.GetMetadataAsync(
            //                        "Newtonsoft.Json",
            //                        includePrerelease: true,
            //                        includeUnlisted: false,
            //                        cache,
            //                        logger,
            //                        cancellationToken);
            IEnumerable<IPackageSearchMetadata> packages = await resourceMeta.GetMetadataAsync(
                                  packageName,
                                  includePrerelease: true,
                                  includeUnlisted: false,
                                  cache,
                                  logger,
                                  cancellationToken);
            


            foreach (IPackageSearchMetadata package in packages)
            {
                if (package.Identity.Version.OriginalVersion == version.OriginalVersion)
                {
                    var targetFrameworks = package.DependencySets.ToList();
                    foreach(NuGet.Packaging.PackageDependencyGroup framework in targetFrameworks)
                    {
                        framework.TargetFramework.ToString();
                    }
                    return targetFrameworks;                    
                }
            }           
            return null;
           
        }
    
    private async Task<NuGetVersion> GetLatestPackageVersion(string packageName)
        {
            ILogger logger = NullLogger.Instance;
            CancellationToken cancellationToken = CancellationToken.None;

            SourceCacheContext cache = new SourceCacheContext();
            //SourceRepository repository = Repository.Factory.GetCoreV3("https://api.nuget.org/v3/index.json");
            SourceRepository repository = Repository.Factory.GetCoreV3("https://usnugetproxy.am.bm.net/v3/index.json");
            FindPackageByIdResource resource = await repository.GetResourceAsync<FindPackageByIdResource>();

            try
            {
                // Get all versions of the package
                IEnumerable<NuGetVersion> versions = await resource.GetAllVersionsAsync(
                    packageName,
                    cache,
                    logger,
                    cancellationToken);

                // Find the latest version
                NuGetVersion latestVersion = versions.OrderByDescending(v => v).FirstOrDefault();

                return latestVersion;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                return null;
            }
        }        

        private async void LoadGridViewData()
        {
            //string packageName = "Newtonsoft.Json"; // Replace with the package name
            
            try
            {
                
               
                DataTable table = new DataTable();
                table.Columns.Add("Dependency", typeof(string));
                table.Columns.Add("Current Version", typeof(string));
                table.Columns.Add(".NetFramework8.0 Support", typeof(string));
                table.Columns.Add(".NetFramework2.0 Support", typeof(string));
                table.Columns.Add(".NetFramework2.1 Support", typeof(string));
                table.Columns.Add("Supporting Framework", typeof(string));

                

                // Add a new row to the DataTable
                foreach (string str in references)
                {
                    string packageName = str;
                    var strTargetFramework = string.Empty;
                    var isNF20Support = false;
                    var isNF80Support = false;
                    var isNF21Support = false;
                    var currentVersion = string.Empty;
                    var latestVersion = await GetLatestPackageVersion(packageName);
                    var targetFrameworks = await GetPackageTargetFrameworks(packageName, latestVersion);
                    foreach (var item in targetFrameworks)
                    {
                        strTargetFramework += item.TargetFramework.ToString();

                        strTargetFramework += " |";                        
                    }
                    //if (strTargetFramework.Contains("8.0") || strTargetFramework.Contains("2.0") || strTargetFramework.Contains("2.1"))
                    //{
                    //    isNF20Support = true;
                    //}
                    if (strTargetFramework.Contains("8.0"))
                    {
                        isNF80Support = true;
                    }
                    if (strTargetFramework.Contains("2.0"))
                    {
                        isNF20Support = true;
                    }
                    if (strTargetFramework.Contains("2.1"))
                    {
                        isNF21Support = true;
                    }
                    foreach (var item in currentVersions.Where(i => i.Key == packageName))
                    {
                        //use name items
                        currentVersion = item.Value;
                    }
                    
                    // Get the XmlDocument from wherever you have it
                    // XmlDocument packagesConfigXml = GetPackagesConfigXml(); // Implement this method

                    // Get the version for the current package
                    // string version = GetVersionsFromPackagesConfig(packagesConfigXml, packageName);
                    //table.Rows.Add(str, version ?? "Not Found", "Available", strTargetFramework);
                    table.Rows.Add(str, currentVersion, isNF80Support ? "Available" : "Not Available", isNF20Support ? "Available" : "Not Available", isNF21Support ? "Available" : "Not Available", strTargetFramework);
                }
                dataGridView1.DataSource = table;
                dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
    private async void button2_Click(object sender, EventArgs e)
        {
            string packageName = textBox2.Text.Trim();
            var isNF20Support = false;
            var isNF80Support = false;
            var isNF21Support = false;
            // Check if the package name is not empty
            if (string.IsNullOrEmpty(packageName))
            {
                MessageBox.Show("Please enter a package name first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                var getVersion = await GetLatestPackageVersion(packageName);
                var targetFrameworks = await GetPackageTargetFrameworks(packageName, getVersion);

                var strTargetFramework = string.Empty;
                foreach (var item in targetFrameworks)
                {
                    strTargetFramework += item.TargetFramework.ToString();

                    strTargetFramework += " |";
                }
                if (strTargetFramework.Contains("8.0"))
                {
                    isNF80Support = true;
                }
                if (strTargetFramework.Contains("2.0"))
                {
                    isNF20Support = true;
                }
                if (strTargetFramework.Contains("2.1"))
                {
                    isNF21Support = true;
                }
                DataTable table = new DataTable();
                table.Columns.Add("Dependency", typeof(string));
                table.Columns.Add("Current Version", typeof(string));
                table.Columns.Add(".NetFramework8.0 Support", typeof(string));
                table.Columns.Add(".NetFramework2.0 Support", typeof(string));
                table.Columns.Add(".NetFramework2.1 Support", typeof(string));
                table.Columns.Add("Supporting Framework", typeof(string));

                table.Rows.Add(packageName, getVersion, isNF80Support ? "Available" : "Not Available", isNF20Support ? "Available" : "Not Available", isNF21Support ? "Available" : "Not Available", string.Join(", ", strTargetFramework));
                dataGridView1.DataSource = table;
                dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                //Display package information
                //MessageBox.Show($"Package Name: {packageName}\n" +
                //                $"Current Version: {getVersion}\n" +
                //                $"Target Frameworks: {string.Join(", ", strTargetFramework)}",
                //                "Package Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Project Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFilePath = openFileDialog.FileName;

                ExcelReadUpdate.Library.ExcelReadUpdate excelReadUpdate = new ExcelReadUpdate.Library.ExcelReadUpdate();
                excelReadUpdate.ReadExcelToPrepareJSONAndUpdateExcel(selectedFilePath);

               // MessageBox.Show("Compatability Check Complete !!");
            }
        }
    }
    }
