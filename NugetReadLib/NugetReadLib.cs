using NuGet.Common;
using NuGet.Protocol;
using NuGet.Protocol.Core.Types;
using NuGet.Versioning;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace NugetRead.Library
{
    public class NugetReadLib
    {
        public NugetReadLib() { }

        public async Task<NuGetVersion> FetchLatestNugetVersion(string packageName)
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


        public async Task<IEnumerable<NuGet.Packaging.PackageDependencyGroup>> GetNugetTargetFrameworks(string packageName, NuGetVersion version)
        {
            ILogger logger = NullLogger.Instance;
            CancellationToken cancellationToken = CancellationToken.None;

            SourceCacheContext cache = new SourceCacheContext();
            SourceRepository repository = Repository.Factory.GetCoreV3("https://usnugetproxy.am.bm.net/v3/index.json");
            PackageMetadataResource resourceMeta = await repository.GetResourceAsync<PackageMetadataResource>();

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
                    foreach (NuGet.Packaging.PackageDependencyGroup framework in targetFrameworks)
                    {
                        framework.TargetFramework.ToString();
                    }
                    return targetFrameworks;
                }
            }

            return null;
        }


    }
}
