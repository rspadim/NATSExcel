using System;
using Xunit;

namespace NATSExcel.Tests
{
    /// <summary>
    /// Makes the unit suite hermetic with respect to the machine's configuration.
    /// <c>AppConfig.Current</c> normally reads %USERPROFILE%\NATSExcel.json (or the
    /// add-in folder / C:\Windows), so a developer (or a leftover E2E run) with a
    /// non-default config there would change the runtime behavior and fail tests
    /// for the wrong reason. This fixture pins a known default for the whole
    /// "config-seam" collection.
    ///
    /// Mirrors RedisExcel.Tests\HermeticConfigFixture.cs, using the frozen
    /// AppConfig test seams from PLAN.md §1.4:
    /// <c>AppConfig.CurrentOverrideForTests</c> and
    /// <c>AppConfig.Sanitize(new ConfigRoot())</c>. Because the production sources
    /// are linked into this assembly (not referenced), the internal seams are
    /// visible here without InternalsVisibleTo.
    ///
    /// Tests that exercise parsing call the loader directly and are unaffected;
    /// the write-mode seams still take precedence over the configured values.
    /// </summary>
    [CollectionDefinition("config-seam")]
    public class ConfigSeamCollection : ICollectionFixture<HermeticConfigFixture>
    {
    }

    public sealed class HermeticConfigFixture : IDisposable
    {
        private readonly AppConfig.Root _previous;

        public HermeticConfigFixture()
        {
            _previous = AppConfig.CurrentOverrideForTests;
            AppConfig.CurrentOverrideForTests = AppConfig.Sanitize(new AppConfig.Root());
        }

        public void Dispose()
        {
            AppConfig.CurrentOverrideForTests = _previous;
        }
    }
}
