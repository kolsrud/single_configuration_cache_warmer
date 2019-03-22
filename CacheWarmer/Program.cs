using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using Qlik.Engine;

namespace CacheWarmer
{
    class Program
    {
        class Settings
        {
            public readonly string XmlFilePath;

            public Settings(string[] args)
            {
                if (!args.Any())
                    throw new ArgumentException("No argument provided.", nameof(args));

                XmlFilePath = args.First();
            }
        }

        static void Main(string[] args)
        {
            var settings = new Settings(args);
            var uris = ReadUrisFromExcel(settings.XmlFilePath);
            var configurations = uris.Select(uri => new SingleConfiguration(uri));
            ApplyConfigurations(settings, configurations);
        }

        private static void ApplyConfigurations(Settings settings, IEnumerable<SingleConfiguration> configurations)
        {
            foreach (var configurationGroup in configurations.GroupBy(config => config.Url + "/" + config.VirtualProxy))
            {
                Console.WriteLine("Warming cache for node: {0}", configurationGroup.Key);
                var serverConfiguration = configurationGroup.First();
                var location = Location.FromUri(serverConfiguration.Uri.Scheme + "://" + serverConfiguration.Uri.Host);
                location.VirtualProxyPath = serverConfiguration.VirtualProxy;
                location.IsVersionCheckActive = false;
                location.AsNtlmUserViaProxy(serverConfiguration.Scheme == "https", certificateValidation: false);
                WarmCacheForNode(location, configurationGroup);
            }
        }

        private static void WarmCacheForNode(ILocation location, IEnumerable<SingleConfiguration> configurations)
        {
            foreach (var configurationGroup in configurations.GroupBy(config => config.AppId))
            {
                Console.WriteLine("  - Warming cache for app: {0}", configurationGroup.Key);
                var appIdentifier = location.AppWithIdOrDefault(configurationGroup.First().AppId);
                if (appIdentifier == null)
                {
                    Console.WriteLine("    - Warning: App not found.");
                    return;
                }
                using (var app = location.App(appIdentifier, Session.Random))
                {
                    WarmCacheForApp(app, configurationGroup);
                    Console.WriteLine("    - Completed cache warming for app: {0}", configurationGroup.Key);
                }
            }
        }

        private static void WarmCacheForApp(IApp app, IEnumerable<SingleConfiguration> configurations)
        {
            Console.WriteLine("    - Applying {0} configurations...", configurations.Count());
            foreach (var configuration in configurations)
            {
                var sheet = app.GetGenericObject(configuration.SheetId);
                var children = sheet.GetChildInfos().Select(childInfo => app.GetGenericObject(childInfo.Id));
                var allObjects = new[] {sheet}.Concat(children);
                ApplySelections(app, configuration);
                Console.Write("    - Warming cache... ");
                Task.WhenAll(allObjects.Select(obj => obj.GetLayoutAsync())).Wait();
                Console.WriteLine(" Done!");
            }
        }

        private static void ApplySelections(IApp app, SingleConfiguration configuration)
        {
            app.ClearAll();
            configuration.Selections.ToList().ForEach(selection => ApplySelection(app, selection));
        }

        private static void ApplySelection(IApp app, Selection selection)
        {
            Console.WriteLine("    - Applying selections to field '{0}' ({1} values)...", selection.Field, selection.Values.Count);
            var field = app.GetField(selection.Field);
            field.SelectValues(selection.Values.Select(MakeFieldValue));
        }

        private static FieldValue MakeFieldValue(string value)
        {
            return double.TryParse(value, out var d)
                ? new FieldValue {IsNumeric = true, Number = d}
                : new FieldValue {IsNumeric = false, Text = value};
        }

        private static IEnumerable<Uri> ReadUrisFromExcel(string xlsx)
        {
            var uris = new List<Uri>();
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(xlsx)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //selet starting row here
                {
                    var cells = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Where(c => c.Value != null)
                        .Select(c => c.Value.ToString());
                    foreach (var cell in cells)
                    {
                        try
                        {
                            var uri = new Uri(cell);
                            uris.Add(uri);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
            }

            return uris;
        }
    }
}
