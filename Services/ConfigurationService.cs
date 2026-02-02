using System.IO;
using System.Text.Json;
using LauraAssetBuildReview.Models;

namespace LauraAssetBuildReview.Services;

/// <summary>
/// Service for saving and loading processing configurations.
/// </summary>
public class ConfigurationService
{
    private const string ConfigFileName = "processing_config.json";

    public void SaveConfiguration(ProcessingConfiguration config, string? directory = null)
    {
        var configPath = GetConfigPath(directory);
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        var json = JsonSerializer.Serialize(config, options);
        File.WriteAllText(configPath, json);
    }

    public ProcessingConfiguration? LoadConfiguration(string? directory = null)
    {
        var configPath = GetConfigPath(directory);
        
        if (!File.Exists(configPath))
            return null;

        try
        {
            var json = File.ReadAllText(configPath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            return JsonSerializer.Deserialize<ProcessingConfiguration>(json, options);
        }
        catch
        {
            return null;
        }
    }

    public ProcessingConfiguration GetDefaultConfiguration()
    {
        return new ProcessingConfiguration
        {
            EanColumn = "C",
            StatusColumn = "G",
            DropdownColumn = "G",
            StartRow = 3,
            WorksheetIndex = 1,
            MinEanDigits = 14,
            MaxEanDigits = 14,
            AllowNonNumericEans = false,
            ReferenceFiles = new List<ReferenceFileConfig>
            {
                new ReferenceFileConfig
                {
                    EanColumn = "C",
                    StartRow = 1,
                    WorksheetIndex = 1,
                    Priority = 1
                },
                new ReferenceFileConfig
                {
                    EanColumn = "C",
                    StartRow = 1,
                    WorksheetIndex = 1,
                    Priority = 2
                }
            },
            AutoMapFilenames = true,
            ComparisonColumn = "G",
            ComparisonStartRow = 3
        };
    }

    private string GetConfigPath(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory))
        {
            // Save in application directory or user's AppData
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "LauraAssetBuildReview");
            
            if (!Directory.Exists(appDataPath))
                Directory.CreateDirectory(appDataPath);
            
            return Path.Combine(appDataPath, ConfigFileName);
        }

        return Path.Combine(directory, ConfigFileName);
    }
}
