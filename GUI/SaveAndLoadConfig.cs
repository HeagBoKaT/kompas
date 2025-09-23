using System;
using System.IO;
using System.Text.Json;

namespace GUI;

public interface SaveAndLoadConfig
{   public static bool load = false;
    public static void SaveSettingConfig()
    {
        ConfigSettings? config;
        try
        {   
            if (load) return;
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            Directory.CreateDirectory(localFolder);
            string settingsPath = Path.Combine(localFolder, "settings.json");
            if (File.Exists(settingsPath))
            {
                var oldConfig = File.ReadAllText(settingsPath);
                config = JsonSerializer.Deserialize<ConfigSettings>(oldConfig);
            }
            else
            {
                config = new ConfigSettings();
            }
            
            config.saveAllStatus = Utility.saveAllStatus;
            config.saveOldVersionStatus = Utility.saveOldVersionStatus;
            config.addStampCustomStatus = Drawing.AddStampCustomStatus;
            config.signStampStatus = Drawing.SignStampStatus;
            config.savedPdfStatus = Drawing.SavedPdfStatus;
            config.autoPlaceStampStatus = Drawing.AutoPlaceStampStatus;
            config.closeDocStatus = Drawing.CloseDocStatus;
            config.target = Drawing.Target;
            config.tabActive = MainWindow.tabActive;
            
            var newConfig = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(settingsPath, newConfig);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
        
    }

    public static void LoadSettingConfig()
    {
        try
        {
            load = true;
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            string settingsPath = Path.Combine(localFolder, "settings.json");
            if (!File.Exists(settingsPath))
            {
                MainWindow.tabActive = 0;
                Utility.saveAllStatus = false;
                Utility.saveOldVersionStatus = false;
                Drawing.AddStampCustomStatus = false;
                Drawing.SignStampStatus = false;
                Drawing.SavedPdfStatus = false;
                Drawing.AutoPlaceStampStatus = false;
                Drawing.CloseDocStatus = false;
                Drawing.Target = "SHU";
                return;
            }
            string json = File.ReadAllText(settingsPath);
            var config = JsonSerializer.Deserialize<ConfigSettings>(json);
            MainWindow.tabActive = config?.tabActive ?? 0;
            Utility.saveAllStatus = config?.saveAllStatus ?? false;
            Utility.saveOldVersionStatus = config?.saveOldVersionStatus ?? false;
            Drawing.AddStampCustomStatus = config?.addStampCustomStatus ?? false;
            Drawing.SignStampStatus = config?.signStampStatus ?? false;
            Drawing.SavedPdfStatus = config?.savedPdfStatus ?? false;
            Drawing.AutoPlaceStampStatus = config?.autoPlaceStampStatus ?? false;
            Drawing.CloseDocStatus = config?.closeDocStatus ?? false;
            Drawing.Target = config?.target ?? "SHU";
            load = false;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }
    private class ConfigSettings
    {
        public bool? addStampCustomStatus { get; set; }
        public bool? signStampStatus { get; set; }
        public bool? savedPdfStatus { get; set; }
        public bool? autoPlaceStampStatus { get; set; }
        public bool? closeDocStatus { get; set; }
        public string? target { get; set; }
        public bool? saveAllStatus { get; set; }
        public bool? saveOldVersionStatus { get; set; }
        public int tabActive { get; set; }
        
    }
}