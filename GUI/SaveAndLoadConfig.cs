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
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            Directory.CreateDirectory(localFolder);
            string settingsPath = Path.Combine(localFolder, "settings.json");
            string defaultConfigPath = Path.Combine(AppContext.BaseDirectory, "Assets", "settings.json");
            if (!File.Exists(settingsPath) && File.Exists(defaultConfigPath))
            {
                File.Copy(defaultConfigPath, settingsPath);
            }
            if (File.Exists(settingsPath))
            {
                var oldConfig = File.ReadAllText(settingsPath);
                config = JsonSerializer.Deserialize<ConfigSettings>(oldConfig);
            }
            else
            {
                config = new ConfigSettings();
            }
            
            config.SaveAllStatus = Utility.saveAllStatus;
            config.SaveOldVersionStatus = Utility.saveOldVersionStatus;
            config.AddStampCustomStatus = Drawing.AddStampCustomStatus;
            config.SignStampStatus = Drawing.SignStampStatus;
            config.SavedPdfStatus = Drawing.SavedPdfStatus;
            config.AutoPlaceStampStatus = Drawing.AutoPlaceStampStatus;
            config.CloseDocStatus = Drawing.CloseDocStatus;
            config.SilentCheckBoxStatus = Drawing.SilentCheckBoxStatus;
            config.Target = Drawing.Target;
            config.TabActive = MainWindow.tabActive;
            
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
                Drawing.SilentCheckBoxStatus = false;
                Drawing.Target = "SHU";
                return;
            }
            string json = File.ReadAllText(settingsPath);
            var config = JsonSerializer.Deserialize<ConfigSettings>(json);
            MainWindow.tabActive = config?.TabActive ?? 0;
            Utility.saveAllStatus = config?.SaveAllStatus ?? false;
            Utility.saveOldVersionStatus = config?.SaveOldVersionStatus ?? false;
            Drawing.AddStampCustomStatus = config?.AddStampCustomStatus ?? false;
            Drawing.SignStampStatus = config?.SignStampStatus ?? false;
            Drawing.SavedPdfStatus = config?.SavedPdfStatus ?? false;
            Drawing.AutoPlaceStampStatus = config?.AutoPlaceStampStatus ?? false;
            Drawing.CloseDocStatus = config?.CloseDocStatus ?? false;
            Drawing.SilentCheckBoxStatus = config?.SilentCheckBoxStatus ?? false;
            Drawing.Target = config?.Target ?? "SHU";
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
        public bool? AddStampCustomStatus { get; set; }
        public bool? SignStampStatus { get; set; }
        public bool? SavedPdfStatus { get; set; }
        public bool? AutoPlaceStampStatus { get; set; }
        public bool? CloseDocStatus { get; set; }
        public string? Target { get; set; }
        public bool? SaveAllStatus { get; set; }
        public bool? SilentCheckBoxStatus { get; set; }
        public bool? SaveOldVersionStatus { get; set; }
        public int TabActive { get; set; }
        
    }
}