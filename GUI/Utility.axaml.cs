using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using KompasAPI7;
using System;
using System.Text.Json;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Path = System.IO.Path;

namespace GUI;

public partial class Utility : UserControl
{
    public static bool saveAllStatus = false;
    public static bool saveOldVersionStatus = false;
    public Utility()
    {
        InitializeComponent();
        Load();
    }

    private void Load()
    {
        SaveAndLoadConfig.LoadSettingConfig();
        SaveAll.IsChecked = saveAllStatus;
        saveOldVersion.IsChecked = saveOldVersionStatus;
    }
    
    private void SaveOldVersion_OnChecked(object? sender, RoutedEventArgs e)
    {
        oldVersion.IsEnabled = true;
        saveOldVersionStatus = true;
        SaveAndLoadConfig.SaveSettingConfig();
    }

    private void SaveOldVersion_OnUnchecked(object? sender, RoutedEventArgs e)
    {
        oldVersion.IsEnabled = false;
        saveOldVersionStatus = false;
        SaveAndLoadConfig.SaveSettingConfig();
    }

    private void Button_OnClick(object? sender, RoutedEventArgs e)
    {
        SaveAndLoadConfig.SaveSettingConfig();
        var oldVersionValue = ((ComboBoxItem)oldVersion.SelectedItem).Content.ToString();
        var saveAll = SaveAll.IsChecked == true;
        var oldVersionActive = saveOldVersion.IsChecked == true;
        IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
        int total = app.Documents.Count;
        Console.WriteLine(saveAll);
        if (!oldVersionActive) return;
        if (saveAll)
        {
            for (int i = total - 1; i >= 0; i--)
            {
                IKompasDocument kompasDocument = app.Documents[i];
                SaveOldVersion(kompasDocument, oldVersionValue);
            }
        }
        else
        {
            IKompasDocument kompasDocument = app.ActiveDocument;
            SaveOldVersion(kompasDocument, oldVersionValue);
        }
        
    }

    private void SaveOldVersion(IKompasDocument kompasDocument, string oldVersionValue)
    {
        var path = Path.Combine(kompasDocument.Path, kompasDocument.Name);
        var version = oldVersionValue switch { "21" => 27, "22" => 28, "23" => 29 };
        ((IKompasDocument1)kompasDocument).SaveAsEx(path, version);
        kompasDocument.Close(0);
    }

    private void SaveAll_OnChecked(object? sender, RoutedEventArgs e)
    {
        saveAllStatus = true;
        SaveAndLoadConfig.SaveSettingConfig();
    }

    private void SaveAll_OnUnchecked(object? sender, RoutedEventArgs e)
    {
        saveAllStatus = false;
        SaveAndLoadConfig.SaveSettingConfig();
    }
}