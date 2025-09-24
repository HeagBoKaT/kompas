
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Kompas6Constants;
using KompasAPI7;
using System.Text.Json;
using System.Xml.Linq;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using Microsoft.Win32;
using Avalonia.Media;
using Kompas6API5;
using GUI;

namespace GUI;



public partial class MainWindow : Window
{
    public static int tabActive;
    private bool saveTab = false;
    private Logger _logger;
    
    public MainWindow()
    {   
        InitializeComponent();
        _logger = new Logger("app.log");
        LoadSettingConfig();
        try
        {
            const string registryKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";
            using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(registryKeyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    if (value != null && value is int intValue)
                    {
                        if (intValue == 1)
                        {
                            mainWindow.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                            var brush = App.Current.Resources["BackgroundBrush"] as SolidColorBrush;
                            brush.Color = Color.FromArgb(235, 229, 229, 229);
                        }

                        else
                        {
                            var brush = App.Current.Resources["BackgroundBrush"] as SolidColorBrush;
                            brush.Color = Color.FromArgb(235, 72, 72, 72);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex.Message);
        }
        
        
    }

    private void LoadSettingConfig()
    {
        try
        {
            SaveAndLoadConfig.LoadSettingConfig();
            Console.WriteLine(tabActive);
            Control.SelectedIndex = tabActive;
            saveTab = true;
        }
        catch (Exception ex)
        {
            _logger.Error(ex.Message);
        }
        
    }

    private void TabControl_OnSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        try
        {
            if (sender is TabControl tabControl && tabControl.SelectedItem is TabItem tabItem)
            {
                tabActive = tabControl.SelectedIndex;
                if (saveTab) SaveAndLoadConfig.SaveSettingConfig();
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex.Message);
        }
       
        
    }
}