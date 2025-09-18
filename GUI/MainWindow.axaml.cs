
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Kompas6Constants;
using KompasAPI7;
using System.Text.Json;

namespace GUI;

enum Target
{
    VOL,
    SHU,
    QAR
}

public partial class MainWindow : Window
{
    List<int> id_stamp = new List<int>() { 110, 111, 114 };
    private Target _target = Target.VOL;
    private bool _isLoading = false;
    private Dictionary<string, List<string>> _stamps = new();
    private bool _badDocument = false;
    private int _badCount = 0;

    public MainWindow()
    {
        InitializeComponent();
        LoadNames();
        LoadState();
    }

    private void LoadState()
    {
        _isLoading = true;
        try
        {
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            string settingsPath = Path.Combine(localFolder, "settings.json");
            string json = File.ReadAllText(settingsPath);
            var config = JsonSerializer.Deserialize<ConfigSettings>(json);
            add_stamp.IsChecked = config?.add_stamp ?? false;
            sign_stamp.IsChecked = config?.sign_stamp ?? false;
            saved_pdf.IsChecked = config?.saved_pdf ?? false;
            auto_place.IsChecked = config?.auto_place ?? false;
            close_doc.IsChecked = config?.close_doc ?? false;
            if (!Enum.TryParse(config?.target ?? "SHU", out _target)) _target = Target.SHU;
            rbVol.IsChecked = _target == Target.VOL;
            rbShu.IsChecked = _target == Target.SHU;
            rbQar.IsChecked = _target == Target.QAR;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
        finally
        {
            _isLoading = false;
        }
    }

    private void SaveState()
    {
        try
        {
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            Directory.CreateDirectory(localFolder);
            string settingsPath = Path.Combine(localFolder, "settings.json");

            var cfg = new ConfigSettings
            {
                add_stamp = add_stamp.IsChecked == true,
                sign_stamp = sign_stamp.IsChecked == true,
                saved_pdf = saved_pdf.IsChecked == true,
                auto_place = auto_place.IsChecked == true,
                close_doc = close_doc.IsChecked == true,
                target = _target.ToString()
            };
            string json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(settingsPath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }

    private void LoadNames()
    {
        try
        {
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KompasTweaker");
            Directory.CreateDirectory(localFolder);

            string configPath = Path.Combine(localFolder, "config.json");
            string defaultConfigPath = Path.Combine(AppContext.BaseDirectory, "Assets", "config.json");

            // если файла нет — копируем из Assets
            if (!File.Exists(configPath) && File.Exists(defaultConfigPath))
            {
                File.Copy(defaultConfigPath, configPath);
            }

            string json = File.ReadAllText(configPath);
            var config = JsonSerializer.Deserialize<ConfigName>(json);
            if (config?.names != null)
            {
                foreach (var n in config.names)
                {
                    name1.Items.Add(new ComboBoxItem { Content = n });
                    name2.Items.Add(new ComboBoxItem { Content = n });
                    name3.Items.Add(new ComboBoxItem { Content = n });
                }
            }

            if (config?.stamps != null)
            {
                _stamps = config.stamps;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    private void OnAnyToggleChanged(object? sender, RoutedEventArgs e)
    {
        if (_isLoading) return; // во время загрузки не сохраняем
        SaveState();
    }

    void RadioButton_Checked(object? sender, RoutedEventArgs e)
    {
        if (sender is RadioButton rb && rb.IsChecked == true)
        {
            var tag = rb.Tag?.ToString() ?? "VOL";
            if (!Enum.TryParse<Target>(tag, out var parsed)) parsed = Target.VOL;
            _target = parsed;
            if (_isLoading)
                return;
            SaveState();
        }
    }

    void Radiobutton_add_click(object? sender, RoutedEventArgs e)
    {
        string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KompasTweaker");
        string path = Path.Combine(localFolder, "config.json");
        Console.WriteLine(path);
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }


    private class ConfigSettings
    {
        public bool? add_stamp { get; set; }
        public bool? sign_stamp { get; set; }
        public bool? saved_pdf { get; set; }
        public bool? auto_place { get; set; }
        public bool? close_doc { get; set; }
        public string? target { get; set; }
    }

    private class ConfigName
    {
        public List<string>? names { get; set; }
        public Dictionary<string, List<string>>? stamps { get; set; }
    }

    static bool Check_pos(double x, double y, IKompasDocument2D1 document2D1)
    {
        var collision = document2D1.FindObject(x, y, 10, null);
        if (collision != null)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    // void Test_button_click(object? sender, RoutedEventArgs e)
    // {
    //     IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
    //     IKompasDocument kompasDocument = app.ActiveDocument;
    //     IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;
    //     IKompasDocument2D1 document2D1 = (IKompasDocument2D1)kompasDocument2D;
    //     double x = kompasDocument2D.LayoutSheets[0].Format.FormatWidth;
    //     IDrawingDocument drawing = (IDrawingDocument)kompasDocument2D;
    //     TechnicalDemand technicalDemand = drawing.TechnicalDemand;
    //     IText textTechnical = technicalDemand.Text;
    //     bool moved_tt = false;
    //     bool free_place = false;
    //     if ((bool)auto_place.IsChecked)
    //     {
    //         if (technicalDemand != null)
    //         {
    //             try
    //             {
    //                 double[] technicalPos = (double[])technicalDemand.BlocksGabarits;
    //                 for (int i = 0; i < technicalPos.Length; i++)
    //                 {
    //                     Console.WriteLine(i + ":" + technicalPos[i]);
    //                 }
    //                 if (technicalPos[0] >= x - 190 && technicalPos[1] >= 71 && technicalPos[1] <= 96) // проверяю если тт над рамкой иначе проверить свободное место над рамкой
    //                 {
    //                     Console.WriteLine("Orientir top");
    //                     for (int i = 0; i < 10; i++) // проверяю верхнюю границу
    //                     {
    //                         if (technicalPos[1] <= 85)
    //                         {
    //                             technicalPos[1] = 85;
    //                         }
    //                         free_place = Check_pos(technicalPos[0] + 18 * i, technicalPos[1] + technicalDemand.Text.Count * 7 + 2, document2D1);

    //                         Console.WriteLine(i + ":" + free_place + "|" + (technicalPos[0] + 18 * i) + ":" + (technicalPos[1] + technicalDemand.Text.Count * 7 + 2));
    //                         if (!free_place)
    //                         {
    //                             moved_tt = true;
    //                             break;
    //                         }
    //                     }
    //                     if (!moved_tt)
    //                     {
    //                         Console.WriteLine("Move");
    //                         technicalDemand.BlocksGabarits = new double[4] { technicalPos[0], technicalPos[1] + 12, technicalPos[2], technicalPos[1] + technicalDemand.Text.Count * 7 + 22 };
    //                         technicalDemand.Update();
    //                         moved_tt = true;

    //                     }
    //                     else
    //                     {
    //                         Console.WriteLine("Top in collision, dont move, checked left");
    //                         double x_center = _target switch { "VOL" => 253, "SHU" => 242, "QAR" => 254 };
    //                         double size = _target switch { "VOL" => 116, "SHU" => 93, "QAR" => 118 };
    //                         for (int i = 0; i < 10; i++)
    //                         {
    //                             free_place = Check_pos(x - 190 - size + (size / 10) * i, 21, document2D1);
    //                             Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
    //                             if (!free_place) break;
    //                         }
    //                         Console.WriteLine(free_place);
    //                         if (free_place)
    //                         {
    //                             Console.WriteLine("Left stamp");
    //                         }

    //                     }

    //                 }
    //                 else
    //                 {
    //                     double size = _target switch { "VOL" => 116, "SHU" => 93, "QAR" => 118 };
    //                     for (int i = 0; i < 10; i++)
    //                     {
    //                         free_place = Check_pos(x - 190 + (size / 10) * i, 82, document2D1);
    //                         Console.WriteLine(free_place + ":" + (x - 190 + size / 10));
    //                         if (!free_place) break;
    //                     }
    //                     Console.WriteLine(free_place);
    //                     if (free_place)
    //                     {
    //                         Console.WriteLine("Free stamp top. TT xz");
    //                     }
    //                 }


    //             }
    //             catch (Exception ex)
    //             {
    //                 Console.WriteLine(ex);
    //             }
    //         }
    //         else
    //         {
    //             double size = _target switch { "VOL" => 116, "SHU" => 93, "QAR" => 118 };
    //             for (int i = 0; i < 10; i++)
    //             {
    //                 free_place = Check_pos(x - 190 + (size / 10) * i, 82, document2D1);
    //                 Console.WriteLine(free_place + ":" + (x - 190 + size / 10));
    //                 if (!free_place) break;
    //             }
    //             Console.WriteLine(free_place);
    //             if (free_place)
    //             {
    //                 Console.WriteLine("Free stamp top. TT xz");
    //             }
    //         }
    //     }
    // }

    private void Button_Click(object? sender, RoutedEventArgs e)
    {
        LoadNames();
        button_run.IsEnabled = false;
        var addStamp = add_stamp.IsChecked == true;
        var needSign = sign_stamp.IsChecked == true;
        var needPdf = saved_pdf.IsChecked == true;
        var auto_paced_stamp = auto_place.IsChecked == true;
        var needClose = close_doc.IsChecked == true;

        bool shouldFillTitle = addStamp
                               || ((name1.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0
                                   || (name2.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0
                                   || (name3.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0);
        _badCount = 0;

        IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
        var oldHide = app.HideMessage;
        app.HideMessage = (ksHideMessageEnum)1;
        try
        {
            int total = app.Documents.Count;
            progressBar.Value = 0;
            if (total > 1) progressBar.Maximum = total + 1;
            else
            {
                progressBar.Maximum = total;
            }

            var case_text = new[]
            {
                (name1.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "",
                (name2.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "",
                (name3.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "",
            };
            IKompasDocument kompasDocument = app.Documents[0];
            for (int i = total - 1; i >= 0; i--)
            {
                _badDocument = false;
                kompasDocument = app.Documents[i];
                kompasDocument.Active = true;

                // IKompasDocument kompasDocument = app.ActiveDocument;
                if (kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentPart ||
                    kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentAssembly ||
                    kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentFragment)
                {
                    kompasDocument.Close((DocumentCloseOptions)1);
                    continue;
                }

                if (kompasDocument.Type == KompasAPIObjectTypeEnum.ksObjectSpecificationDocument)
                {
                    Set_text_stamp(app, kompasDocument, case_text);
                    SavePDF(kompasDocument);
                    continue;
                }

                IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;

                if (addStamp)
                {
                    Set_holc_stamp(kompasDocument2D, auto_paced_stamp);
                }

                if (shouldFillTitle || addStamp)
                {
                    Set_text_stamp(app, kompasDocument, case_text);
                }

                if (needSign)
                {
                    Set_sign(app, kompasDocument2D);
                }

                if (needPdf && _badDocument == false)
                {
                    SavePDF(kompasDocument);
                }

                var Views = kompasDocument2D.ViewsAndLayersManager.Views;
                var signView = Views.View["Подписи"];
                signView?.Delete();

                if (needClose && !_badDocument) kompasDocument.Close((DocumentCloseOptions)1);
                progressBar.Value = progressBar.Value + 1;
            }
        }
        finally
        {
            app.HideMessage = oldHide;
            progressBar.Value = progressBar.Value + 1;
            button_run.IsEnabled = true;
        }
    }

    void Set_holc_stamp(IKompasDocument2D kompasDocument2D, bool auto_paced_stamp)
    {
        Views Views = kompasDocument2D.ViewsAndLayersManager.Views;
        var view = Views.View["Штамп"] ?? Views.Add((LtViewType)1);
        view.X = 0;
        view.Y = 0;
        view.Number = 255;
        view.Name = "Штамп";
        view.Current = true;
        view.Update();
        view = Views.View["Штамп"];
        view.Layers.Layer[0].Color = 255;
        view.Layers.Layer[0].Update();
        var draw_cont = (IDrawingContainer)view;
        if (view.ObjectCount > 0)
        {
            dynamic obj = draw_cont.Objects[0];
            obj[0].Delete();
        }

        var ins_manager = (IInsertionsManager)kompasDocument2D;
        double x = kompasDocument2D.LayoutSheets[0].Format.FormatWidth;
        IDrawingDocument drawing = (IDrawingDocument)kompasDocument2D;
        TechnicalDemand technicalDemand = drawing.TechnicalDemand;
        IText textTechnical = technicalDemand.Text;
        IKompasDocument2D1 document2D1 = (IKompasDocument2D1)kompasDocument2D;
        var format = kompasDocument2D.LayoutSheets[0].Format;


        var frwPath = Path.Combine(AppContext.BaseDirectory, "Assets", "frame", _target.ToString() + ".frw");
        var ins_definition = ins_manager.AddDefinition(0, "Штамп", frwPath);
        var ins_obj = draw_cont.InsertionObjects;
        var ins = ins_obj.Add(ins_definition);
        Console.WriteLine(technicalDemand.IsCreated);
        bool moved_tt = false;
        bool free_place = false;
        double x_center = _target switch
        {
            Target.VOL => 250, Target.SHU => 238, Target.QAR => 252, _ => throw new NotImplementedException()
        };
        double size = _target switch
        {
            Target.VOL => 116, Target.SHU => 95, Target.QAR => 118, _ => throw new NotImplementedException()
        };
        Console.WriteLine(technicalDemand.IsCreated);
        if (auto_paced_stamp)
        {
            if (technicalDemand.IsCreated)
            {
                try
                {
                    double[] technicalPos = (double[])technicalDemand.BlocksGabarits;
                    for (int i = 0; i < technicalPos.Length; i++)
                    {
                        Console.WriteLine(i + ":" + technicalPos[i]);
                    }

                    if (technicalPos[0] >= x - 190 && technicalPos[1] >= 71 &&
                        technicalPos[1] <= 96) // проверяю если тт над рамкой и двигаю
                    {
                        Console.WriteLine("Orientir top");
                        for (int i = 0; i < 10; i++) // проверяю верхнюю границу для тт
                        {
                            if (technicalPos[1] <= 85)
                            {
                                technicalPos[1] = 85;
                            }

                            free_place = Check_pos(technicalPos[0] + 18 * i,
                                technicalPos[1] + technicalDemand.Text.Count * 7 + 2, document2D1);

                            Console.WriteLine(i + ":" + free_place + "|" + (technicalPos[0] + 18 * i) + ":" +
                                              (technicalPos[1] + technicalDemand.Text.Count * 7 + 2));
                            if (!free_place)
                            {
                                moved_tt = true;
                                break;
                            }
                        }

                        if (!moved_tt)
                        {
                            Console.WriteLine("Move");
                            technicalDemand.BlocksGabarits = new double[4]
                            {
                                technicalPos[0], technicalPos[1] + 12, technicalPos[2],
                                technicalPos[1] + technicalDemand.Text.Count * 7 + 22
                            };
                            technicalDemand.Update();
                            moved_tt = true;
                            ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                            ins.Update();
                        }
                        else
                        {
                            if ((format.VerticalOrientation && (format.Format == ksDocumentFormatEnum.ksFormatA4 ||
                                                                format.Format == ksDocumentFormatEnum.ksFormatA3)))
                            {
                                _badCount++;
                                _badDocument = true;
                                return;
                            }

                            Console.WriteLine("Top in collision, dont move, checked left");
                            for (int i = 0; i < 10; i++)
                            {
                                free_place = Check_pos(x - 190 - size + (size / 10) * i, 17, document2D1);
                                Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                                if (!free_place) break;
                            }

                            Console.WriteLine("TT enable, left pos, dont move" + free_place);
                            if (free_place)
                            {
                                Console.WriteLine("Left stamp");
                                ins.SetPlacement(x - x_center, 16, 0, false);
                                ins.Update();
                            }
                            else
                            {
                                _badDocument = true;
                                _badCount++;
                            }
                        }
                    }
                    else // проверяю свободное место над рамкой если тт не над рамкой
                    {
                        if (technicalPos[0] <= x - 190 && technicalPos[1] <= 21)
                        {
                            _badDocument = true;
                            _badCount++;
                            return;
                        }

                        for (int i = 0; i < 10; i++)
                        {
                            free_place = Check_pos(x - 188 + (size / 10) * i, 82, document2D1);
                            Console.WriteLine(free_place + ":" + (x - 188 + size / 10));
                            if (!free_place) break;
                        }

                        Console.WriteLine("TT unknow pos, top pos" + free_place);
                        if (free_place)
                        {
                            Console.WriteLine("Free stamp top. TT xz");
                            ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                            ins.Update();
                        }
                        else // слева от рамки
                        {
                            if ((format.VerticalOrientation && (format.Format == ksDocumentFormatEnum.ksFormatA4 ||
                                                                format.Format == ksDocumentFormatEnum.ksFormatA3)))
                            {
                                _badCount++;
                                _badDocument = true;
                                return;
                            }

                            for (int i = 0; i < 10; i++)
                            {
                                free_place = Check_pos(x - 190 - size + (size / 10) * i, 17, document2D1);
                                Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                                if (!free_place) break;
                            }

                            if (free_place)
                            {
                                Console.WriteLine("Left stamp");
                                ins.SetPlacement(x - x_center, 16, 0, false);
                                ins.Update();
                            }
                            else
                            {
                                _badDocument = true;
                                _badCount++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            else // если тт нет то проверяю просто
            {
                for (int i = 0; i < 10; i++)
                {
                    free_place = Check_pos(x - 188 + (size / 10) * i, 82, document2D1);
                    Console.WriteLine(free_place + ":" + (x - 190 + size / 10));
                    if (!free_place) break;
                }

                Console.WriteLine("TT not in doc, top pos" + free_place);
                if (free_place) // над рамкой
                {
                    Console.WriteLine("Free stamp top. TT xz");
                    ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                    ins.Update();
                }
                else // слева от рамки
                {
                    if ((format.VerticalOrientation && (format.Format == ksDocumentFormatEnum.ksFormatA4 ||
                                                        format.Format == ksDocumentFormatEnum.ksFormatA3)))
                    {
                        _badCount++;
                        _badDocument = true;
                        return;
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        free_place = Check_pos(x - 190 - size + (size / 10) * i, 17, document2D1);
                        Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                        if (!free_place) break;
                    }

                    Console.WriteLine("TT not in doc, left pos" + free_place);
                    if (free_place)
                    {
                        Console.WriteLine("Left stamp");
                        ins.SetPlacement(x - x_center, 16, 0, false);
                        ins.Update();
                    }
                    else
                    {
                        _badDocument = true;
                        _badCount++;
                    }
                }
            }
        }
        else
        {
            ins.SetPlacement(x - 190 + size / 2, 82, 0, false);
            ins.Update();
        }
    }

    public void Set_text_stamp(IApplication app, IKompasDocument kompasDocument, string[] case_text)
    {
        var format = kompasDocument.LayoutSheets[0];
        Console.WriteLine("Stamp");
        if (case_text[0] != null || case_text[1] != null || case_text[2] != null)
        {
            for (int i = 0; i < 3; i++)
            {
                if (case_text[i] == "") continue;
                var text = format.Stamp.Text[id_stamp[i]];
                text.Clear();
                var textLine = text.Add();
                var textItem = textLine.Add();
                textItem.Str = case_text[i];
            }
        }

        if (add_stamp.IsChecked == true)
        {
            var text = format.Stamp.Text[9];
            text.Clear();
            if (_stamps.TryGetValue(_target.ToString(), out var lines))
            {
                foreach (var line in lines)
                {
                    var textLine = text.Add();
                    textLine.Str = line;
                }
            }
        }

        format.Stamp.Update();
    }

    public void Set_sign(IApplication app, IKompasDocument2D kompasDocument2D)
    {
        int[] y = { 29, 24, 9 };
        var view_manager = kompasDocument2D.ViewsAndLayersManager;
        var Views = view_manager.Views;
        var view = Views.Add((LtViewType)1);
        view.X = 0;
        view.Y = 0;
        view.Number = 254;
        view.Name = "Подписи";
        view.Current = true;
        view.Update();
        for (int i = 0; i < 3; i++)
        {
            var format = kompasDocument2D.LayoutSheets[0];
            var text = format.Stamp.Text[id_stamp[i]];
            if (text.Str != "")
            {
                var path = AppContext.BaseDirectory + "Assets\\sign\\" + text.Str + ".png";
                view = Views.View["Подписи"];
                var img_view_cont = (IDrawingContainer)view;
                var img_view = img_view_cont.Rasters;
                var img_view_add = img_view.Add();
                img_view_add.SetPlacement(format.Format.FormatWidth - 150, y[i], 0, false);
                img_view_add.FileName = path;
                img_view_add.Scale = 0.045;
                img_view_add.Update();
            }
        }
    }

    private static void SavePDF(IKompasDocument kompasDocument)
    {
        string outDir = Path.Combine(kompasDocument.Path, "Чертежи в pdf");
        Directory.CreateDirectory(outDir);
        var local_path = kompasDocument.Path + "\\Чертежи в pdf\\" +
                         Path.GetFileNameWithoutExtension(kompasDocument.Name) + ".pdf";
        kompasDocument.SaveAs(local_path);
    }
}