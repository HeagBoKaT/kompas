
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
                        // Console.WriteLine("Light");
                        mainWindow.Background = new SolidColorBrush(Color.FromArgb(235, 229, 229, 229));
                        mainWindow.Foreground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
                    }

                    else mainWindow.Background = new SolidColorBrush(Color.FromArgb(235, 72, 72, 72));
                }
            }
        }
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
        // Console.WriteLine(path);
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

    private void Test_button_click(object? sender, RoutedEventArgs e)
    {
        try
        {
            IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
            app.GetSystemVersion(out int major, out int minor, out int patch, out int revision );
            Console.WriteLine(major + "." + minor + "." + patch + "." + revision);
            SystemSettings settings = app.SystemSettings; //Интерфейс настроек
            settings.EnablesAddSystemDelimersInMarking = true; // отображение разделителей и спец символов
            settings.EnableAddFilesToRecentList = true;
            IKompasDocument document = app.ActiveDocument;
            if (document != null)
            {
                int start_pos = 0;
                IKompasDocument2D kompasDocument2D = (IKompasDocument2D)document;
                IDrawingDocument drawingDocument = (IDrawingDocument)kompasDocument2D;
                ITechnicalDemand technicalDemand = drawingDocument.TechnicalDemand;
                Console.WriteLine(technicalDemand.Text.Str);
                if (major == 22)
                {
                    while (technicalDemand.Text.Str.IndexOf("^(#31~") != -1)
                    {
                        
                    }
                }
                
                IViews views = kompasDocument2D.ViewsAndLayersManager.Views;
                int t = 1;
                string current = "";
                int pos1 = 0;
                var result = new List<string>();
                for (int i = 0; i < technicalDemand.Text.Count; i++)
                {
                    var currentLine = technicalDemand.Text.TextLine[i];
                    int pos = 0;

                    if (currentLine.Numbering == ksTextNumberingEnum.ksTNumbNumber)
                    {
                        if (!string.IsNullOrEmpty(current))
                        {
                            result.Add(current.Trim());
                        }

                        current = currentLine.Str;
                    }
                    else
                    {
                        if (current != null)
                        {
                            current += " " + currentLine.Str;
                        }
                    }
                }
                if (current != null && current.Length > 0)
                    result.Add(current.ToString().Trim());

                technicalDemand.Text.Clear();
                string tt = "";
                string end1 = String.Empty;
                string end2 = String.Empty;
                int pos4;
                if (major == 24)
                {
                    end1 = ";~";
                    end2 = ")";
                    pos4 = 2;
                }
                else
                {
                    end1 = "~";
                    end2 = ")";
                    pos4 = 1;
                }
                for (int i = 0; i < result.Count; i++)
                {
                            while (result[i].IndexOf("^(#31~") != -1)
                            {
                                result[i].Replace("^(#31~", string.Empty);
                            }
                }
                
                for (int i=0; i < result.Count; i++)
                {
                    start_pos = 0;
                    
                    
                    while ((pos1 = result[i].IndexOf("^(", start_pos)) != -1)
                    {
                        Console.WriteLine(result[i].Length);
                        int pos2 = result[i].IndexOf(end1, pos1+8);
                        string target = result[i].Substring(pos1, pos2 - pos1 + pos4);
                        int pos3 = target.IndexOf(end2);
                        string insert = target.Substring(0, pos3 + 1);
                        Console.WriteLine($"Позиция: {pos1}");
                        Console.WriteLine($"Что заменяем: {target}");
                        Console.WriteLine($"Основа поиска: {insert}");
                        start_pos = pos2 + 1;
                        foreach (IView view in views)
                        {
                            // Console.WriteLine($"__{view.Name}__");
                            // Console.WriteLine($"Число объектов вида: {view.ObjectCount}");
                            ISymbols2DContainer drawingContainer = (ISymbols2DContainer)view;
                            var leaders = drawingContainer.Leaders;
                            foreach (IBaseLeader leader in leaders)
                            {
                                // Console.WriteLine(leader.Type);
                                switch (leader.Type)
                                {
                                    case KompasAPIObjectTypeEnum.ksObjectMarkLeader:
                                    {

                                        if (((IMarkLeader)leader).Designation.Str.Contains(insert))
                                        {
                                            Console.WriteLine(((IMarkLeader)leader).Designation.Str);
                                            result[i] = result[i].Replace(target,
                                                ((IMarkLeader)leader).Designation.Str);
                                        }

                                        break;
                                    }
                                    case KompasAPIObjectTypeEnum.ksObjectLeader:
                                    {
                                        if (((ILeader)leader).TextOnShelf.Str.Contains(insert))
                                        {
                                            // Console.WriteLine(((ILeader)leader).TextOnShelf.Str);
                                            result[i] = result[i].Replace(target,
                                                    ((ILeader)leader).TextOnShelf.Str);
                                        }

                                        break;
                                    }
                                    case KompasAPIObjectTypeEnum.ksObjectPositionLeader:
                                    {
                                        if (((IPositionLeader)leader).Positions.Str.Contains(insert))
                                        {
                                            // Console.WriteLine(((IPositionLeader)leader).Positions.Str);
                                            result[i] = result[i].Replace(target, ((IPositionLeader)leader).Positions.Str);
                                        }

                                        break;
                                    }
                                }
                            }

                        }
                    }

                    if (i < result.Count - 1)
                    {
                        tt += result[i] + "\n";
                    }
                    else
                    {
                        tt+=result[i];
                    }
                    
                    
                }
                Console.WriteLine(tt);
                technicalDemand.Text.Str = tt;
                technicalDemand.Update();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            throw;
        }

        //     IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
        //     SystemSettings settings = app.SystemSettings; //Интерфейс настроек
        //     settings.EnablesAddSystemDelimersInMarking = true; // отображение разделителей и спец символов
        //     IKompasDocument kompasDocument = app.ActiveDocument;
        //     IPropertyKeeper propertyKeeper = ((IPropertyKeeper)(IKompasDocument2D)kompasDocument); // Интерфейс получения/редактирования значения свойств
        //     var meta = propertyKeeper.Properties; //Метаданные объекта (позволяет вывести в xml файле не знаю на сколько нужно но там есть отдельный id documentNumber
        //     IPropertyMng propertyMng = (IPropertyMng)app; // Менеджер свойств
        //     var properties = propertyMng.GetProperties(kompasDocument); //свойства документа
        //     var props = ((IEnumerable)properties).Cast<_Property>().ToList(); //приведение к списку свойств
        //     for (int i = 0; i < props.Count; i++)
        //     {
        //         propertyKeeper.GetPropertyValue(props[i], out object value, true, out var source );
        //         Console.WriteLine($"{source}: {value}");
        //     }
        //     //И самое важное
        //     propertyKeeper.SetPropertyValue(props[0], "VL.39А-SP1.80002878.00.20$|$|$|$|$| $|СБ", false);
        //
        //     // propertyKeeper.SetPropertyValue(prope, "VL.39А-SP1.80002878.00.20 ГЧ", false);
        //     // XDocument doc = XDocument.Parse(properties);
        //     // XElement? element = doc
        //     //     .Descendants("property")
        //     //     .FirstOrDefault(x => (string?)x.Attribute("id") == "documentNumber");
        //     // Console.WriteLine(element.Attribute("value").Value);
        //     // element.SetAttributeValue("value", "ГЧ");
        //     // properties = doc.ToString();
        //
        //     // Console.WriteLine(propertyMng);
        //
        //
        //
        //
        //     // for (int i=0; i< 1320; i++)
        //     // {
        //     //     if (stamp.Text[i].Str != "")
        //     //     {
        //     //         Console.WriteLine($"{i}:{stamp.Text[i].Str}");
        //     //     }
        //     //     
        //     // }
        //
        //
        //
        //
        //
        //
    }
    

    private void Button_Click(object? sender, RoutedEventArgs e)
    {
        Stopwatch sw = new Stopwatch();
        sw.Start();
        button_run.IsEnabled = false;
        var addStamp = add_stamp.IsChecked == true;
        var needSign = sign_stamp.IsChecked == true;
        var needPdf = saved_pdf.IsChecked == true;
        var auto_paced_stamp = auto_place.IsChecked == true;
        var needClose = close_doc.IsChecked == true;
        var oldSaved = saveOldVersion.IsChecked == true;
        var oldVersionValue = ((ComboBoxItem)oldVersion.SelectedItem).Content.ToString();
        bool shouldFillTitle = addStamp
                               || ((name1.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0
                                   || (name2.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0
                                   || (name3.SelectedItem as ComboBoxItem)?.Content?.ToString()?.Length > 0);
        _badCount = 0;
        IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
        var oldHide = app.HideMessage;
        app.Visible = silentCheckBox.IsChecked != true;
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
                    DocCloseAndSave(kompasDocument, oldSaved, oldVersionValue);
                    continue;
                }


                if (addStamp)
                {
                    Set_holc_stamp(kompasDocument, auto_paced_stamp);
                }

                if (shouldFillTitle)
                {
                    Set_text_stamp(app, kompasDocument, case_text);
                }

                if (needSign)
                {
                    Set_sign(app, kompasDocument);
                }

                if (needPdf)
                {
                    SavePDF(kompasDocument);
                }

                if (needClose && !_badDocument)
                {
                    DocCloseAndSave(kompasDocument, oldSaved, oldVersionValue);
                    progressBar.Value = progressBar.Value + 1;
                }
                
            }
        }
        finally
        {
            app.HideMessage = oldHide;
            progressBar.Value = progressBar.Value + 1;
            button_run.IsEnabled = true;
            if (!app.Visible) app.Visible = !app.Visible;
        }

        sw.Stop();
        Console.WriteLine($"Time: {sw.ElapsedMilliseconds} ms");
    }

    private void DocCloseAndSave(IKompasDocument kompasDocument, bool oldSaved, string oldVersionValue)
    {
        if (!oldSaved)
        {
            kompasDocument.Close((DocumentCloseOptions)1);
        }
        else
        {   
            var path  = Path.Combine(kompasDocument.Path, kompasDocument.Name);
            var version = oldVersionValue switch { "21" => 27, "22" => 28, "23" => 29 };
            ((IKompasDocument1)kompasDocument).SaveAsEx(path, version);
            kompasDocument.Close(0);
        }
    }


    private void SavePDF_spec(IKompasDocument kompasDocument)
    {
        string outDir = Path.Combine(kompasDocument.Path, "Чертежи в pdf");
        Directory.CreateDirectory(outDir);
        var localPath = kompasDocument.Path + "\\Чертежи в pdf\\" +
                        Path.GetFileNameWithoutExtension(kompasDocument.Name) + ".pdf";
        IKompasDocument1 kompasDocument1 = (IKompasDocument1)kompasDocument;
        kompasDocument1.SaveAsEx(localPath, 0);
        var format = kompasDocument.LayoutSheets[0];
        var stamp = format.Stamp;
        var doc = PdfReader.Open(localPath, PdfDocumentOpenMode.Modify);
        var page = doc.Pages[0];
        using var gfx = XGraphics.FromPdfPage(page);
        double x = 190; // pt, от верхнего левого угла
        double[] y = { 768, 780, 810 }; // pt, вниз
        double width = 44; // pt
        for (int i = 0; i < 3; i++)
        {
            var text = stamp.Text[id_stamp[i]];
            if (text.Str == "")
            {
                continue;
            }
            var signPath = Path.Combine(AppContext.BaseDirectory, "Assets", "sign", text.Str + ".png");
            using var img = XImage.FromFile(signPath);
            double height = img.PixelHeight * width / img.PixelWidth;
            // Console.WriteLine($"{text.Str}: {img.PixelWidth}x{img.PixelHeight} -> {width}x{height}");
            gfx.DrawImage(img, x - (width / 2), y[i] - (height / 2), width, height);

        }

        doc.Save(localPath);
    }

    void Set_holc_stamp(IKompasDocument kompasDocument, bool auto_paced_stamp)
    {
        // Установка дополнительного штампа (фрейма) на чертёж.
        // Логика:
        // 1. Пропускаем спецификации (там используем другой метод сохранения PDF).
        // 2. Подготавливаем / создаём отдельное представление (View) "Штамп".
        // 3. Зачищаем предыдущее содержимое штампа (если уже вставлялся).
        // 4. Вставляем .frw-файл выбранного формата (_target) как объект вставки.
        // 5. Если включён авто-режим — пытаемся разместить штамп:
        //    a) Сначала предпочитаем позицию над основной рамкой (top, Y≈82)
        //    b) Если занято или мешает TechnicalDemand (ТТ) — пробуем слева (left, Y≈16)
        //    c) Для вертикальных форматов A4/A3 блокируем смещение влево (расцениваем как плохой документ)
        //    d) Пытаемся при необходимости сдвинуть сам TechnicalDemand вверх, если он перекрывает область штампа
        // 6. Если авто-режим выключен — просто ставим в верх (над рамкой)
        // Переменные:
        //  _badDocument / _badCount — индикаторы невозможности корректного размещения штампа
        //  technicalDemand — блок технических требований, который может мешать
        //  size / x_center — габариты и центральная координата штампа зависят от выбранного Target

        if (kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentSpecification)
        {
            // Для спецификаций штамп не ставим
            return;
        }

        IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;
        Views Views = kompasDocument2D.ViewsAndLayersManager.Views;
        // Получаем существующий View "Штамп" либо создаём новый (тип 1 — обычный вид)
        var view = Views.View["Штамп"] ?? Views.Add((LtViewType)1);
        view.X = 0;
        view.Y = 0;
        view.Name = "Штамп";
        view.Current = true;
        view.Update();
        view = Views.View["Штамп"];
        // Красим нулевой слой (вероятно для визуального выделения — 255 = красный)
        view.Layers.Layer[0].Color = 255;
        view.Layers.Layer[0].Update();
        var draw_cont = (IDrawingContainer)view;
        if (view.ObjectCount > 0)
        {
            // Если ранее что-то вставляли — удаляем первый объект (избегаем дублирования)
            dynamic obj = draw_cont.Objects[0];
            obj[0].Delete();
        }

        var ins_manager = (IInsertionsManager)kompasDocument2D;
        // Ширина формата (для вычисления смещения от правого края)
        double x = kompasDocument2D.LayoutSheets[0].Format.FormatWidth;
        IDrawingDocument drawing = (IDrawingDocument)kompasDocument2D;
        TechnicalDemand technicalDemand = drawing.TechnicalDemand;
        IText textTechnical = technicalDemand.Text;
        IKompasDocument2D1 document2D1 = (IKompasDocument2D1)kompasDocument2D;
        var format = kompasDocument2D.LayoutSheets[0].Format;


        // Путь к нужному фрейму согласно выбранной цели (VOL/SHU/QAR)
        var frwPath = Path.Combine(AppContext.BaseDirectory, "Assets", "frame", _target.ToString() + ".frw");
        // Создаём определение и сам объект вставки
        var ins_definition = ins_manager.AddDefinition(0, "Штамп", frwPath);
        var ins_obj = draw_cont.InsertionObjects;
        var ins = ins_obj.Add(ins_definition);
        // Console.WriteLine(technicalDemand.IsCreated);
        // Флаги работы алгоритма
        bool moved_tt = false;
        bool free_place = false;
        // Предопределённые смещения центра и ширина для разных видов штампа
        double x_center = _target switch
        {
            Target.VOL => 250, Target.SHU => 238, Target.QAR => 252, _ => throw new NotImplementedException()
        };
        double size = _target switch
        {
            Target.VOL => 116, Target.SHU => 95, Target.QAR => 118, _ => throw new NotImplementedException()
        };
        // Console.WriteLine(technicalDemand.IsCreated);
        if (auto_paced_stamp)
        {
            // Автоматическое размещение штампа
            if (technicalDemand.IsCreated)
            {
                try
                {
                    // Получаем габариты блока TechnicalDemand (массив: x1, y1, x2, y2)
                    double[] technicalPos = (double[])technicalDemand.BlocksGabarits;
                    // for (int i = 0; i < technicalPos.Length; i++)
                    // {
                    // Console.WriteLine(i + ":" + technicalPos[i]);
                    // }

                    if (technicalPos[0] >= x - 190 && technicalPos[1] >= 71 &&
                        technicalPos[1] <= 96) // проверяю если тт над рамкой и двигаю
                    {
                        // Сценарий: TechnicalDemand прямо над рамкой — нужно подвинуть его выше, если есть место
                        // Console.WriteLine("Orientir top");
                        for (int i = 0; i < 10; i++) // проверяю верхнюю границу для тт
                        {
                            if (technicalPos[1] <= 85)
                            {
                                // Принудительно ограничиваем минимальный Y (чтобы не наезжал на штамп)
                                technicalPos[1] = 85;
                            }

                            // Проверяем отсутствие коллизий по точкам над планируемой областью перемещения ТТ
                            free_place = Check_pos(technicalPos[0] + 18 * i,
                                technicalPos[1] + technicalDemand.Text.Count * 7 + 2, document2D1);

                            // Console.WriteLine(i + ":" + free_place + "|" + (technicalPos[0] + 18 * i) + ":" +
                            //                   (technicalPos[1] + technicalDemand.Text.Count * 7 + 2));
                            if (!free_place)
                            {
                                // Коллизия — придётся ставить штамп в альтернативном месте
                                moved_tt = true;
                                break;
                            }
                        }

                        if (!moved_tt)
                        {
                            // Есть место чтобы поднять TechnicalDemand — сдвигаем его вверх и ставим штамп сверху
                            // Console.WriteLine("Move");
                            technicalDemand.BlocksGabarits = new double[4]
                            {
                                technicalPos[0], technicalPos[1] + 12, technicalPos[2],
                                technicalPos[1] + technicalDemand.Text.Count * 7 + 22
                            };
                            technicalDemand.Update();
                            moved_tt = true;
                            // Ставим штамп по верхней позиции (x - 188 + size/2, Y=82)
                            ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                            ins.Update();
                        }
                        else
                        {
                            // Не удалось двинуть TechnicalDemand — пробуем левую позицию, если формат позволяет
                            if ((format.VerticalOrientation && (format.Format == ksDocumentFormatEnum.ksFormatA4 ||
                                                                format.Format == ksDocumentFormatEnum.ksFormatA3)))
                            {
                                // Для вертикальных A4/A3 — отказ
                                _badCount++;
                                _badDocument = true;
                                return;
                            }

                            // Console.WriteLine("Top in collision, dont move, checked left");
                            for (int i = 0; i < 10; i++)
                            {
                                // Проверяем сектор слева от рамки (Y=17)
                                free_place = Check_pos(x - 190 - size + (size / 10) * i, 17, document2D1);
                                // Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                                if (!free_place) break;
                            }

                            // Console.WriteLine("TT enable, left pos, dont move" + free_place);
                            if (free_place)
                            {
                                // Есть место слева — ставим штамп
                                // Console.WriteLine("Left stamp");
                                ins.SetPlacement(x - x_center, 16, 0, false);
                                ins.Update();
                            }
                            else
                            {
                                // Ни сверху ни слева — помечаем документ как проблемный
                                _badDocument = true;
                                _badCount++;
                            }
                        }
                    }
                    else // проверяю свободное место над рамкой если тт не над рамкой
                    {
                        // TechnicalDemand не в верхней зоне — сначала пытаемся поставить штамп сверху

                        for (int i = 0; i < 10; i++)
                        {
                            // Проверяем верхнюю позицию под штамп (Y=82)
                            free_place = Check_pos(x - 188 + (size / 10) * i, 82, document2D1);
                            // Console.WriteLine(free_place + ":" + (x - 188 + size / 10));
                            if (!free_place) break;
                        }

                        // Console.WriteLine("TT unknow pos, top pos" + free_place);
                        if (free_place)
                        {
                            // Верх свободен — ставим штамп
                            // Console.WriteLine("Free stamp top. TT xz");
                            ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                            ins.Update();
                        }
                        else // слева от рамки
                        {
                            if (technicalPos[0] <= x - 190 && technicalPos[1] <= 21)
                            {
                                // ТТ находится в зоне будущего левого штампа. Пытаемся приподнять его вместо немедленного отказа.
                                try
                                {
                                    double
                                        desiredY =
                                            28; // минимальное безопасное положение верхней границы ТТ над левым штампом
                                    if (technicalPos[1] < desiredY)
                                    {
                                        bool canMoveTT = true;
                                        // Проверяем отсутствие коллизий по линии предполагаемого верхнего края ТТ после подъёма
                                        for (int i = 0; i < 10; i++)
                                        {
                                            if (!Check_pos(technicalPos[0] + 18 * i,
                                                    desiredY + technicalDemand.Text.Count * 7 + 2, document2D1))
                                            {
                                                canMoveTT = false;
                                                break;
                                            }
                                        }

                                        if (canMoveTT)
                                        {
                                            technicalDemand.BlocksGabarits = new double[4]
                                            {
                                                technicalPos[0], desiredY, technicalPos[2],
                                                desiredY + technicalDemand.Text.Count * 7 + 22
                                            };
                                            technicalDemand.Update();
                                            technicalPos[1] =
                                                desiredY; // обновляем локально чтобы последующая логика видела новое положение
                                            // Console.WriteLine("Left area: TT moved up to free stamp zone");
                                        }
                                        else
                                        {
                                            // Не удалось сдвинуть — помечаем как проблемный документ
                                            _badDocument = true;
                                            _badCount++;
                                            return;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Console.WriteLine("Error moving TT for left placement: " + ex.Message);
                                    _badDocument = true;
                                    _badCount++;
                                    return;
                                }
                            }

                            // Верх занят — пробуем слева
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
                                // Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                                if (!free_place) break;
                            }

                            if (free_place)
                            {
                                // Console.WriteLine("Left stamp");
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
                    // Любая ошибка логики — просто логируем
                    Console.WriteLine(ex);
                }
            }
            else // если тт нет то проверяю просто
            {
                // TechnicalDemand отсутствует — проверяем только зоны размещения
                for (int i = 0; i < 10; i++)
                {
                    free_place = Check_pos(x - 188 + (size / 10) * i, 82, document2D1);
                    // Console.WriteLine(free_place + ":" + (x - 190 + size / 10));
                    if (!free_place) break;
                }

                // Console.WriteLine("TT not in doc, top pos" + free_place);
                if (free_place) // над рамкой
                {
                    // Верх свободен — ставим штамп
                    // Console.WriteLine("Free stamp top. TT xz");
                    ins.SetPlacement(x - 188 + size / 2, 82, 0, false);
                    ins.Update();
                }
                else // слева от рамки
                {
                    // Верх занят — пробуем слева
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
                        // Console.WriteLine(free_place + ":" + (x - 190 - size + size / 10));
                        if (!free_place) break;
                    }

                    // Console.WriteLine("TT not in doc, left pos" + free_place);
                    if (free_place)
                    {
                        // Лево свободно — ставим штамп
                        // Console.WriteLine("Left stamp");
                        ins.SetPlacement(x - x_center, 16, 0, false);
                        ins.Update();
                    }
                    else
                    {
                        // Нет доступных зон
                        _badDocument = true;
                        _badCount++;
                    }
                }
            }
        }
        else
        {
            // Авторазмещение выключено — используем стандартную верхнюю позицию
            ins.SetPlacement(x - 190 + size / 2, 82, 0, false);
            ins.Update();
        }
    }

    public void Set_text_stamp(IApplication app, IKompasDocument kompasDocument, string[] case_text)
    {
        var layotSheets = kompasDocument.LayoutSheets[0];
        if (layotSheets.LayoutStyleNumber == 16 &&
            kompasDocument.DocumentType != DocumentTypeEnum.ksDocumentSpecification)
        {
            layotSheets.LayoutStyleNumber = 16;
        }

        layotSheets.Update();
        // Console.WriteLine("Stamp");
        if (case_text[0] != null || case_text[1] != null || case_text[2] != null)
        {
            for (int i = 0; i < 3; i++)
            {
                if (case_text[i] == "") continue;
                var text = layotSheets.Stamp.Text[id_stamp[i]];
                text.Clear();
                var textLine = text.Add();
                textLine.Align = 0;
                var textItem = textLine.Add();
                textItem.Str = case_text[i];
            }
        }

        if (add_stamp.IsChecked == true)
        {
            var text = layotSheets.Stamp.Text[9];
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

        layotSheets.Stamp.Update();
    }

    public void Set_sign(IApplication app, IKompasDocument kompasDocument)
    {
        if (kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentSpecification)
        {
            SavePDF_spec(kompasDocument);
            return;
        }

        IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;
        int[] y = { 29, 24, 9 };
        var view_manager = kompasDocument2D.ViewsAndLayersManager;
        var Views = view_manager.Views;
        var view = Views.Add((LtViewType)1);
        view.X = 0;
        view.Y = 0;
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

    private void SavePDF(IKompasDocument kompasDocument)
    {
        if (kompasDocument.DocumentType == DocumentTypeEnum.ksDocumentSpecification)
        {
            SavePDF_spec(kompasDocument);
            return;
        }

        if (_badDocument == false)
        {
            string outDir = Path.Combine(kompasDocument.Path, "Чертежи в pdf");
            Directory.CreateDirectory(outDir);
            var local_path = kompasDocument.Path + "\\Чертежи в pdf\\" +
                             Path.GetFileNameWithoutExtension(kompasDocument.Name) + ".pdf";
            kompasDocument.SaveAs(local_path);
        }

        IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;
        var Views = kompasDocument2D.ViewsAndLayersManager.Views;
        var signView = Views.View["Подписи"];
        signView?.Delete();
    }

    private void SaveOldVersion_OnChecked(object? sender, RoutedEventArgs e)
    {
        oldVersion.IsEnabled = true;
    }

    private void SaveOldVersion_OnUnchecked(object? sender, RoutedEventArgs e)
    {
        oldVersion.IsEnabled = false;
    }
}