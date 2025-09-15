
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Kompas6Constants;
using KompasAPI7;
using Kompas6API5;
using Kompas6API2D5COM;
using KAPITypes;

namespace GUI;

public partial class MainWindow : Window
{
    List<int> id_stamp = new List<int>() { 110, 111, 114 };
    private string _target = "VOL";
    // List<string> name = new List<string>() { "Фадеев", "Жуков", "Пащенко", "Харланов", "Шурпач" };
    public MainWindow()
    {

        InitializeComponent();
    }


    void RadioButton_Checked(object? sender, RoutedEventArgs e)
    {
        if (sender is RadioButton rb && rb.IsChecked == true)
        {
            _target = rb.Tag?.ToString() ?? "VOL";
        }

    }

    // void Test_button_click(object? sender, RoutedEventArgs e)
    // {
    //     IApplication? app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
    //     IKompasDocument kompasDocument = app.ActiveDocument;
    //     DocumentFrames frames = kompasDocument.DocumentFrames;
    //     Debug.WriteLine(frames[0].Caption);
    //     var kompas = (KompasObject)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.5");
    //     var a = (ksSpcDocument)kompas.SpcActiveDocument();
    //     Debug.WriteLine(a.ksGetSpcDocumentPagesCount());


    //     // ksObjectSpecificationDocument
    //     // ksObjectDrawingDocument

    // }

    async void Button_Click(object? sender, RoutedEventArgs e)
    {
        button_run.IsEnabled = false;

        var addStamp = add_stamp.IsChecked == true;
        var needSign = sign_stamp.IsChecked == true;
        var needPdf = saved_pdf.IsChecked == true;
        var auto_paced_stamp = auto_place.IsChecked == true;

        bool sign_check = true;
        IApplication? app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
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
            await Task.Run(() =>
            {
                for (int i = 0; i < total; i++)
                {
                    app.Documents[i].Active = true;
                    IKompasDocument kompasDocument = app.ActiveDocument;
                    if (kompasDocument.Type == KompasAPIObjectTypeEnum.ksObjectSpecificationDocument)
                    {
                        Set_text_stamp(app, kompasDocument, case_text);
                        SavePDF(kompasDocument);
                        return;
                    }
                    IKompasDocument2D kompasDocument2D = (IKompasDocument2D)kompasDocument;
                    if (addStamp)
                    {
                        Set_holc_stamp(kompasDocument2D, auto_paced_stamp);
                    }
                    if (sign_check)
                    {
                        Set_text_stamp(app, kompasDocument, case_text);
                    }
                    if (needSign)
                    {
                        Set_sign(app, kompasDocument2D);
                    }
                    if (needPdf)
                    {
                        SavePDF(kompasDocument);
                    }
                    var view_manager = kompasDocument2D.ViewsAndLayersManager;
                    var Views = view_manager.Views;
                    var view = Views.View["Подписи"];
                    if (view != null)
                    {
                        view.Delete();
                    }
                    Dispatcher.UIThread.Post(() => progressBar.Value = i + 1);
                }

            });
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
        var view = Views.Add((LtViewType)1);
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

        IKompasDocument2D1 document2D1 = (IKompasDocument2D1)kompasDocument2D;
        var collision1 = document2D1.FindObject(kompasDocument2D.LayoutSheets[0].Format.FormatWidth - 180, 85, 21, null);
        var collision2 = document2D1.FindObject(kompasDocument2D.LayoutSheets[0].Format.FormatWidth - 70, 85, 21, null);
        var collision3 = document2D1.FindObject(kompasDocument2D.LayoutSheets[0].Format.FormatWidth - 140, 85, 21, null);
        IDrawingDocument drawing = (IDrawingDocument)kompasDocument2D;
        TechnicalDemand technical = drawing.TechnicalDemand;
        double[] techical_pos = (double[])technical.BlocksGabarits;
        var ins_definition = ins_manager.AddDefinition(0, "Штамп", @"\\server\\Чтение и запись\\Обмен файлами\\" + _target + ".frw");
        double x, y;
        bool left_side = false;
        if (collision1 != null || collision2 != null || collision3 != null && kompasDocument2D.LayoutSheets[0].Format.VerticalOrientation != true)
        {
            x = kompasDocument2D.LayoutSheets[0].Format.FormatWidth - _target switch { "VOL" => 255, "SHU" => 245, "QAR" => 255, _ => throw new NotImplementedException() };
            y = 25;
            left_side = true;
        }
        else
        {
            x = kompasDocument2D.LayoutSheets[0].Format.FormatWidth - _target switch { "VOL" => 130, "SHU" => 137, "QAR" => 129, _ => throw new NotImplementedException() };
            y = 85;
        }
        var ins_obj = draw_cont.InsertionObjects;
        var ins = ins_obj.Add(ins_definition);
        ins.SetPlacement(x, y, 0, false);
        ins.Update();
        if (techical_pos[1] < 100 && auto_paced_stamp && !left_side)
        {
            technical.AutoPlacement = true;
            technical.Update();
        }
    }
    public void Set_text_stamp(IApplication app, IKompasDocument kompasDocument, string[] case_text)
    {
        var format = kompasDocument.LayoutSheets[0];
        if (case_text[0] == null && case_text[1] == null && case_text[2] == null) return;
        else
        {
            for (int i = 0; i < 3; i++)
            {
                if (case_text[i] == null) continue;
                var text = format.Stamp.Text[id_stamp[i]];
                text.Clear();
                var textLine = text.Add();
                var textItem = textLine.Add();
                textItem.Str = case_text[i];

            }
            format.Stamp.Update();

        }
        return;
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
                var path = @"\\server\\Чтение и запись\\Подписи\\" + text.Str + ".png";
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

    void SavePDF(IKompasDocument kompasDocument)
    {
        var outDir = Path.Combine(kompasDocument.Path, "Чертежи в pdf");
        Directory.CreateDirectory(outDir);
        var local_path = kompasDocument.Path + "\\Чертежи в pdf\\" + Path.GetFileNameWithoutExtension(kompasDocument.Name) + ".pdf";
        kompasDocument.SaveAs(local_path);
    }
}