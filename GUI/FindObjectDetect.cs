using System;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;

namespace GUI;

public class FindObjectDetect
{
    public static void CheckFreePlaceForTT(IKompasDocument kompasDocument)
    {
        TechnicalDemand demand = ((IDrawingDocument)kompasDocument).TechnicalDemand;
        double[] gabarite = (double[])demand.BlocksGabarits;
        IKompasDocument2D1? document2D1 = (IKompasDocument2D1)kompasDocument;
        // отображение координат
        for (int i = 0; i < gabarite.Length; i++)
        {
            Console.WriteLine(i + ":" + gabarite[i]);
        }

        if (gabarite.Length > 4)
        {
            Console.WriteLine("Сложное ТТ");
        }
        else
        {
            double deltaX = gabarite[2] - gabarite[0];
            Console.WriteLine($"deltaX: {deltaX}");
            int steps = (int)(Math.Ceiling(deltaX / 10));
            double x = gabarite[0];
            double y = gabarite[1] + demand.Text.Count * 7 + 2;
            DrawingLine(kompasDocument,x, y, x + deltaX, y);

            for (int i = 0; i < steps; i++)
            {
                var t = document2D1.FindObjects(gabarite[0] + 10 * i, y + 10, 22, null);
                Console.WriteLine($"");
            }
        }
    }

    private static void DrawingLine(IKompasDocument kompasDocument, double x1, double y1, double x2, double y2)
    {

        Views views = ((IKompasDocument2D)kompasDocument).ViewsAndLayersManager.Views;
        ((ILayer)(IView)views.View["Тест"]).Delete();
        IView view = views.View["Тест"] ?? views.Add(LtViewType.vt_Normal);
        view.Name = "Тест";
        view.X = 0;
        view.Y = 0;
        view.Current = true;
        view.Update();
        IDrawingContainer container = (IDrawingContainer)view;
        var line = container.LineSegments.Add();
        line.X1 = x1;
        line.Y1 = y1;
        line.X2 = x2;
        line.Y2 = y2;
        line.Update();
    }
    
}