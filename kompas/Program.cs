using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using KompasAPI7;
using HeagBoKaT;

class Program
{
    public static void Main()
    {
        IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
        IKompasDocument2D kompasDocument2D = (IKompasDocument2D)app.ActiveDocument;
        Views views = kompasDocument2D.ViewsAndLayersManager.Views;
        for (int i = 0; i < views.Count; i++)
        {
            IView currentView = views.View[i];
            ISymbols2DContainer symbols2DContainer = currentView as ISymbols2DContainer;
            for (int j = 0; j < symbols2DContainer.LineDimensions.Count; j++)
            {
                LineDimension lineDimension = symbols2DContainer.LineDimensions.LineDimension[j];
                IDimensionText idimensionText = lineDimension as IDimensionText;
                Console.Write("Автоматический размер: " + idimensionText.AutoNominalValue + "; ");
                Console.WriteLine(idimensionText.NominalText.Str);
            }

        }
    }
}