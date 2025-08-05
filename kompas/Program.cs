using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using KompasAPI7;
using HeagBoKaT;
using Kompas6Constants;


namespace Kompas
{
    class Program
    {
        static void Main()
        {
            var t = Check();
            foreach (var dim in t)
            {
                Console.WriteLine($"Тип размера: {dim.name}|{dim.vault}|Ручной:{dim.nominal}");
            }
            Console.ReadKey();
        }

        static List<(string type, string name, float vault, bool nominal)> Check()
        {  
            List<(string type, string name, float vault, bool nominal)> listDimension = new List<(string type, string name, float vault, bool nominal)>();
            IApplication app = (IApplication)HeagBoKaT.HeagBoKaT.GetActiveObject("KOMPAS.Application.7");
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)app.ActiveDocument;
            Views views = kompasDocument2D.ViewsAndLayersManager.Views;
            app.HideMessage = ksHideMessageEnum.ksHideMessageYes;
            for (int i = 0; i < views.Count; i++)
            {
                IView currentView = views.View[i];
                ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)currentView;
                for (int j = 0; j < symbols2DContainer.LineDimensions.Count; j++)
                {
                    LineDimension lineDimension = symbols2DContainer.LineDimensions.LineDimension[j];
                    IDimensionText idimensionText = (IDimensionText)lineDimension;
                    if (idimensionText != null && !idimensionText.AutoNominalValue)
                    {
                        float.TryParse(idimensionText.NominalText.Str, out float nominal);
                        listDimension.Add(
                            ("LineDimension", "Линеный размер", nominal, !idimensionText.AutoNominalValue));
                    }


                }

                for (int j = 0; j < symbols2DContainer.AngleDimensions.Count; j++)
                {
                    IAngleDimension angleDimension = symbols2DContainer.AngleDimensions.AngleDimension[j];
                    IDimensionText angleDimensionText = (IDimensionText)angleDimension;
                    if (angleDimensionText != null && !angleDimensionText.AutoNominalValue)
                    {
                        string degree = angleDimensionText.NominalText.Str.Split("@1~")[0];
                        int degreeInt = int.Parse(degree);
                        string minute = angleDimensionText.NominalText.Str.Split("@1~")[1].Split("'")[0];
                        int minuteInt = int.Parse(minute);
                        if (minuteInt != 0 || minuteInt != 15 || minuteInt != 30 || minuteInt != 45)
                        {
                            listDimension.Add(("AngleDimension", "Угловой размер",
                                (float)Math.Round(degreeInt + (((float)minuteInt * (1f / 60f))), 2),
                                !angleDimensionText.AutoNominalValue));
                        }
                    }
                }


            }
            return listDimension;
        }
        
    }
    
}