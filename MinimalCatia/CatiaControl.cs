using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinimalCatia
{
    class CatiaControl
    {
        CatiaControl()
        {
            try
            {

                CatiaConnection cc = new CatiaConnection();

                // Finde Catia Prozess
                if (cc.CATIALaeuft())
                {

                    Console.WriteLine("0");

                    // Öffne ein neues Part
//                    cc.ErzeugePart();
                    Console.WriteLine("1");

                    // Erstelle eine Skizze
//                    cc.ErstelleLeereSkizze();
                    Console.WriteLine("2");

                    // Generiere ein Profil
//                    cc.ErzeugeProfil(20, 10);
                    Console.WriteLine("3");

                    // Extrudiere Balken
//                    cc.ErzeugeBalken(300);
                    Console.WriteLine("4");

                    // cc.setMaterial();

                    // cc.Screenshot("test");
                    Console.WriteLine("5");

                    cc.openFile();
                    Console.WriteLine("6");
                    // cc.changeUserParameter(2);

                    cc.FEM();
                    Console.WriteLine("7");

                }
                else
                {
                    Console.WriteLine("Laufende Catia Application nicht gefunden");
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message, "Exception aufgetreten");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.Source);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.InnerException);
            }
            Console.WriteLine("Fertig - Taste drücken.");
            Console.ReadKey();

        }

        static void Main(string[] args)
        {
            new CatiaControl();
        }
    }
}
