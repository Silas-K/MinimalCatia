using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace EigenesCatia
{
    public class EigenesCatiaControl
    {

        public EigenesCatiaControl()
        {
            try
            {
                EigeneCatiaConnection cc = new EigeneCatiaConnection();

                if (cc.CatiaLaeuft())
                {
                    Console.WriteLine("0");

                    //Part öffnen
                    cc.ErzeugePart();
                    Console.WriteLine("1");

                    cc.ErzeugeLeereSkizze();
                    Console.WriteLine("2");

                    //cc.ErzeugeRechteckProfil(60, 150);
                    cc.ErzeugeKreisProfil(50);
                    Console.WriteLine("3");

                    cc.ErzeugeBalken(600);
                    Console.WriteLine("4");
                }
                else
                {
                    Console.WriteLine("Laufende Catia Application nicht gefunden");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception aufgetreten");
            }
            Console.WriteLine("Fertig - Taste drücken.");
            Console.ReadKey();
        }


        static void Main(string[] args)
        {
            new EigenesCatiaControl();
        }
    }
}
