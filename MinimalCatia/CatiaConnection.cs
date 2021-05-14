﻿using System;
using System.Threading;
using System.Windows;
using INFITF;
using MECMOD;
using PARTITF;


namespace MinimalCatia
{
    class CatiaConnection
    {
        INFITF.Application hsp_catiaApp;
        MECMOD.PartDocument hsp_catiaPart;
        MECMOD.Sketch hsp_catiaSkizze;
        ShapeFactory SF;

        public bool CATIALaeuft()
        {
            try
            {
                object catiaObject = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "CATIA.Application");
                hsp_catiaApp = (INFITF.Application) catiaObject;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public Boolean ErzeugePart()
        {
            INFITF.Documents catDocuments1 = hsp_catiaApp.Documents;
            hsp_catiaPart = catDocuments1.Add("Part") as MECMOD.PartDocument;
            return true;
        }

        public void ErstelleLeereSkizze()
        {
            // geometrisches Set auswaehlen und umbenennen
            HybridBodies catHybridBodies1 = hsp_catiaPart.Part.HybridBodies;
            SF = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;

            HybridBody catHybridBody1;
            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden! " + Environment.NewLine +
                    "Ein PART manuell erzeugen und ein darauf achten, dass 'Geometisches Set' aktiviert ist.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            catHybridBody1.set_Name("Profile");
            // neue Skizze im ausgewaehlten geometrischen Set anlegen
            Sketches catSketches1 = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = hsp_catiaPart.Part.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;
            hsp_catiaSkizze = catSketches1.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystem();

            // Part aktualisieren
            hsp_catiaPart.Part.Update();
        }

        internal void ErzeugeZylinder(double d = 8, double l=60)
        {
            // Hauptkoerper in Bearbeitung definieren
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;
            
            // Skizze umbenennen
            hsp_catiaSkizze.set_Name("Kreis");

            // Skizze oeffnen
            Factory2D catFactory2D1 = hsp_catiaSkizze.OpenEdition();

            double H0 = 0;
            double V0 = 0;
            Point2D Ursprung = catFactory2D1.CreatePoint(H0, V0);
            Circle2D Kreis = catFactory2D1.CreateCircle(H0, V0, d/2, 0, 0);
            Kreis.CenterPoint = Ursprung;

            hsp_catiaSkizze.CloseEdition();

            Reference RefmySchaft = hsp_catiaPart.Part.CreateReferenceFromObject(hsp_catiaSkizze);

            Pad myPad = SF.AddNewPadFromRef(RefmySchaft, l);
            hsp_catiaPart.Part.Update();

            Console.WriteLine("GeometricElement");
            GeometricElements elements = hsp_catiaPart.Part.GeometricElements;
            int ii = 0;
            foreach (GeometricElement element in elements)
            {
                Console.WriteLine(ii + " " + element.GeometricType);
                ii++;
            }

//            Console.WriteLine("HybridBodies");
//            HybridBodies hbs = hsp_catiaPart.Part.HybridBodies;
//            for (int kk = 1; kk <= hbs.Count; kk++)
//            {
//                Console.WriteLine(kk + " " + hbs.Item(kk) + " " + hbs.Item(kk).get_Name());
//            }

            Console.WriteLine("Selection");
            Selection sel = hsp_catiaPart.Selection;
            sel.Add(hsp_catiaPart.Part);
            sel.Search("Topologie.Teilfläche,sel");
            for (int jj = 1; jj <= sel.Count2; jj++)
            {
                Face e = (Face) sel.Item(jj).Value;
                Console.WriteLine(jj);
                Console.WriteLine(e.DisplayName + " ");
//                if (jj==3)
//                {
//                    hsp_catiaPart.Selection.Clear();
//                    hsp_catiaPart.Selection.Add(e);
//                    hsp_catiaPart.Part.Update();
//                }
            }

            // Face aussenFlaeche = (Face)sel.Item(3).Value;
            Reference RefMantelflaeche = hsp_catiaPart.Part.CreateReferenceFromBRepName(
                "RSur:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;1)));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", myPad);

//            Face frontFlaeche = (Face)sel.Item(1).Value;
            Reference RefFrontflaeche = hsp_catiaPart.Part.CreateReferenceFromBRepName(
                "RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", myPad);

            PARTITF.Thread thread1 = SF.AddNewThreadWithOutRef();
            thread1.Side = CatThreadSide.catRightSide;
            thread1.Diameter = 8.000000;
            thread1.Depth = 50.000000;
            thread1.LateralFaceElement = RefMantelflaeche;
            thread1.LimitFaceElement = RefFrontflaeche;

            //            Face frontFlaeche = (Face)sel.Item(1).Value;
            //            selectoinSplit = frontFlaeche.DisplayName.Split(';');
            //            String frontFlaecheName = selectoinSplit[0] + ";" + selectoinSplit[1] + ";" + selectoinSplit[2] +
            //                ";None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)";
            //            Console.WriteLine(frontFlaecheName);
            //            Reference RefFrontflaeche = hsp_catiaPart.Part.CreateReferenceFromBRepName(
            //                frontFlaecheName, myPad);

            thread1.CreateUserStandardDesignTable("Metric_Thick_Pitch", @"C:\Program Files\Dassault Systemes\B28\win_b64\resources\standard\thread\Metric_Thick_Pitch.xml");
            thread1.Diameter = 8.000000;
            thread1.Pitch = 1.250000;

            hsp_catiaPart.Part.Update();
        }

        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            hsp_catiaSkizze.SetAbsoluteAxisData(arr);
        }

        public void ErzeugeProfil(Double b, Double h)
        {
            // Skizze umbenennen
            hsp_catiaSkizze.set_Name("Rechteck");

            // Rechteck in Skizze einzeichnen
            // Skizze oeffnen
            Factory2D catFactory2D1 = hsp_catiaSkizze.OpenEdition();

            // Rechteck erzeugen

            // erst die Punkte
            Point2D catPoint2D1 = catFactory2D1.CreatePoint(-50, 50);
            Point2D catPoint2D2 = catFactory2D1.CreatePoint(50, 50);
            Point2D catPoint2D3 = catFactory2D1.CreatePoint(50, -50);
            Point2D catPoint2D4 = catFactory2D1.CreatePoint(-50, -50);

            // dann die Linien
            Line2D catLine2D1 = catFactory2D1.CreateLine(-50, 50, 50, 50);
            catLine2D1.StartPoint = catPoint2D1;
            catLine2D1.EndPoint = catPoint2D2;

            Line2D catLine2D2 = catFactory2D1.CreateLine(50, 50, 50, -50);
            catLine2D2.StartPoint = catPoint2D2;
            catLine2D2.EndPoint = catPoint2D3;

            Line2D catLine2D3 = catFactory2D1.CreateLine(50, -50, -50, -50);
            catLine2D3.StartPoint = catPoint2D3;
            catLine2D3.EndPoint = catPoint2D4;

            Line2D catLine2D4 = catFactory2D1.CreateLine(-50, -50, -50, 50);
            catLine2D4.StartPoint = catPoint2D4;
            catLine2D4.EndPoint = catPoint2D1;

            // Skizzierer verlassen
            hsp_catiaSkizze.CloseEdition();
            // Part aktualisieren
            hsp_catiaPart.Part.Update();
        }

        public void ErzeugeBalken(Double l)
        {
            // Hauptkoerper in Bearbeitung definieren
            hsp_catiaPart.Part.InWorkObject = hsp_catiaPart.Part.MainBody;

            // Block(Balken) erzeugen
            ShapeFactory catShapeFactory1 = (ShapeFactory)hsp_catiaPart.Part.ShapeFactory;
            Pad catPad1 = catShapeFactory1.AddNewPad(hsp_catiaSkizze, l);

            // Block umbenennen
            catPad1.set_Name("Balken");

            // Part aktualisieren
            hsp_catiaPart.Part.Update();
        }



    }
}
