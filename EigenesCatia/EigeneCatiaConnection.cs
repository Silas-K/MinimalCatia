using INFITF;
using MECMOD;
using PARTITF;
using System;
using System.Windows;

namespace EigenesCatia
{
    public class EigeneCatiaConnection
    {
        INFITF.Application catiaApp;
        MECMOD.PartDocument catiaPart;
        Sketch catiaProfil;


        public bool CatiaLaeuft()
        {
            try
            {
                object catiaObject = System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
                catiaApp = (INFITF.Application)catiaObject;
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }


        public void ErzeugePart()
        {

            Documents catDocuments = catiaApp.Documents;

            catiaPart = (PartDocument)catDocuments.Add("Part");


        }

        public void ErzeugeLeereSkizze()
        {
            HybridBodies catHybridBodies1 = catiaPart.Part.HybridBodies;
            HybridBody catHybridBody1;

            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {

                MessageBox.Show("Kein geometrisches Set gefunden! " + Environment.NewLine +
                    "Ein PART manuell erzeugen und einmal darauf achten, dass 'Geometisches Set' aktiviert ist.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            catHybridBody1.set_Name("Profile");

            Sketches catSketches = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = catiaPart.Part.OriginElements;
            Reference catRef = (Reference)catOriginElements.PlaneYZ;
            catiaProfil = catSketches.Add(catRef);


            ErzeugeAchsensystem();

            catiaPart.Part.Update();
        }

        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            catiaProfil.SetAbsoluteAxisData(arr);
        }

        public void ErzeugeRechteckProfil(double b, double h)
        {
            b /= 2;
            h /= 2;
            catiaProfil.set_Name("Rechteck ");

            Factory2D catFactory2D = catiaProfil.OpenEdition();

            Point2D catP2D_1 = catFactory2D.CreatePoint(-b, h);
            Point2D catP2D_2 = catFactory2D.CreatePoint(b, h);
            Point2D catP2D_3 = catFactory2D.CreatePoint(b, -h);
            Point2D catP2D_4 = catFactory2D.CreatePoint(-b, -h);

            Line2D catL2D_1 = catFactory2D.CreateLine(-b, h, b, h);
            catL2D_1.StartPoint = catP2D_1;
            catL2D_1.EndPoint = catP2D_2;

            Line2D catL2D_2 = catFactory2D.CreateLine(b, h, b, -h);
            catL2D_2.StartPoint = catP2D_2;
            catL2D_2.EndPoint = catP2D_3;

            Line2D catL2D_3 = catFactory2D.CreateLine(b, -h, -b, -h);
            catL2D_3.StartPoint = catP2D_3;
            catL2D_3.EndPoint = catP2D_4;

            Line2D catL2D_4 = catFactory2D.CreateLine(-b, -h, -b, h);
            catL2D_4.StartPoint = catP2D_4;
            catL2D_4.EndPoint = catP2D_1;


            catiaProfil.CloseEdition();
            catiaPart.Part.Update();


        }

        public void ErzeugeKreisProfil(double d)
        {
            catiaProfil.set_Name("Kreis");

            Factory2D catFactory2D = catiaProfil.OpenEdition();


            Point2D center = catFactory2D.CreatePoint(0, 0);
            Circle2D circle = catFactory2D.CreateClosedCircle(0, 0, d / 2);
            circle.CenterPoint = center;

            catiaProfil.CloseEdition();
            catiaPart.Part.Update();

        }

        public void ErzeugeBalken(double l)
        {
            //Hauptkörper in Bearbeitung
            catiaPart.Part.InWorkObject = catiaPart.Part.MainBody;
            ShapeFactory shapeFac = (ShapeFactory)catiaPart.Part.ShapeFactory;


            //BLOCK
            Pad newPad = shapeFac.AddNewPad(catiaProfil, l);
            newPad.set_Name("Block");



            catiaPart.Part.Update();




            //FASE

            Reference reference1 = catiaPart.Part.CreateReferenceFromName("");

            Chamfer chamfer = shapeFac.AddNewChamfer(
                reference1,
                CatChamferPropagation.catTangencyChamfer,
                CatChamferMode.catLengthAngleChamfer,
                CatChamferOrientation.catNoReverseChamfer,
                1,
                45);

            Reference reference2 = catiaPart.Part.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", newPad);

            /*("REdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;3)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1);*/

            //"RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1)

            chamfer.AddElementToChamfer(reference2);
            chamfer.Mode = CatChamferMode.catLengthAngleChamfer;
            chamfer.Propagation = CatChamferPropagation.catTangencyChamfer;
            chamfer.Orientation = CatChamferOrientation.catNoReverseChamfer;

            catiaPart.Part.Update();


            Reference reference3 = catiaPart.Part.CreateReferenceFromName("");

            Chamfer chamfer2 = shapeFac.AddNewChamfer(
                reference3,
                CatChamferPropagation.catTangencyChamfer,
                CatChamferMode.catLengthAngleChamfer,
                CatChamferOrientation.catNoReverseChamfer,
                1,
                45);

            Reference reference4 = catiaPart.Part.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.1;1);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", newPad);

            /*("REdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;3)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1);*/

            //"RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1)
            chamfer2.AddElementToChamfer(reference4);
            chamfer2.Mode = CatChamferMode.catLengthAngleChamfer;
            chamfer2.Propagation = CatChamferPropagation.catTangencyChamfer;
            chamfer2.Orientation = CatChamferOrientation.catNoReverseChamfer;


            catiaPart.Part.Update();
        }

    }
}
