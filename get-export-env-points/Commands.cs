using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AcadPlugin
{
    public class Commands
    {
        [CommandMethod("TEST")]
        public void Test()
        {

            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                #region Create selection set filter
                // Cria uma array do tipo TypedValue de tamanho 1
                TypedValue[] acTypValAr = new TypedValue[2];

                // Atribui o valor do index 0, este valor é um par ordenado
                // DxfCode.Start é o tipo (Objeto) e o "LWPOLYLINE" especifica que o 
                // objeto a ser selecionado é uma polyline.
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "LANG-PAREDE"), 1);
                
                // Atribui o array de par ordenado ao filtro de seleção
                SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
                #endregion

                // Cria o selection set e pede ao usuário para selecionar
                // através do método GetSelection
                PromptSelectionResult acSSPrompt = ed.GetSelection(acSelFtr);

                // Se a seleção estiver OK, quer dizer que 
                // os objetos foram selecionados
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;
                    StringBuilder[] coords = new StringBuilder[2];

                    StringBuilder csv = new StringBuilder();
                    Polyline objPoly;
                    MText objMText;
                    string desc;
                    Point2d pt = new Point2d();
                    int numVert;
                    int cont = 0;

                    csv.AppendLine("id;tipo;desc;coord x;coord y");
                    ObjectId[] idArray = new ObjectId[acSSet.Count];
                    idArray = acSSet.GetObjectIds();
                    
                    foreach(ObjectId objID in idArray)
                    {
                        objPoly = tr.GetObject(objID, OpenMode.ForRead) as Polyline;

                        Point3dCollection arPts = new Point3dCollection();
                        GetCoordsOf(objPoly, arPts);

                        // Iterar com cada ponto da polyline
                        numVert = objPoly.NumberOfVertices;


                        for (int j = 0; j < numVert; j++)
                        {
                            arPts.Add(objPoly.GetPoint3dAt(j));
                            
                        }

                        

                        acTypValAr = new TypedValue[2];
                        acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 0);
                        acTypValAr.SetValue(new TypedValue((int)DxfCode.Text, "[A-Z]*"), 1);

                        acSelFtr = new SelectionFilter(acTypValAr);
                        acSSPrompt = ed.SelectWindowPolygon(arPts, acSelFtr);

                        desc = "Sem descrição";
                        if(acSSPrompt.Status == PromptStatus.OK)
                        {
                            acSSet = acSSPrompt.Value;
                            ObjectId[] idMText = new ObjectId[acSSet.Count];
                            idMText = acSSet.GetObjectIds();

                            objMText = tr.GetObject(idMText[0], OpenMode.ForRead) as MText;
                            desc = objMText.Text;
                        }


                        if (numVert > 4) numVert--;
                        for (int j = 0; j < numVert; j++)
                        {
                            pt = objPoly.GetPoint2dAt(j);

                            xCoord.Append(pt.X.ToString("n2") + ",");
                            yCoord.Append(pt.Y.ToString("n2") + ",");
                        }

                        csv.AppendLine(cont + ";" + objPoly.Layer.ToString().Substring(5) + ";" + desc + ";" + xCoord + ";" + yCoord + ";");
                        cont++;
                    }



                    // Mudar o local de salvar **************************************************************************************************************************************************************************************
                    // **************************************************************************************************************************************************************************************
                    if (Directory.Exists("C:/Users/" + Environment.UserName + "." + Environment.UserDomainName))
                    {
                        File.WriteAllText("C:/Users/" + Environment.UserName + "." + Environment.UserDomainName + "/Desktop/test" + DateTime.Now.Millisecond.ToString() + ".csv",
                        csv.ToString());
                    }
                    else
                    {
                        File.WriteAllText("C:/Users/" + Environment.UserName + "/Desktop/test" + DateTime.Now.Millisecond.ToString() + ".csv",
                        csv.ToString());
                    }
                    // **************************************************************************************************************************************************************************************
                    // **************************************************************************************************************************************************************************************




                }

                tr.Commit();
            }
        }

        public void GetCoordsOf(Polyline objPoly, Point3dCollection arPts)
        {
         
        }
    }
}
