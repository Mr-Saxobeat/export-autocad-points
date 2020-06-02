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
        // EXPORTA PONTOS DO AMBIENTE PARA UM ARQUIVO .CSV
        // Este comando lê polylines selecionadas que representam as
        // paredes do ambiente e escreve num arquivo .csv 
        // as informações na seguinte ordem: 
        // - um index do ambiente (id),
        // - o tipo (parede, ou outra coisa), 
        // - a descrição (texto encontrado dentro do polyline), 
        // - coordenadas x de seus pontos e 
        // - coordenadas y de seus pontos
        [CommandMethod("EPA")]
        public void Main()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                #region Create selection set filter
                // Cria uma array do tipo TypedValue de tamanho 1
                TypedValue[] acTypValAr = new TypedValue[1];

                // Atribui o valor do index 0, este valor é um par ordenado
                // DxfCode.Start é o tipo (Objeto) e o "LWPOLYLINE" especifica que o 
                // objeto a ser selecionado é uma polyline.
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                //acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "LANG-PAREDE"), 1);
                
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
                    string curDwgPath = AcAp.GetSystemVariable("DWGPREFIX").ToString();    
                    string csvPath = curDwgPath + "\\pontos-do-ambiente.csv";
                    FileStream fileCsv = new FileStream(csvPath, FileMode.Create);
                    StreamWriter strWrt = new StreamWriter(fileCsv, Encoding.UTF8);
                    SelectionSet acSSet = acSSPrompt.Value;
                    Polyline objPoly;
                    Point3dCollection acPt3dCol;
                    MText objMText;
                    string sTipo;
                    string sDesc = "";
                    int id = 0;

                    strWrt.WriteLine("id;tipo;desc;coord x;coord y");

                    foreach (SelectedObject aObj in acSSet)
                    {

                        objPoly = tr.GetObject(aObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if(objPoly.Layer.Length > 5)
                        {
                            sTipo = objPoly.Layer.ToString().Substring(5);
                            acPt3dCol = GetCoordsOf(objPoly);

                            acTypValAr = new TypedValue[2];
                            acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 0);
                            acTypValAr.SetValue(new TypedValue((int)DxfCode.Text, "[A-Z]*"), 1);

                            acSelFtr = new SelectionFilter(acTypValAr);
                            acSSPrompt = ed.SelectWindowPolygon(acPt3dCol, acSelFtr);

                            if (acSSPrompt.Status == PromptStatus.OK)
                            {
                                SelectionSet acSSInsideEnv = acSSPrompt.Value;
                                ObjectId[] idMText = new ObjectId[acSSInsideEnv.Count];
                                idMText = acSSInsideEnv.GetObjectIds();

                                objMText = tr.GetObject(idMText[0], OpenMode.ForRead) as MText;
                                sDesc = objMText.Text;
                            }

                            // Escreve os dados do ambiente no arquivo
                            strWrt.Write(id + ";" + sTipo + ";" + sDesc + ";");
                            foreach(Point3d pt in acPt3dCol) strWrt.Write(" " + pt.X.ToString("n2") + ",");
                            strWrt.Write(";");
                            foreach (Point3d pt in acPt3dCol) strWrt.Write(" " + pt.Y.ToString("n2") + ",");
                            strWrt.WriteLine(";");

                            id++;
                        }
                    }

                    ed.WriteMessage("Arquivo criado em: " + csvPath);
                    strWrt.Close();
                }
                tr.Commit();
            }
        }

        public Point3dCollection GetCoordsOf(Polyline objPoly)
        {
            Point3d pt;
            Point3dCollection point3DCollection = new Point3dCollection();
            int numVert;

            numVert = objPoly.NumberOfVertices;

            // Iterar com cada ponto da polyline e concatenar
            // todas coordenadas X e Y numa array.
            if (numVert > 4) numVert--; // Se o ambiente tem mais que 4 pontos, o último não é necessário, pois é o mesmo ponto inicial
            for (int i = 0; i < numVert; i++)
            {
                pt = objPoly.GetPoint3dAt(i);
                point3DCollection.Add(pt);
            }

            return point3DCollection;
        }
    }
}
