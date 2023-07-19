using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using System.Linq;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.IO;
using AcGi = Autodesk.AutoCAD.GraphicsInterface;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MgdAcApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using MgdAcDocument = Autodesk.AutoCAD.ApplicationServices.Document;
using AcWindowsNS = Autodesk.AutoCAD.Windows;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace insereNAP
{
    public class ThisDrawing
    {
        public static Document AcadDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        public static Point2d pt = Point2d.Origin; // initialize the start point to (0,0)
        public string teste;
        public static Entity blRefer;
        public static BlockReference Sup;
        public static bool criou = false;
        public static Point3d point = Point3d.Origin;
        public static string nome;
        public class BlockMoving : EntityJig

        {
            #region Fields

            public int mCurJigFactorNumber = 1;

            private Point3d mPosition = new Point3d();    // Factor #1

            #endregion

            #region Constructors

            public BlockMoving(Entity ent)
                : base(ent)
            {
            }

            #endregion

            #region Overrides

            protected override bool Update()
            {
                switch (mCurJigFactorNumber)
                {
                    case 1:
                        (Entity as BlockReference).Position = mPosition;
                        break;
                    default:
                        return false;
                }

                return true;
            }

            protected override SamplerStatus Sampler(JigPrompts prompts)
            {
                switch (mCurJigFactorNumber)
                {
                    case 1:
                        JigPromptPointOptions prOptions1 = new JigPromptPointOptions("\nClique no ponto de inserção da etiqueta:");
                        PromptPointResult prResult1 = prompts.AcquirePoint(prOptions1);
                        if (prResult1.Status == PromptStatus.Cancel) return SamplerStatus.Cancel;

                        if (prResult1.Value.Equals(mPosition))
                        {
                            return SamplerStatus.NoChange;
                        }
                        else
                        {
                            mPosition = prResult1.Value;
                            return SamplerStatus.OK;
                        }
                    default:
                        break;
                }

                return SamplerStatus.OK;
            }

            #endregion

            #region Method to Call

            public static bool Jig(BlockReference ent)
            {
                try
                {
                    Editor ed = MgdAcApplication.DocumentManager.MdiActiveDocument.Editor;
                    BlockMoving jigger = new BlockMoving(ent);
                    PromptResult pr;

                    do
                    {
                        pr = ed.Drag(jigger);
                        jigger.mCurJigFactorNumber++;
                    } while (pr.Status != PromptStatus.Cancel &&
                                pr.Status != PromptStatus.Error &&
                                pr.Status != PromptStatus.Keyword &&
                                jigger.mCurJigFactorNumber <= 1);


                    return pr.Status == PromptStatus.OK;
                }
                catch
                {
                    return false;
                }
            }

            #endregion
        }
        public static string GetXData(Entity _entity, string _IdXdata)
        {
            string _valuexdata = "";
            if (_entity.XData != null)
            {
                int _count = _entity.XData.AsArray().Length;
                TypedValue[] _xdata = _entity.XData.AsArray();
                for (int i = 0; i < _count; i++)
                {
                    if (_xdata[i].Value.ToString() == _IdXdata)
                    {
                        _valuexdata = _xdata[i + 1].Value.ToString();
                        break;
                    }
                    i++;
                }
            }
            return _valuexdata;
        }

        public static void SetXData(Entity _entity, string _IdXdata, string _value)
        {
            if (_entity.XData != null)
            {
                if (GetXData(_entity, _IdXdata) != "")
                {
                    int _count = _entity.XData.AsArray().Length;
                    TypedValue[] _xdata = _entity.XData.AsArray();
                    List<TypedValue> _lstxData = new List<TypedValue>();
                    for (int i = 0; i < _count; i++)
                    {
                        if (_xdata[i].Value.ToString() == _IdXdata)
                        {
                            _lstxData.Add(_xdata[i]);
                            _lstxData.Add(new TypedValue(1000, _value));
                        }
                        else
                        {
                            _lstxData.Add(_xdata[i]);
                            _lstxData.Add(_xdata[i + 1]);
                        }
                        i++;
                    }

                    TypedValue[] _tpValues = _lstxData.ToArray();
                    _entity.UpgradeOpen();
                    _entity.XData = new ResultBuffer(_tpValues);
                    _entity.DowngradeOpen();
                }
                else
                {
                    List<TypedValue> _tpValues = new List<TypedValue>();


                    //if (_entity.XData.AsArray().ToList().Find(x => x.TypeCode == 1001 && x.Value.ToString() == _IdXdata.ToString()) == null)
                    if (!_entity.XData.AsArray().ToList().Exists(x => x.TypeCode == 1001 && x.Value.ToString() == _IdXdata.ToString()))
                    {
                        _tpValues.Add(new TypedValue(1001, _IdXdata));
                        _tpValues.Add(new TypedValue(1000, _value));

                        foreach (TypedValue tv in _entity.XData)
                            _tpValues.Add(tv);
                    }
                    else
                    {
                        List<TypedValue> _aux = _entity.XData.AsArray().ToList();

                        foreach (TypedValue tv in _aux)
                        {
                            int index = _aux.IndexOf(tv);

                            if (index > 0)
                            {
                                if (_aux[index - 1].TypeCode == 1001 && _aux[index - 1].Value.ToString() == _IdXdata)
                                {
                                    _tpValues.Add(new TypedValue(1000, _value));
                                }
                                else
                                {
                                    _tpValues.Add(tv);
                                }
                            }
                            else if (tv.TypeCode == 1001)
                            {
                                _tpValues.Add(tv);
                            }

                        }
                    }
                    ResultBuffer _resBulffer = new ResultBuffer(_tpValues.ToArray());

                    _entity.UpgradeOpen();
                    _entity.XData = _resBulffer;
                    _entity.DowngradeOpen();
                }
            }
            else
            {
                ResultBuffer _newData = new ResultBuffer(new TypedValue(1001, _IdXdata),
                    new TypedValue(1000, _value));
                _entity.UpgradeOpen();
                _entity.XData = _newData;
                _entity.DowngradeOpen();
            }
        }

        [CommandMethod("InserirNAP")]
        public static void InserirNAP()
        {
            InserirBloco("GPON-CTO", "NET-NAP");
        }

        [CommandMethod("InserirCXEMENDA")]
        public static void InserirCXEMENDA()
        {
            InserirBloco("GPON-CEO", "NET-CX-EMENDA");

            //InserirCaixa("GPON-CEO", "NET-CX-EMENDA");
        }
                
        [CommandMethod("Unifilar")]//Exemplo de selecção de uma entidade através dos "SelectionSet".
        public void Unifilar()
        {
            //Setup
            Document doc = AcadDoc;//1º Abrimos o documento activo e passamo-lo para a variável "doc" que é do tipo "Document".
            Editor ed = doc.Editor;//2º Abrimos o editor do documento actual e passamo-lo para a variável "ed" que é do tipo Editor.
            Database db = AcadDoc.Database;//3º Abrimos a base de dados do documento actual e passamo-lo para a variável "db" que é do tipo Database.
            /*
             * A classe "Database" representa o ficheiro do desenho AutoCAD, cada
             * objecto "Database" contêm as várias variáveis de cabeçalho "header variables",
             * Tabelas de simbolos "Symbol tables", "Table Records", entidades e objectos que
             * são o que fazem um desenho. Esta classe contem funções que permitem o acesso a
             * todos os "Symbol Tables", a ler e a escrever um ficheiro DWG, a aceder e a modificar
             * os defaults da base de dados, a executar várias operações ao nivel da base de dados
             * tais como "wblock" e "deepCloneObjects", e a aceder e modificar todas as variáveis de cabeçalho "header variables".
             */
            //Opções para a selecção de uma unica entidade com selection sets
            PromptSelectionOptions Opso = new PromptSelectionOptions();//Opso é um novo objecto do tipo "PromptSelectionOptions"
            Opso.SingleOnly = true;//A propriedade "SingleOnly" indica se a selecção é feita só para uma entidade ou não, é do tipo boolean.
            Opso.MessageForAdding = "\nSelecione um objeto: ";
            PromptSelectionResult psr = ed.GetSelection(Opso);
            
            //se der erro
            if (psr.Status == PromptStatus.Error)
            {
                //retorna
                return;
            }
            //se cancelou a ação 
            if (psr.Status == PromptStatus.Cancel)
            {
                //retorna
                return;
            }           

            SelectionSet ss = psr.Value;//O resultado da selecção unica que está na variável "psr" é passado para uma variável "ss" do tipo "SelectionSet".
            ObjectId[] idarray = ss.GetObjectIds();//Passamos agora o "Object Id" do objecto seleccionado para uma variável "idarray" do tipo ObjectId[], sendo esta um array de ObjectId's.

            Transaction tr = db.TransactionManager.StartTransaction();//Damos incio ás transacções com a base de dados actual do documento AutoCAD
            /*
             * A partir do momento que começamos as Transacções com a base de dados podemos começar a alterar e a ler os objectos do desenho
             * no final de todas as operações temos que terminar as transações o que neste caso será com o método "tr.Commit()" se as Transações
             * não forem fechadas ou se usar o método "tr.Abort()" todas as alterações que efectuamos não serão gravadas na base de dados.
             *  A utilização do método "Abort()" é útil quando se quer abortar as operações efectuadas por um qualquer erro ou porque o
             * utilizador decidiu cancelar as operações que estava a efectuar.
             */
            using (tr)
            {
                foreach (ObjectId objid in idarray)
                {                    

                    DBObject obj = tr.GetObject(objid, OpenMode.ForWrite, true);//Recolhemos o objecto que foi passado do array de ObjectID[] e abrimo-lo no modo de escrita.
                    Entity entity = (Entity)obj;//Aqui fazemos uma conversão explicita de obj que é um objecto do tipo DBObject para uma classe Entity que deriva da classe DBOject.
                    ed.WriteMessage("\nVocê selecionou: {0}", entity.Layer);//Aqui escrevemos na linha de comandos o tipo de entidade seleccionada.
                    
                    //se for cabo de fibras                    
                    if (entity.Layer.Contains("FIBRA"))
                    {
                        int numr = entity.Layer.Length;
                        string qtdFibrasTexto = entity.Layer.Substring(9, numr -9).ToString();
                        int qtdFibras = int.Parse(qtdFibrasTexto);
                   
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        //tamanho
                        double length = 25;
                        //espaço para iniciar o desenho X, Y
                        pt += new Vector2d(3 , 70);

                        //cria o objeto polyline
                        Polyline po = new Polyline();
                        po.Layer = "NET-NOVO EQUIPAMENTO";
                        //começa no ponto indice = 0(primeiro ponto) começa no pt = (posição 0) rotação = 0
                        po.AddVertexAt(0, pt, 0.0, 0.0, 0.0); // add o primeiro vertice
                        pt += new Vector2d(length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        po.AddVertexAt(1, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto

                        //adiciona o objeto na tela 
                        btr.AppendEntity(po);
                        tr.AddNewlyCreatedDBObject(po, true);

                        //objeto antigo de suporte
                        Polyline sup = po;

                        // cria um laço para construir as 12 linhas 
                        for (int i = 0; i < qtdFibras; i++)
                        {
                            // distancia do offset da polyline
                            DBObjectCollection acDbObjColl = sup.GetOffsetCurves(0.5);
                            //Para cada entidade 
                            foreach (Entity acEnt in acDbObjColl)
                            {
                                //se nao for a primeira
                                if (i != 0)
                                {
                                    //adiciona a ultima na variavel de suporte
                                    sup = (Polyline)acEnt;
                                }

                                // Adiciona o objeto offset na tela 
                                btr.AppendEntity(acEnt);
                                tr.AddNewlyCreatedDBObject(acEnt, true);
                            }
                        }
                        //volta a distancia em y inicial
                        pt -= new Vector2d(0, 70);
                        blRefer = (Entity)po;
                        InserirBloco2("GeponD-NFB", "NET-NUM-FIBRA", po, qtdFibras.ToString());
                    }
                    #region se for OLT
                    if (entity.Layer == "SITE CLARO_3" || entity.Layer == "SITE CLARO")
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        double length = 78;
                        pt += new Vector2d(length * 2, 0.0);

                        Polyline rec = new Polyline();
                        //seta a layer para 0
                        rec.Layer = "0";
                        //seta a cor para azul
                        rec.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 255);
                        //começa no ponto indice = 0(primeiro ponto) começa no pt = (posição 0) rotação = 0
                        rec.AddVertexAt(0, pt, 0.0, 0.0, 0.0); // add o primeiro vertice
                        pt += new Vector2d(length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(1, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(2, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(-length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(3, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, -length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(4, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto                                             

                        btr.AppendEntity(rec);
                        tr.AddNewlyCreatedDBObject(rec, true);

                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(rec.GeometricExtents.MinPoint.X, rec.GeometricExtents.MaxPoint.Y + 0.300, 0);
                            acText.Height = 1.57;
                            acText.TextString = "OLT";
                            acText.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(0, 0, 255);

                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                        }
                        blRefer = (Entity)rec;
                        pt += new Vector2d(length, 0.0);
                       

                    }
                    #endregion
                    #region Se for CX Emenda
                    if (entity.Layer == "NET-CX-EMENDA")
                    {
                        //faz a verificação para ver se tem splitters
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //procura todos os atributos do bloco
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {
                                if (attRef.Tag == "NOME")
                                {
                                    nome = attRef.TextString;
                                }
                            }
                        }
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        double length = 78;
                        pt += new Vector2d(length * 2, 0.0);

                        Polyline rec = new Polyline();
                        //seta o layer para 0
                        rec.Layer = "0";
                        //seta a cor para branca
                        rec.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
                        //começa no ponto indice = 0(primeiro ponto) começa no pt = (posição 0) rotação = 0
                        rec.AddVertexAt(0, pt, 0.0, 0.0, 0.0); // add o primeiro vertice
                        pt += new Vector2d(length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(1, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(2, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(-length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(3, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, -length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(4, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto                     

                        //adiciona na tela 
                        btr.AppendEntity(rec);
                        tr.AddNewlyCreatedDBObject(rec, true);

                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(rec.GeometricExtents.MinPoint.X, rec.GeometricExtents.MaxPoint.Y + 0.300, 0);
                            acText.Height = 1.57;
                            acText.TextString = nome;
                            acText.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                        }
                        pt += new Vector2d(length, 0.0);
                        blRefer = (Entity)rec;
                    }
                    #endregion
                    #region se for MDU
                    if (entity.Layer == "Caixa externa exclusiva MDU")
                    {
                        //faz a verificação para ver se tem splitters
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //procura todos os atributos do bloco
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {                                
                                if (attRef.Tag == "NOME")
                                {
                                    nome = attRef.TextString;
                                }
                            }
                        }

                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        double length = 78;
                        pt += new Vector2d(length * 2, 0.0);

                        Polyline rec = new Polyline();
                        //seta a layer para 0
                        rec.Layer = "0";
                        //seta a cor para branco 
                        rec.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
                        //começa no ponto indice = 0(primeiro ponto) começa no pt = (posição 0) rotação = 0
                        rec.AddVertexAt(0, pt, 0.0, 0.0, 0.0); // add o primeiro vertice
                        pt += new Vector2d(length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(1, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(2, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(-length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(3, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, -length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(4, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto                     

                        btr.AppendEntity(rec);
                        tr.AddNewlyCreatedDBObject(rec, true);

                        //cria a caixa de texto acima do retangulo
                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(rec.GeometricExtents.MinPoint.X, rec.GeometricExtents.MaxPoint.Y + 0.300, 0);
                            acText.Height = 1.57;
                            acText.TextString = nome;
                            acText.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            //adiciona na tela
                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                        } 
                        pt += new Vector2d(length, 0.0); 
                        blRefer = (Entity)rec;
                    }
                    #endregion
                    #region Se for Primaria
                    if (entity.Layer == "NET-PRIMARIA")
                    {
                        
                        //faz a verificação para ver se tem splitters
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //procura todos os atributos do bloco
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {
                                //se encontrou a tag splitter
                                if (attRef.Tag == "QTS_SPLITTER")
                                {
                                    //pega a  quantidade 
                                    teste = attRef.TextString;
                                }
                                if (attRef.Tag == "NOME")
                                {
                                    nome = attRef.TextString;
                                }
                            }
                        }


                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        double length = 78;
                        pt += new Vector2d(length * 2, 0.0);

                        Polyline rec = new Polyline();
                        //seta a layer para 0
                        rec.Layer = "0";
                        //seta a cor para branco
                        rec.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);
                        //começa no ponto indice = 0(primeiro ponto) começa no pt = (posição 0) rotação = 0
                        rec.AddVertexAt(0, pt, 0.0, 0.0, 0.0); // add o primeiro vertice
                        pt += new Vector2d(length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(1, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(2, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(-length, 0.0); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(3, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto
                        pt += new Vector2d(0, -length); // almenta o tamanho para a posição do segundo ponto
                        rec.AddVertexAt(4, pt, 0.0, 0.0, 0.0);//add o segundo vertice no segundo ponto                      

                        //adiciona na tela 
                        btr.AppendEntity(rec);
                        tr.AddNewlyCreatedDBObject(rec, true);

                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(rec.GeometricExtents.MinPoint.X, rec.GeometricExtents.MaxPoint.Y + 0.300, 0);
                            acText.Height = 1.57;
                            acText.TextString = nome;
                            acText.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                        }
                        pt += new Vector2d(length, 0.0);
 
                        //guarda o bloco em uma variavel suporte
                        blRefer = (Entity)rec;
                        #region se tiver splitter
                        //se tiver splitter inseri o bloco dentro da caixa com a quantidade
                        if (teste != "") InserirBloco2("GeponD-SPLITTER", "NET-PRIMARIA-Diagrama", rec, teste);
                        #endregion
                    }
                    #endregion
                    #region se for spliter azul
                    if (entity.Layer == "NET-SPLITTER-SEC")
                    {
                        //verifica a quantidade de entradas
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //busca os atributos 
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            //procura a tag dentro dos atributos
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {
                                //se axou a tag 
                                if (attRef.Tag == "QTS_SPLITTER")
                                {
                                    //pega a quantidade
                                   teste = attRef.TextString;
                                }
                            }
                        }                        
                        //inseri o bloco com a  referencia 
                        InserirBloco2("GPON-AT_MDU", "NET-SPLITTER-SEC", entity, teste);
                        //guarda o bloco em uma variavel suporte
                    }
                    #endregion
                    #region se for splitter vermelho
                    if (entity.Layer == "NET-SPLITTER-PRI")
                    {
                        //verifica a quantidade de entradas
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //busca os atributos 
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            //procura a tag dentro dos atributos
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {
                                //se axou a tag 
                                if (attRef.Tag == "QTS_SPLITTER")
                                {
                                    //pega a quantidade
                                    teste = attRef.TextString;
                                }
                            }
                        }    
                        //inseri o bloco com a  referencia 
                        InserirBloco2("GPON-AT_MDU", "NET-SPLITTER-PRI", entity, teste);
                    }
                    #endregion
                    #region se for NAP
                    if (entity.Layer == "NET-NAP")
                    {
                        //verifica a quantidade de entradas
                        BlockReference _blockRef = null;
                        ObjectId _objID = db.GetObjectId(false, objid.Handle, 0);
                        _blockRef = tr.GetObject(_objID, OpenMode.ForWrite) as BlockReference;
                        //busca os atributos 
                        foreach (ObjectId entId in _blockRef.AttributeCollection)
                        {
                            DBObject obj1 = entId.GetObject(OpenMode.ForRead);
                            //procura a tag dentro dos atributos
                            using (AttributeReference attRef = obj1 as AttributeReference)
                            {
                                //se axou a tag 
                                if (attRef.Tag == "QTS_SPLITTER")
                                {
                                    //pega a quantidade
                                    teste = attRef.TextString;
                                }
                                //se axou a tag 
                                if (attRef.Tag == "NOME")
                                {
                                    //pega a quantidade
                                    nome = attRef.TextString;
                                }
                            }
                        }    
                        //inseri o bloco com a  referencia 
                        InserirBloco2("GPON-CTO", "NET-NAP", entity, teste);
                        
                    }
                    #endregion
                }                    
                tr.Commit();//Aqui Terminamos as nossas transacções.
            }

            Unifilar();
        }

        [CommandMethod("InserirMDU")]
        public static void InserirMDU()
        {
            InserirBloco("GPON-CEO-MDU", "Caixa externa exclusiva MDU");
        }

        [CommandMethod("InserirPRIMARIA")]
        public static void InserirPRIMARIA()
        {
            InserirBloco("GPON-CEOS", "NET-PRIMARIA");
        }

        [CommandMethod("InserirRESERVA")]
        public static void InserirRESERVA()
        {
            InserirRESERVA("GPON-R_ST01", "NET-SobraTecnica");
        }

        public static void InserirBloco2(String nomeBloco, String nomeLayer, Entity entity, string texto)
        {

            Document doc = AcadDoc;//1º Abrimos o documento activo e passamo-lo para a variável "doc" que é do tipo "Document".
            Editor ed = doc.Editor;//2º Abrimos o editor do documento actual e passamo-lo para a variável "ed" que é do tipo Editor.
            Database db = AcadDoc.Database;//3º Abrimos a base de dados do documento actual e passamo-lo para a variável "db" que é do tipo Database.
            Transaction tr = db.TransactionManager.StartTransaction();
            //pega o valor do x e y da posicao
            double x = pt.X;
            double y = pt.Y;

            using (tr)
            {
                //Get the block definition "Check".
                string blockName = nomeBloco;

                BlockTable bt = db.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                BlockTableRecord blockDef = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;

                //Also open modelspace - we'll be adding our BlockReference to it
                BlockTableRecord btr = bt[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                //Point3d point = new Point3d();
                //Autodesk.AutoCAD.Internal.Utils.EntLast();
                //point recebe as cordenadas do pt
                point = new Point3d(x , y , 0);
                //incrementa
                point += new Vector3d(10, 10, 0);
                //adiciona um bloco novo 
                BlockReference blockRef = new BlockReference(point, bt[nomeBloco]);
                //volta a posição anterior 
                point -= new Vector3d(0, 10, 0);
                //pega os novos valores 
                double Xp = point.X;
                double Yp = point.Y;
                //altera o valor da variavel original 
                pt = new Point2d(Xp, Yp);
                
                //cria o spliter dentro de caixa de emenda
                #region Se for Splitter
                if (nomeBloco == "GeponD-SPLITTER")
                {
                    int qtd = int.Parse(texto);
                    //Create new BlockReference, and link it to our block definition
                    Point3d point2 = new Point3d(entity.GeometricExtents.MaxPoint.X - 39, entity.GeometricExtents.MaxPoint.Y / qtd, 0);
                    
                    for (int j = 0; j < qtd; j++)
                    {
                        BlockReference blockRef1 = new BlockReference(point2, bt[nomeBloco]);
                        BlockTableRecord blockDef1 = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        using (blockRef1)
                        {                            
                            blockRef1.Layer = nomeLayer;
                            blockRef1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            //Add the block reference to modelspace
                            btr.AppendEntity(blockRef1);
                            tr.AddNewlyCreatedDBObject(blockRef1, true);

                            //Iterate block definition to find all non-constant
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef1)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;

                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    //This is a non-constant AttributeDefinition
                                    //Create a new AttributeReference
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, blockRef1.BlockTransform);
                                        if (attRef.Tag == "NOME")
                                        {
                                            //pega a  definição do bloco e copia para o novo 
                                            attRef.TextString = "P-"+(j+1);
                                        }
                                        //Add the AttributeReference to the BlockReference
                                        blockRef1.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);

                                        //cria a linha  
                                        using (Polyline po = new Polyline())
                                        {
                                            Point2d pt2 = new Point2d(blockRef1.GeometricExtents.MinPoint.X, (blockRef1.GeometricExtents.MinPoint.Y + 2));
                                            po.AddVertexAt(0, pt2, 0, 0, 0);
                                            pt2 += new Vector2d(-5, 0);
                                            po.AddVertexAt(1, pt2, 0, 0, 0);
                                            po.Layer = "NET-NOVO EQUIPAMENTO";

                                            //adiciona na tela
                                            btr.AppendEntity(po);
                                            tr.AddNewlyCreatedDBObject(po, true);
                                        }
                                        //cria a linha  
                                        using (Polyline po = new Polyline())
                                        {
                                            Point2d pt2 = new Point2d((blockRef1.GeometricExtents.MaxPoint.X -0.5), (blockRef1.GeometricExtents.MaxPoint.Y - 0.25));
                                            po.AddVertexAt(0, pt2, 0, 0, 0);
                                            pt2 += new Vector2d(5, 0);
                                            po.AddVertexAt(1, pt2, 0, 0, 0);
                                            po.Layer = "NET-NOVO EQUIPAMENTO";

                                            //adiciona na tela
                                            btr.AppendEntity(po);
                                            tr.AddNewlyCreatedDBObject(po, true);

                                            Polyline sup = po;

                                            // Step through the new objects created
                                            for (int i = 0; i < 8; i++)
                                            {
                                                // Offset the polyline a given distance
                                                DBObjectCollection acDbObjColl = sup.GetOffsetCurves(0.5);
                                                foreach (Entity acEnt in acDbObjColl)
                                                {
                                                    if (i != 0)
                                                    {
                                                        sup = (Polyline)acEnt;
                                                    }
                                                    // Add each offset object
                                                    btr.AppendEntity(acEnt);
                                                    tr.AddNewlyCreatedDBObject(acEnt, true);
                                                }
                                            }
                                        }
                                    }
                                }
                                point2 =(Point3d)blockRef1.GeometricExtents.MinPoint + new Vector3d(0,10,0);                                
                            }                            
                        }
                        //Sup = blockRef1;
                    }
                    //guarda o bloco em uma variavel suporte
                    //blRefer = (Entity)Sup;
                }
                #endregion

                #region Se for Splitter
                if (nomeBloco == "GeponD-NFB")
                {
                    int qtd = int.Parse(texto);
                    //Create new BlockReference, and link it to our block definition
                    Point3d point2 = new Point3d(entity.GeometricExtents.MinPoint.X + 3, entity.GeometricExtents.MaxPoint.Y, 0);

                    for (int j = 0; j < qtd; j++)
                    {
                        BlockReference blockRef1 = new BlockReference(point2, bt[nomeBloco]);
                        BlockTableRecord blockDef1 = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        using (blockRef1)
                        {
                            blockRef1.Layer = nomeLayer;
                            blockRef1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            //Add the block reference to modelspace
                            btr.AppendEntity(blockRef1);
                            tr.AddNewlyCreatedDBObject(blockRef1, true);

                            //Iterate block definition to find all non-constant
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef1)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;

                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    //This is a non-constant AttributeDefinition
                                    //Create a new AttributeReference
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, blockRef1.BlockTransform);
                                        if (attRef.Tag == "NFB")
                                        {
                                            //pega a  definição do bloco e copia para o novo 
                                            attRef.TextString = (j + 1).ToString();
                                        }
                                        //Add the AttributeReference to the BlockReference
                                        blockRef1.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);
                                        
                                    }
                                }
                                point2 = new Point3d(blockRef1.GeometricExtents.MinPoint.X+0.25, blockRef1.GeometricExtents.MinPoint.Y -0.25, 0);
                            }
                        }
                        //Sup = blockRef1;
                    }
                    point2 = new Point3d(entity.GeometricExtents.MaxPoint.X - 3, entity.GeometricExtents.MaxPoint.Y, 0);
                    for (int j = 0; j < qtd; j++)
                    {
                        BlockReference blockRef1 = new BlockReference(point2, bt[nomeBloco]);
                        BlockTableRecord blockDef1 = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;
                        using (blockRef1)
                        {
                            blockRef1.Layer = nomeLayer;
                            blockRef1.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255);

                            //Add the block reference to modelspace
                            btr.AppendEntity(blockRef1);
                            tr.AddNewlyCreatedDBObject(blockRef1, true);

                            //Iterate block definition to find all non-constant
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef1)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;

                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    //This is a non-constant AttributeDefinition
                                    //Create a new AttributeReference
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, blockRef1.BlockTransform);
                                        if (attRef.Tag == "NFB")
                                        {
                                            //pega a  definição do bloco e copia para o novo 
                                            attRef.TextString = (j + 1).ToString();
                                        }
                                        //Add the AttributeReference to the BlockReference
                                        blockRef1.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);

                                    }
                                }
                                point2 = new Point3d(blockRef1.GeometricExtents.MinPoint.X + 0.25, blockRef1.GeometricExtents.MinPoint.Y - 0.25, 0);
                            }
                        }
                        //Sup = blockRef1;
                    }
                    //guarda o bloco em uma variavel suporte
                    //blRefer = (Entity)Sup;
                }
                #endregion
                //cria a nap
                #region Se for NAP
                else if (nomeBloco == "GPON-CTO")
                {
                    using (blockRef)
                    {
                        blockRef.Rotation = 4.71238898038;
                        blockRef.Layer = nomeLayer;
                        blockRef.Color = entity.Color;

                        //Add the block reference to modelspace
                        btr.AppendEntity(blockRef);
                        tr.AddNewlyCreatedDBObject(blockRef, true);

                         //Iterate block definition to find all non-constant
                        // AttributeDefinitions
                        foreach (ObjectId id in blockDef)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition attDef = obj as AttributeDefinition;

                            if ((attDef != null) && (!attDef.Constant))
                            {
                                //This is a non-constant AttributeDefinition
                                //Create a new AttributeReference
                                using (AttributeReference attRef = new AttributeReference())
                                {
                                    attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                    if (attRef.Tag == "QTS_SPLITTER")
                                        //pega a  definição do bloco e copia para o novo 
                                        attRef.TextString = texto;
                                    //se axou a tag nome
                                    if (attRef.Tag == "NOME")
                                    {
                                        //coloca o nome do bloco 
                                        attRef.TextString = nome;
                                        //rotaciona o texto
                                        attRef.Rotation = 1.57079632679;
                                    }
                                    //Add the AttributeReference to the BlockReference
                                    blockRef.AttributeCollection.AppendAttribute(attRef);
                                    tr.AddNewlyCreatedDBObject(attRef, true);
                                }
                            }
                        }
                        //cria a linha antes 
                        using (Polyline po = new Polyline())
                        {
                            Point2d pt2 = new Point2d(blockRef.GeometricExtents.MinPoint.X + 1.9, blockRef.GeometricExtents.MaxPoint.Y -3);
                            po.AddVertexAt(0, pt2, 0, 0, 0);
                            pt2 += new Vector2d(0, 9.5);
                            po.AddVertexAt(1, pt2, 0, 0, 0);
                            po.Layer = "NET-NOVO EQUIPAMENTO";

                            //adiciona na tela
                            btr.AppendEntity(po);
                            tr.AddNewlyCreatedDBObject(po, true);
                        }                        
                        }
                    criou = true;
                    
                }
                #endregion

                //cria o spliter
                #region Se for Splitter
                else
                    {
                        using (blockRef)
                        {
                            blockRef.Layer = nomeLayer;
                            blockRef.Color = entity.Color;

                            //Add the block reference to modelspace
                            btr.AppendEntity(blockRef);
                            tr.AddNewlyCreatedDBObject(blockRef, true);

                            //Iterate block definition to find all non-constant
                            // AttributeDefinitions
                            foreach (ObjectId id in blockDef)
                            {
                                DBObject obj = id.GetObject(OpenMode.ForRead);
                                AttributeDefinition attDef = obj as AttributeDefinition;

                                if ((attDef != null) && (!attDef.Constant))
                                {
                                    //This is a non-constant AttributeDefinition
                                    //Create a new AttributeReference
                                    using (AttributeReference attRef = new AttributeReference())
                                    {
                                        attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                        if (attRef.Tag == "QTS_SPLITTER")
                                            //pega a  definição do bloco e copia para o novo 
                                            attRef.TextString = texto;
                                        //Add the AttributeReference to the BlockReference
                                        blockRef.AttributeCollection.AppendAttribute(attRef);
                                        tr.AddNewlyCreatedDBObject(attRef, true);

                                        //cria a linha antes 
                                        using (Polyline po = new Polyline())
                                        {
                                            Point2d pt2 = new Point2d(blockRef.GeometricExtents.MinPoint.X, blockRef.GeometricExtents.MinPoint.Y + 0.625);
                                            po.AddVertexAt(0, pt2, 0, 0, 0);
                                            pt2 += new Vector2d(0, 1.5);
                                            po.AddVertexAt(1, pt2, 0, 0, 0);
                                            po.Color = entity.Color;
                                            po.Layer = nomeLayer;

                                            //adiciona na tela
                                            btr.AppendEntity(po);
                                            tr.AddNewlyCreatedDBObject(po, true);
                                        }
                                        //se for 1 entrada
                                        if (attRef.TextString == "1")
                                        {
                                            //cria a linha antes 
                                            using (Polyline po = new Polyline())
                                            {
                                                Point2d pt2 = new Point2d(blockRef.GeometricExtents.MinPoint.X, blockRef.GeometricExtents.MinPoint.Y + 1.25);
                                                po.AddVertexAt(0, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(-1.5, 0);
                                                po.AddVertexAt(1, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(0, 5);
                                                po.AddVertexAt(2, pt2, 0, 0, 0);
                                                po.Layer = "NET-NOVO EQUIPAMENTO";

                                                //adiciona na tela
                                                btr.AppendEntity(po);
                                                tr.AddNewlyCreatedDBObject(po, true);

                                            }
                                        }
                                        //se for 2 entradas
                                        if (attRef.TextString == "2")
                                        {
                                            //cria a linha antes 
                                            using (Polyline po = new Polyline())
                                            {
                                                Point2d pt2 = new Point2d(blockRef.GeometricExtents.MinPoint.X, blockRef.GeometricExtents.MinPoint.Y + 1.2);
                                                po.AddVertexAt(0, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(-1.5, 0);
                                                po.AddVertexAt(1, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(0, 5);
                                                po.AddVertexAt(2, pt2, 0, 0, 0);
                                                po.Layer = "NET-NOVO EQUIPAMENTO";

                                                //adiciona na tela
                                                btr.AppendEntity(po);
                                                tr.AddNewlyCreatedDBObject(po, true);
                                                // Offset the polyline a given distance
                                                DBObjectCollection acDbObjColl = po.GetOffsetCurves(0.5);
                                                foreach (Entity acEnt in acDbObjColl)
                                                {
                                                    // Add each offset object
                                                    btr.AppendEntity(acEnt);
                                                    tr.AddNewlyCreatedDBObject(acEnt, true);
                                                }
                                            }
                                        }
                                        //se for 2 entradas
                                        if (attRef.TextString == "3")
                                        {
                                            //cria a linha antes 
                                            using (Polyline po = new Polyline())
                                            {
                                                Point2d pt2 = new Point2d(blockRef.GeometricExtents.MinPoint.X, blockRef.GeometricExtents.MinPoint.Y + 0.75);
                                                po.AddVertexAt(0, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(-1.5, 0);
                                                po.AddVertexAt(1, pt2, 0, 0, 0);
                                                pt2 += new Vector2d(0, 5);
                                                po.AddVertexAt(2, pt2, 0, 0, 0);
                                                po.Layer = "NET-NOVO EQUIPAMENTO";

                                                //adiciona na tela
                                                btr.AppendEntity(po);
                                                tr.AddNewlyCreatedDBObject(po, true);

                                                Polyline sup = po;

                                                // Step through the new objects created
                                                for (int i = 0; i < 3; i++)
                                                {
                                                    // Offset the polyline a given distance
                                                    DBObjectCollection acDbObjColl = sup.GetOffsetCurves(0.5);
                                                    foreach (Entity acEnt in acDbObjColl)
                                                    {
                                                        if (i != 0)
                                                        {
                                                            sup = (Polyline)acEnt;
                                                        }
                                                        // Add each offset object
                                                        btr.AppendEntity(acEnt);
                                                        tr.AddNewlyCreatedDBObject(acEnt, true);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //guarda o bloco em uma variavel suporte
                            //blRefer = (Entity)blockRef;
                            criou = true;
                        }
                    }
                
    
                    //Our work here is done
                    tr.Commit();
            }
                #endregion

        }

        private static void InserirBloco(String nomeBloco, String nomeLayer)
        {
            int linha = 0;

            try
            {

                Database _AcadDB = AcadDoc.Database;
                linha = 1;

                BlockReference _blockReserva;
                Point3d _pontoSeleca;
                Polyline _caboLine = new Polyline();
                Polyline _pwLine = new Polyline();

                bool inseriu = false;

                using (Transaction _acTrans = _AcadDB.TransactionManager.StartTransaction())
                {
                    linha = 2;
                    BlockTable _blckTable = _acTrans.GetObject(_AcadDB.BlockTableId, OpenMode.ForRead) as BlockTable;


                    linha = 3;
                    if (_blckTable.Has(nomeBloco))
                    {
                        linha = 4;
                        //Aguarda a selecão de um ponto no mapa.
                        PromptPointResult _pointres = AcadDoc.Editor.GetPoint("Selecione o ponto de inserção");
                        Cursor.Current = Cursors.WaitCursor;

                        if (_pointres.Status != PromptStatus.OK) return;

                        _pontoSeleca = _pointres.Value;

                        linha = 5;
                        _blockReserva = new BlockReference(_pontoSeleca, _blckTable[nomeBloco]);

                        linha = 6;
                        PromptAngleOptions _doub = new Autodesk.AutoCAD.EditorInput.PromptAngleOptions("Selecione a rotação:");

                        _doub.BasePoint = _pontoSeleca;
                        _doub.UseBasePoint = true;

                        //Aguarda a seleção do valor de rotação do bloco.
                        PromptDoubleResult _dbresult = AcadDoc.Editor.GetAngle(_doub);
                        Cursor.Current = Cursors.WaitCursor;
                        Double _angPt = _dbresult.Value;

                        linha = 7;
                        BlockTableRecord _ModelSpace = _acTrans.GetObject(_AcadDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        linha = 8;
                        _ModelSpace.AppendEntity(_blockReserva);
                        _acTrans.AddNewlyCreatedDBObject(_blockReserva, true);

                        // if (BlockMoving.Jig(_blockReserva))
                        {
                            _pontoSeleca = new Point3d(_blockReserva.Position.X, _blockReserva.Position.Y, _blockReserva.Position.Z);
                            Point3d Pt1 = new Point3d(_blockReserva.Position.X + 0.2, _blockReserva.Position.Y + 0.2, _blockReserva.Position.Z);
                            Point3d Pt2 = new Point3d(_blockReserva.Position.X - 0.2, _blockReserva.Position.Y - 0.2, _blockReserva.Position.Z);


                            linha = 9;
                            TypedValue[] _values = { new TypedValue(0, "LWPOLYLINE") };

                            SelectionFilter _selfilter = new SelectionFilter(_values);
                            PromptSelectionResult _resSel = AcadDoc.Editor.SelectCrossingWindow(Pt1, Pt2, _selfilter);

                            SelectionSet _selSets = _resSel.Value;

                            if (_selSets != null)
                            {
                                ObjectId _objID = _selSets.GetObjectIds()[0];

                                int indice = -1;
                                _pwLine = _acTrans.GetObject(_objID, OpenMode.ForRead) as Polyline;

                                _blckTable = _acTrans.GetObject(_AcadDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                                LayerTable _layerTable = _acTrans.GetObject(_AcadDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                                //// Entity _pwLineEnt = _acTrans.GetObject(_objID, OpenMode.ForRead) as Entity;

                                double _angleFix = _angPt;

                                //// string _layName = _pwLine.Layer;
                                double Lwheight = Convert.ToDouble(0.2);
                                bool _inVertex = false;

                                _blockReserva.Rotation = _angleFix;
                                _blockReserva.Layer = nomeLayer;

                                int _idx = 0;

                                linha = 10;
                                BlockTableRecord _blockDefinition = _blckTable[nomeBloco].GetObject(OpenMode.ForRead) as BlockTableRecord;

                                linha = 11;
                                foreach (ObjectId EntId in _blockDefinition)
                                {

                                    linha = 12;
                                    DBObject obj = EntId.GetObject(OpenMode.ForRead);
                                    AttributeDefinition attDef = obj as AttributeDefinition;

                                    if ((attDef != null) && (!attDef.Constant))
                                    {
                                        using (AttributeReference attRef = new AttributeReference())
                                        {
                                            attRef.SetAttributeFromBlock(attDef, _blockReserva.BlockTransform);

                                            if (attRef.Tag.Equals("NOME"))
                                            {
                                                attRef.TextString = nomeLayer;

                                                if (nomeLayer.Equals("NET-NAP"))
                                                {
                                                    if (_angleFix > 1.5 && _angleFix <= 3.14)
                                                    {
                                                        attRef.Rotation = 6.28 - Math.Abs(_angleFix - 3.14);
                                                    }
                                                    else if (_angleFix > 3.14 && _angleFix <= 4.5)
                                                    {
                                                        attRef.Rotation = _angleFix - 3.14;
                                                    }
                                                }
                                            }

                                            _blockReserva.AttributeCollection.AppendAttribute(attRef);
                                            _acTrans.AddNewlyCreatedDBObject(attRef, true);
                                            _idx++;
                                        }
                                    }
                                }



                                #region quebracabo

                                if (nomeLayer.Equals("NET-CX-EMENDA") || nomeLayer.Equals("NET-NAP"))
                                {
                                    inseriu = true;
                                }

                                #endregion



                                _acTrans.Commit();
                            }

                            /*comando testes*/
                            if (inseriu) trim(_blockReserva, _pwLine, _pontoSeleca);
                        }
                    }
                }

            }
            catch (System.Exception erro)
            {
                MessageBox.Show("Ocorreu um erro durante o bloco.\nErro: linha: " + linha.ToString() + " - " + erro.Message + " - " + erro.StackTrace + erro.InnerException, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private static void InserirRESERVA(String nomeBloco, String nomeLayer)
        {
            try
            {

                Database _AcadDB = AcadDoc.Database;

                PromptPointOptions _selOpt = new PromptPointOptions("Selecione um cabo");
                Cursor.Current = Cursors.Default;

                PromptPointResult _result = AcadDoc.Editor.GetPoint(_selOpt);
                Cursor.Current = Cursors.WaitCursor;

                if (_result.Status == PromptStatus.OK)
                {

                    Point3d _pontoSeleca = new Point3d(_result.Value.X, _result.Value.Y, _result.Value.Z);
                    Point2d _ptSelecao = new Point2d(_result.Value.X, _result.Value.Y);
                    Point3d Pt1 = new Point3d(_result.Value.X + 0.5, _result.Value.Y + 0.5, 0);//;/PolarPoint(_pontoSeleca, (5 * Math.PI) / 4, 0.71);
                    Point3d Pt2 = new Point3d(_result.Value.X - 0.5, _result.Value.Y - 0.5, 0);//PolarPoint(_pontoSeleca, Math.PI / 4 , 0.71);

                    List<TypedValue> _tvfilter = new List<TypedValue>();

                    _tvfilter.Add(new TypedValue(0, "LWPOLYLINE"));

                    SelectionFilter _filter = new SelectionFilter(_tvfilter.ToArray());

                    PromptSelectionResult _resSel = AcadDoc.Editor.SelectCrossingWindow(Pt1, Pt2, _filter);

                    SelectionSet _selSets = _resSel.Value;

                    if (_selSets != null)
                    {
                        using (Transaction _acTrans = _AcadDB.TransactionManager.StartTransaction())
                        {

                            ObjectId _objId = _selSets.GetObjectIds()[0];
                            BlockTableRecord _CurrenSpace = _acTrans.GetObject(_AcadDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline _caboLine = _acTrans.GetObject(_objId, OpenMode.ForRead) as Polyline;
                            string layerCabo = _caboLine.Layer;
                            string Prompt = "Quantidade";
                            string Titulo = "Informe a Quantidade - Layer Cabo: " + layerCabo;


                            string Resultado = Microsoft.VisualBasic.Interaction.InputBox(Prompt, Titulo, "15", 150, 150);

                            BlockTable _blckTable = _acTrans.GetObject(_AcadDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                            if (_blckTable.Has(nomeBloco))
                            {

                                //_pontoSeleca = new Point3d(Cursor.Position.X, Cursor.Position.Y, 0.0);
                                BlockReference _blockReserva = new BlockReference(_pontoSeleca, _blckTable[nomeBloco]);

                                PromptAngleOptions _doub = new Autodesk.AutoCAD.EditorInput.PromptAngleOptions("Selecione a rotação:");

                                _doub.BasePoint = _pontoSeleca;
                                _doub.UseBasePoint = true;

                                //Aguarda a seleção do valor de rotação do bloco.
                                PromptDoubleResult _dbresult = AcadDoc.Editor.GetAngle(_doub);
                                Cursor.Current = Cursors.WaitCursor;
                                Double _angPt = _dbresult.Value;

                                BlockTableRecord _ModelSpace = _acTrans.GetObject(_AcadDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                _ModelSpace.AppendEntity(_blockReserva);
                                _acTrans.AddNewlyCreatedDBObject(_blockReserva, true);

                                // if (BlockMoving.Jig(_blockReserva))
                                {
                                    _blckTable = _acTrans.GetObject(_AcadDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                                    LayerTable _layerTable = _acTrans.GetObject(_AcadDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                                    //// Entity _pwLineEnt = _acTrans.GetObject(_objID, OpenMode.ForRead) as Entity;

                                    double _angleFix = _angPt;

                                    //// string _layName = _pwLine.Layer;
                                    double Lwheight = Convert.ToDouble(0.2);
                                    bool _inVertex = false;

                                    _blockReserva.Rotation = _angleFix;
                                    _blockReserva.Layer = nomeLayer;

                                    // Entity _entity = _acTrans.GetObject(_blockReserva.ObjectId, OpenMode.ForRead) as Entity;
                                    // SetXData(_entity, "FIBRA", layerCabo);

                                    int _idx = 0;

                                    BlockTableRecord _blockDefinition = _blckTable[nomeBloco].GetObject(OpenMode.ForRead) as BlockTableRecord;


                                    foreach (ObjectId EntId in _blockDefinition)
                                    {
                                        DBObject obj = EntId.GetObject(OpenMode.ForRead);
                                        AttributeDefinition attDef = obj as AttributeDefinition;

                                        if ((attDef != null) && (!attDef.Constant))
                                        {
                                            using (AttributeReference attRef = new AttributeReference())
                                            {
                                                attRef.SetAttributeFromBlock(attDef, _blockReserva.BlockTransform);


                                                if (true)//_noOpticoSelecionado != null)
                                                {
                                                    if (attRef.Tag.Equals("FIBRA"))
                                                    {
                                                        attRef.TextString = layerCabo;
                                                    }

                                                    if (attRef.Tag.Equals("ST"))
                                                    {
                                                        attRef.TextString = Resultado;
                                                    }

                                                    _blockReserva.AttributeCollection.AppendAttribute(attRef);
                                                    _acTrans.AddNewlyCreatedDBObject(attRef, true);

                                                    _idx++;
                                                }

                                            }
                                        }
                                    }


                                    _acTrans.Commit();


                                }
                            }

                        }
                    }
                }

            }
            catch (System.Exception erro)
            {
                MessageBox.Show("Ocorreu um erro durante a inserção da reserva técnica.\nErro: " + erro.Message + " - " + erro.StackTrace + erro.InnerException, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool Inter(Point3d Pt1, Point3d Pt2, Point3d Pt3, Point3d Pt4)
        {

            double Det;
            double s, t;

            Det = ((Pt4.X - Pt3.X) * (Pt2.Y - Pt1.Y)) - ((Pt4.Y - Pt3.Y) * (Pt2.X - Pt1.X));

            if (Det == 0)
            {
                return false;
            }

            if (Det != 0)
            {
                s = (((Pt4.X - Pt3.X) * (Pt3.Y - Pt1.Y)) - ((Pt4.Y - Pt3.Y) * (Pt3.X - Pt1.X))) / Det;
                t = (((Pt2.X - Pt1.X) * (Pt3.Y - Pt1.Y)) - ((Pt2.Y - Pt1.Y) * (Pt3.X - Pt1.X))) / Det;
                if ((s > 0) && (s < 1) && (t > 0) && (t < 1))
                    return true;
            }

            return false;
        }

        public static bool Inter(Point2d Pt1, Point2d Pt2, Point2d Pt3, Point2d Pt4)
        {

            double Det;
            double s, t;

            Det = ((Pt4.X - Pt3.X) * (Pt2.Y - Pt1.Y)) - ((Pt4.Y - Pt3.Y) * (Pt2.X - Pt1.X));

            if (Det == 0)
            {
                return false;
            }

            if (Det != 0)
            {
                s = (((Pt4.X - Pt3.X) * (Pt3.Y - Pt1.Y)) - ((Pt4.Y - Pt3.Y) * (Pt3.X - Pt1.X))) / Det;
                t = (((Pt2.X - Pt1.X) * (Pt3.Y - Pt1.Y)) - ((Pt2.Y - Pt1.Y) * (Pt3.X - Pt1.X))) / Det;
                if ((s > 0) && (s < 1) && (t > 0) && (t < 1))
                    return true;
            }

            return false;
        }

        public static List<Polyline> FilletAt(Polyline _polyline, int _idxVertex, Point3d PtInicial, Point3d PtFinal, double weightL, bool intVertex)
        {
            Database _AcadDb = AcadDoc.Database;
            List<Polyline> _lstPolyline = new List<Polyline>();
            long _idNoopticoA = 0;
            long _idNoopticoB = 0;

            Entity _pwLinePrincipal = _polyline.ObjectId.GetObject(OpenMode.ForRead) as Entity;

            List<Point2d> _vertexP1 = new List<Point2d>();
            List<Point2d> _vertexP2 = new List<Point2d>();

            if (intVertex == false)
            {
                for (int i = 0; i < _polyline.NumberOfVertices; i++)
                {
                    if (_idxVertex >= i)
                    {
                        if (_idxVertex == i)
                        {
                            _vertexP1.Add(_polyline.GetPoint2dAt(i));
                            _vertexP1.Add(new Point2d(PtInicial.X, PtInicial.Y));
                            _vertexP2.Add(new Point2d(PtFinal.X, PtFinal.Y));

                        }
                        else
                        {
                            _vertexP1.Add(_polyline.GetPoint2dAt(i));
                        }
                    }
                    else
                    {
                        _vertexP2.Add(_polyline.GetPoint2dAt(i));
                    }
                }
            }
            else
            {
                for (int i = 0; i < _polyline.NumberOfVertices; i++)
                {
                    if (_idxVertex >= i)
                    {
                        if (_idxVertex == i)
                        {
                            _vertexP1.Add(_polyline.GetPoint2dAt(i));
                            //_vertexP1.Add(new Point2d(PtInicial.X, PtInicial.Y));
                            _vertexP2.Add(new Point2d(PtFinal.X, PtFinal.Y));

                        }
                        else
                        {
                            _vertexP1.Add(_polyline.GetPoint2dAt(i));
                        }
                    }
                    else
                    {
                        _vertexP2.Add(_polyline.GetPoint2dAt(i));
                    }
                }
            }

            Polyline _pline1 = new Polyline();
            foreach (Point2d _pt in _vertexP1)
                _pline1.AddVertexAt(_pline1.NumberOfVertices, _pt, 0, weightL, weightL);

            Polyline _pline2 = new Polyline();
            foreach (Point2d _pt in _vertexP2)
                _pline2.AddVertexAt(_pline2.NumberOfVertices, _pt, 0, weightL, weightL);

            _lstPolyline.Add(_pline1);
            _lstPolyline.Add(_pline2);


            return _lstPolyline;
        }

        public static Point2d PolarPoint(Point2d pPt, double dAng, double dDist)
        {
            return new Point2d(pPt.X + dDist * Math.Cos(dAng), pPt.Y + dDist * Math.Sin(dAng));

        }

        public static Point3d PolarPoint(Point3d pPt, double dAng, double dDist)
        {
            return new Point3d(pPt.X + dDist * Math.Cos(dAng), pPt.Y + dDist * Math.Sin(dAng), pPt.Z);

        }

        #region testes

        private static void trim(BlockReference _blockReserva, Polyline line, Point3d ponto)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = AcadDoc;

            Editor ed = doc.Editor;

            String strHandle1, strEntName1, strHandle2, strEntName2, strCommand;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                Point3dCollection pts = new Point3dCollection();

                _blockReserva.IntersectWith(line, Intersect.ExtendArgument, pts, 0, 0);

                if (pts.Count == 0) return;


                Point3d pini = ponto;
                Point3d pfim;

                double length = 1.5; //Metade da Caixa de emenda

                if (_blockReserva.Layer.Equals("NET-NAP")) //Nap 
                {
                    length = 1.25;
                }

                pfim = new Point3d(ponto.X + Math.Cos(_blockReserva.Rotation) * length, ponto.Y + Math.Sin(_blockReserva.Rotation) * length, 0);

                Circle ci = new Circle();
                ci.Radius = length;
                ci.Center = pfim;


                Point2d pmax = new Point2d(ci.GeometricExtents.MaxPoint.X, ci.GeometricExtents.MaxPoint.Y);
                Point2d pmin = new Point2d(ci.GeometricExtents.MinPoint.X, ci.GeometricExtents.MinPoint.Y);


                //pmax = new Point2d(pfim.X, pfim.Y);
                //pmin = new Point2d(ponto.X, ponto.Y);

                btr.AppendEntity(ci);
                tr.AddNewlyCreatedDBObject(ci, true);



                strHandle1 = ci.ObjectId.Handle.ToString();
                strEntName1 = " (handent \"" + strHandle1 + "\")";

                //strHandle2 = line.ObjectId.Handle.ToString();
                //strEntName2 = " (handent \"" + strHandle2 + "\")";

                strCommand = "._trim" + strEntName1 + "\n f " + pmax.ToString().Replace("(", "").Replace(")", "") + " " + pmin.ToString().Replace("(", "").Replace(")", "") + "\n\n\n";

                strCommand += "._erase" + strEntName1 + "\n\n";

                AcadDoc.SendStringToExecute(strCommand, true, false, false);

                ed.Regen();
                tr.Commit();
            }


        }

        /*Comando teste para trim em linhas
         * obs: falta adaptar para polylines
         * 
         */
        [CommandMethod("ipLine")]
        public static void IntersectTestIn()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = AcadDoc;


            Editor ed = doc.Editor;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a Line >>");
                peo.SetRejectMessage("\nYou have to select the Line only >>");
                peo.AddAllowedClass(typeof(Line), false);

                PromptEntityResult res;
                res = ed.GetEntity(peo);

                if (res.Status != PromptStatus.OK) return;


                Entity ent = (Entity)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                if (ent == null) return;

                Line line = (Line)ent as Line;
                if (line == null) return;

                peo = new PromptEntityOptions("\nSelect a Rectangle: ");
                peo.SetRejectMessage("\nYou have to select the Polyline only >>");
                peo.AddAllowedClass(typeof(Polyline), false);

                res = ed.GetEntity("\nSelect a Rectangle: ");

                if (res.Status != PromptStatus.OK) return;

                ent = (Entity)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                if (ent == null) return;

                Polyline pline = (Polyline)ent as Polyline;
                if (line == null) return;

                Point3dCollection pts = new Point3dCollection();
                pline.IntersectWith(line, Intersect.ExtendArgument, pts, 0, 0);

                if (pts.Count == 0) return;

                List<Point3d> points = new List<Point3d>();
                points.Add(line.StartPoint);
                points.Add(line.EndPoint);

                foreach (Point3d p in pts)
                    points.Add(p);

                DBObjectCollection objs = line.GetSplitCurves(pts);
                List<DBObject> lstobj = new List<DBObject>();

                foreach (DBObject dbo in objs)
                    lstobj.Add(dbo);

                points.Sort(delegate (Point3d p1, Point3d p2)
                {
                    return Convert.ToDouble(
                    Convert.ToDouble(line.GetParameterAtPoint(p1))).CompareTo(
                    Convert.ToDouble(line.GetParameterAtPoint(p2)));
                }
                );

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                Line ln0 = lstobj[1] as Line;// middle
                Line ln1 = new Line(points[0], points[1]);
                Line ln2 = new Line(points[2], points[3]);

                if (!ln0.IsWriteEnabled)
                    ln0.UpgradeOpen();
                ln0.Dispose();

                if (!ln1.IsWriteEnabled)
                    ln1.UpgradeOpen();
                btr.AppendEntity(ln1);

                tr.AddNewlyCreatedDBObject(ln1, true);

                if (!ln2.IsWriteEnabled)
                    ln2.UpgradeOpen();

                btr.AppendEntity(ln2);
                tr.AddNewlyCreatedDBObject(ln2, true);

                if (!line.IsWriteEnabled)
                    line.UpgradeOpen();
                line.Erase();
                line.Dispose();

                ed.Regen();
                tr.Commit();
            }
        }

        [CommandMethod("opLine")]
        public static void IntersectTest()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = AcadDoc;

            Editor ed = doc.Editor;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a Line >>");
                peo.SetRejectMessage("\nYou have to select the Line only >>");
                peo.AddAllowedClass(typeof(Line), false);
                PromptEntityResult res;
                res = ed.GetEntity(peo);
                if (res.Status != PromptStatus.OK)
                    return;
                Entity ent = (Entity)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                if (ent == null)
                    return;
                Line line = (Line)ent as Line;
                if (line == null) return;
                peo = new PromptEntityOptions("\nSelect a Rectangle: ");
                peo.SetRejectMessage("\nYou have to select the Polyline only >>");
                peo.AddAllowedClass(typeof(Polyline), false);
                res = ed.GetEntity("\nSelect a Rectangle: ");
                if (res.Status != PromptStatus.OK)
                    return;
                ent = (Entity)tr.GetObject(res.ObjectId, OpenMode.ForRead);
                if (ent == null)
                    return;
                Polyline pline = (Polyline)ent as Polyline;
                if (line == null) return;
                Point3dCollection pts = new Point3dCollection();

                pline.IntersectWith(line, Intersect.ExtendArgument, pts, 0, 0);

                if (pts.Count == 0) return;
                DBObjectCollection lstobj = line.GetSplitCurves(pts);

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                Line ln0 = lstobj[1] as Line;// middle
                Line ln1 = lstobj[0] as Line;

                Line ln2 = lstobj[2] as Line;
                if (!ln0.IsWriteEnabled)
                    ln0.UpgradeOpen();
                btr.AppendEntity(ln0);
                tr.AddNewlyCreatedDBObject(ln0, true);
                if (!ln1.IsWriteEnabled)
                    ln1.UpgradeOpen();
                ln1.Dispose();

                if (!ln2.IsWriteEnabled)
                    ln2.UpgradeOpen();
                ln2.Dispose();

                if (!line.IsWriteEnabled)
                    line.UpgradeOpen();
                line.Erase();
                line.Dispose();
                ed.Regen();
                tr.Commit();
            }
        }

        #endregion
    }
}
