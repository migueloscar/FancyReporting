using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FancyReporting.Models;
using Newtonsoft.Json;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using System.IO;

namespace FancyReporting
{
    public abstract class Informe
    {
        public JsonConfiguracion configuracion;
        List<JsonDatos> listaDatos;

        List<RepGenTabla.resultado> listFunTblRes;
        List<RepGenTabla.validacion> listFunTblVal;
        List<RepGenParrafos.generarIntro> listFunParIntro;
        List<RepGenParrafos.generarCriterio> listFunParrCri;

        List<string> listTitTblRes;
        List<string[]> listParConectores;

        public List<DateTime> listPeriodos;
        List<string> listDetalles;
        List<DatoPreparado> listDatosPrep;
        List<string> listCodGraEst;

        Dictionary<string, string[]> dicConfTabla;

        public Informe()
        {

            configuracion = new JsonConfiguracion();
            listFunTblRes = new List<RepGenTabla.resultado>();
            listTitTblRes = new List<string>();

            listFunParIntro = new List<RepGenParrafos.generarIntro>();
            listFunTblVal = new List<RepGenTabla.validacion>();
            listFunParrCri = new List<RepGenParrafos.generarCriterio>();
            listParConectores = new List<string[]>();

            listDatosPrep = new List<DatoPreparado>();
            dicConfTabla = new Dictionary<string, string[]>();
            listCodGraEst = new List<string>();
        }

        public abstract void cargandoCadenasDefecto();
        public abstract void agregandoFuncionesResultado();
        public abstract void agregandoTitColRes();
        public abstract void agregandoIntrosParrafos();
        public abstract void agregandoCriComentario();
        public abstract void agregandoConector();
        public abstract void agregandoValidaciones();

        public void ejecutarValidaciones() {
            for(int iValidaciones = 0 ; iValidaciones < listFunTblVal.Count ; iValidaciones++)
            {
                List<string> mensajesErrores=listFunTblVal[iValidaciones](listaDatos);
                for(int iError = 0 ; iError < mensajesErrores.Count;)
                {
                    throw new System.ArgumentException(mensajesErrores[iError]);
                }
            }
        }

        public void inicializarInforme(string listaDatosJson, string configuracionJson)
        {
            cargandoCadenasDefecto();
            parseandoJsons(listaDatosJson, configuracionJson);
            remplazarCadenas(configuracion.cadenas);
            obtenerPeriodos();
            agregandoValidaciones();
        
        }

        public void InformeEjecutarValidaciones()
        {
            ejecutarValidaciones();
        }

        public void agregarFuncionesPlantilla()
        {
            agregandoFuncionesResultado();
            agregandoTitColRes();
            agregandoIntrosParrafos();
            agregandoCriComentario();
            agregandoConector();
        }

        public MemoryStream generarInformeMemStreamValoresInicializados()
        {
            MemoryStream file = generarWordMemStream();
            return file;
        }

        public MemoryStream generarInformeMemStream(string listaDatosJson, string configuracionJson)
        {
            inicializarInforme(listaDatosJson, configuracionJson);
            agregarFuncionesPlantilla();

            MemoryStream file =generarWordMemStream();
            return file;
        }

        public void parseandoJsons(string listaDatosJson, string configuracionJson)
        {
            listaDatos = JsonConvert.DeserializeObject<List<JsonDatos>>(listaDatosJson);
            configuracion = JsonConvert.DeserializeObject<JsonConfiguracion>(configuracionJson);
            Console.WriteLine(configuracion);
        }

        public void remplazarCadenas(Dictionary<string, string> nuevasCadenas)
        {
            List<string> keys= nuevasCadenas.Keys.ToList();
            foreach (string key 
                in keys)
            {
                if (configuracion.cadenas.Keys.Contains(key))
                {
                    configuracion.cadenas[key] = nuevasCadenas[key];
                }
            }
        }

        public void agregarFuncionValidacion(RepGenTabla.validacion funcion)
        {
            listFunTblVal.Add(funcion);
        }
        public void agregarFuncionesColResultado(RepGenTabla.resultado funcion)
        {
            listFunTblRes.Add(funcion);
        }
        public void agregarColResultadoTitulo(string tituloColResultado)
        {
            listTitTblRes.Add(tituloColResultado);
        }
        public void agregarIntParrafo(RepGenParrafos.generarIntro funcion)
        {
            listFunParIntro.Add(funcion);
        }
        public void agregarCriComentario(RepGenParrafos.generarCriterio funcion)
        {
            listFunParrCri.Add(funcion);
        }
        public void agregarConector(string[] arrayConector)
        {
            listParConectores.Add(arrayConector);
        }
        public void obtenerPeriodos()
        {
            listPeriodos = listaDatos.GroupBy(x => x.periodo).Select(group => group.Key).ToList();
            listPeriodos.Sort();//sort no devuelve nada
        }
        public MemoryStream generarWordMemStream()
        {
            var stream = new MemoryStream();
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                crearPartes(wordDocument);
            }
            return stream;
        }
        public void crearPartes(WordprocessingDocument wordDocument)
        {
            string formatoFechas = configuracion.cadenas["formatoFecha"];

            //llena las listas periodos, detalles
            prepararDatos();

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            Document document = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid wp14" } };
            #region markup

            document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            #endregion
            
            Body body = new Body();

            body.Append(generarSubtituloGeneral());
            body.Append(generarTabla());
            body.Append(agregarFuente(configuracion.cadenas["fuente"]));
            body.Append(generarParrafosDescripcion(mainDocumentPart));

            //setting del documento
            DocumentSettingsPart documentSettingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>("rId2");
            generarConfiguracion(documentSettingsPart);

            //estilos
            StyleDefinitionsPart styleDefinitionsPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
            generarEstilos(styleDefinitionsPart);

            //fuentes
            FontTablePart fontTablePart = mainDocumentPart.AddNewPart<FontTablePart>("rId8");
            generarFuentes(fontTablePart);

            document.Append(body);

            mainDocumentPart.Document = document;
        }

        private void generarConfiguracion(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "160" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00645E4C" };
            Rsid rsid1 = new Rsid() { Val = "00645E4C" };
            Rsid rsid2 = new Rsid() { Val = "00B87D82" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "4959F6D5" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{97814B57-DC16-4A95-90E5-2BCBEBA0226C}" };

            settings1.Append(zoom1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }
        private void generarFuentes(FontTablePart fontTablePart)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Arial Narrow" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0606020202030204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "00000287", UnicodeSignature1 = "00000800", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000009F", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Arial" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);

            fontTablePart.Fonts = fonts1;
        }

        public void generarEstilos(StyleDefinitionsPart styleDefinitionsPart)
        {
            Styles styles = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid" } };
            #region markup
            styles.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            #endregion
            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            Style style5 = new Style() { Type = StyleValues.Table, StyleId = "TableGrid" };
            StyleName styleName5 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn1 = new BasedOn() { Val = "TableNormal" };
            UIPriority uIPriority4 = new UIPriority() { Val = 39 };
            Rsid rsid3 = new Rsid() { Val = "00565867" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines2);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            styleTableProperties2.Append(tableBorders1);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(uIPriority4);
            style5.Append(rsid3);
            style5.Append(styleParagraphProperties1);
            style5.Append(styleTableProperties2);

            styles.Append(docDefaults1);
            styles.Append(style5);

            styleDefinitionsPart.Styles = styles;
        }

        public void  prepararDatos()
        {
            listDetalles = listaDatos.GroupBy(x => x.detalle).Select(group => group.Key).ToList();

            foreach (string detalle in listDetalles)
            {
                listDatosPrep.Add(new DatoPreparado()
                {
                    detalle = detalle,
                    periodos = listaDatos.Where(x => x.detalle.Equals(detalle)).Select(y =>
                         new PeriodoValor()
                         {
                             periodo = y.periodo,
                             valor = y.valor
                         }
                        ).OrderBy(d => d.periodo).ToList()
                });
            }
        }

        public Paragraph generarSubtituloGeneral()
        {
            string formatoFechas = configuracion.cadenas["formatoFecha"];
            string dato = configuracion.cadenas["dato"];

            string res = configuracion.cadenas["subTituloGeneral"] + " ";
            res += configuracion.cadenas["artPlural"] + " ";
            res +=  RepGenParrafos.pluralizarPalabra(dato) + " ";
            res += string.Format(configuracion.cadenas["entreFechasFormato3"],
                                 listPeriodos.First().ToString(formatoFechas),
                                 listPeriodos.Last().ToString(formatoFechas)
                                        );
            return RepGenParrafos.agregarSubTituloGeneral(res);
        }

        public Paragraph agregarFuente(string cadena)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Both };

            paragraphProperties.Append(justification);

            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow" };
            FontSize fontSize = new FontSize() { Val = "16" };

            runProperties.Append(runFonts);
            runProperties.Append(fontSize);

            Run run = new Run();
            Text text = new Text();
            text.Text = cadena;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            return paragraph;
        }
        public Table generarTabla()
        {
            int numDetalles = 1;
            string[,] arrayDatos = RepGenTabla.generarArrayTabla(configuracion, listFunTblRes, listTitTblRes,
                listDatosPrep, listPeriodos, listDetalles, out dicConfTabla);

            //rellenando configuracion tabla
            dicConfTabla["fuente"]=new string[] { configuracion.cadenas["fuente"] };
            dicConfTabla["numColumns"]= new string[] { numDetalles.ToString(), listPeriodos.Count.ToString(), listFunTblRes.Count.ToString() };
            dicConfTabla["anchosColumns"] = new string[] { configuracion.cadenas["anchoPrimera"], configuracion.cadenas["anchoPeriodos"], configuracion.cadenas["anchoResultados"] };

            return RepGenTabla.generarTabla(arrayDatos, dicConfTabla);
        }
        public List<Paragraph> generarParrafosDescripcion(MainDocumentPart mainDocumentPart)
        {
            List<Paragraph> listaParrafos = new List<Paragraph>();

            foreach (GrupoDato gr in configuracion.grupos)
            {
                listaParrafos.Add(RepGenParrafos.agregarSubTitulo(gr.descripcion));

                List<DatosRelevantes> detallesGrupoCom = configuracion.datos.Where(x => x.grupo.Equals(gr.grupo) && x.comentario)//obtenemos todos los detalles de un grupo
                                            .OrderBy(y => y.orden).ToList()
                                            .Select(x => new DatosRelevantes()
                                            {
                                                detalle = x.detalle,
                                                orden = x.orden
                                            }).ToList();//obtenemos solo los nombres de los detalles.

                if (detallesGrupoCom.Count != 0)
                {
                    for (int iParDes = 0; iParDes < listFunParIntro.Count; iParDes++)
                    {
                        List<DatosRelevantes> datosPeriodo = (from detGru in detallesGrupoCom
                                                              join datPre in listDatosPrep on detGru.detalle equals datPre.detalle
                                                              select new DatosRelevantes()
                                                              {
                                                                  detalle = detGru.detalle,
                                                                  orden = detGru.orden,
                                                                  comparacion = listFunTblRes[iParDes](datPre.periodos),
                                                                  tendecia = listFunParrCri[iParDes](datPre.periodos)
                                                              }).ToList();

                        string intro = listFunParIntro[iParDes](listPeriodos);

                        listaParrafos.Add(
                            RepGenParrafos.generarDescripcion
                            (
                                configuracion.cadenas,
                                listPeriodos, 
                                datosPeriodo, 
                                iParDes, 
                                intro, 
                                listParConectores[iParDes]
                            )
                        );
                    }
                }
                //generar grafico estadistico
                List<string> detallesGrupoGrafEst = configuracion.datos.Where(x => x.grupo.Equals(gr.grupo) && x.graficoEstadistico)//obtenemos todos los detalles de un grupo
                                            .Select(x => x.detalle).ToList();//obtenemos solo los nombres de los detalles.

                if (detallesGrupoGrafEst.Count != 0)
                {
                    List<DatoPreparado> datosGrafico = listDatosPrep.Where(x => detallesGrupoGrafEst.Contains(x.detalle)).ToList();

                    listaParrafos.Add(RepGenGraficoEstadistico.genIntroGraEst(configuracion.cadenas, datosGrafico.Count, gr.descripcion));

                    string[] codGraEst = generarNuevoCodigoGrafico();
                    listCodGraEst.Add(codGraEst[0]);

                    listaParrafos.Add(RepGenGraficoEstadistico.genParrafoGrafico(codGraEst[0], codGraEst[1]));
                    //generacion de graficos estadisticos
                    string tituloGraEsta=RepGenGraficoEstadistico.generarTituloGrafico(configuracion.cadenas,listPeriodos,datosGrafico);
                    RepGenGraficoEstadistico.generarGraficoEstadistico(mainDocumentPart,codGraEst[0], datosGrafico, tituloGraEsta,configuracion.cadenas);
                }
            }
            return listaParrafos;
        }
        public string[] generarNuevoCodigoGrafico()
        {
            //codigo del chart, numero de chart
            return new string[] { "chart" + (listCodGraEst.Count + 1).ToString(), (listCodGraEst.Count + 1).ToString() };
        }
    }

}
