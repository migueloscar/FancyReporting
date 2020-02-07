using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using FancyReporting.Models;

namespace FancyReporting
{
    public static class HerGenGraficoEstadistico
    {
        public static Paragraph generarParrafoGrafico(string codigoChart, string numChart)
        {
            Paragraph paragraph = new Paragraph();
           
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            VerticalTextAlignment verticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            paragraphMarkRunProperties.Append(verticalTextAlignment);
            paragraphProperties.Append(paragraphMarkRunProperties);

            Run run1 = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties.Append(noProof);
            runProperties.Append(verticalTextAlignment2);

            Drawing drawing = new Drawing();

            Wp.Inline inline = new Wp.Inline() { 
                DistanceFromTop = (UInt32Value)0U, 
                DistanceFromBottom = (UInt32Value)0U, 
                DistanceFromLeft = (UInt32Value)0U, 
                DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 5681207L, Cy = 3200400L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 15240L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)(UInt32.Parse(numChart)), Name = codigoChart };
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic = new A.Graphic();
            graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = codigoChart };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData.Append(chartReference1);

            graphic.Append(graphicData);

            inline.Append(extent1);
            inline.Append(effectExtent1);
            inline.Append(docProperties1);
            inline.Append(nonVisualGraphicFrameDrawingProperties1);
            inline.Append(graphic);

            drawing.Append(inline);

            run1.Append(runProperties);
            run1.Append(drawing);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run1);

            return paragraph;
        }
        public static void generarGraficoEstadistico(MainDocumentPart mainDocumentPart,
                                                     string codigoChart, 
                                                     List<DatoPreparado> datosGrafico,
                                                     string tituloGraEsta,
                                                     Dictionary<string,string> cadenas)
        {
            ChartPart chartPart = mainDocumentPart.AddNewPart<ChartPart>(codigoChart);
            GenerateChartPart1Content(chartPart, datosGrafico, tituloGraEsta, cadenas);

            EmbeddedPackagePart embeddedPackagePart1 = chartPart.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
            GenerateEmbeddedPackagePart1Content(embeddedPackagePart1);
        }

        private static C.LineChartSeries generarSeries(
            Dictionary<string, string> cadenas,
            DatoPreparado datoPreparado,
            int numFila)//numfila inicia desde el 2, la primera es para los periodos
        {
            int numPeriodos = datoPreparado.periodos.Count;
            string formtFecPeriodos = cadenas["formtGraficoPer"];
            string formtNumGra= cadenas["formatoNumGrafico"];

            List<C.LineChartSeries> listSeries = new List<C.LineChartSeries>();
            C.LineChartSeries lineChartSeries1 = new C.LineChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)((UInt32)numFila) };
            C.Order order1 = new C.Order() { Val = (UInt32Value)((UInt32)numFila) };

            C.SeriesText seriesText1 = new C.SeriesText();
            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula { Text = "Sheet1!$A$"+numFila };

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue() { Text = datoPreparado.detalle };

            stringPoint1.Append(numericValue1);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            seriesText1.Append(stringReference1);

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

            A.Outline outline5 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill9.Append(schemeColor18);
            A.Round round1 = new A.Round();

            //outline5.Append(solidFill9);
            outline5.Append(round1);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties2.Append(outline5);
            chartShapeProperties2.Append(effectList7);

            C.Marker marker1 = new C.Marker();
            C.Symbol symbol1 = new C.Symbol() { Val = C.MarkerStyleValues.Auto };
            C.Size size1 = new C.Size() { Val = 5 };
            

            A.SolidFill solidFill10 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill10.Append(schemeColor19);

            A.Outline outline6 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill11.Append(schemeColor20);

            outline6.Append(solidFill11);
            A.EffectList effectList8 = new A.EffectList();

            marker1.Append(symbol1);
            marker1.Append(size1);

            C.DataLabels dataLabels1 = new C.DataLabels();

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();
            A.NoFill noFill3 = new A.NoFill();

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill4 = new A.NoFill();

            outline7.Append(noFill4);
            A.EffectList effectList9 = new A.EffectList();

            chartShapeProperties4.Append(noFill3);
            chartShapeProperties4.Append(outline7);
            chartShapeProperties4.Append(effectList9);

            C.TextProperties textProperties2 = new C.TextProperties();

            A.BodyProperties bodyProperties3 = new A.BodyProperties()
            {
                Rotation = 0,
                UseParagraphSpacing = true,
                VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                LeftInset = 38100,
                TopInset = 19050,
                RightInset = 38100,
                BottomInset = 19050,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = true
            };

            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties3.Append(shapeAutoFit1);
            A.ListStyle listStyle3 = new A.ListStyle();
            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties()
            {
                FontSize = 900,
                Bold = false,
                Italic = false,
                Underline = A.TextUnderlineValues.None,
                Strike = A.TextStrikeValues.NoStrike,
                Kerning = 1200,
                Baseline = 0
            };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor21.Append(luminanceModulation11);
            schemeColor21.Append(luminanceOffset3);

            solidFill12.Append(schemeColor21);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill12);
            defaultRunProperties3.Append(latinFont5);
            defaultRunProperties3.Append(eastAsianFont5);
            defaultRunProperties3.Append(complexScriptFont5);

            paragraphProperties4.Append(defaultRunProperties3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties3);

            textProperties2.Append(bodyProperties3);
            textProperties2.Append(listStyle3);
            textProperties2.Append(paragraph4);
            C.DataLabelPosition dataLabelPosition1 = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Top };
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = true };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };
            C.ShowLeaderLines showLeaderLines1 = new C.ShowLeaderLines() { Val = false };
           
            C15.ShowLeaderLines showLeaderLines2 = new C15.ShowLeaderLines() { Val = true };

            C15.LeaderLines leaderLines1 = new C15.LeaderLines();

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();

            A.Outline outline8 = new A.Outline() { 
                Width = 9525, 
                CapType = A.LineCapValues.Flat, 
                CompoundLineType = A.CompoundLineValues.Single, 
                Alignment = A.PenAlignmentValues.Center 
            };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor22.Append(luminanceModulation12);
            schemeColor22.Append(luminanceOffset4);

            solidFill13.Append(schemeColor22);
            A.Round round2 = new A.Round();

            outline8.Append(solidFill13);
            outline8.Append(round2);
            A.EffectList effectList10 = new A.EffectList();

            chartShapeProperties5.Append(outline8);
            chartShapeProperties5.Append(effectList10);

            leaderLines1.Append(chartShapeProperties5);

            dataLabels1.Append(chartShapeProperties4);
            dataLabels1.Append(textProperties2);
            dataLabels1.Append(dataLabelPosition1);
            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            dataLabels1.Append(showLeaderLines1);

            int columnaInicio = 66;//B
            char columnaFinal = (char)(columnaInicio + numPeriodos);

            string celdaInicio = string.Format("${0}${1}", columnaInicio.ToString(), numFila);
            string celdaFinal = string.Format("${0}${1}", columnaFinal.ToString(), numFila);

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();
            C.StringReference stringReference2 = new C.StringReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = string.Format( "Sheet1!{0}:{1}", celdaInicio, celdaFinal);

            
            C.StringCache stringCache = new C.StringCache();
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)((UInt32)numPeriodos) };

            stringCache.Append(pointCount2);


            for (int i = 0; i < numPeriodos; i++)
            {
                C.StringPoint stringPoint = new C.StringPoint() { Index = (UInt32Value)((UInt32)i) };
                C.NumericValue numericValue = new C.NumericValue();
                numericValue.Text = datoPreparado.periodos[i].periodo.ToString(formtFecPeriodos);
                stringPoint.Append(numericValue);
                stringCache.Append(stringPoint);
            }

            stringReference2.Append(formula2);
            stringReference2.Append(stringCache);

            categoryAxisData1.Append(stringReference2);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula3 = new C.Formula();
            formula3.Text = string.Format("Sheet1!{0}:{1}", celdaInicio, celdaFinal);

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode() { Text = formtNumGra };
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)((UInt32)numPeriodos) };

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount3);

            for (int i = 0; i < numPeriodos; i++)
            {
                C.NumericPoint numericPoint = new C.NumericPoint() { Index = (UInt32Value)((UInt32)i) };
                C.NumericValue numericValue = new C.NumericValue();
                numericValue.Text = datoPreparado.periodos[i].valor.ToString();
                numericPoint.Append(numericValue);
                numberingCache1.Append(numericPoint);
            }
            numberReference1.Append(formula3);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);
            C.Smooth smooth1 = new C.Smooth() { Val = false };

            lineChartSeries1.Append(index1);
            lineChartSeries1.Append(order1);
            lineChartSeries1.Append(seriesText1);
            lineChartSeries1.Append(chartShapeProperties2);
            lineChartSeries1.Append(marker1);
            lineChartSeries1.Append(dataLabels1);
            lineChartSeries1.Append(categoryAxisData1);
            lineChartSeries1.Append(values1);
            lineChartSeries1.Append(smooth1);
            return lineChartSeries1;
        }

        // Generates content of chartPart1.
        private static void GenerateChartPart1Content(ChartPart chartPart, 
                                                        List<DatoPreparado> datosGrafico, 
                                                        string tituloGraEsta,
                                                        Dictionary<string,string> cadenas)
        {

            int numDatos = datosGrafico.Count;
            string tituloEjeY = cadenas["graficoEjeY"];
            string formtNumGra = cadenas["formatoNumGrafico"];

            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };
            C.Chart chart1 = new C.Chart();

            #region Titulo

            C.Title title1 = new C.Title();

            C.ChartText chartText1 = new C.ChartText();

            C.RichText richText1 = new C.RichText();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { 
                Rotation = 0, 
                UseParagraphSpacing = true, 
                VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, 
                Vertical = A.TextVerticalValues.Horizontal, 
                Wrap = A.TextWrappingValues.Square, 
                Anchor = A.TextAnchoringTypeValues.Center, 
                AnchorCenter = true 
            };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { 
                FontSize = 1200, 
                Bold = false, 
                Italic = false, 
                Underline = A.TextUnderlineValues.None, 
                Strike = A.TextStrikeValues.NoStrike, 
                Kerning = 1200, 
                Spacing = 0, 
                Baseline = 0 
            };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor16.Append(luminanceModulation9);
            schemeColor16.Append(luminanceOffset1);

            solidFill7.Append(schemeColor16);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill7);
            defaultRunProperties1.Append(latinFont3);
            defaultRunProperties1.Append(eastAsianFont3);
            defaultRunProperties1.Append(complexScriptFont3);

            paragraphProperties2.Append(defaultRunProperties1);

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { 
                Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Baseline = 0 };
            A.EffectList effectList4 = new A.EffectList();

            runProperties2.Append(effectList4);
            A.Text text1 = new A.Text();
            text1.Text = tituloGraEsta;

            run2.Append(runProperties2);
            run2.Append(text1);

            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() {
                Language = "en-US", FontSize = 1200 };
            A.EffectList effectList5 = new A.EffectList();

            endParagraphRunProperties1.Append(effectList5);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);
            paragraph2.Append(endParagraphRunProperties1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph2);

            chartText1.Append(richText1);
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline4.Append(noFill2);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties1.Append(noFill1);
            chartShapeProperties1.Append(outline4);
            chartShapeProperties1.Append(effectList6);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { 
                Rotation = 0, 
                UseParagraphSpacing = true, 
                VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, 
                Vertical = A.TextVerticalValues.Horizontal, 
                Wrap = A.TextWrappingValues.Square, 
                Anchor = A.TextAnchoringTypeValues.Center, 
                AnchorCenter = true 
            };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { 
                FontSize = 1200, 
                Bold = false, 
                Italic = false, 
                Underline = A.TextUnderlineValues.None, 
                Strike = A.TextStrikeValues.NoStrike, 
                Kerning = 1200, 
                Spacing = 0, 
                Baseline = 0 
            };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor17.Append(luminanceModulation10);
            schemeColor17.Append(luminanceOffset2);

            solidFill8.Append(schemeColor17);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill8);
            defaultRunProperties2.Append(latinFont4);
            defaultRunProperties2.Append(eastAsianFont4);
            defaultRunProperties2.Append(complexScriptFont4);

            paragraphProperties3.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(endParagraphRunProperties2);

            textProperties1.Append(bodyProperties2);
            textProperties1.Append(listStyle2);
            textProperties1.Append(paragraph3);

            title1.Append(chartText1);
            title1.Append(overlay1);
            title1.Append(chartShapeProperties1);
            title1.Append(textProperties1);

            #endregion
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

            #region areaTrabajo

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();

            C.LineChart lineChart1 = new C.LineChart();
            C.Grouping grouping1 = new C.Grouping() { Val = C.GroupingValues.Standard };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            lineChart1.Append(grouping1);
            lineChart1.Append(varyColors1);

            for (int i=0;i<numDatos;i++)
            {
                lineChart1.Append(generarSeries(cadenas,datosGrafico[i], i + 2));//numFila inicia desde la fila2
            }
            #region DataLabels

            C.DataLabels dataLabels5 = new C.DataLabels();
            C.DataLabelPosition dataLabelPosition5 = new C.DataLabelPosition() { Val = C.DataLabelPositionValues.Top };
            C.ShowLegendKey showLegendKey5 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue5 = new C.ShowValue() { Val = true };
            C.ShowCategoryName showCategoryName5 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName5 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent5 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize5 = new C.ShowBubbleSize() { Val = false };

            dataLabels5.Append(dataLabelPosition5);
            dataLabels5.Append(showLegendKey5);
            dataLabels5.Append(showValue5);
            dataLabels5.Append(showCategoryName5);
            dataLabels5.Append(showSeriesName5);
            dataLabels5.Append(showPercent5);
            dataLabels5.Append(showBubbleSize5);
            #endregion

            C.ShowMarker showMarker1 = new C.ShowMarker() { Val = true };
            C.Smooth smooth5 = new C.Smooth() { Val = false };
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)1489934448U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)1488627904U };

            lineChart1.Append(dataLabels5);
            lineChart1.Append(showMarker1);
            lineChart1.Append(smooth5);
            lineChart1.Append(axisId1);
            lineChart1.Append(axisId2);

            #region categoryAxis

            C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)1489934448U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling1.Append(orientation1);
            C.Delete delete1 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties18 = new C.ChartShapeProperties();
            A.NoFill noFill11 = new A.NoFill();

            A.Outline outline21 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill29 = new A.SolidFill();

            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor38.Append(luminanceModulation19);
            schemeColor38.Append(luminanceOffset11);

            solidFill29.Append(schemeColor38);
            A.Round round9 = new A.Round();

            outline21.Append(solidFill29);
            outline21.Append(round9);
            A.EffectList effectList23 = new A.EffectList();

            chartShapeProperties18.Append(noFill11);
            chartShapeProperties18.Append(outline21);
            chartShapeProperties18.Append(effectList23);

            C.TextProperties textProperties6 = new C.TextProperties();
            A.BodyProperties bodyProperties7 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor39.Append(luminanceModulation20);
            schemeColor39.Append(luminanceOffset12);

            solidFill30.Append(schemeColor39);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill30);
            defaultRunProperties7.Append(latinFont9);
            defaultRunProperties7.Append(eastAsianFont9);
            defaultRunProperties7.Append(complexScriptFont9);

            paragraphProperties8.Append(defaultRunProperties7);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(endParagraphRunProperties7);

            textProperties6.Append(bodyProperties7);
            textProperties6.Append(listStyle7);
            textProperties6.Append(paragraph8);
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)1488627904U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
            C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
            C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
            C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId3);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(numberingFormat1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(chartShapeProperties18);
            categoryAxis1.Append(textProperties6);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);


            #endregion


            #region valuesAxis

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)1488627904U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling2.Append(orientation2);
            C.Delete delete2 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties19 = new C.ChartShapeProperties();

            A.Outline outline22 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor40.Append(luminanceModulation21);
            schemeColor40.Append(luminanceOffset13);

            solidFill31.Append(schemeColor40);
            A.Round round10 = new A.Round();

            outline22.Append(solidFill31);
            outline22.Append(round10);
            A.EffectList effectList24 = new A.EffectList();

            chartShapeProperties19.Append(outline22);
            chartShapeProperties19.Append(effectList24);

            majorGridlines1.Append(chartShapeProperties19);

            #region milesSoles
            C.Title title2 = new C.Title();


            C.ChartText chartText2 = new C.ChartText();

            C.RichText richText2 = new C.RichText();
            A.BodyProperties bodyProperties8 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor41.Append(luminanceModulation22);
            schemeColor41.Append(luminanceOffset14);

            solidFill32.Append(schemeColor41);
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill32);
            defaultRunProperties8.Append(latinFont10);
            defaultRunProperties8.Append(eastAsianFont10);
            defaultRunProperties8.Append(complexScriptFont10);

            paragraphProperties9.Append(defaultRunProperties8);

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US" };
            A.Text text2 = new A.Text();
            text2.Text = tituloEjeY;

            run3.Append(runProperties3);
            run3.Append(text2);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run3);


            richText2.Append(bodyProperties8);
            richText2.Append(listStyle8);
            richText2.Append(paragraph9);

            chartText2.Append(richText2);

            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties20 = new C.ChartShapeProperties();
            A.NoFill noFill12 = new A.NoFill();

            A.Outline outline23 = new A.Outline();
            A.NoFill noFill13 = new A.NoFill();

            outline23.Append(noFill13);
            A.EffectList effectList25 = new A.EffectList();

            chartShapeProperties20.Append(noFill12);
            chartShapeProperties20.Append(outline23);
            chartShapeProperties20.Append(effectList25);

            C.TextProperties textProperties7 = new C.TextProperties();
            A.BodyProperties bodyProperties9 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor42.Append(luminanceModulation23);
            schemeColor42.Append(luminanceOffset15);

            solidFill33.Append(schemeColor42);
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill33);
            defaultRunProperties9.Append(latinFont11);
            defaultRunProperties9.Append(eastAsianFont11);
            defaultRunProperties9.Append(complexScriptFont11);

            paragraphProperties10.Append(defaultRunProperties9);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(endParagraphRunProperties8);

            textProperties7.Append(bodyProperties9);
            textProperties7.Append(listStyle9);
            textProperties7.Append(paragraph10);

            title2.Append(chartText2);
            title2.Append(overlay2);
            title2.Append(chartShapeProperties20);
            title2.Append(textProperties7);

            #endregion


            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = formtNumGra, SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties21 = new C.ChartShapeProperties();
            A.NoFill noFill14 = new A.NoFill();

            A.Outline outline24 = new A.Outline();
            A.NoFill noFill15 = new A.NoFill();

            outline24.Append(noFill15);
            A.EffectList effectList26 = new A.EffectList();

            chartShapeProperties21.Append(noFill14);
            chartShapeProperties21.Append(outline24);
            chartShapeProperties21.Append(effectList26);

            C.TextProperties textProperties8 = new C.TextProperties();
            A.BodyProperties bodyProperties10 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor43.Append(luminanceModulation24);
            schemeColor43.Append(luminanceOffset16);

            solidFill34.Append(schemeColor43);
            A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill34);
            defaultRunProperties10.Append(latinFont12);
            defaultRunProperties10.Append(eastAsianFont12);
            defaultRunProperties10.Append(complexScriptFont12);

            paragraphProperties11.Append(defaultRunProperties10);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(endParagraphRunProperties9);

            textProperties8.Append(bodyProperties10);
            textProperties8.Append(listStyle10);
            textProperties8.Append(paragraph11);
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)1489934448U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

            valueAxis1.Append(axisId4);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(title2);
            valueAxis1.Append(numberingFormat2);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(chartShapeProperties21);
            valueAxis1.Append(textProperties8);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);
            #endregion

            plotArea1.Append(layout1);
            plotArea1.Append(lineChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            #endregion

            #region leyenda
            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
            C.Overlay overlay3 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties22 = new C.ChartShapeProperties();
            A.NoFill noFill18 = new A.NoFill();

            A.Outline outline26 = new A.Outline();
            A.NoFill noFill19 = new A.NoFill();

            outline26.Append(noFill19);
            A.EffectList effectList28 = new A.EffectList();

            chartShapeProperties22.Append(noFill18);
            chartShapeProperties22.Append(outline26);
            chartShapeProperties22.Append(effectList28);

            C.TextProperties textProperties9 = new C.TextProperties();
            A.BodyProperties bodyProperties11 = new A.BodyProperties()
            {
                Rotation = 0,
                UseParagraphSpacing = true,
                VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                Vertical = A.TextVerticalValues.Horizontal,
                Wrap = A.TextWrappingValues.Square,
                Anchor = A.TextAnchoringTypeValues.Center,
                AnchorCenter = true
            };
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor44.Append(luminanceModulation25);
            schemeColor44.Append(luminanceOffset17);

            solidFill35.Append(schemeColor44);
            A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties11.Append(solidFill35);
            defaultRunProperties11.Append(latinFont13);
            defaultRunProperties11.Append(eastAsianFont13);
            defaultRunProperties11.Append(complexScriptFont13);

            paragraphProperties12.Append(defaultRunProperties11);
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(endParagraphRunProperties10);

            textProperties9.Append(bodyProperties11);
            textProperties9.Append(listStyle11);
            textProperties9.Append(paragraph12);

            legend1.Append(legendPosition1);
            legend1.Append(overlay3);
            legend1.Append(chartShapeProperties22);
            legend1.Append(textProperties9);
            #endregion

            chart1.Append(title1);
            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(chart1);
            chartPart.ChartSpace = chartSpace1;
        }

        // Generates content of embeddedPackagePart1.
        private static void GenerateEmbeddedPackagePart1Content(EmbeddedPackagePart embeddedPackagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(embeddedPackagePart1Data);
            embeddedPackagePart1.FeedData(data);
            data.Close();
        }

        #region Binary Data
        private static string embeddedPackagePart1Data = "UEsDBBQABgAIAAAAIQDdK4tYbAEAABAFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACslE1PwzAMhu9I/IcqV9Rm44AQWrcDH0eYxPgBoXHXaG0Sxd7Y/j1u9iGEyiq0Xmq1id/nrR1nMts2dbKBgMbZXIyzkUjAFk4bu8zFx+IlvRcJkrJa1c5CLnaAYja9vposdh4w4WyLuaiI/IOUWFTQKMycB8srpQuNIn4NS+lVsVJLkLej0Z0snCWwlFKrIaaTJyjVuqbkecuf904C1CiSx/3GlpUL5X1tCkXsVG6s/kVJD4SMM+MerIzHG7YhZCehXfkbcMh749IEoyGZq0CvqmEbclvLLxdWn86tsvMiHS5dWZoCtCvWDVcgQx9AaawAqKmzGLNGGXv0fYYfN6OMYTywkfb/onCPD+J+g4zPyy1EmR4g0q4GHLrsUbSPXKkA+p0CT8bgBn5q95VcfXIFJLVh6LZH0XN8Prfz4DzyBAf4fxeOI9pmp56FIJCB05B2HfYTkaf/4rZDe79o0B1sGe+z6TcAAAD//wMAUEsDBBQABgAIAAAAIQC1VTAj9AAAAEwCAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJJNT8MwDIbvSPyHyPfV3ZAQQkt3QUi7IVR+gEncD7WNoyQb3b8nHBBUGoMDR3+9fvzK2908jerIIfbiNKyLEhQ7I7Z3rYaX+nF1ByomcpZGcazhxBF21fXV9plHSnkodr2PKqu4qKFLyd8jRtPxRLEQzy5XGgkTpRyGFj2ZgVrGTVneYviuAdVCU+2thrC3N6Dqk8+bf9eWpukNP4g5TOzSmRXIc2Jn2a58yGwh9fkaVVNoOWmwYp5yOiJ5X2RswPNEm78T/XwtTpzIUiI0Evgyz0fHJaD1f1q0NPHLnXnENwnDq8jwyYKLH6jeAQAA//8DAFBLAwQUAAYACAAAACEA0nHpxwMDAADZBgAADwAAAHhsL3dvcmtib29rLnhtbKxVbU/bMBD+Pmn/IfJ3EztvTSJaREijIW0TGox9nEziUIskjhyHliH++85uU8aLJsRWtXbvznnuuRdfDo82bePccjUI2c0RPSDI4V0pK9Fdz9H3iwLHyBk06yrWyI7P0R0f0NHi44fDtVQ3V1LeOADQDXO00rpPXXcoV7xlw4HseQeWWqqWaRDVtTv0irNqWHGu28b1CInclokObRFS9RYMWdei5Lksx5Z3eguieMM00B9Woh8mtLZ8C1zL1M3Y41K2PUBciUboOwuKnLZMT687qdhVA2FvaOhsFHwj+FECizd5AtMLV60olRxkrQ8A2t2SfhE/JS6lT1KweZmDtyEFruK3wtRwz0pF72QV7bGiRzBK/hmNQmvZXkkhee9EC/fcPLQ4rEXDL7et67C+/8paU6kGOQ0b9LISmldzNANRrvkThRr7bBQNWD2P+gS5i307nymn4jUbG30BjTzBw82IosQLzcmNSqdkn2nlwP/T/DM4PGe34B6CrHbdeQr41P/ZlSqNCf15T+OIBLOiwMQnAQ4ooTgu8iWOlnno5TNCvSh5gBypKC0lG/VqF5sBn6MgfMX0hW0mCyXpKKpHIvdk98Fmf7ZMtgcTkLnFl4Kvh8csGNHZ/BBdJdcQhEdCiOtukmdBDOLaWn+ISq/gSBwle90nLq5XQJn6VgnlNtTm6D6mx3EUZgX2PJLjIKEhPk5OTnCeQVr8wqNJNLOU3D842YEB3OzudLbI52aIUJhMZrd5Ro5KjQ91WlETlDs9VrKmhKKazRy0xZYNPxe/ONSunqNje55v9OdBLw5hd0YlgCwNyPGMJAEmSz/EQZx4OA58D58EubcMZ8t8mYWmXGb8pf9jCEBr0TCd5qrhvGJKXyhW3sA0/sbrjA3QYdvwgOefZLMwzogPFIOCFtBbCcFZFgU4zAs/nNH8ZBkWj2RNMiDyZyP7baMmdu3TnOlRwbsASFs5NWux0+6V9Vaxq9oTB+m33N6m7dN/O1hc2oK+6se1eTCrrZ47ZW/xGwAA//8DAFBLAwQUAAYACAAAACEAgT6Ul/MAAAC6AgAAGgAIAXhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArFJNS8QwEL0L/ocwd5t2FRHZdC8i7FXrDwjJtCnbJiEzfvTfGyq6XVjWSy8Db4Z5783Hdvc1DuIDE/XBK6iKEgR6E2zvOwVvzfPNAwhi7a0egkcFExLs6uur7QsOmnMTuT6SyCyeFDjm+CglGYejpiJE9LnShjRqzjB1Mmpz0B3KTVney7TkgPqEU+ytgrS3tyCaKWbl/7lD2/YGn4J5H9HzGQlJPA15ANHo1CEr+MFF9gjyvPxmTXnOa8Gj+gzlHKtLHqo1PXyGdCCHyEcffymSc+WimbtV7+F0QvvKKb/b8izL9O9m5MnH1d8AAAD//wMAUEsDBBQABgAIAAAAIQB9mjrGXAMAAE8JAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1snFbbjts2EH0P0H8g9G5JFHU1bAe+BdmHAEHbpM80RdvCSqJK0t5dBPn3Dsm11pKzhRHBF0pzeGYOZzjU7ONzU6Mzl6oS7dzDfugh3jJRVu1h7n37+9Mk95DStC1pLVo+91648j4u/vgwexLyUR051wgYWjX3jlp30yBQ7MgbqnzR8RYseyEbquFWHgLVSU5LO6mpgygM06ChVes5hqm8h0Ps9xXjG8FODW+1I5G8phriV8eqUxe2ht1D11D5eOomTDQdUOyqutIvltRDDZs+HFoh6a4G3c84pgw9S/hE8CUXN/b5jaemYlIosdc+MAcu5lv5RVAElPVMt/rvosFxIPm5Mgl8o4p+LySc9FzRGxn5TbK0JzPLJaenqpx7PyKyKdZZvpxsw3Q5iaM0nixxlk+2WUzydBUXqyT86S1mZQUZNqqQ5Pu5t8TTbeIFi5mtn+8Vf1JXY6Tp7i9ec6Y5+MAeMuW5E+LRAB/gUQiMygIMI2W6OvM1r+u5t4Lg1L/WBwzBQdB7uB5fvH2yBf1VopLv6anWf4qnz7w6HDW4jf0EhJpKmZYvG64YlCi49okNnIkaSOAXNZXZa1Bi9NkFW5X6CKPQj9IiITgCGnZSWjT/vFpMXP1MSI2dGYNMZy98nBVpnv3/PLDaefDfzyMpifJ3HAYuYrscG6rpYibFE4IyhdBVR82mx9N3FUPABrs0YDsFlkJBFs6LcBacYWnZK2J1i8BDxPoWEQ0Rm1sEGSK2t4i4RwQgrFcHy3u/OgMeqktG6hwissqzLM+LkTZnJ9ZekCQJ/RHDZoBIs5yMEdtrBA4TyKif/1qcKfa7U2fAQ3HpSJxDOHGExPEo9LWzO3GERBHxR/I3I0QY+WnRX9koh0NwWKS/Bg/SCbvkfsUGPFQ8CmHlEE4xxhEZldna2Z1ijDOc+SOGzRCRY+KPyn17jcizIrwCDKTB3rtfmgEPpb2ViN2pK4d4rdQ4GeV6fW3OSTiW5cxOeB7nqQ8HfH+NE3kNLkgevQN2al1Ddh2oowf+hcpD1SpU871tr5mHpOvAoQ9jLTrTdE0z3AkNTfRyd4S3Dw4tCRqyh/ZC6MsNNFc4P2r+lUqtEBMn07Yx7Mj+KZJTc3jJhxLbI+INDo2sfxVa/AcAAP//AwBQSwMEFAAGAAgAAAAhAMEXEL5OBwAAxiAAABMAAAB4bC90aGVtZS90aGVtZTEueG1s7FnNixs3FL8X+j8Mc3f8NeOPJd7gz2yT3SRknZQctbbsUVYzMpK8GxMCJTn1UiikpZdCbz2U0kADDb30jwkktOkf0SfN2COt5SSbbEpadg2LR/69p6f3nn5683Tx0r2YekeYC8KSll++UPI9nIzYmCTTln9rOCg0fE9IlIwRZQlu+Qss/Evbn35yEW3JCMfYA/lEbKGWH0k52yoWxQiGkbjAZjiB3yaMx0jCI58Wxxwdg96YFiulUq0YI5L4XoJiUHt9MiEj7A2VSn97qbxP4TGRQg2MKN9XqrElobHjw7JCiIXoUu4dIdryYZ4xOx7ie9L3KBISfmj5Jf3nF7cvFtFWJkTlBllDbqD/MrlMYHxY0XPy6cFq0iAIg1p7pV8DqFzH9ev9Wr+20qcBaDSClaa22DrrlW6QYQ1Q+tWhu1fvVcsW3tBfXbO5HaqPhdegVH+whh8MuuBFC69BKT5cw4edZqdn69egFF9bw9dL7V5Qt/RrUERJcriGLoW1ane52hVkwuiOE94Mg0G9kinPUZANq+xSU0xYIjflWozuMj4AgAJSJEniycUMT9AIsriLKDngxNsl0wgSb4YSJmC4VCkNSlX4rz6B/qYjirYwMqSVXWCJWBtS9nhixMlMtvwroNU3IC+ePXv+8Onzh789f/To+cNfsrm1KktuByVTU+7Vj1///f0X3l+//vDq8Tfp1CfxwsS//PnLl7//8Tr1sOLcFS++ffLy6ZMX333150+PHdrbHB2Y8CGJsfCu4WPvJothgQ778QE/ncQwQsSSQBHodqjuy8gCXlsg6sJ1sO3C2xxYxgW8PL9r2bof8bkkjpmvRrEF3GOMdhh3OuCqmsvw8HCeTN2T87mJu4nQkWvuLkqsAPfnM6BX4lLZjbBl5g2KEommOMHSU7+xQ4wdq7tDiOXXPTLiTLCJ9O4Qr4OI0yVDcmAlUi60Q2KIy8JlIITa8s3eba/DqGvVPXxkI2FbIOowfoip5cbLaC5R7FI5RDE1Hb6LZOQycn/BRyauLyREeoop8/pjLIRL5jqH9RpBvwoM4w77Hl3ENpJLcujSuYsYM5E9dtiNUDxz2kySyMR+Jg4hRZF3g0kXfI/ZO0Q9QxxQsjHctwm2wv1mIrgF5GqalCeI+mXOHbG8jJm9Hxd0grCLZdo8tti1zYkzOzrzqZXauxhTdIzGGHu3PnNY0GEzy+e50VciYJUd7EqsK8jOVfWcYAFlkqpr1ilylwgrZffxlG2wZ29xgngWKIkR36T5GkTdSl045ZxUep2ODk3gNQLlH+SL0ynXBegwkru/SeuNCFlnl3oW7nxdcCt+b7PHYF/ePe2+BBl8ahkg9rf2zRBRa4I8YYYICgwX3YKIFf5cRJ2rWmzulJvYmzYPAxRGVr0Tk+SNxc+Jsif8d8oedwFzBgWPW/H7lDqbKGXnRIGzCfcfLGt6aJ7cwHCSrHPWeVVzXtX4//uqZtNePq9lzmuZ81rG9fb1QWqZvHyByibv8uieT7yx5TMhlO7LBcW7Qnd9BLzRjAcwqNtRuie5agHOIviaNZgs3JQjLeNxJj8nMtqP0AxaQ2XdwJyKTPVUeDMmoGOkh3UrFZ/QrftO83iPjdNOZ7msupqpCwWS+XgpXI1Dl0qm6Fo9796t1Ot+6FR3WZcGKNnTGGFMZhtRdRhRXw5CFF5nhF7ZmVjRdFjRUOqXoVpGceUKMG0VFXjl9uBFveWHQdpBhmYclOdjFae0mbyMrgrOmUZ6kzOpmQFQYi8zII90U9m6cXlqdWmqvUWkLSOMdLONMNIwghfhLDvNlvtZxrqZh9QyT7liuRtyM+qNDxFrRSInuIEmJlPQxDtu+bVqCLcqIzRr+RPoGMPXeAa5I9RbF6JTuHYZSZ5u+HdhlhkXsodElDpck07KBjGRmHuUxC1fLX+VDTTRHKJtK1eAED5a45pAKx+bcRB0O8h4MsEjaYbdGFGeTh+B4VOucP6qxd8drCTZHMK9H42PvQM65zcRpFhYLysHjomAi4Ny6s0xgZuwFZHl+XfiYMpo17yK0jmUjiM6i1B2ophknsI1ia7M0U8rHxhP2ZrBoesuPJiqA/a9T903H9XKcwZp5memxSrq1HST6Yc75A2r8kPUsiqlbv1OLXKuay65DhLVeUq84dR9iwPBMC2fzDJNWbxOw4qzs1HbtDMsCAxP1Db4bXVGOD3xric/yJ3MWnVALOtKnfj6yty81WYHd4E8enB/OKdS6FBCb5cjKPrSG8iUNmCL3JNZjQjfvDknLf9+KWwH3UrYLZQaYb8QVINSoRG2q4V2GFbL/bBc6nUqD+BgkVFcDtPr+gFcYdBFdmmvx9cu7uPlLc2FEYuLTF/MF7Xh+uK+XNl8ce8RIJ37tcqgWW12aoVmtT0oBL1Oo9Ds1jqFXq1b7w163bDRHDzwvSMNDtrVblDrNwq1crdbCGolZX6jWagHlUo7qLcb/aD9ICtjYOUpfWS+APdqu7b/AQAA//8DAFBLAwQUAAYACAAAACEAcfcYcecCAAB1BwAADQAAAHhsL3N0eWxlcy54bWysVW1v0zAQ/o7Ef7C8r6ROsqS0VZKJtouENCakDYmvbuK01vwSOe5IQfx3zklfUg3BGPtS25fzc8/dc74mV60U6JGZhmuV4mDkY8RUoUuu1in+cp97E4waS1VJhVYsxTvW4Kvs7ZuksTvB7jaMWQQQqknxxtp6RkhTbJikzUjXTMGXShtJLRzNmjS1YbRs3CUpSOj7YyIpV7hHmMniOSCSmodt7RVa1tTyFRfc7josjGQx+7hW2tCVAKptENECtcHYhKg1hyCd9UkcyQujG13ZEeASXVW8YE/pTsmU0OKEBMgvQwpi4odnubfmhUgRMeyRO/lwllRa2QYVeqssiAlEXQlmD0p/U7n75Iy9V5Y039EjFWAJMMmSQgttkAXpoHKdRVHJeo8FFXxluHOrqORi15tDZ+jU3vtJDrV3RuJ47JcGLnEhjqxCRwAMWQLyWWZUDge039/vagivoNN6mM7vL95rQ3dBGA8ukC5glqy0KaGzT/U4mLJEsMoCUcPXG7daXcPvSlsL6mdJyelaKypcKj3IcQPpFEyIO9f9X6sz7LZCaitzaT+WKYZ35Ipw2EIi+22P1x8c/hCtxx7ARlCsf4dFbXXEf8Xbl89J6Rgb0boWu9utXDGTdzNg31dnjKLXweyqCHUbiHMmzbHIyHV1im8dIQHvY18otNpyYbn6jSyAWbbnQsM5S3qpB4p3qbhZt9AlhLh4d3Hhj3y/68vuBnE40GtuOnX9c6QIzVKyim6FvT9+TPFp/4mVfCvDo9dn/qhtB5Hi0/7GNXMwdgFZa28aeIGwoq3hKf5xPX8/XV7noTfx5xMvumSxN43nSy+OFvPlMp/6ob/4OZiR/zEhu5EOfRtEs0bAHDX7ZPcp3p1sKR4cevpduYD2kPs0HPsf4sD38ks/8KIxnXiT8WXs5XEQLsfR/DrO4wH3+IWT1CdB0M9kRz6eWS6Z4Oqg1UGhoRVEguMfkiAHJcjp/zL7BQAA//8DAFBLAwQUAAYACAAAACEAjVYa1OUAAACRAQAAFAAAAHhsL3NoYXJlZFN0cmluZ3MueG1sdJDBSgMxEIbvgu8Q5u5m24O2JUkPFQ9eBNHeQzLdDWwmayZb9O2NCwXZ6nHmm+//YdT+Mw7ijJlDIg2rpgWB5JIP1Gl4f3u624DgYsnbIRFq+EKGvbm9UcxFVJdYQ1/KuJOSXY/RcpNGpEpOKUdb6pg7yWNG67lHLHGQ67a9l9EGAuHSREXDFsRE4WPCw2U2ioNRc8WOR+tqdc1gzGcEI5QsRsmfi/nKPKJrxLpdPSzBiysz2PxnXIGLsV0ax9AhFeTl/hVPgSy5YH26gsf6zPDH/nnyoRrD7zhZX2q+AQAA//8DAFBLAwQUAAYACAAAACEAqJz1ALwAAAAlAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzhI/BCsIwEETvgv8Q9m7SehCRpr2I0KvoB6zptg22SchG0b834EVB8DTsDvtmp2oe8yTuFNl6p6GUBQhyxnfWDRrOp8NqC4ITug4n70jDkxiaermojjRhykc82sAiUxxrGFMKO6XYjDQjSx/IZaf3ccaUxziogOaKA6l1UWxU/GRA/cUUbachtl0J4vQMOfk/2/e9NbT35jaTSz8iVMLLRBmIcaCkQcr3ht9SyvwsqLpSX+XqFwAAAP//AwBQSwMEFAAGAAgAAAAhALuQc7kyAgAAOQQAABQAAAB4bC90YWJsZXMvdGFibGUxLnhtbJyT0W6bMBSG7yftHZDvHWxswEQlFQYjVZpWae0ewCUmQQOMbKdNNO3dZ5I2aZVeTLuDA//n/5z/+OZ2P/TBszK202MO8AKBQI2NXnfjJgc/H2vIQGCdHNey16PKwUFZcLv6+uXGyadeBV492hxsnZuWYWibrRqkXehJjf5Lq80gnX81m9BORsm13Srlhj6MEErCQXYjOBGWQ/MvkEGaX7sJNnqYpOueur5zhyMLBEOzvNuM2syucrA3wd6QN/jeXMGHrjHa6tYtPCzUbds16sojpqFRz908mguK/CcrObO8r27tZ+2ZZrmbH38jUpc15QgmKWKQljSDhWAFJDHOaFwLXvLkDwhGOfjmHucevXrd2amXh+8fika1OSjwUsQgcNrJ3v7QLw9b/eLTRWB1iq3U/W4YbdDo3ehyEH+sX9yRV3s4iak3QmCaCG+PRzFkCakhTdIMEZ4mBSdnewEIPxxzxEVzs2+4NPa6iMcwLkoCKakj3y2KYJJVXGQl40VcnXGVahZBhHD6GXaO+IxlglUpYxhixrHHkhIWRYohyvxwGSlIjS/Y+8YdsewzLH2PLTNCSUVTSDPBIY05g5yTAqaFqAkuq5oKfuX2U6yP5OKW8EJEqRAQo7KCVKAKZmVVQRLFRY2ThIk6O2Pf3GY+delktW/v/Nqg2Xt4vIivib4O/sEdenU3tjqwPvm6M9adfpg1x9o3eVWa98SZblL+PvvtmpUn0bn67rzVXwAAAP//AwBQSwMEFAAGAAgAAAAhAC/OUmFTAQAAhwIAABEACAFkb2NQcm9wcy9jb3JlLnhtbCCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIySX0+DMBTF3038DqTvUApzugZY/JM9ubjEGY1vTXvHGqGQtrjt21tgQ8yM8fH2nvu759w0me/LwvsEbWSlUkSCEHmgeCWkylP0sl74N8gzlinBikpBig5g0Dy7vEh4TXmlYaWrGrSVYDxHUobyOkVba2uKseFbKJkJnEK55qbSJbOu1DmuGf9gOeAoDKe4BMsEswy3QL8eiOiIFHxA1o0uOoDgGAooQVmDSUDwt9aCLs2vA11npCylPdQu09HumC143xzUeyMH4W63C3ZxZ8P5J/ht+fjcRfWlam/FAWWJ4JRrYLbS2VLmDRTek+FMe7d5I7UGb8UaVyZ4pGtvWjBjl+78Gwni7vD36Lncbe1C9qtBeM427UOeOq/x/cN6gbIoJDOfED8iazKlVzGdTN5bNz/m2xj9Q3n09E/iNY1mdBKPiCdAluCzr5N9AQAA//8DAFBLAwQUAAYACAAAACEAYUkJEIkBAAARAwAAEAAIAWRvY1Byb3BzL2FwcC54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACckkFv2zAMhe8D+h8M3Rs53VAMgaxiSFf0sGEBkrZnTaZjobIkiKyR7NePttHU2XrqjeR7ePpESd0cOl/0kNHFUInlohQFBBtrF/aVeNjdXX4VBZIJtfExQCWOgOJGX3xSmxwTZHKABUcErERLlFZSom2hM7hgObDSxNwZ4jbvZWwaZ+E22pcOAsmrsryWcCAINdSX6RQopsRVTx8NraMd+PBxd0wMrNW3lLyzhviW+qezOWJsqPh+sOCVnIuK6bZgX7Kjoy6VnLdqa42HNQfrxngEJd8G6h7MsLSNcRm16mnVg6WYC3R/eG1XovhtEAacSvQmOxOIsQbb1Iy1T0hZP8X8jC0AoZJsmIZjOffOa/dFL0cDF+fGIWACYeEccefIA/5qNibTO8TLOfHIMPFOONuBbzpzzjdemU/6J3sdu2TCkYVT9cOFZ3xIu3hrCF7XeT5U29ZkqPkFTus+DdQ9bzL7IWTdmrCH+tXzvzA8/uP0w/XyelF+LvldZzMl3/6y/gsAAP//AwBQSwECLQAUAAYACAAAACEA3SuLWGwBAAAQBQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQC1VTAj9AAAAEwCAAALAAAAAAAAAAAAAAAAAKUDAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDScenHAwMAANkGAAAPAAAAAAAAAAAAAAAAAMoGAAB4bC93b3JrYm9vay54bWxQSwECLQAUAAYACAAAACEAgT6Ul/MAAAC6AgAAGgAAAAAAAAAAAAAAAAD6CQAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECLQAUAAYACAAAACEAfZo6xlwDAABPCQAAGAAAAAAAAAAAAAAAAAAtDAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAi0AFAAGAAgAAAAhAMEXEL5OBwAAxiAAABMAAAAAAAAAAAAAAAAAvw8AAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECLQAUAAYACAAAACEAcfcYcecCAAB1BwAADQAAAAAAAAAAAAAAAAA+FwAAeGwvc3R5bGVzLnhtbFBLAQItABQABgAIAAAAIQCNVhrU5QAAAJEBAAAUAAAAAAAAAAAAAAAAAFAaAAB4bC9zaGFyZWRTdHJpbmdzLnhtbFBLAQItABQABgAIAAAAIQConPUAvAAAACUBAAAjAAAAAAAAAAAAAAAAAGcbAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0MS54bWwucmVsc1BLAQItABQABgAIAAAAIQC7kHO5MgIAADkEAAAUAAAAAAAAAAAAAAAAAGQcAAB4bC90YWJsZXMvdGFibGUxLnhtbFBLAQItABQABgAIAAAAIQAvzlJhUwEAAIcCAAARAAAAAAAAAAAAAAAAAMgeAABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQABgAIAAAAIQBhSQkQiQEAABEDAAAQAAAAAAAAAAAAAAAAAFIhAABkb2NQcm9wcy9hcHAueG1sUEsFBgAAAAAMAAwAEwMAABEkAAAAAA==";

        private static System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
        #endregion
    }
}
