using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using FancyReporting.Models;

namespace FancyReporting
{
    public class HerGenTabla
    {
        public static Table generarTabla(string[,] array, Dictionary<string, string[]> confTabla)
        {
            int arrayFilas = array.GetLength(0);

            Table tabla = new Table();

            TableProperties tableProperties = new TableProperties();
            TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook = new TableLook() { Val = "04A0" };

            TableCellMarginDefault tableCellMarginDefault = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin = new TableCellLeftMargin() { Width = 100, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin = new TableCellRightMargin() { Width = 100, Type = TableWidthValues.Dxa };

            tableCellMarginDefault.Append(tableCellLeftMargin);
            tableCellMarginDefault.Append(tableCellRightMargin);

            tableProperties.Append(tableStyle);
            tableProperties.Append(tableWidth);
            tableProperties.Append(tableCellMarginDefault);
            tableProperties.Append(tableLook);

            TableGrid tableGrid = addGridColumns(confTabla["numColumns"] , confTabla["anchosColumns"]);

            tabla.Append(tableProperties);
            tabla.Append(tableGrid);

            int numColDetalle = int.Parse(confTabla["numColumns"][0]);
            int numColPeriodos =   int.Parse(confTabla["numColumns"][1]);
            int numColResultados = int.Parse(confTabla["numColumns"][2]);

            string anchoColDetalle = confTabla["anchosColumns"][0];
            string anchoColPeriodos = confTabla["anchosColumns"][1];
            string anchoColResultados = confTabla["anchosColumns"][2];

            //cabecera tabla
            int numRowsCabecera = 2;
            int numCellsRow = numColDetalle+ numColPeriodos+ numColResultados;
            int iRows = 0;

            TableRow priFilCab = new TableRow();
            
            int iCell = 0;

            //detalle
            TableCell celdaCabDetalleIni = celdaCabecera(array[iRows, iCell], anchoColDetalle);
            celdaCabMerVertical(celdaCabDetalleIni, true);
            priFilCab.Append(celdaCabDetalleIni);
            iCell++;

            //periodos
            TableCell celda = celdaCabecera(array[iRows, iCell], anchoColPeriodos);
            if(numColPeriodos > 1)
            {
                celdaCabMerLeft(celda, numColPeriodos);
            }
            priFilCab.Append(celda);
            
            iCell += numColPeriodos;
            
            //resultados
            TableCell celda1 = celdaCabecera(array[iRows, iCell], anchoColResultados);
            if (numColResultados > 1)//hacer merge si solo hay una columna
            {
                celdaCabMerLeft(celda1, numColResultados);
            }
            else
            {
                celdaCabMerVertical(celda1, true);
            }
            priFilCab.Append(celda1);
            
            tabla.Append(priFilCab);
            
            iRows++;
            
            //segunda fila cabecera
            TableRow secFilCab = new TableRow();
            iCell = 0;
            //detalle
            TableCell celdaCabDetalleFin = celdaCabecera("", anchoColDetalle);
            celdaCabMerVertical(celdaCabDetalleFin, false);
            secFilCab.Append(celdaCabDetalleFin);
            iCell++;
            //periodos
            while (iCell < numColDetalle+numColPeriodos)
            {
                secFilCab.Append(celdaCabecera(array[iRows, iCell], anchoColPeriodos));
                iCell++;
            }

            //resultados
            if (numColResultados > 1)//hacer merge si solo hay una columna
            {
                while (iCell < numColDetalle + numColPeriodos + numColResultados)
                {
                    secFilCab.Append(celdaCabecera(array[iRows, iCell], anchoColResultados));
                    iCell++;
                }
            }
            else
            {
                array[iRows, iCell] = "";
                TableCell celdaResTit = celdaCabecera(array[iRows, iCell], anchoColResultados);
                celdaCabMerVertical(celdaResTit, false);
                secFilCab.Append(celdaResTit);
            }
            tabla.Append(secFilCab);

            //cuerpo            
            for (int iRow = numRowsCabecera; iRow < arrayFilas; iRow++)
            {
                TableRow row = new TableRow();
                iCell = 0;

                if (confTabla["FilasGrupo"].Contains((iRow+1).ToString()))
                {
                    //detalle
                    row.Append(celdaCuerpoGrupo(array[iRow, iCell],numCellsRow));
                }
                else
                { 
                    row.Append(celdaCuerpoDetalle(array[iRow, iCell]));
                    iCell++;
                    //periodos
                    while (iCell < numColDetalle + numColPeriodos)
                    {
                        row.Append(celdaCuerpoContenido(array[iRow, iCell]));
                        iCell++;
                    }
                    //resultados
                    while (iCell < numColDetalle + numColPeriodos + numColResultados)
                    {
                        row.Append(celdaCuerpoResultado(array[iRow, iCell]));
                        iCell++;
                    }
                }
                tabla.Append(row);
            }
            
            return tabla;
        }
        
        public static TableGrid addGridColumns(string[] numColumnas, string[] widthColumns)
        {
            TableGrid tableGrid = new TableGrid();
            
            for (int i = 0; i < int.Parse(numColumnas[0]); i++)
            {
                tableGrid.Append(new GridColumn() { Width = widthColumns[0] });
            }
            for (int i = 0; i < int.Parse(numColumnas[1]); i++)
            {
                tableGrid.Append(new GridColumn() { Width = widthColumns[1] });
            }
            for (int i = 0; i < int.Parse(numColumnas[2]); i++)
            {
                tableGrid.Append(new GridColumn() { Width = widthColumns[2] });
            }
            return tableGrid;
        }

        public static TableCell celdaCabecera(string texto,string ancho)
        {
            TableCell cellCabecera = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties();

            TableCellWidth tableCellWidth = new TableCellWidth() { Width = ancho, Type = TableWidthUnitValues.Dxa };
            Shading shading = new Shading() { Val = ShadingPatternValues.Clear, 
                                              Color = "auto", 
                                              Fill = "AEAAAA", 
                                              ThemeFill = ThemeColorValues.Background2, 
                                              ThemeFillShade = "BF" };
            TableCellVerticalAlignment tableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties.Append(tableCellWidth);
            tableCellProperties.Append(shading);
            tableCellProperties.Append(tableCellVerticalAlignment);

            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Center };
            paragraphProperties.Append(justification);
            
            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            Bold bold = new Bold();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow", ComplexScript = "Arial" };
            FontSize fontSize = new FontSize() { Val = "16" };
            BoldComplexScript boldComplexScript = new BoldComplexScript();

            runProperties.Append(runFonts);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize);
            
            Text text = new Text();
            text.Text = texto;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            cellCabecera.Append(tableCellProperties);
            cellCabecera.Append(paragraph);
            return cellCabecera;
        }

        public static void celdaCabMerVertical(TableCell cellCabecera, bool iniMerge)
        {
            cellCabecera.GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().InsertAfterSelf<VerticalMerge>(new VerticalMerge()
            {
                Val = iniMerge ? MergedCellValues.Restart : MergedCellValues.Continue
            });
        }

        public static void celdaCabMerLeft(TableCell cellCabecera,int numColumnas)
        {
            cellCabecera.GetFirstChild<TableCellProperties>().GetFirstChild<TableCellWidth>().InsertAfterSelf<GridSpan>(new GridSpan()
            {
                Val = numColumnas
            });
        }

        public static TableCell celdaCuerpoDetalle(string texto)
        {
            TableCell tableCell = new TableCell();

            TableCellProperties tableCellProperties = new TableCellProperties();
            TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Nil };

            tableCellProperties.Append(tableCellWidth);

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Left };
            
            paragraphProperties.Append(justification);

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow", ComplexScript = "Arial" };
            FontSize fontSize = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "16" };

            runProperties.Append(runFonts);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);

            Text text = new Text();
            text.Text = texto;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            tableCell.Append(tableCellProperties);
            tableCell.Append(paragraph);
            return tableCell;
        }

        public static TableCell celdaCuerpoContenido(string texto)
        {
            TableCell tableCell = new TableCell();

            TableCellProperties tableCellProperties = new TableCellProperties();
            TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Nil };

            tableCellProperties.Append(tableCellWidth);

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Right };
            paragraphProperties.Append(justification);

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow", ComplexScript = "Arial" };
            FontSize fontSize2 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "16" };

            runProperties.Append(runFonts2);
            runProperties.Append(fontSize2);
            runProperties.Append(fontSizeComplexScript);

            Text text = new Text();
            text.Text = texto;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            tableCell.Append(tableCellProperties);
            tableCell.Append(paragraph);
            return tableCell;
        }

        public static TableCell celdaCuerpoResultado(string texto)
        {
            TableCell tableCell = new TableCell();

            TableCellProperties tableCellProperties = new TableCellProperties();
            TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Nil };

            tableCellProperties.Append(tableCellWidth);

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Right };
            paragraphProperties.Append(justification);

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            Bold bold = new Bold();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow", ComplexScript = "Arial" };
            FontSize fontSize2 = new FontSize() { Val = "16" };
            BoldComplexScript boldComplexScript = new BoldComplexScript();

            runProperties.Append(runFonts2);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize2);

            Text text = new Text();
            text.Text = texto;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            tableCell.Append(tableCellProperties);
            tableCell.Append(paragraph);
            return tableCell;
        }
        public static TableCell celdaCuerpoGrupo(string texto,int numColumnas)
        {
            TableCell tableCell = new TableCell();

            TableCellProperties tableCellProperties = new TableCellProperties();
            TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Nil };
            GridSpan gridSpan = new GridSpan() { Val = numColumnas };

            tableCellProperties.Append(tableCellWidth);
            tableCellProperties.Append(gridSpan);

            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Left };
            paragraphProperties.Append(justification);

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            Bold bold = new Bold();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow" };
            FontSize fontSize2 = new FontSize() { Val = "16" };
            BoldComplexScript boldComplexScript = new BoldComplexScript();

            runProperties.Append(runFonts2);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize2);
            

            Text text = new Text();
            text.Text = texto;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            tableCell.Append(tableCellProperties);
            tableCell.Append(paragraph);

            return tableCell;
        }


    }

}
