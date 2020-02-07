using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace FancyReporting
{
    public static class HerGenParrafo
    {

        public static Paragraph agregarParrafo(List<Run> listaRun)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Both };

            paragraphProperties.Append(justification);

            paragraph.Append(paragraphProperties);
            paragraph.Append(listaRun);
            return paragraph;
        }

        public static Paragraph agregarParrafo(string texto)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Both };

            paragraphProperties.Append(justification);

            paragraph.Append(paragraphProperties);
            paragraph.Append(addTextNormal(texto));
            return paragraph;
        }

        public static Run agregarEspacio()
        {
            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "20" };
            Languages language = new Languages() { Val = "es-ES" };

            runProperties.Append(runFonts);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);
            runProperties.Append(language);

            Text text = new Text() {  Space = SpaceProcessingModeValues.Preserve };
            text.Text = " ";

            run.Append(runProperties);
            run.Append(text);
            return run;
        }

        public static Run addTextNormal(string cadena)
        {
            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            FontSize fontSize = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "20" };
            Languages language = new Languages() { Val = "es-ES" };

            runProperties.Append(runFonts);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);
            runProperties.Append(language);

            Text text = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text.Text = cadena;

            run.Append(runProperties);
            run.Append(text);
            return run;
        }

        public static Run addTextBold(string cadena)
        {
            Run run = new Run();
            
            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold = new Bold();
            BoldComplexScript boldComplexScript = new BoldComplexScript();
            FontSize fontSize = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "20" };
            Languages language = new Languages() { Val = "es-ES" };

            runProperties.Append(runFonts);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);
            runProperties.Append(language);

            Text text = new Text();
            text.Text = cadena;

            run.Append(runProperties);
            run.Append(text);
            return run;
        }

        public static Paragraph agregarSubTitulo(string cadena)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold = new Bold();
            BoldComplexScript boldComplexScript = new BoldComplexScript();
            FontSize fontSize = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "20" };
            Underline underline = new Underline();
            Languages language = new Languages() { Val = "es-ES" };

            Text text = new Text();

            text.Text = cadena;

            runProperties.Append(runFonts);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);
            runProperties.Append(underline);
            runProperties.Append(language);

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(run);
            return paragraph;
        }

        public static Paragraph agregarSubTituloGeneral(string cadena)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            Justification justification = new Justification() { Val = JustificationValues.Both };

            paragraphProperties.Append(justification);

            RunProperties runProperties = new RunProperties();
            RunFonts runFonts = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Bold bold = new Bold();
            BoldComplexScript boldComplexScript = new BoldComplexScript();
            FontSize fontSize = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = "20" };
            Underline underline = new Underline() { Val = UnderlineValues.Single };
            Languages language = new Languages() { Val = "es-ES" };

            runProperties.Append(runFonts);
            runProperties.Append(bold);
            runProperties.Append(boldComplexScript);
            runProperties.Append(fontSize);
            runProperties.Append(fontSizeComplexScript);
            runProperties.Append(underline);
            runProperties.Append(language);

            Run run = new Run();
            Text text = new Text();
            text.Text = cadena;

            run.Append(runProperties);
            run.Append(text);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            return paragraph;
        }
    }
}

