using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlUtilities
{
    public class ClsOpenXmlUtilities
    {
        public static void InsertPicture(WordprocessingDocument wordprocessingDocument, string fileName)
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        }

        public static void AddStyle(MainDocumentPart mainPart, string styleId, string styleName,
            string fontName, int fontSize, string rgbColor, bool isBold, bool isItalic, bool isUnderLined)
        {
            // we have to set the properties
            RunProperties rPr = new RunProperties();
            Color color = new Color() { Val = rgbColor };
            RunFonts rFont = new RunFonts();
            rFont.Ascii = fontName;
            rPr.Append(color);
            rPr.Append(rFont);
            if(isBold) rPr.Append(new Bold());
            if(isItalic) rPr.Append(new Italic());
            if(isUnderLined) rPr.Append(new Underline() { Val = UnderlineValues.Single });
            rPr.Append(new FontSize() { Val = (fontSize * 2).ToString() });

            Style style = new Style();
            style.StyleId = styleId;
            if (styleName == null || styleName.Length == 0) styleName = styleId;
            style.Append(new Name() { Val = styleName }); //this is the name of the new style
            style.Append(rPr); //we are adding properties previously defined

            // we have to add style that we have created to the StylePart
            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles();
            stylePart.Styles.Append(style);
            stylePart.Styles.Save(); // we save the style part
        }

        public static Paragraph CreateParagraphWithStyle(string styleId, JustificationValues justification)
        {
            Paragraph paragraph = new Paragraph();         
            ParagraphProperties pp = new ParagraphProperties();
            paragraph.Append(pp);
            // we set the style
            pp.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
            // we set the alignement
            pp.Justification = new Justification() { Val = justification };       
            return paragraph;
        }

        public static void AddTextParagraph(Paragraph paragraph, string content)
        {
            Run r = new Run();
            Text t = new Text(content);
            r.Append(t);
            paragraph.Append(r);
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }
    }
}
