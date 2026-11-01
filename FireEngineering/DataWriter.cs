using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FireEngineering;

public static class WordTableWriter
{
    public static void WriteCombinedDoc(string outputFile, List<ProtocolDataModel> records)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputFile, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            Body body = mainPart.Document.Body;

            Table table = body.Elements<Table>().FirstOrDefault();
            if (table == null)
                throw new Exception("Expected table with headers not found.");

            var counter = 0;
            foreach (var rec in records)
            {
                counter++;
                TableRow tr = new TableRow();

                tr.Append(
                    MakeTextCell(counter.ToString() + ".", "200"),
                    MakeTextCell(rec.ChangeDate, "1200"),
                    MakeTextCell(rec.Type.ToString(), "1000"),
                    MakeTextCell(rec.BrandAndModel, "1500"),
                    MakeTextCell(rec.Serial, "1000"),
                    MakeTextCell(rec.Changed, "1500"),
                    MakeTextCell(rec.StickerId.ToString(), "1000"),
                    MakeTextCell(rec.ChangeType, "1000"),
                    MakeTextCell(rec.FutureChangeDate, "1200"),
                    //TODO -- read rotation!!!
                    MakeImageCell(mainPart, rec.Signature, "1000")
                );

                table.Append(tr);
            }

            mainPart.Document.Save();
        }
    }

    private static TableCell MakeTextCell(string text, string width)
    {
        // fontSize is in half-points → 20 = 10pt, 24 = 12pt
        var runProps = new RunProperties(
            new FontSize { Val = "16" }
        );

        var run = new Run();
        run.Append(runProps);
        run.Append(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });

        var para = new Paragraph(run);

        var cell = new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
            ),
            para
        );

        return cell;
    }

    private static TableCell MakeImageCell(MainDocumentPart mainPart, (byte[]? data, int rotation) image, string width)
    {
        TableCell cell = new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width }
            )
        );

        if (image.data == null || image.data.Length == 0)
        {
            cell.Append(new Paragraph(new Run(new Text(""))));
            return cell;
        }

        // Resize and convert to JPEG
        image.data = ResizeAndConvertToJpeg(image.data);

        // Add image part
        var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
        using (var stream = new MemoryStream(image.data))
        {
            imagePart.FeedData(stream);
        }

        string relationshipId = mainPart.GetIdOfPart(imagePart);

        // scale image size relative to column width
        long cx = 400000L; // width in EMU
        long cy = 200000L;  // height in EMU

        var element = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = (UInt32Value)1U, Name = "Signature" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Signature.png" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = cx, Cy = cy })
                                { Rotation = image.rotation },
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
        );

        cell.Append(new Paragraph(new Run(element)));
        return cell;
    }

    private static byte[] ResizeAndConvertToJpeg(byte[] original)
    {
        using (var input = new MemoryStream(original))
        using (var image = System.Drawing.Image.FromStream(input))
        {
            double scale = 0.5;
            int newWidth = (int)(image.Width * scale);
            int newHeight = (int)(image.Height * scale);

            using var bitmap = new System.Drawing.Bitmap(newWidth, newHeight);
            using var g = System.Drawing.Graphics.FromImage(bitmap);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.DrawImage(image, 0, 0, newWidth, newHeight);

            using (var output = new MemoryStream())
            {
                var encoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                    .First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);

                var parameters = new System.Drawing.Imaging.EncoderParameters(1);
                parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality, 80L); // 80% quality

                bitmap.Save(output, encoder, parameters);
                return output.ToArray();
            }
        }
    }

}
