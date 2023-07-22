using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using System;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.AnalisisMateriaPrima.General
{
    public class ProtocoloMateriaPrima_PDF
    {
        public string GenerarReporte(PlantillaProtocoloDTO datosReporte)
        {
            string reporte = null;

            MemoryStream ms = new MemoryStream();

            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Protocolo de materia prima");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(5, 15, 30, 15);

            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(110)
               .SetHeight(44)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            #region estilos

            Style estiloTitulo = new Style().SetFontSize(13).SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.CENTER);

            Style estiloTituloAnalisis = new Style().SetFontSize(12).SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.RIGHT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            Style estiloCabeceraTitulo = new Style().SetFontSize(9).SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK).SetBorder(Border.NO_BORDER);

            Style estiloCabeceraTexto = new Style().SetFontSize(9).SetFontColor(ColorConstants.BLACK)
                .SetBorder(Border.NO_BORDER);

            Style estiloDetalleTitulo = new Style().SetFontSize(9).SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK).SetTextAlignment(TextAlignment.CENTER).SetBackgroundColor(new DeviceRgb(191, 191, 191));

            Style estiloDetalleTexto = new Style().SetFontSize(9).SetFontColor(ColorConstants.BLACK);
            #endregion

            #region titulo

            Table tituloTable = new Table(new float[] { 20, 60, 20 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(10);

            Cell tituloCelda = new Cell(1, 1).Add(img);
            tituloTable.AddCell(tituloCelda.SetBorder(Border.NO_BORDER));

            tituloCelda = new Cell(1, 1).Add(new Paragraph("PROTOCOLO DE ANÁLISIS\nDE MATERIAS PRIMAS Y MATERIALES")).AddStyle(estiloTitulo);
            tituloTable.AddCell(tituloCelda.SetBorder(Border.NO_BORDER));

            tituloCelda = new Cell(1, 1).Add(new Paragraph("N°: " + datosReporte.Cabecera.Analisis)).AddStyle(estiloTituloAnalisis);
            tituloTable.AddCell(tituloCelda.SetBorder(Border.NO_BORDER));

            document.Add(tituloTable);

            #endregion

            #region cabecera

            Table cabeceraTable = new Table(new float[] { 12, 23, 16, 20, 17, 12 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(10);

            Cell cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Descripción:").SetMargins(6, 0, 0, 6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 5).Add(new Paragraph(datosReporte.Cabecera.Item).SetMarginTop(6)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Proveedor:").SetMarginLeft(6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 3).Add(new Paragraph(datosReporte.Cabecera.Proveedor)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Lote:")).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.Lote)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Cantidad:").SetMarginLeft(6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.Cantidad.ToString("#,##0.##"))).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Unidad medida:")).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.Unidad)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Fecha de Vcto:")).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.FechaVencimiento.ToString("dd/MM/yyyy"))).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("N° Ingreso:").SetMargins(0, 0, 6, 6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.NroIngreso).SetMarginBottom(6)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("F. Análisis:").SetMarginBottom(6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.FechaAnalisis.ToString("dd/MM/yyyy")).SetMarginBottom(6)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph("Fecha de reanálisis:").SetMarginBottom(6)).AddStyle(estiloCabeceraTitulo);
            cabeceraTable.AddCell(cabeceraCelda);

            cabeceraCelda = new Cell(1, 1).Add(new Paragraph(datosReporte.Cabecera.FechaReAnalisis.ToString("dd/MM/yyyy")).SetMarginBottom(6)).AddStyle(estiloCabeceraTexto);
            cabeceraTable.AddCell(cabeceraCelda);

            document.Add(cabeceraTable.SetBorder(new SolidBorder(0.5f)));

            #endregion

            #region detalle

            Table detalleTable = new Table(new float[] { 29, 43, 12, 16 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(10);

            Cell detalleCelda = new Cell(1, 1).Add(new Paragraph("Pruebas Efectuadas")).AddStyle(estiloDetalleTitulo);
            detalleTable.AddHeaderCell(detalleCelda);

            detalleCelda = new Cell(1, 1).Add(new Paragraph("Especificaciones")).AddStyle(estiloDetalleTitulo);
            detalleTable.AddHeaderCell(detalleCelda);

            detalleCelda = new Cell(1, 1).Add(new Paragraph("Resultados")).AddStyle(estiloDetalleTitulo);
            detalleTable.AddHeaderCell(detalleCelda);

            detalleCelda = new Cell(1, 1).Add(new Paragraph("Metodología")).AddStyle(estiloDetalleTitulo);
            detalleTable.AddHeaderCell(detalleCelda);

            foreach (PlantillaDetalleProtocoloDTO prueba in datosReporte.Detalle)
            {
                detalleCelda = new Cell(1, 1).Add(new Paragraph(prueba.Prueba)).AddStyle(estiloDetalleTexto);
                detalleTable.AddCell(detalleCelda);

                detalleCelda = new Cell(1, 1).Add(new Paragraph(prueba.Especificacion).SetPaddingLeft(2)).AddStyle(estiloDetalleTexto);
                detalleTable.AddCell(detalleCelda);

                detalleCelda = new Cell(1, 1).Add(new Paragraph(prueba.Valor)).AddStyle(estiloDetalleTexto).SetTextAlignment(TextAlignment.CENTER);
                detalleTable.AddCell(detalleCelda);

                detalleCelda = new Cell(1, 1).Add(new Paragraph(prueba.Metodologia)).AddStyle(estiloDetalleTexto);
                detalleTable.AddCell(detalleCelda);
            }

            document.Add(detalleTable);
            #endregion

            #region observaciones

            Paragraph seccionObservacion = new Paragraph("OBSERVACIONES:").SetFontSize(10).SetMarginTop(10).SetFont(fuenteNegrita);
            document.Add(seccionObservacion);

            string observaciones = datosReporte.Detalle[0].Comentarios ?? "";

            seccionObservacion = new Paragraph(observaciones).SetFontSize(9).SetMarginTop(2).SetBorder(Border.NO_BORDER)
                .SetBorderBottom(new SolidBorder(0.5f)).SetMinHeight(13);

            document.Add(seccionObservacion);

            #endregion

            #region conclusion

            Table conclusionTable = new Table(new float[] { 20, 28, 4, 28, 4, 16 }).SetWidth(UnitValue.CreatePercentValue(100))
               .SetFixedLayout().SetMarginTop(13).SetBorder(new SolidBorder(0.5f));

            string conclusion = datosReporte.Detalle[0].ConclusionFlag ?? "";

            Cell conclusionCell = new Cell(1, 6).Add(new Paragraph("")).SetFontSize(9).SetBorder(Border.NO_BORDER).SetHeight(2);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph("CONCLUSIÓN:")).SetFontSize(10).SetFont(fuenteNegrita)
                    .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetBorder(Border.NO_BORDER).SetPaddingLeft(7).SetHeight(14);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph("APROBADO:")).SetFont(fuenteNegrita).SetFontSize(9)
                    .SetTextAlignment(TextAlignment.RIGHT).SetBorder(Border.NO_BORDER).SetHeight(14).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph(conclusion == "S" ? "X" : "")).SetFont(fuenteNegrita).SetFontSize(9)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetTextAlignment(TextAlignment.CENTER).SetHeight(14);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph("RECHAZADO:").SetFont(fuenteNegrita).SetFontSize(9)
                    .SetTextAlignment(TextAlignment.RIGHT)).SetBorder(Border.NO_BORDER).SetHeight(14).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph(conclusion == "N" ? "X" : "")).SetFont(fuenteNegrita).SetFontSize(9)
                    .SetTextAlignment(TextAlignment.CENTER).SetHeight(14).SetVerticalAlignment(VerticalAlignment.MIDDLE);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph("")).SetFontSize(9).SetBorder(Border.NO_BORDER).SetHeight(14);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 6).Add(new Paragraph("")).SetFontSize(9).SetBorder(Border.NO_BORDER).SetHeight(2);
            conclusionTable.AddCell(conclusionCell);

            document.Add(conclusionTable);

            #endregion

            document.Close();

            byte[] file = ms.ToArray();

            if (file == null || file.Length == 0)
            {
                pdf.Close();
                writer.Close();
                ms.Close();
                return reporte;
            }

            reporte = Convert.ToBase64String(file, 0, file.Length);

            pdf.Close();
            writer.Close();
            ms.Close();

            return reporte;
        }
    }
}
