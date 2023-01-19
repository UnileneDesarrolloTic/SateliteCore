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
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.RRHH
{
    public class FormatoAutorizacionSobretiempo_PDF
    {
        public DatosFormatoHorasExtrasCabeceraModel Cabecera { get; private set; }
        public List<DatosFormatoHorasExtrasDetalle> Detalle { get; private set; }

        public FormatoAutorizacionSobretiempo_PDF(DatosFormatoHorasExtrasCabeceraModel cabecera, List<DatosFormatoHorasExtrasDetalle> detalle)
        {
            Cabecera = cabecera;
            Detalle = detalle;
        }

        public string Exportar()
        {
            string reporte = null;

            MemoryStream ms = new MemoryStream();

            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Formato de autorización de trabajo en sobretiempo.");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(15, 30, 10, 30);

            string rutaLogoUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image img = new Image(ImageDataFactory.Create(rutaLogoUnilene))
                .SetWidth(115)
                .SetHeight(47)
                .SetMarginBottom(0)
                .SetPadding(0)
                .SetTextAlignment(TextAlignment.CENTER);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Style estiloTitulo = new Style().SetFontSize(12).SetFont(fuenteNegrita).SetFontColor(ColorConstants.BLACK).SetPaddingBottom(10f);
            Style headerDetalle = new Style().SetFontSize(9).SetFont(fuenteNegrita).SetFontColor(ColorConstants.BLACK).SetHorizontalAlignment(HorizontalAlignment.CENTER);
            Style bodyDetalle = new Style().SetFontSize(9).SetFont(fuenteNormal).SetFontColor(ColorConstants.BLACK).SetHorizontalAlignment(HorizontalAlignment.CENTER);
            Style centrado = new Style().SetVerticalAlignment(VerticalAlignment.MIDDLE).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetTextAlignment(TextAlignment.CENTER);
            Paragraph saltoLinea = new Paragraph(new Text("\n"));

            document.Add(img);

            Table headerDocTable = new Table(new float[] { 1, 1, 1, 1, 1, 1 }).UseAllAvailableWidth();
            headerDocTable.SetWidth(UnitValue.CreatePercentValue(100));
            headerDocTable.SetFixedLayout();

            Cell cellHeaderDoc = new Cell(1, 6).Add(new Paragraph("FORMATO DE AUTORIZACIÓN DE TRABAJOS EN SOBRETIEMPO").AddStyle(estiloTitulo).SetUnderline(1f, -3f))
               .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER);

            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 3).Add(new Paragraph("Trabajo en sobre tiempo correspondiente al día: ").SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 3).Add(new Paragraph(Cabecera.FechaRegistro.ToString("dd/MM/yyyy")).SetPaddingLeft(2f));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));


            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("Área: ").SetPaddingTop(2f).SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 5).Add(new Paragraph(Cabecera.Descripcion).SetPaddingTop(2f));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("Persona: ").SetPaddingTop(2f).SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 5).Add(new Paragraph("OBRERO"));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER).SetPaddingTop(2f));

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("Justificación: ").SetPaddingTop(2f).SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 5).Add(new Paragraph(Cabecera.Justificacion).SetPaddingTop(2f).SetFontSize(10));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            Paragraph textoAutorizacion = new Paragraph().SetFont(fuenteNegrita);
            textoAutorizacion.Add(new Text("Autorización: ").SetFont(fuenteNegrita));
            textoAutorizacion.Add(new Text("Los que suscriben el presente documento, manifiestan que voluntariamente han realizado trabajo en sobretiempo, según el siguiente detalle: ").SetFont(fuenteNormal));

            cellHeaderDoc = new Cell(1, 6).Add(textoAutorizacion.SetPaddingTop(4f));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER)).SetMarginBottom(4f).SetFontSize(9);

            document.Add(headerDocTable.SetMarginBottom(2f));

            Table detalleTable = new Table(new float[] { 10f, 50f, 20f, 20f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell cellDetalle = new Cell(1, 1).Add(new Paragraph("N°").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("Apellidos y Nombres").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("N° Horas").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("       Firma       ").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            int indice = 1;


            foreach (DatosFormatoHorasExtrasDetalle d in Detalle)
            {

                cellDetalle = new Cell(1, 1).Add(new Paragraph(indice.ToString()).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(d.NombrePersona).AddStyle(bodyDetalle).SetMarginLeft(2f));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(string.Format("{0:##0.##}", d.horasextras)).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(""));
                detalleTable.AddCell(cellDetalle);

                indice++;
            }

            document.Add(detalleTable.SetMarginBottom(5));

            document.Add(saltoLinea);
            document.Add(saltoLinea);

            Table footerTable = new Table(new float[] { 50f, 50f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell cellFooter = new Cell(1, 1).Add(new Paragraph("-------------------------------------------------------\nFirma y Sello del Gerente o Jefe de Área").SetFontSize(10));
            footerTable.AddCell(cellFooter.SetBorder(Border.NO_BORDER));

            cellFooter = new Cell(1, 1).Add(new Paragraph("---------------------------------------------\nV°B° Recursos Humanos").SetFontSize(10));
            footerTable.AddCell(cellFooter.SetBorder(Border.NO_BORDER)).SetTextAlignment(TextAlignment.CENTER).SetHorizontalAlignment(HorizontalAlignment.CENTER);

            document.Add(footerTable.SetPaddingTop(10));

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
