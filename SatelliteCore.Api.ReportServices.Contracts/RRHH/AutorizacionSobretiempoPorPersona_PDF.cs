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
using SatelliteCore.Api.Models.Report.RRHH;
using System;
using System.Collections.Generic;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.RRHH
{
    public class AutorizacionSobretiempoPorPersona_PDF
    {
        public AutorizacionSobretiempoPersonaDTO DatosReporte { get; private set; }

        public AutorizacionSobretiempoPorPersona_PDF(AutorizacionSobretiempoPersonaDTO datos)
        {
            DatosReporte = datos;
        }


        public string Exportar()
        {
            string reporte = null;

            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Autorización de sobretiempo");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A5);
            document.SetMargins(5, 25, 30, 25);

            int index = 1;
            int cantidadRegistros = DatosReporte.Cabecera.Count;

            foreach (AutoSobretiempoPersonaCabeceraDTO persona in DatosReporte.Cabecera)
            {
                List<AutoSobretiempoPersonaDetalleDTO> detalle = DatosReporte.Detalle.FindAll(x => x.IdPersona == persona.IdPersona);
                GenerarFormato(document, persona, detalle);

                if(index != cantidadRegistros)
                    document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                index++;
            }

            document.Close();

            byte[] file = ms.ToArray();

            if (file == null || file.Length == 0)
                return reporte;

            reporte = Convert.ToBase64String(file, 0, file.Length);

            pdf.Close();
            writer.Close();
            ms.Close();

            return reporte;
        }

        private void GenerarFormato(Document document, AutoSobretiempoPersonaCabeceraDTO cabecera, List<AutoSobretiempoPersonaDetalleDTO> detalle)
        {
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

            headerDocTable.AddCell(cellHeaderDoc).SetMarginTop(9f);

            cellHeaderDoc = new Cell(1, 2).Add(new Paragraph("Apellidos y Nombres: ").SetPaddingTop(2f).SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 4).Add(new Paragraph(cabecera.Nombres).SetPaddingTop(2f));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 2).Add(new Paragraph("Área: ").SetPaddingTop(2f).SetFont(fuenteNegrita));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER));

            cellHeaderDoc = new Cell(1, 4).Add(new Paragraph(cabecera.Area));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER).SetPaddingTop(2f));

            Paragraph textoAutorizacion = new Paragraph().SetFont(fuenteNegrita);
            textoAutorizacion.Add(new Text("Autorización: ").SetFont(fuenteNegrita));
            textoAutorizacion.Add(new Text(", el que suscribe el presente documento, manifiesta que voluntariamente han realizado trabajo en sobretiempo, según el siguiente detalle: ").SetFont(fuenteNormal));

            cellHeaderDoc = new Cell(1, 6).Add(textoAutorizacion.SetPaddingTop(4f));
            headerDocTable.AddCell(cellHeaderDoc.SetBorder(Border.NO_BORDER)).SetMarginBottom(4f).SetFontSize(9);

            document.Add(headerDocTable);

            Table detalleTable = new Table(new float[] { 10f, 25f, 25f, 25f, 15f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();
            //Table detalleTable = new Table(new float[] { 10f, 35f, 20f, 20f, 15f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell cellDetalle = new Cell(2, 1).Add(new Paragraph("N°").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 1).Add(new Paragraph("Fecha").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 2).Add(new Paragraph("Horario de trabajo en sobretiempo").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 1).Add(new Paragraph("N° Horas").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("Hora de Inicio").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("Hora de Termino").AddStyle(headerDetalle)).AddStyle(centrado);
            detalleTable.AddCell(cellDetalle);

            int indice = 1;

            foreach (AutoSobretiempoPersonaDetalleDTO d in detalle)
            {

                cellDetalle = new Cell(1, 1).Add(new Paragraph(indice.ToString()).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(d.FechaRegistro.ToString("dd/MM/yyyy")).AddStyle(centrado).AddStyle(bodyDetalle).SetMarginLeft(2f));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(d.HoraInicio).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(d.HoraFin).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(d.Cant_horas.ToString()).AddStyle(centrado).AddStyle(bodyDetalle));
                detalleTable.AddCell(cellDetalle);

                indice++;
            }

            document.Add(detalleTable);

            document.Add(saltoLinea);
            document.Add(saltoLinea);
            document.Add(saltoLinea);

            Table footerTable = new Table(new float[] { 50f, 50f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell cellFooter = new Cell(1, 1).Add(new Paragraph("-------------------------------------------\nFirma del trabajador").SetFontSize(10));
            footerTable.AddCell(cellFooter.SetBorder(Border.NO_BORDER));

            document.Add(footerTable.SetTextAlignment(TextAlignment.CENTER));
        }

    }
}
