using iText.Barcodes;
using iText.Commons.Actions;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SatelliteCore.Api.ReportServices.Contracts.OrdenServicio
{
    public class ReporteOrdenServicio_PDF
    {
        public DatosReporteOrdenServicioPDF_DTO OrdenServicio { get; set; }

        public ReporteOrdenServicio_PDF(DatosReporteOrdenServicioPDF_DTO ordenServicio)
        {
            OrdenServicio = ordenServicio;
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

            //pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterAutorizacionSobretiempoPorPersona_PDF());

            Document document = new Document(pdf, PageSize.A4.Rotate());
            document.SetMargins(5, 25, 30, 25);

            pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterReporteOrdenServicio());

            GenerarFormato(document, OrdenServicio, pdf);
           

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

        private void GenerarFormato(Document document, DatosReporteOrdenServicioPDF_DTO datosReporte, PdfDocument pdf)
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

            Style estiloTitulo = new Style().SetFontSize(14).SetFont(fuenteNegrita).SetFontColor(ColorConstants.BLACK).SetPaddingBottom(10f);
            Style headerDetalle = new Style().SetFontSize(8).SetFontColor(ColorConstants.BLACK);
            Style detalleStyle = new Style().SetFontSize(7f).SetFontColor(ColorConstants.BLACK);
            Style styleDetalleHeader = new Style().SetFontSize(7.5f).SetFont(fuenteNegrita).SetPadding(3f).SetPaddingTop(5f)
                .SetVerticalAlignment(VerticalAlignment.BOTTOM);
            Style centrado = new Style().SetVerticalAlignment(VerticalAlignment.MIDDLE).SetHorizontalAlignment(HorizontalAlignment.CENTER).SetTextAlignment(TextAlignment.CENTER);
            Style bordesEnY = new Style().SetBorder(Border.NO_BORDER).SetBorderTop(new SolidBorder(1)).SetBorderBottom(new SolidBorder(1));
            Table headerDocTable = new Table(new float[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 , 1, 1 }).UseAllAvailableWidth();
            headerDocTable.SetWidth(UnitValue.CreatePercentValue(100));
            headerDocTable.SetFixedLayout();

            Cell cellImageLogo = new Cell(1, 3).Add(img).SetBorder(Border.NO_BORDER);
            headerDocTable.AddCell(cellImageLogo);

            Cell cellHeaderDoc = new Cell(1, 9).Add(new Paragraph($"Orden de Servicio - {datosReporte.OrdenServicio}").AddStyle(estiloTitulo).SetUnderline(1f, -3f))
               .SetVerticalAlignment(VerticalAlignment.BOTTOM).SetBorder(Border.NO_BORDER).SetPaddingLeft(100f);

            headerDocTable.AddCell(cellHeaderDoc);
            document.Add(headerDocTable);


            Barcode128 code128 = new Barcode128(pdf);
            code128.SetFont(null);
            code128.SetCode(datosReporte.OrdenServicio);
            code128.SetCodeType(Barcode128.CODE128);

            Image code128Image = new Image(code128.CreateFormXObject(pdf)).SetHeight(37f).SetWidth(140f).SetAutoScale(false);

            code128Image.SetFixedPosition(1, 675, 550);

            document.Add(code128Image);

            Table tableCabecera = new Table(new float[] { 34f, 33f, 33f }).UseAllAvailableWidth();
            tableCabecera.SetWidth(UnitValue.CreatePercentValue(100));
            tableCabecera.SetFixedLayout();

            Paragraph textoTransportista = new Paragraph().SetFont(fuenteNegrita);
            textoTransportista.Add(new Text("TRANSPORTISTA: ").SetFont(fuenteNegrita));
            textoTransportista.Add(new Text(datosReporte.Transportista).SetFont(fuenteNormal));


            Cell cellCabecera = new Cell(1, 1).Add(textoTransportista.AddStyle(headerDetalle)).SetBorder(Border.NO_BORDER)
               .SetVerticalAlignment(VerticalAlignment.BOTTOM);

            tableCabecera.AddCell(cellCabecera);

            Paragraph textoPlaca = new Paragraph().SetFont(fuenteNegrita);
            textoPlaca.Add(new Text("PLACA: ").SetFont(fuenteNegrita));
            textoPlaca.Add(new Text("_______________").SetFont(fuenteNormal));

            cellCabecera = new Cell(1, 1).Add(textoPlaca.AddStyle(headerDetalle)).SetBorder(Border.NO_BORDER)
               .SetVerticalAlignment(VerticalAlignment.BOTTOM).AddStyle(centrado);

            tableCabecera.AddCell(cellCabecera);

            Paragraph fecha = new Paragraph().SetFont(fuenteNegrita);
            fecha.Add(new Text("FECHA: ").SetFont(fuenteNegrita));
            fecha.Add(new Text(datosReporte.Fecha).SetFont(fuenteNormal));

            cellCabecera = new Cell(1, 1).Add(fecha.AddStyle(headerDetalle)).SetBorder(Border.NO_BORDER)
               .SetVerticalAlignment(VerticalAlignment.BOTTOM).SetHorizontalAlignment(HorizontalAlignment.RIGHT).SetTextAlignment(TextAlignment.RIGHT);

            tableCabecera.AddCell(cellCabecera);

            document.Add(tableCabecera);

            Table tablaDetalle = new Table(new float[] { 9f, 30f, 30f, 10f, 9f, 6f, 7f }).UseAllAvailableWidth();
            tablaDetalle.SetWidth(UnitValue.CreatePercentValue(100));
            tablaDetalle.SetFixedLayout();
            tablaDetalle.SetMarginTop(5f).SetMarginTop(10f);

            Cell detalleHeader = new Cell(1, 1).Add(new Paragraph("GUIA").SetFont(fuenteNegrita))
               .AddStyle(styleDetalleHeader).AddStyle(centrado).AddStyle(bordesEnY).SetBorderLeft(new SolidBorder(1));
                
            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("CLIENTE").SetFont(fuenteNegrita))
               .AddStyle(styleDetalleHeader).AddStyle(centrado).AddStyle(bordesEnY);

            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("DIRECCIÓN DESTINO").SetFont(fuenteNegrita))
               .AddStyle(styleDetalleHeader).AddStyle(centrado).AddStyle(bordesEnY);

            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("DEPARTAMENTO").SetFont(fuenteNegrita))
               .AddStyle(styleDetalleHeader).AddStyle(centrado).AddStyle(bordesEnY);

            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("FACTURA").SetFont(fuenteNegrita))
               .AddStyle(styleDetalleHeader).AddStyle(centrado).AddStyle(bordesEnY);

            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("PESO").SetFont(fuenteNegrita))
               .SetVerticalAlignment(VerticalAlignment.BOTTOM).AddStyle(styleDetalleHeader).AddStyle(centrado)
               .AddStyle(bordesEnY);

            tablaDetalle.AddCell(detalleHeader);

            detalleHeader = new Cell(1, 1).Add(new Paragraph("BULTOS").SetFont(fuenteNegrita))
               .SetVerticalAlignment(VerticalAlignment.BOTTOM).AddStyle(styleDetalleHeader).AddStyle(centrado)
               .AddStyle(bordesEnY).SetBorderRight(new SolidBorder(1));

            tablaDetalle.AddCell(detalleHeader);

            int indice = 1;

            Cell cellDetalle;

            foreach (DatosReporteOrdenServicioDetallePDF_DTO detalle in datosReporte.Detalle)
            {
                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Guia).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Cliente).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Destino).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Departamento).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Factura).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Peso.ToString()).AddStyle(centrado).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f).SetPaddingBottom(3f);
                tablaDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.Bultos.ToString()).AddStyle(centrado).AddStyle(detalleStyle))
                    .SetBorder(Border.NO_BORDER).SetBorderBottom(new DottedBorder(1)).SetPaddingTop(5f);
                tablaDetalle.AddCell(cellDetalle);

                indice++;
            }

            document.Add(tablaDetalle.SetMarginBottom(15f));

            Table tableObservaciones = new Table(new float[] { 18f, 20f, 20f, 21f, 21f }).UseAllAvailableWidth().SetBorder(new SolidBorder(1));
            tableObservaciones.SetWidth(UnitValue.CreatePercentValue(100));         
            tableObservaciones.SetFixedLayout();
            tableObservaciones.SetMarginTop(5f).SetBorder(Border.NO_BORDER);

            Table tableObsVerificacion = new Table(new float[] { 60f, 40f }).UseAllAvailableWidth().SetBorder(new SolidBorder(1));
            tableObsVerificacion.SetWidth(UnitValue.CreatePercentValue(100));
            tableObsVerificacion.SetFixedLayout();

            Cell cellVerificacion = new Cell(1, 2).Add(new Paragraph("VERIFICACIÓN DE LIMPIEZA DE UNIDAD DE TRANSPORTE")
                .SetFontSize(6.5f).SetFont(fuenteNegrita)).SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER).SetPadding(5f);
                
            tableObsVerificacion.AddCell(cellVerificacion);

            cellVerificacion = new Cell(1, 1).Add(new Paragraph("CONFORME: ").SetFontSize(6.5f))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER).SetPaddingTop(4F);

            tableObsVerificacion.AddCell(cellVerificacion);

            Paragraph rectangulo = new Paragraph() .SetHeight(8) .SetWidth(8)
                    .SetBackgroundColor(ColorConstants.WHITE) .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f));

            cellVerificacion = new Cell(1, 1).Add(rectangulo).SetBorder(Border.NO_BORDER);
                
            tableObsVerificacion.AddCell(cellVerificacion);

            cellVerificacion = new Cell(1, 1).Add(new Paragraph("NO CONFORME: ").SetFontSize(6.5f))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER).SetPaddingTop(4F);

            tableObsVerificacion.AddCell(cellVerificacion);           

            cellVerificacion = new Cell(1, 1).Add(rectangulo).SetBorder(Border.NO_BORDER);
            tableObsVerificacion.AddCell(cellVerificacion);

            cellVerificacion = new Cell(1, 2).Add(new Paragraph("").SetFontSize(6.5f)).SetBorder(Border.NO_BORDER).SetHeight(5f);

            tableObsVerificacion.AddCell(cellVerificacion);

            cellVerificacion = new Cell(1, 2).Add(new Paragraph("________________________________\nV°B° ALMACEN").SetFont(fuenteNegrita).SetFontSize(6.5f))
                .SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER).SetMarginTop(8f).SetPaddingTop(8f);

            tableObsVerificacion.AddCell(cellVerificacion);


            Cell cellObservaciones = new Cell(6, 1).Add(tableObsVerificacion).SetBorder(Border.NO_BORDER);

            tableObservaciones.AddCell(cellObservaciones);


            cellObservaciones = new Cell(1, 4).Add(new Paragraph("   OBSERVACIONES:").SetFontSize(7).SetFont(fuenteNegrita))
                .SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1)).SetPaddingTop(5f);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 4).Add(new Paragraph(" ").SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1)).SetHeight(10F);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 4).Add(new Paragraph(" ").SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1)).SetHeight(10F);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 4).Add(new Paragraph(" ").SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(1)).SetHeight(10F);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 4).Add(new Paragraph(" ").SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetHeight(10F);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 1).Add(new Paragraph("________________________\nV°B° ALMACEN").SetFont(fuenteNegrita).SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 1).Add(new Paragraph("________________________\nVIGILANCIA\nRecibido").SetFont(fuenteNegrita).SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 1).Add(new Paragraph("________________________\nTRANSPORTISTA\nRecibido").SetFont(fuenteNegrita).SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);

            tableObservaciones.AddCell(cellObservaciones);

            cellObservaciones = new Cell(1, 1).Add(new Paragraph("________________________\nV°B° OFICINA VENTAS").SetFont(fuenteNegrita).SetFontSize(7))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);

            tableObservaciones.AddCell(cellObservaciones);

            document.Add(tableObservaciones);
        }
    }

    public class FooterReporteOrdenServicio : iText.Kernel.Events.IEventHandler
    {
        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent documentoEvento = (PdfDocumentEvent)@event;
            PdfDocument pdf = documentoEvento.GetDocument();
            PdfPage pagina = documentoEvento.GetPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pagina.NewContentStreamBefore(), pagina.GetResources(), pdf);

            int pageNumber = pdf.GetPageNumber(pagina);
            int numberOfPages = pdf.GetNumberOfPages();

            Style estiloFooter = new Style().SetFontSize(8)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetMargin(0)
                    .SetPadding(0)
                    .SetFontSize(8);

            Table tablaResult = new Table(new float[] { 1f, 1f }).SetWidth(UnitValue.CreatePercentValue(100)).SetMargin(0).SetPadding(0);

            Cell footer = new Cell(1, 1).Add(new Paragraph("F/ALM-006; Versión 01").AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0)
                .SetTextAlignment(TextAlignment.LEFT);

            tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

            footer = new Cell(1, 1).Add(new Paragraph($"Página {pageNumber} de {numberOfPages}").AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0).SetTextAlignment(TextAlignment.RIGHT);

            tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

            Rectangle rectangulo = new Rectangle(15, -20, pagina.GetPageSize().GetWidth() - 70, 50);

            new Canvas(pdfCanvas, rectangulo).Add(tablaResult);

        }
    }
}
