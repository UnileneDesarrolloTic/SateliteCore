using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.AnalsisAguja
{
    public class FlexionAguja
    {
        public string GenerarReporte(string loteAnalisis, ObtenerAnalisisAgujaModel cabecera, List<AnalisisAgujaFlexionEntity> detalle)
        {

            string reporte = null;

            string fechaRegistro = detalle[0].FechaRegistro.ToString("dd/MM/yyyy hh:mm");

            MemoryStream ms = new MemoryStream();

            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Análisis de aguja prueba de flexión");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(5, 15, 30, 15);

            pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterFlexionAgujaEventHandler());

            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(110)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            document.Add(img);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

            Style estiloTitulo = new Style()
                .SetFontSize(14)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);

            Style estiloHeaderDG = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloSubTitulo = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloResultadoNegrita = new Style()
                .SetFont(fuenteNegrita)
                .SetFontSize(9)
                .SetFontColor(ColorConstants.BLACK);


            Paragraph header = new Paragraph("PRUEBA DE FLEXIÓN DE AGUJAS").AddStyle(estiloTitulo);

            document.Add(header);

            Table tablaDatosGenerales = new Table(6).UseAllAvailableWidth();

            Cell cellDG = new Cell(1, 1).Add(new Paragraph("Proveedor:")
                .AddStyle(estiloHeaderDG))
                .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 2).Add(new Paragraph(cabecera.Proveedor)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("Aguja:")
                .AddStyle(estiloHeaderDG))
                .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 2).Add(new Paragraph(cabecera.DescripcionItem)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("Fecha análisis:")
              .AddStyle(estiloHeaderDG))
              .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(fechaRegistro)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("N° Análisis:")
              .AddStyle(estiloHeaderDG))
              .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(loteAnalisis)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("Cantidad:")
               .AddStyle(estiloHeaderDG))
               .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.CantidadPruebas.ToString())
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("Ord. Compra:")
                .AddStyle(estiloHeaderDG))
                .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.OrdenCompra)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("Item:")
                .AddStyle(estiloHeaderDG))
                .SetTextAlignment(TextAlignment.RIGHT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.Item)
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.CENTER);

            tablaDatosGenerales.AddCell(cellDG);


            document.Add(tablaDatosGenerales);
            document.Add(saltoLinea);


            Table tablaDetalleCiclos = new Table(new float[] { 1, 1, 1, 1, 1 });
            tablaDetalleCiclos.SetWidth(UnitValue.CreatePercentValue(100));
            tablaDetalleCiclos.SetFixedLayout();

            Cell cellTablaCicloDetalle;

            List<AnalisisAgujaFlexionEntity> listaCiclos = detalle.Where(x => x.TipoRegistro == 1).ToList();

            int maxIndice = listaCiclos.Max(x => x.Llave);
            List<int> gruposCiclos = GenerarListaGrupos(maxIndice);
            decimal valor = 0;


            foreach (int grupo in gruposCiclos)
            {

                Table tablaCiclos = new Table(2);

                Cell cellCiclo = new Cell(1, 1).Add(new Paragraph("").AddStyle(estiloTexto))
                    .SetBorder(Border.NO_BORDER)
                    .SetWidth(50)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaCiclos.AddHeaderCell(cellCiclo);

                cellCiclo = new Cell(1, 1).Add(new Paragraph("Ciclo").AddStyle(estiloTexto))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetWidth(35)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaCiclos.AddHeaderCell(cellCiclo);

                for (var i = 1; i <= 10; i++)
                {
                    valor = 0;

                    valor = listaCiclos.Where(x => x.Llave == (grupo + (i - 1))).FirstOrDefault().Valor;

                    cellCiclo = new Cell(1, 1).Add(new Paragraph("Ciclo " + (grupo + i)).AddStyle(estiloTexto))
                    .SetTextAlignment(TextAlignment.LEFT)
                    .SetPadding(0)
                    .SetPaddingLeft(5)
                    .SetMargin(0);

                    tablaCiclos.AddCell(cellCiclo);

                    cellCiclo = new Cell(1, 1).Add(new Paragraph(valor == 0 ? "" : Convert.ToInt32(valor).ToString()).AddStyle(estiloTexto))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetPadding(0)
                        .SetMargin(0);

                    tablaCiclos.AddCell(cellCiclo)
                        .SetMarginBottom(5)
                        .SetHorizontalAlignment(HorizontalAlignment.CENTER);

                }

                cellTablaCicloDetalle = new Cell(1, 1).Add(tablaCiclos).SetBorder(Border.NO_BORDER);
                tablaDetalleCiclos.AddCell(cellTablaCicloDetalle);
            }

            document.Add(tablaDetalleCiclos.SetMarginBottom(8));


            Paragraph subTitulo = new Paragraph("Resultado de la flexión - SERIE 400:").AddStyle(estiloSubTitulo);
            document.Add(subTitulo);

            Table resultadoFlexion = new Table(new float[] { 1, 1, 1, 1 });
            resultadoFlexion.SetWidth(UnitValue.CreatePercentValue(100));
            resultadoFlexion.SetFixedLayout();

            Cell cellTablaResultado;

            List<AnalisisAgujaFlexionEntity> listaResumenCiclos = detalle.Where(x => x.TipoRegistro == 2).ToList();

            int cantidadCiclosPorResumen = 0;

            foreach (AnalisisAgujaFlexionEntity resumen in listaResumenCiclos)
            {
                cantidadCiclosPorResumen = 0;
                cantidadCiclosPorResumen = detalle.Where(x => x.TipoRegistro == 1 && x.Valor == resumen.Llave).ToList().Count();

                Table tablaResultado = new Table(new float[] { 1, 1, 1 });
                tablaResultado.SetWidth(UnitValue.CreatePercentValue(100));

                Cell cellResultado = new Cell(1, 1).Add(new Paragraph($" {resumen.Llave} ciclos =")
                    .AddStyle(estiloResultadoNegrita))
                    .SetBorder(Border.NO_BORDER)
                    .SetTextAlignment(TextAlignment.RIGHT)
                    .SetWidth(48)
                    .SetPadding(0)
                    .SetMargin(0)
                    .SetPaddingRight(3);

                tablaResultado.AddCell(cellResultado);

                cellResultado = new Cell(1, 1).Add(new Paragraph(cantidadCiclosPorResumen.ToString()).AddStyle(estiloTexto))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetWidth(20)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaResultado.AddCell(cellResultado);

                cellResultado = new Cell(1, 1).Add(new Paragraph(resumen.Valor.ToString() + " %").AddStyle(estiloTexto).SetPaddingRight(2))
                    .SetTextAlignment(TextAlignment.RIGHT)
                    .SetWidth(40)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaResultado.AddCell(cellResultado);

                cellTablaResultado = new Cell(1, 1).Add(tablaResultado)
                    .SetBorder(Border.NO_BORDER)
                    .SetTextAlignment(TextAlignment.CENTER);

                resultadoFlexion.AddCell(cellTablaResultado);
            }


            document.Add(resultadoFlexion);

            document.Add(saltoLinea.SetMarginBottom(8));


            Table tableResponsable = new Table(new float[] { 1, 1 }).UseAllAvailableWidth();
            tableResponsable.SetWidth(UnitValue.CreatePercentValue(100));
            tableResponsable.SetFixedLayout();
            document.Add(saltoLinea);

            Cell cellResponsable = new Cell(1, 1).Add(new Paragraph($"REALIZADO POR: __________________")
                .AddStyle(estiloResultadoNegrita))
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetPadding(0)
                .SetMargin(0);

            tableResponsable.AddCell(cellResponsable);


            cellResponsable = new Cell(1, 1).Add(new Paragraph($"V° B°: __________________")
                .AddStyle(estiloResultadoNegrita))
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetPadding(0)
                .SetMargin(0);

            tableResponsable.AddCell(cellResponsable);
            document.Add(tableResponsable);

            document.Add(saltoLinea);

            //Paragraph footer = new Paragraph("F/CDC-078, Versión 01").AddStyle(estiloSubTitulo).SetMargin(0).SetPadding(0).SetFontSize(8);
            //document.Add(footer);
            //footer = new Paragraph("Vigente desde: 28/06/2019").AddStyle(estiloSubTitulo).SetMargin(0).SetPadding(0).SetFontSize(8);
            //document.Add(footer);

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

        private List<int> GenerarListaGrupos(int valorMax)
        {
            List<int> listaGrupos = new List<int>();

            for (int i = 0; i <= valorMax; i += 10)
                listaGrupos.Add(i);

            return listaGrupos;

        }

    }

    public class FooterFlexionAgujaEventHandler : IEventHandler
    {
        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent documentoEvento = (PdfDocumentEvent)@event;
            PdfDocument pdf = documentoEvento.GetDocument();
            PdfPage pagina = documentoEvento.GetPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pagina.NewContentStreamBefore(), pagina.GetResources(), pdf);


            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            Style estiloFooter = new Style().SetFontSize(8)
                    .SetFont(fuenteNegrita)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetMargin(0)
                    .SetPadding(0)
                    .SetFontSize(8)
                    ;
            Table tablaResult = new Table(1).SetWidth(UnitValue.CreatePercentValue(100)).SetMargin(0).SetPadding(0);

            Cell footer = new Cell(1,1).Add(new Paragraph("F/CDC-078, Versión 01").AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0);

            tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

            footer = new Cell(1,1).Add(new Paragraph("Vigente desde: 28/06/2019").AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0);

            tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

            Rectangle rectangulo = new Rectangle(15, 0, pagina.GetPageSize().GetWidth() - 70, 50);

            new Canvas(pdfCanvas, rectangulo).Add(tablaResult);

        }
    }

}
