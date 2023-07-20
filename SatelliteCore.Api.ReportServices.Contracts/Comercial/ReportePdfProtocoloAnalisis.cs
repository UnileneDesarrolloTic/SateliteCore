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
using SatelliteCore.Api.Models.Report.Comercial;
using System;
using System.Collections.Generic;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.Comercial
{
    public class ReportePdfProtocoloAnalisis
    {
        private readonly int idioma;
        private readonly List<ProtocoloReportModel> protocolos;

        private Image imgFirmaLiliaHurtado;
        private Image imgFirmaMilagrosMunoz;
        //private Image imgFooter;
        private Image imglogoUnilene;
        private Image imgConclusion;
        private Image imgConclusionIngles;

        private Style estiloTitulo;
        private Style estiloCabeceraSubtituloLote;
        private Style estiloCabeceraSubtituloCodsut;
        private Style estiloTexto;
        private Style estiloNegrita;
        private Style estiloTextoDetalleProtocolo;
        private Style estiloNegritaDetalleProtocolo;

        private Color bgColorfondo;
        string versionProtocolo = "";

        public ReportePdfProtocoloAnalisis(int idioma, List<ProtocoloReportModel> protocolos, string versionProtocolo)
        {
            this.idioma = idioma;
            this.protocolos = protocolos;
            this.versionProtocolo = versionProtocolo;
            InicializarObjetosComunes();
        }

        public string GenerarReporte()
        {
            string reporte = null;

            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Protocolo de análisis");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(5, 15, 30, 15);

            pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterRevisionProtocolo(this.versionProtocolo));
            int contarpagina = 0;
            if (idioma == 1)
            {
                foreach (ProtocoloReportModel protocolo in protocolos)
                {
                    contarpagina++;
                    AgregarReporteProtocoloEspaniol(document, protocolo);

                    if (protocolos.Count != contarpagina)
                        document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }

            }
            else
            {
                foreach (ProtocoloReportModel protocolo in protocolos)
                {
                    contarpagina++;
                    AgregarReporteProtocoloIngles(document, protocolo);

                    if (protocolos.Count != contarpagina)
                        document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }

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

        private void AgregarReporteProtocoloEspaniol(Document document, ProtocoloReportModel protocolo )
        {
            /*Table PiePagina = new Table(2).UseAllAvailableWidth().SetFixedLayout();

            Cell cellFPiePagina = new Cell(1, 1).Add(imgFooter)
                .SetFixedPosition(0f, document.GetBottomMargin() - 2, 0f)
                .SetBorder(Border.NO_BORDER);

            PiePagina.AddCell(cellFPiePagina);
            document.Add(PiePagina);*/

            Table tablaDatosTitulo = new Table(3).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellTitulo = new Cell(1, 1).Add(imglogoUnilene)
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaDatosTitulo.AddCell(cellTitulo);

            cellTitulo = new Cell(1, 1).Add(new Paragraph("PROTOCOLO DE ANÁLISIS")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°: " + protocolo.Cabecera.OrdenFabricacion + " \n").AddStyle(estiloCabeceraSubtituloLote))
                            .Add(new Paragraph(protocolo.Cabecera.NumeroParte + " \n").AddStyle(estiloCabeceraSubtituloCodsut))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetBorder(Border.NO_BORDER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);

            Table tablaDatosdeCabecera = new Table(6).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout().SetPaddingBottom(3);

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Producto:").AddStyle(estiloNegrita))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetHorizontalAlignment(HorizontalAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBorderRight(Border.NO_BORDER)
               .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(protocolo.Cabecera.ItemDescripcion).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(""))
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(""))
               .SetBorderBottom(Border.NO_BORDER)
               .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Presentación:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(protocolo.Cabecera.Presentacion).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Fabricación:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            string formatoFecha = (protocolo.Cabecera.OrdenFabricacion.StartsWith("PE") ? "MM-yyyy" : "dd-MM-yyyy");

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaFabricacion.ToString(formatoFecha)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("N° de Lote:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.Lote).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Tamaño de Lote : ").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorder(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(string.Format("{0:###,##0.##}", protocolo.Cabecera.TamanoLote)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetWidth(60.9f)
                .SetBorder(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Expira:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaExpiracion.ToString(formatoFecha)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Marca:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(protocolo.Cabecera.Marca).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Análisis:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaAnalisis.ToString("dd-MM-yyyy")).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            document.Add(tablaDatosdeCabecera);

            Table tablaDetalleProtocolo = new Table(9).UseAllAvailableWidth();
            tablaDetalleProtocolo.SetFixedLayout();

            Cell cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph("Pruebas Efectuadas").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph("Especificaciones").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Resultados").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Metodologia")
            .AddStyle(estiloNegritaDetalleProtocolo))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER)
            .SetBackgroundColor(bgColorfondo);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);


            //FOREARCH
            foreach (ProtocoloDetalleModel detalle in protocolo.Detalle)
            {
                cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph(detalle.Prueba).AddStyle(estiloTextoDetalleProtocolo))
                   .SetTextAlignment(TextAlignment.LEFT)
                   .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.TOP)
                   .SetPaddingBottom(3)
                   .SetBorderLeft(Border.NO_BORDER)
                   .SetBorderBottom(Border.NO_BORDER)
                   .SetBorderTop(Border.NO_BORDER)
                   .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph(detalle.Especificacion).AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(detalle.Resultado ?? "").AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(detalle.Metodologia).AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);
            }

            document.Add(tablaDetalleProtocolo);

            Table tablaTecnicaPropia = new Table(6).UseAllAvailableWidth();
            tablaTecnicaPropia.SetFixedLayout();

            Cell cellTecnicaPropia = new Cell(1, 6)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.Leyenda).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.MetodosEsterilizacion).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-10)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaTecnicaPropia.AddCell(cellTecnicaPropia);


            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.NormasISO).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-10)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            document.Add(tablaTecnicaPropia);

            Table tablaobservacion = new Table(1).UseAllAvailableWidth();
            tablaobservacion.SetFixedLayout();

            Cell cellObservarcion = new Cell(1, 1).Add(new Paragraph("OBSERVACION: ").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(imgConclusion)
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            document.Add(tablaobservacion);

            Table tablaFirmaResponsables = new Table(2).UseAllAvailableWidth();
            tablaFirmaResponsables.SetFixedLayout();

            Cell cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaLiliaHurtado)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetPaddingLeft(50)
            .SetBorder(Border.NO_BORDER);
            tablaFirmaResponsables.AddCell(cellFirmaResponsables);

            cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaMilagrosMunoz)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetPaddingLeft(50)
            .SetBorder(Border.NO_BORDER);
            tablaFirmaResponsables.AddCell(cellFirmaResponsables);

            document.Add(tablaFirmaResponsables);
        }

        private void AgregarReporteProtocoloIngles(Document document, ProtocoloReportModel protocolo)
        {
            /*Table PiePagina = new Table(2).UseAllAvailableWidth().SetFixedLayout();

            Cell cellFPiePagina = new Cell(1, 1).Add(imgFooter)
                .SetFixedPosition(0f, document.GetBottomMargin() - 2, 0f)
                .SetBorder(Border.NO_BORDER);

            PiePagina.AddCell(cellFPiePagina);
            document.Add(PiePagina);*/

            Table tablaDatosTitulo = new Table(3).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellTitulo = new Cell(1, 1).Add(imglogoUnilene)
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaDatosTitulo.AddCell(cellTitulo);

            cellTitulo = new Cell(1, 1).Add(new Paragraph("ANALYSIS PROTOCOL")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°: " + protocolo.Cabecera.OrdenFabricacion + " \n").AddStyle(estiloCabeceraSubtituloLote))
                            .Add(new Paragraph(protocolo.Cabecera.NumeroParte + " \n").AddStyle(estiloCabeceraSubtituloCodsut))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetBorder(Border.NO_BORDER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);

            Table tablaDatosdeCabecera = new Table(8).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout().SetPaddingBottom(3);

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Product:").AddStyle(estiloNegrita))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetHorizontalAlignment(HorizontalAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBorderRight(Border.NO_BORDER)
               .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.ItemDescripcion).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(""))
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(""))
               .SetBorderBottom(Border.NO_BORDER)
               .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Presentation:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 4).Add(new Paragraph(protocolo.Cabecera.Presentacion).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 2).Add(new Paragraph("Date of manufacture:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            string formatoFecha = (protocolo.Cabecera.OrdenFabricacion.StartsWith("PE") ? "MM-yyyy" : "dd-MM-yyyy");

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaFabricacion.ToString(formatoFecha)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Batch No:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.Lote).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Lot size : ").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorder(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(string.Format("{0:###,##0.##}", protocolo.Cabecera.TamanoLote)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetWidth(60.9f)
                .SetBorder(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(""))
            .SetWidth(60.9f)
            .SetBorder(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 2).Add(new Paragraph("Expiration date:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaExpiracion.ToString(formatoFecha)).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Brand:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 4).Add(new Paragraph(protocolo.Cabecera.Marca).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 2).Add(new Paragraph("Date of analysis:").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(protocolo.Cabecera.FechaAnalisis.ToString("dd-MM-yyyy")).AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            document.Add(tablaDatosdeCabecera);

            Table tablaDetalleProtocolo = new Table(9).UseAllAvailableWidth();
            tablaDetalleProtocolo.SetFixedLayout();

            Cell cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph("Tests performed").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph("Specifications").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Results").AddStyle(estiloNegritaDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBackgroundColor(bgColorfondo);

            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Methodology")
            .AddStyle(estiloNegritaDetalleProtocolo))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER)
            .SetBackgroundColor(bgColorfondo);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);


            //FOREARCH
            foreach (ProtocoloDetalleModel detalle in protocolo.Detalle)
            {
                cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph(detalle.Prueba).AddStyle(estiloTextoDetalleProtocolo))
                   .SetTextAlignment(TextAlignment.LEFT)
                   .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.TOP)
                   .SetPaddingBottom(3)
                   .SetBorderLeft(Border.NO_BORDER)
                   .SetBorderBottom(Border.NO_BORDER)
                   .SetBorderTop(Border.NO_BORDER)
                   .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph(detalle.Especificacion).AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.JUSTIFIED)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(detalle.Resultado ?? "").AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(detalle.Metodologia).AddStyle(estiloTextoDetalleProtocolo))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.TOP)
                    .SetPaddingBottom(3)
                    .SetBorderLeft(Border.NO_BORDER)
                    .SetBorderBottom(Border.NO_BORDER)
                    .SetBorderTop(Border.NO_BORDER)
                    .SetBorderRight(Border.NO_BORDER);

                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);
            }

            document.Add(tablaDetalleProtocolo);

            Table tablaTecnicaPropia = new Table(6).UseAllAvailableWidth();
            tablaTecnicaPropia.SetFixedLayout();

            Cell cellTecnicaPropia = new Cell(1, 6)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.Leyenda).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.MetodosEsterilizacion).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-10)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaTecnicaPropia.AddCell(cellTecnicaPropia);


            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(protocolo.Cabecera.NormasISO).AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetMarginTop(-10)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            cellTecnicaPropia = new Cell(1, 1)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);

            tablaTecnicaPropia.AddCell(cellTecnicaPropia);

            document.Add(tablaTecnicaPropia);

            Table tablaobservacion = new Table(1).UseAllAvailableWidth();
            tablaobservacion.SetFixedLayout();

            Cell cellObservarcion = new Cell(1, 1).Add(new Paragraph("OBSERVATIONS: ").AddStyle(estiloNegrita))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);

            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(imgConclusionIngles)
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            document.Add(tablaobservacion);

            Table tablaFirmaResponsables = new Table(2).UseAllAvailableWidth();
            tablaFirmaResponsables.SetFixedLayout();

            Cell cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaLiliaHurtado)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetPaddingLeft(50)
            .SetBorder(Border.NO_BORDER);
            tablaFirmaResponsables.AddCell(cellFirmaResponsables);

            cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaMilagrosMunoz)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetPaddingLeft(50)
            .SetBorder(Border.NO_BORDER);
            tablaFirmaResponsables.AddCell(cellFirmaResponsables);

            document.Add(tablaFirmaResponsables);
        }

        private void InicializarObjetosComunes()
        {
            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            string rutaFirmaLiliaHurtado = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLiliaHurtadoDias.jpg");
            string rutaFirmaMilagrosMunoz = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaMilagrosMunozTafur.jpg");
            //string rutaImagenFooter = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Protocolo.png");
            string rutaLogoUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            string conclusion = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\ConclusionProtocolo.png");
            string conclusionIngles = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\ConclusionProtocoloIngles.png");

            bgColorfondo = new DeviceRgb(217, 217, 217);

            imgFirmaLiliaHurtado = new Image(ImageDataFactory.Create(rutaFirmaLiliaHurtado))
             .SetWidth(150)
             .SetHeight(65)
             .SetMarginBottom(0)
             .SetPadding(0)
             .SetTextAlignment(TextAlignment.CENTER);

            imgFirmaMilagrosMunoz = new Image(ImageDataFactory.Create(rutaFirmaMilagrosMunoz))
             .SetWidth(150)
             .SetHeight(65)
             .SetMarginBottom(0)
             .SetPadding(0)
             .SetTextAlignment(TextAlignment.CENTER);

            /*imgFooter = new Image(ImageDataFactory
             .Create(rutaImagenFooter))
             .SetWidth(400)
             .SetHeight(15)
             .SetMarginBottom(0)
             .SetPadding(0)
             .ScaleAbsolute(50f, 50f)
             .SetTextAlignment(TextAlignment.CENTER);*/

            imglogoUnilene = new Image(ImageDataFactory
               .Create(rutaLogoUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            imgConclusion = new Image(ImageDataFactory
               .Create(conclusion))
               .SetWidth(560)
               .SetHeight(40)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            imgConclusionIngles = new Image(ImageDataFactory
               .Create(conclusionIngles))
               .SetWidth(560)
               .SetHeight(40)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            

            estiloTitulo = new Style()
               .SetFontSize(14)
               .SetFont(fuenteNegrita)
               .SetMarginTop(-3)
               .SetFontColor(ColorConstants.BLACK)
               .SetTextAlignment(TextAlignment.CENTER);

            estiloCabeceraSubtituloLote = new Style()
               .SetFontSize(14)
               .SetFont(fuenteNegrita)
               .SetMarginTop(-3)
               .SetFontColor(ColorConstants.BLACK)
               .SetTextAlignment(TextAlignment.RIGHT);

            estiloCabeceraSubtituloCodsut = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNormal)
                .SetMarginTop(9)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.RIGHT);

            estiloTexto = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            estiloNegrita = new Style()
               .SetFontSize(10)
               .SetFont(fuenteNegrita)
               .SetFontColor(ColorConstants.BLACK);

            estiloTextoDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            estiloNegritaDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);
        }
    }

    public class FooterRevisionProtocolo : IEventHandler
    {
        String header;

        public FooterRevisionProtocolo(String header)
        {
            this.header = header;
        }

        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent documentoEvento = (PdfDocumentEvent)@event;
            PdfDocument pdf = documentoEvento.GetDocument();
            PdfPage pagina = documentoEvento.GetPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pagina.NewContentStreamAfter(), pagina.GetResources(), pdf);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            Style estiloFooter = new Style().SetFontSize(8)
                    .SetFont(fuenteNegrita)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetMargin(0)
                    .SetPadding(0)
                    .SetFontSize(8);

            Table tablaResult = new Table(1).SetWidth(UnitValue.CreatePercentValue(100)).SetMargin(0).SetPadding(0);

            Cell footer = new Cell(1, 1).Add(new Paragraph(this.header).AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0);

            tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

            Rectangle rectangulo = new Rectangle(15, -20, pagina.GetPageSize().GetWidth() - 70, 50);

            new Canvas(pdfCanvas, rectangulo).Add(tablaResult);

        }
    }
}
