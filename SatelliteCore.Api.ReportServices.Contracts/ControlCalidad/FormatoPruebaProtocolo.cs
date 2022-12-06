using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.IO;

using System.Globalization;

namespace SatelliteCore.Api.ReportServices.Contracts.ControlCalidad
{
    public class FormatoPruebaProtocolo
    {
        public string ReporteFormatoPruebaProtocoloEspaniol(IEnumerable<DatosFormatoProtocoloPruebaModel> ListadoProtocolo, DatosFormatoNumeroLoteProtocoloModel Cabecera, bool Opcion)
        {
           
            string reporte = null;
            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("PROTOCOLO DE ANÁLISIS");
            docInfo.SetAuthor("UNILENE");

            Document document = new Document(pdf, PageSize.A4);
            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

            Color bgColour = new DeviceRgb(192, 192, 192);

            NumberFormatInfo formato = new CultureInfo("en-US").NumberFormat;
            formato.CurrencyGroupSeparator = ".";
            formato.NumberDecimalSeparator = ",";

            document.SetMargins(5, 15, 30, 15);
            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            string Conclusion = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\ConclusionProtocolo.png");
            string imagenFlooter = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Protocolo.png");
            string FirmaLiliaHurtado = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLiliaHurtadoDias.jpg");
            string FirmaFirmaMilagrosMunoz = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaMilagrosMunozTafur.jpg");

            Color bgColorfondo = new DeviceRgb(217, 217, 217);

           
            Image imagenFlooterss = new Image(ImageDataFactory
             .Create(imagenFlooter))
             .SetWidth(400)
             .SetHeight(15)
             .SetMarginBottom(0)
             .SetPadding(0)
             .ScaleAbsolute(50f, 50f)
             .SetTextAlignment(TextAlignment.CENTER);

            Image imgFirmaLiliaHurtado = new Image(ImageDataFactory
              .Create(FirmaLiliaHurtado))
              .SetWidth(150)
              .SetHeight(65)
              .SetMarginBottom(0)
              .SetPadding(0)
              .SetTextAlignment(TextAlignment.CENTER);

            Image imgFirmaFirmaMilagrosMunoz = new Image(ImageDataFactory
              .Create(FirmaFirmaMilagrosMunoz))
              .SetWidth(150)
              .SetHeight(65)
              .SetMarginBottom(0)
              .SetPadding(0)
              .SetTextAlignment(TextAlignment.CENTER);

            Image imgConclusion = new Image(ImageDataFactory
            .Create(Conclusion))
            .SetWidth(560)
            .SetHeight(40)
            .SetMarginBottom(0)
            .SetPadding(0)
            .SetTextAlignment(TextAlignment.LEFT);

            Table PiePagina = new Table(2).UseAllAvailableWidth();
            PiePagina.SetFixedLayout();


            Cell cellFPiePagina = new Cell(1, 1).Add(imagenFlooterss)
            .SetFixedPosition(0f, document.GetBottomMargin() - 2, 0f)
            .SetBorder(Border.NO_BORDER);

            PiePagina.AddCell(cellFPiePagina);

            document.Add(PiePagina);

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);


            Style estiloTitulo = new Style()
                .SetFontSize(14)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-3)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);


            Style estiloCabeceraSubtituloLote = new Style()
               .SetFontSize(14)
               .SetFont(fuenteNegrita)
               .SetMarginTop(-3)
               .SetFontColor(ColorConstants.BLACK)
               .SetTextAlignment(TextAlignment.RIGHT);

            Style estiloCabeceraSubtituloCodsut = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNormal)
                .SetMarginTop(9)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.RIGHT);


            Style estiloCabeceraVariable = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegrita = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTextoDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegritaDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);




            //CABECERA 
            Table tablaDatosTitulo = new Table(3).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellTitulo = new Cell(1, 1).Add(img)
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

            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°:"+ Cabecera.REFERENCIANUMERO+ " \n").AddStyle(estiloCabeceraSubtituloLote)).Add(new Paragraph(Cabecera.NUMERODEPARTE+" \n").AddStyle(estiloCabeceraSubtituloCodsut))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);


            Table tablaDatosdeCabecera = new Table(6).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout().SetPaddingBottom(3);
             /*
                .SetBorderBottomLeftRadius(new BorderRadius(0.5f))
                .SetBorderBottomRightRadius(new BorderRadius(0.5f))
                .SetBorderTopLeftRadius(new BorderRadius(0.5f))
                .SetBorderTopRightRadius(new BorderRadius(0.5f));
              */

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Producto:")
           .AddStyle(estiloCabeceraVariable))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetHorizontalAlignment(HorizontalAlignment.CENTER)
           .SetVerticalAlignment(VerticalAlignment.MIDDLE)
           .SetBorderTopLeftRadius(new BorderRadius(0.5f))
           .SetBorderRight(Border.NO_BORDER)
           .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.ITEMDESCRIPCION)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);



            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Presentación")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.Presentacion)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Fabricación:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAPRODUCCION.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("N° de Lote:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.ORDENFABRICACION)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Tamaño de Lote : ")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.CANTIDADPRODUCIDA.ToString("#,##0"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);



            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Expira:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAEXPIRACION.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Marca:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.MARCA)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Fecha de Análisis:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAANALISIS.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            document.Add(tablaDatosdeCabecera);

            Table tablaDetalleProtocolo = new Table(9).UseAllAvailableWidth();
            tablaDetalleProtocolo.SetFixedLayout();

            Cell cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph("Pruebas Efectuadas")
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

            cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph("Especificaciones")
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

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Resultados")
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
            foreach (DatosFormatoProtocoloPruebaModel item in ListadoProtocolo)
            {
                cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph(item.DESCRIPCION_PRUEBA)
               .AddStyle(estiloTextoDetalleProtocolo))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetHorizontalAlignment(HorizontalAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.TOP)
               .SetPaddingBottom(3)
               .SetBorderLeft(Border.NO_BORDER)
               .SetBorderBottom(Border.NO_BORDER)
               .SetBorderTop(Border.NO_BORDER)
               .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph(item.ESPECIFICACION)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.JUSTIFIED)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetPaddingBottom(3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.RESULTADO)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetPaddingBottom(3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.DESCRIPCION_METODOLOGIA)
                .AddStyle(estiloTextoDetalleProtocolo))
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

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.TECNICA)
            .AddStyle(estiloTextoDetalleProtocolo))
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



            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.METODO)
            .AddStyle(estiloTextoDetalleProtocolo))
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


            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.DETALLE)
            .AddStyle(estiloTextoDetalleProtocolo))
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

            Cell cellObservarcion = new Cell(1, 1).Add(new Paragraph("OBSERVACION: ")
            .AddStyle(estiloNegrita))
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

            if (Opcion)
            {
                Table tablaFirmaResponsables = new Table(2).UseAllAvailableWidth();
                tablaFirmaResponsables.SetFixedLayout();


                Cell cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaLiliaHurtado)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPaddingLeft(50)
                .SetBorder(Border.NO_BORDER);
                tablaFirmaResponsables.AddCell(cellFirmaResponsables);

                cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaFirmaMilagrosMunoz)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPaddingLeft(50)
                .SetBorder(Border.NO_BORDER);
                tablaFirmaResponsables.AddCell(cellFirmaResponsables);

                document.Add(tablaFirmaResponsables);

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


        public string ReporteFormatoPruebaProtocoloIngles(IEnumerable<DatosFormatoProtocoloPruebaModel> ListadoProtocolo, DatosFormatoNumeroLoteProtocoloModel Cabecera, bool Opcion)
        {

            string reporte = null;
            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("ANALYSIS PROTOCOL");
            docInfo.SetAuthor("UNILENE");

            Document document = new Document(pdf, PageSize.A4);
            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

            Color bgColour = new DeviceRgb(192, 192, 192);

            NumberFormatInfo formato = new CultureInfo("en-US").NumberFormat;
            formato.CurrencyGroupSeparator = ".";
            formato.NumberDecimalSeparator = ",";


            document.SetMargins(5, 15, 30, 15);
            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            string Conclusion = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\ConclusionProtocoloIngles.png");
            string imagenFlooter = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Protocolo.png");
            string FirmaLiliaHurtado = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLiliaHurtadoDias.jpg");
            string FirmaFirmaMilagrosMunoz = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaMilagrosMunozTafur.jpg");

            Color bgColorfondo = new DeviceRgb(217, 217, 217);


            Image imagenFlooterss = new Image(ImageDataFactory
             .Create(imagenFlooter))
             .SetWidth(400)
             .SetHeight(15)
             .SetMarginBottom(0)
             .SetPadding(0)
             .ScaleAbsolute(50f, 50f)
             .SetTextAlignment(TextAlignment.CENTER);

            Image imgFirmaLiliaHurtado = new Image(ImageDataFactory
              .Create(FirmaLiliaHurtado))
              .SetWidth(150)
              .SetHeight(65)
              .SetMarginBottom(0)
              .SetPadding(0)
              .SetTextAlignment(TextAlignment.CENTER);

            Image imgFirmaFirmaMilagrosMunoz = new Image(ImageDataFactory
              .Create(FirmaFirmaMilagrosMunoz))
              .SetWidth(150)
              .SetHeight(65)
              .SetMarginBottom(0)
              .SetPadding(0)
              .SetTextAlignment(TextAlignment.CENTER);

            Image imgConclusion = new Image(ImageDataFactory
            .Create(Conclusion))
            .SetWidth(560)
            .SetHeight(40)
            .SetMarginBottom(0)
            .SetPadding(0)
            .SetTextAlignment(TextAlignment.LEFT);

            Table PiePagina = new Table(2).UseAllAvailableWidth();
            PiePagina.SetFixedLayout();


            Cell cellFPiePagina = new Cell(1, 1).Add(imagenFlooterss)
            .SetFixedPosition(0f, document.GetBottomMargin() - 2, 0f)
            .SetBorder(Border.NO_BORDER);

            PiePagina.AddCell(cellFPiePagina);

            document.Add(PiePagina);

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);


            Style estiloTitulo = new Style()
                .SetFontSize(14)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-3)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);


            Style estiloCabeceraSubtituloLote = new Style()
               .SetFontSize(14)
               .SetFont(fuenteNegrita)
               .SetMarginTop(-3)
               .SetFontColor(ColorConstants.BLACK)
               .SetTextAlignment(TextAlignment.RIGHT);

            Style estiloCabeceraSubtituloCodsut = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNormal)
                .SetMarginTop(9)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.RIGHT);


            Style estiloCabeceraVariable = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegrita = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTextoDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegritaDetalleProtocolo = new Style()
                .SetFontSize(7)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);




            //CABECERA 
            Table tablaDatosTitulo = new Table(3).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellTitulo = new Cell(1, 1).Add(img)
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

            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°:" + Cabecera.REFERENCIANUMERO + " \n").AddStyle(estiloCabeceraSubtituloLote)).Add(new Paragraph(Cabecera.NUMERODEPARTE + " \n").AddStyle(estiloCabeceraSubtituloCodsut))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);


            Table tablaDatosdeCabecera = new Table(6).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout().SetPaddingBottom(3);
            /*
               .SetBorderBottomLeftRadius(new BorderRadius(0.5f))
               .SetBorderBottomRightRadius(new BorderRadius(0.5f))
               .SetBorderTopLeftRadius(new BorderRadius(0.5f))
               .SetBorderTopRightRadius(new BorderRadius(0.5f));
             */

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Product:")
           .AddStyle(estiloCabeceraVariable))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetHorizontalAlignment(HorizontalAlignment.CENTER)
           .SetVerticalAlignment(VerticalAlignment.MIDDLE)
           .SetBorderTopLeftRadius(new BorderRadius(0.5f))
           .SetBorderRight(Border.NO_BORDER)
           .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.ITEMDESCRIPCION)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTitulo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);



            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Presentation:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.Presentacion)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Date of manufacture")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAPRODUCCION.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Batch No.:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.ORDENFABRICACION)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Lot size: ")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.CANTIDADPRODUCIDA.ToString("#,##0"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);



            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Expiration date:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAEXPIRACION.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Brand:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(Cabecera.MARCA)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Date of analysis:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(Cabecera.FECHAANALISIS.ToString("dd/MM/yyyy"))
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            document.Add(tablaDatosdeCabecera);

            Table tablaDetalleProtocolo = new Table(9).UseAllAvailableWidth();
            tablaDetalleProtocolo.SetFixedLayout();

            Cell cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph("Tests carried ou ")
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

            cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph("Specifications")
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

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Results")
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
            foreach (DatosFormatoProtocoloPruebaModel item in ListadoProtocolo)
            {
                cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph(item.DESCRIPCION_PRUEBA)
               .AddStyle(estiloTextoDetalleProtocolo))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetHorizontalAlignment(HorizontalAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.TOP)
               .SetPaddingBottom(3)
               .SetBorderLeft(Border.NO_BORDER)
               .SetBorderBottom(Border.NO_BORDER)
               .SetBorderTop(Border.NO_BORDER)
               .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph(item.ESPECIFICACION)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.JUSTIFIED)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetPaddingBottom(3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.RESULTADO)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetPaddingBottom(3)
                .SetBorderLeft(Border.NO_BORDER)
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderTop(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER);
                tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.DESCRIPCION_METODOLOGIA)
                .AddStyle(estiloTextoDetalleProtocolo))
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

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.TECNICA)
            .AddStyle(estiloTextoDetalleProtocolo))
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



            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.METODO)
            .AddStyle(estiloTextoDetalleProtocolo))
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


            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Cabecera.DETALLE)
            .AddStyle(estiloTextoDetalleProtocolo))
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

            Cell cellObservarcion = new Cell(1, 1).Add(new Paragraph("OBSERVATIONS")
            .AddStyle(estiloNegrita))
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

            if (Opcion)
            {
                Table tablaFirmaResponsables = new Table(2).UseAllAvailableWidth();
                tablaFirmaResponsables.SetFixedLayout();


                Cell cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaLiliaHurtado)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPaddingLeft(50)
                .SetBorder(Border.NO_BORDER);
                tablaFirmaResponsables.AddCell(cellFirmaResponsables);

                cellFirmaResponsables = new Cell(1, 1).Add(imgFirmaFirmaMilagrosMunoz)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetPaddingLeft(50)
                .SetBorder(Border.NO_BORDER);
                tablaFirmaResponsables.AddCell(cellFirmaResponsables);

                document.Add(tablaFirmaResponsables);

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
    }
}
