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
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.Actaverifacioncc
{
    public class ActaVerificacioncc
    {

        public string fechaActual = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

        public string GenerarReporteActaVerificacion(List<CReporteGuiaRemisionModel> NumeroGuias, ListarOpcionesImprimir dato, string versionProtocolo)
        {
            //DOCUMENTO QUE TIENE TRUE 
            bool documentoActaCC = dato.Acta;
            bool documentoCondiciones = dato.Condicion;
            bool documentoProtocolo = dato.Protocolo;
            bool documentoRsanitario = dato.Carta;
            bool documentoBPA = dato.practicas;
            bool documentoManufactura = dato.Manufactura;

            string fechahoy = DateTime.Today.ToLongDateString().ToString();
            string[] separarfecha = fechahoy.Split(',');
            string fechaFinal = "Lima," + separarfecha[1];

            int contador = 0;
            int contadoArray = NumeroGuias.Count;
            string reporte = null;


            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Licitaciones");
            docInfo.SetAuthor("Sistema Licitaciones");

            Document document = new Document(pdf, documentoActaCC ? PageSize.A4.Rotate() : PageSize.A4);
            document.SetMargins(5, 15, 30, 15);

            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            string Conclusion = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\ConclusionProtocolo.png");
            string FirmaLiliaHurtado = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLiliaHurtadoDias.jpg");
            string FirmaFirmaMilagrosMunoz = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaMilagrosMunozTafur.jpg");

            string CertificadoBPA = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\CERTIFICADOBPA1310-21.jpg");
            string ManufacturaP1 = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\CERTIFICADOBPMN0762021P1.jpg");
            string ManufacturaP2 = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\CERTIFICADOBPMN0762021P2.jpg");


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

            Image imgCertificadoBPA = new Image(ImageDataFactory
            .Create(CertificadoBPA))
            .ScaleAbsolute(550f, 850f)
            .SetMarginBottom(0)
            .SetPadding(0)
            .SetTextAlignment(TextAlignment.RIGHT);

            Image imgManufacturaP1 = new Image(ImageDataFactory
            .Create(ManufacturaP1))
            .ScaleAbsolute(550f, 850f)
            .SetMarginBottom(0)
            .SetPadding(0)
            .SetTextAlignment(TextAlignment.RIGHT);

            Image imgManufacturaP2 = new Image(ImageDataFactory
            .Create(ManufacturaP2))
            .ScaleAbsolute(550f, 850f)
            .SetMarginBottom(0)
            .SetPadding(0)
            .SetTextAlignment(TextAlignment.RIGHT);

            foreach (CReporteGuiaRemisionModel cabecera in NumeroGuias)
            {
                string Entrega = "";


                switch (cabecera.NumeroEntrega)
                {
                    case "1":
                        Entrega = "1ERA ENTREGA";
                        break;
                    case "2":
                        Entrega = "2DA ENTREGA";
                        break;
                    case "3":
                        Entrega = "3ERA ENTREGA";
                        break;
                    case "4":
                        Entrega = "4TA ENTREGA";
                        break;
                    case "5":
                        Entrega = "5TA ENTREGA";
                        break;
                    case "6":
                        Entrega = "6TA ENTREGA";
                        break;
                    case "7":
                        Entrega = "7MA ENTREGA";
                        break;
                    case "8":
                        Entrega = "8VA ENTREGA";
                        break;
                    case "9":
                        Entrega = "9NA ENTREGA";
                        break;
                    case "10":
                        Entrega = "10MA ENTREGA";
                        break;
                    case "11":
                        Entrega = "11VA ENTREGA";
                        break;
                    case "12":
                        Entrega = "12VA ENTREGA";
                        break;
                    default:
                        Entrega = "NO HAY ENTREGA";
                        break;
                }

                if (documentoActaCC)
                    GenerarPdf(document, cabecera, Entrega, fechaFinal);

                if (documentoCondiciones)
                {
                    if (documentoActaCC)
                        document.Add(new AreaBreak(PageSize.A4));

                    Compromiso(document, cabecera, rutaUnilene, fechaFinal);
                    document.Add(new AreaBreak(PageSize.A4));
                    Condiciones(document, cabecera, rutaUnilene, fechaFinal);
                }

                if (documentoRsanitario)
                {
                    if (documentoActaCC || documentoCondiciones)
                        document.Add(new AreaBreak(PageSize.A4));

                    Carta(document, cabecera, rutaUnilene, fechaFinal);
                }


                if (cabecera.DetalleGuia[0].DetalleProtocolo.Count > 0)
                    if (documentoProtocolo)
                    {
                        if (documentoActaCC || documentoCondiciones || documentoRsanitario)
                            document.Add(new AreaBreak(PageSize.A4));

                        Protocolo(document, cabecera, rutaUnilene, fechaFinal, Conclusion, imgFirmaLiliaHurtado, imgFirmaFirmaMilagrosMunoz);
                    }

                if (documentoBPA)
                {
                    if (documentoActaCC || documentoCondiciones || documentoRsanitario)
                        document.Add(new AreaBreak(PageSize.A4));

                    buenaspracticasalmacenamiento(document, imgCertificadoBPA);
                }

                if (documentoManufactura)
                {
                    if (documentoActaCC || documentoCondiciones || documentoRsanitario)
                        document.Add(new AreaBreak(PageSize.A4));

                    DocumentoManufactura(document, imgManufacturaP1, imgManufacturaP2);
                }


                if (contador < contadoArray - 1)
                {
                    document.Add(new AreaBreak());
                }


                contador++;

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

        public void GenerarPdf(Document document, CReporteGuiaRemisionModel cabecera, string Entrega, string fechaFinal)
        {
            Color bgColour = new DeviceRgb(161, 205, 241);
            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            document.Add(img);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));

            Style estiloTitulo = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-8)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);

            Style estiloCabecera = new Style()
                .SetFontSize(6)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloDetalle = new Style()
                .SetFontSize(6)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(6)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloFechaVerificacion = new Style()
                .SetFontSize(6)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Paragraph titulo1 = new Paragraph("ACTA DE VERIFICACION CUALITATIVA Y CUANTITATIVA").AddStyle(estiloTitulo);
            Paragraph titulo2 = new Paragraph(cabecera.DescripcionProceso).AddStyle(estiloTitulo);
            document.Add(titulo1);
            document.Add(titulo2);

            Table tablaDatosGenerales = new Table(10).UseAllAvailableWidth();
            tablaDatosGenerales.SetFixedLayout().SetFontSize(9);

            Cell cellDG = new Cell(1, 7).Add(new Paragraph("")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloCabecera))
            .SetFont(fuenteNegrita)
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 2).Add(new Paragraph(fechaFinal)
             .AddStyle(estiloCabecera))
             .SetFont(fuenteNegrita)
             .SetTextAlignment(TextAlignment.RIGHT)
             .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("CONTRATISTA ")
             .AddStyle(estiloCabecera))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(new SolidBorder(1))
             .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph("UNILENE S.A.C")
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);


            cellDG = new Cell(1, 1).Add(new Paragraph("TIPO DE ADJUDICACIÓN ")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(new SolidBorder(1))
              .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph(cabecera.DescripcionProceso + " - " + cabecera.DescripcionComercialDetalle)
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("ORDEN DE COMPRA N° ")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(new SolidBorder(1))
              .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph(cabecera.OrdenCompra + "- Pecosa N°: " + cabecera.Pecosa)
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("Contrato N°")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(new SolidBorder(1))
              .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph(cabecera.Contrato)
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("Entrega N°")
           .AddStyle(estiloCabecera))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetBorder(new SolidBorder(1))
           .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph(Entrega)
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph("Usuario")
           .AddStyle(estiloCabecera))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetBorder(new SolidBorder(1))
           .SetBackgroundColor(bgColour);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 9).Add(new Paragraph(cabecera.ClienteNombre + " EN SALUD - " + cabecera.Region)
            .AddStyle(estiloCabecera))
            .SetTextAlignment(TextAlignment.LEFT);

            tablaDatosGenerales.AddCell(cellDG);

            document.Add(tablaDatosGenerales);


            //Tabla 2 
            Table tablaDatosParrafo = new Table(10).UseAllAvailableWidth();
            tablaDatosParrafo.SetFixedLayout().SetFontSize(10);

            Cell cellPresentacion = new Cell(1, 10).Add(new Paragraph("En la fecha los Representantes del Almacen y EL CONTRATISTA proceden a dar conformidad a los siguientes productos correspondiente a la orden de Compra referida :")
              .AddStyle(estiloTexto))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);
            tablaDatosParrafo.AddCell(cellPresentacion);
            document.Add(tablaDatosParrafo);

            //Detalle Item 

            Table tablaDatosDetalle = new Table(30).UseAllAvailableWidth();
            tablaDatosDetalle.SetFixedLayout().SetFontSize(9);

            //bloque 1
            Cell cellDetalle = new Cell(2, 1).Add(new Paragraph("ITEM")
              .AddStyle(estiloDetalle))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 5).Add(new Paragraph("NOMBRE DEL PRODUCTO(DCI)")
             .AddStyle(estiloDetalle))
             .SetTextAlignment(TextAlignment.CENTER)
             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
             .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 1).Add(new Paragraph("UNIDAD DE MEDIDA")
             .AddStyle(estiloDetalle).SetFontSize(5))
             .SetTextAlignment(TextAlignment.CENTER)
             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
             .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 4).Add(new Paragraph("PRESENTACION")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            cellDetalle = new Cell(2, 2).Add(new Paragraph("CANT. SOLICITADA")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            cellDetalle = new Cell(2, 2).Add(new Paragraph("CANT. RECEPCIONADA")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(2, 2).Add(new Paragraph("GUIA REMISION")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            cellDetalle = new Cell(1, 4).Add(new Paragraph("LOTE")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            cellDetalle = new Cell(2, 3).Add(new Paragraph("REGISTRO SANITARIO")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            cellDetalle = new Cell(2, 2).Add(new Paragraph("N° DE PROTOCOLO DE ANALISIS")
            .AddStyle(estiloCabecera))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 4).Add(new Paragraph("LABORATORIO DE CONTROL DE CALIDAD")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);


            //bloque 2 


            cellDetalle = new Cell(1, 2).Add(new Paragraph("N°")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 2).Add(new Paragraph("F.V")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);



            cellDetalle = new Cell(1, 2).Add(new Paragraph("N° DE ACTA DE MUESTREO")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);

            tablaDatosDetalle.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 2).Add(new Paragraph("N° INFORME DE ENSAYO")
            .AddStyle(estiloDetalle))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour);
            tablaDatosDetalle.AddCell(cellDetalle);



            //FOREARCH
            foreach (DReportGuiaRemisionModel detalle in cabecera.DetalleGuia)
            {
                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.NumeroItem.ToString())
                  .AddStyle(estiloDetalle))
                  .SetTextAlignment(TextAlignment.CENTER)
                  .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 5).Add(new Paragraph(detalle.Descripcion)
                 .AddStyle(estiloDetalle))
                 .SetTextAlignment(TextAlignment.CENTER)
                 .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 1).Add(new Paragraph(detalle.UnidadCodigo)
                .AddStyle(estiloDetalle))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 4).Add(new Paragraph(detalle.CaractervaluesDescripcion)
                .AddStyle(estiloDetalle))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.CantidadGRD.ToString())
                   .AddStyle(estiloDetalle))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.CantidadGRD.ToString())
                   .AddStyle(estiloDetalle))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.Guia)
                   .AddStyle(estiloDetalle))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.Lote)
                 .AddStyle(estiloDetalle))
                 .SetTextAlignment(TextAlignment.CENTER)
                 .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.FechaExpiracion.ToString("dd'/'MM'/'yyyy"))
                .AddStyle(estiloDetalle).SetFontSize(6))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 3).Add(new Paragraph(detalle.RegistroSanitario)
                .AddStyle(estiloDetalle))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.Protocolo)
                .AddStyle(estiloDetalle))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.NumeroMuestreo == "" || detalle.NumeroMuestreo == "-" ? "NO REQUIERE" : detalle.NumeroMuestreo)
                   .AddStyle(estiloDetalle))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.NumeroEnsayo == "" || detalle.NumeroEnsayo == "-" ? "NO REQUIERE" : detalle.NumeroEnsayo)
                 .AddStyle(estiloDetalle))
                 .SetTextAlignment(TextAlignment.CENTER)
                 .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

            }


            document.Add(tablaDatosDetalle);

            //Tabla 4 
            Table tablaDatosFecha = new Table(1).UseAllAvailableWidth();
            tablaDatosFecha.SetFixedLayout().SetFontSize(9);

            Cell cellFecha = new Cell(1, 1).Add(new Paragraph("la verificación de los productos en almacen se realizo el Dia ...........  Mes ...........  Año ............")
                  .AddStyle(estiloFechaVerificacion))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(new Paragraph("")
             .AddStyle(estiloFechaVerificacion))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")
             .AddStyle(estiloFechaVerificacion))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloFechaVerificacion))
             .SetHeight(3)
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(new SolidBorder(1));
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph("")
                 .AddStyle(estiloFechaVerificacion))
                  .SetHeight(3)
                 .SetTextAlignment(TextAlignment.LEFT)
                 .SetBorder(new SolidBorder(1));
            tablaDatosFecha.AddCell(cellFecha);

            document.Add(tablaDatosFecha);
            document.Add(saltoLinea);
            //Tabla Finalizada
            Table tablaDatosConforme = new Table(1).UseAllAvailableWidth();
            tablaDatosConforme.SetFixedLayout().SetFontSize(9);

            Cell cellConforme = new Cell(1, 1).Add(new Paragraph("FINALIZADA LA VERIFICACIÓN DE LOS PRODUCTOS Y ESTADO CONFORME , SE PROCEDE A LA SUSCRIPCIÓN DE LA PRESENTE ACTA ")
                  .AddStyle(estiloFechaVerificacion))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosConforme.AddCell(cellConforme);
            document.Add(tablaDatosConforme);

            //FIRMA
            Table tablaDatosFirma = new Table(3).UseAllAvailableWidth();
            tablaDatosFirma.SetFixedLayout().SetFontSize(9);

            Cell cellFirma = new Cell(1, 1).Add(new Paragraph("")
                .AddStyle(estiloFechaVerificacion))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHeight(8)
                .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph("")
               .AddStyle(estiloFechaVerificacion))
               .SetTextAlignment(TextAlignment.CENTER)
               .SetHeight(8)
               .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph("")
              .AddStyle(estiloFechaVerificacion))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetHeight(8)
              .SetBorder(Border.NO_BORDER);

            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph(""))
                         .SetTextAlignment(TextAlignment.CENTER)
                         .SetBorder(Border.NO_BORDER)
                         .SetPaddingLeft(20);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph(""))
                         .SetTextAlignment(TextAlignment.CENTER)
                         .SetBorder(Border.NO_BORDER)
                         .SetPaddingLeft(20);
            tablaDatosFirma.AddCell(cellFirma);

            string rutaUnilene2 = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Sello_RepLegal_ErickHartmann.png");

            Image img2 = new Image(ImageDataFactory
               .Create(rutaUnilene2))
               .SetWidth(63)
               .SetHeight(60)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Table tableFirma = new Table(new float[] { 1, 2 }).UseAllAvailableWidth();
            tableFirma.SetWidth(UnitValue.CreatePercentValue(100));
            tableFirma.SetFixedLayout();

            Cell celdaFirma = new Cell(1, 1).Add(img2)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER);

            celdaFirma = new Cell(1, 1).Add(new Paragraph("Firmado digitalmente por:\n HARTMANN BUSTAMANTE ERICK - 20197705249\n " +
                "Motivo: ACTA DE VERIFICACION C.C. \n Fecha: " + fechaActual)
                        .SetFontSize(7))
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER).SetMarginLeft(30).SetMarginRight(8).SetMarginTop(8);

            cellFirma = new Cell(1, 1).Add(tableFirma).SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);


            cellFirma = new Cell(1, 1).Add(new Paragraph("_______________________________________________________________")
                .AddStyle(estiloFechaVerificacion))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph("_______________________________________________________________")
                .AddStyle(estiloFechaVerificacion))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);


            cellFirma = new Cell(1, 1).Add(new Paragraph("_______________________________________________________________")
               .AddStyle(estiloFechaVerificacion))
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph("Firma y Sello de Representante")
             .AddStyle(estiloFechaVerificacion))
             .SetTextAlignment(TextAlignment.CENTER)
             .SetVerticalAlignment(VerticalAlignment.TOP)
             .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            cellFirma = new Cell(1, 1).Add(new Paragraph("Firma y Sello del Representante ALMACEN")
             .AddStyle(estiloFechaVerificacion))
             .SetTextAlignment(TextAlignment.CENTER)
             .SetVerticalAlignment(VerticalAlignment.TOP)
             .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);



            cellFirma = new Cell(1, 1).Add(new Paragraph("Firma y Sello del Representante PROVEEDOR")
            .AddStyle(estiloFechaVerificacion))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.TOP)
            .SetBorder(Border.NO_BORDER);
            tablaDatosFirma.AddCell(cellFirma);

            document.Add(tablaDatosFirma);
            document.Add(saltoLinea);

            Table tablaDatosinformacion = new Table(1).UseAllAvailableWidth();
            tablaDatosinformacion.SetFixedLayout().SetFontSize(9);

            Cell cellInformacion = new Cell(1, 1).Add(new Paragraph("Unilene S.A.C \n Jr Napo 450, Lima 05 - Peru \n Phone: 997509088 \n www.unilene.com \n info@unilene.com")
                  .AddStyle(estiloTexto))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosinformacion.AddCell(cellInformacion);

            document.Add(tablaDatosinformacion);

        }

        public void Compromiso(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
        {


            Color bgColour = new DeviceRgb(161, 205, 241);

            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
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
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-8)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);

            Style estiloCabecera = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estilotextoNegrita = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloFechaVerificacion = new Style()
                .SetFontSize(6)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Paragraph titulo1 = new Paragraph("ANEXO N°9").AddStyle(estiloTitulo).SetMarginTop(10).SetMarginBottom(7);
            Paragraph titulo2 = new Paragraph("DECLARACIÓN JURADA DE COMPROMISO DE CANJE Y/O REPOSICIÓN POR VICIOS OCULTOS").AddStyle(estiloTitulo);
            Paragraph titulo3 = new Paragraph(cabecera.DescripcionProceso).AddStyle(estiloTitulo);

            document.Add(titulo1);
            document.Add(titulo2);
            document.Add(titulo3);
            document.Add(saltoLinea);

            Table tablaDatosGenerales = new Table(1).UseAllAvailableWidth();
            tablaDatosGenerales.SetFixedLayout();

            Cell cellDG = new Cell(1, 1).Add(new Paragraph("Señores: ")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.ClienteNombre + " EN SALUD - " + cabecera.Region)
             .AddStyle(estiloCabecera))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(Border.NO_BORDER);
            tablaDatosGenerales.AddCell(cellDG);

            document.Add(tablaDatosGenerales);

            //Tabla 2 
            Table tablaDatosParrafo = new Table(1).UseAllAvailableWidth();
            tablaDatosParrafo.SetFixedLayout();

            Cell cellPresente = new Cell(1, 1).Add(new Paragraph("Presente")
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            tablaDatosParrafo.AddCell(cellPresente);

            cellPresente = new Cell(1, 1).Add(new Paragraph("De Nuestra consideración: ")
               .AddStyle(estiloTexto))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBorder(Border.NO_BORDER)
               .SetPaddingBottom(20);
            tablaDatosParrafo.AddCell(cellPresente);

            document.Add(tablaDatosParrafo);



            Paragraph LictacionPublica = new Paragraph(cabecera.DescripcionProceso + " - " + cabecera.DescripcionComercialDetalle + " (" + cabecera.CantItems.ToString() + "-ITEMS)-ITEM " + cabecera.DetalleGuia[0].NumeroItem.ToString() + ":" + cabecera.DetalleGuia[0].Descripcion + " . Es perteneciente a la OC " + cabecera.OrdenCompra).AddStyle(estilotextoNegrita);

            //tabla 3
            Table tablaDatosContenido = new Table(1).UseAllAvailableWidth();
            tablaDatosContenido.SetFixedLayout();

            Cell cellContenido = new Cell(1, 1).Add(new Paragraph("Nos es grato hacer llegar a usted, la presente “Declaración  Jurada de Compromiso de Canje y/o Reposición” en representación de  UNILENE S.A.C, por los productos que se nos adjudican en nuestra propuesta presentada al procedimiento de selección ").Add(LictacionPublica)
              .AddStyle(estiloTexto))
              .SetTextAlignment(TextAlignment.JUSTIFIED)
              .SetBorder(Border.NO_BORDER)
               .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("El canje será efectuado en el caso de que el producto haya sufrido alteración de sus características físicas-químicas  sin causa atribuible al usuario o cualquier otro defecto o vicio oculto antes de su fecha de expiración. El producto canjeado tendrá fecha de expiración igual o mayor a la ofertada en el proceso de selección, contada a partir de la fecha de entrega de canje.")
             .AddStyle(estiloTexto))
             .SetTextAlignment(TextAlignment.JUSTIFIED)
             .SetBorder(Border.NO_BORDER)
              .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("El canje se efectuará a solo requerimiento de ustedes, en un plazo no mayor a 60 días calendarios y no generará gastos adicionales a los pactados con vuestra entidad.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            document.Add(tablaDatosContenido);
            document.Add(saltoLinea);

            //tabla 4
            Table tablaDatosFecha = new Table(2).UseAllAvailableWidth();
            tablaDatosFecha.SetFixedLayout();

            Cell cellFecha = new Cell(1, 1).Add(new Paragraph("Atentamente,")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
             .SetPaddingBottom(20);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
             .SetPaddingBottom(20);
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(new Paragraph(fechaFinal)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetMarginTop(10);
            tablaDatosFecha.AddCell(cellFecha);

            //eddie

            //string FirmaLicitaciones = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLicitaciones.png");

            //Image img2 = new Image(ImageDataFactory
            //   .Create(FirmaLicitaciones))
            //   .SetWidth(135)
            //   .SetHeight(50)
            //   .SetMarginBottom(0)
            //   .SetPadding(0)
            //   .SetBorder(Border.NO_BORDER)
            //   .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            //cellFecha = new Cell(1, 1).Add(img2)
            //.SetBorder(Border.NO_BORDER)
            //.SetMarginTop(100);
            //tablaDatosFecha.AddCell(cellFecha);

            // ----------------------------

            string rutaUnilene2 = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Sello_RepLegal_ErickHartmann.png");

            Image img2 = new Image(ImageDataFactory
               .Create(rutaUnilene2))
               .SetWidth(63)
               .SetHeight(60)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Table tableFirma = new Table(new float[] { 1, 2 }).UseAllAvailableWidth();
            tableFirma.SetWidth(UnitValue.CreatePercentValue(100));
            tableFirma.SetFixedLayout();

            Cell celdaFirma = new Cell(1, 1).Add(img2)
                        .SetTextAlignment(TextAlignment.RIGHT)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER);

            celdaFirma = new Cell(1, 1).Add(new Paragraph("Firmado digitalmente por:\n HARTMANN BUSTAMANTE ERICK - 20197705249 \n Motivo: DECLARACIÓN JURADA\n Fecha: " + fechaActual).SetFontSize(7))
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER).SetMarginLeft(30).SetMarginRight(8).SetMarginTop(8);

            cellFecha = new Cell(1, 1).Add(tableFirma).SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);


            // ----------------------------

            document.Add(tablaDatosFecha);
            document.Add(saltoLinea);

            Table tablaDatosinformacion = new Table(1).UseAllAvailableWidth();
            tablaDatosinformacion.SetFixedLayout().SetFontSize(9);

            Cell cellInformacion = new Cell(1, 1).Add(new Paragraph("Unilene S.A.C \n Jr Napo 450, Lima 05 - Peru \n Phone: 997509088 \n www.unilene.com \n info@unilene.com")
                  .AddStyle(estiloTexto))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosinformacion.AddCell(cellInformacion);

            document.Add(tablaDatosinformacion);
            document.Add(saltoLinea);

        }

        public void Condiciones(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
        {
            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            document.Add(img);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));

            Style estiloTitulo = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-3)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);

            Style estiloCabecera = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegrita = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFontColor(ColorConstants.BLACK);


            Paragraph titulo1 = new Paragraph("DECLARACIÓN JURADA DE CONDICIONES ESPECIALES DE EMBALAJE").AddStyle(estiloTitulo).SetMarginTop(10);
            Paragraph titulo2 = new Paragraph(cabecera.DescripcionProceso).AddStyle(estiloTitulo);

            document.Add(saltoLinea);
            document.Add(titulo1);
            document.Add(titulo2);


            Table tablaDatosGenerales = new Table(1).UseAllAvailableWidth();
            tablaDatosGenerales.SetFixedLayout();

            Cell cellDG = new Cell(1, 1).Add(new Paragraph("Señores:")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.ClienteNombre + " EN SALUD - " + cabecera.Region)
             .AddStyle(estiloCabecera))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(Border.NO_BORDER);
            tablaDatosGenerales.AddCell(cellDG);

            document.Add(tablaDatosGenerales);


            //Tabla 2 
            Table tablaDeclaracionJuramento = new Table(1).UseAllAvailableWidth();
            tablaDeclaracionJuramento.SetFixedLayout();

            Paragraph Nombre = new Paragraph("ERICK HARTMANN BUSTAMANTE ").AddStyle(estiloNegrita);
            Paragraph ruc = new Paragraph("20197705249, DECLARO BAJO JURAMENTO ").AddStyle(estiloNegrita);


            Cell cellPresente = new Cell(1, 1).Add(new Paragraph("Presente")
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            tablaDeclaracionJuramento.AddCell(cellPresente);


            cellPresente = new Cell(1, 1).Add(new Paragraph("El que se suscribe, ").Add(Nombre).Add("identificado con DNI Nº 44656730, Representante Legal de UNILENE S.A.C , con RUC Nº ").Add(ruc).Add("la información que a continuación se detalla respecto a las condiciones especiales de embalaje de la: ")
                 .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            tablaDeclaracionJuramento.AddCell(cellPresente);

            cellPresente = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].Descripcion)
               .AddStyle(estiloNegrita))
               .SetTextAlignment(TextAlignment.CENTER)
               .SetBorder(Border.NO_BORDER)
               .SetPadding(20);
            tablaDeclaracionJuramento.AddCell(cellPresente);

            document.Add(tablaDeclaracionJuramento);

            //tabla 3
            Table tablaDatosContenido = new Table(1).UseAllAvailableWidth();
            tablaDatosContenido.SetFixedLayout();

            Cell cellContenido = new Cell(1, 1).Add(new Paragraph("Condiciones Especiales de almacenamiento, embalaje y distribución:")
              .AddStyle(estiloTexto))
              .SetTextAlignment(TextAlignment.JUSTIFIED)
              .SetBorder(Border.NO_BORDER);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("El embalaje de los dispositivos médicos deberá cumplir con los siguientes requisitos:")
             .AddStyle(estiloTexto))
             .SetTextAlignment(TextAlignment.JUSTIFIED)
             .SetBorder(Border.NO_BORDER);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("1. Cajas de cartón nuevas y resistentes que garanticen la integridad, orden, conservación, transporte y adecuado almacenamiento.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("2. Cajas que faciliten su conteo y fácil apilamiento, precisando  el número de cajas apilables.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("3. Cajas debidamente rotuladas indicando nombre del dispositivo médico, cantidad, lote, fecha de vencimiento (según corresponda),  nombre del proveedor, especificaciones para la conservación y almacenamiento.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("4. Dicha información deberá ser indicada en etiquetas. Aplica a caja master, es decir a caja completa del dispositivo médico.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("5. Si corresponde, en las caras laterales debe decir “FRAGIL”, con letras de un tamaño mínimo de 5 cm de alto y en tipo negrita e indicar  con una flecha el sentido correcto para la posición de la caja. ")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("6. Debe descartarse la utilización de cajas de productos comestibles o productos de tocador, entre otros.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("7. Para las dimensiones de la caja de embalaje debe considerarse la paleta (parihuela) estándar definida según NTP vigente. ")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].Temperatura == "" ? "" : "8. " + cabecera.DetalleGuia[0].Temperatura + ". ")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPadding(5);
            tablaDatosContenido.AddCell(cellContenido);


            document.Add(tablaDatosContenido);
            document.Add(saltoLinea);

            Table tablaDatosFecha = new Table(2).UseAllAvailableWidth();
            tablaDatosFecha.SetFixedLayout();

            Cell cellFecha = new Cell(1, 1).Add(new Paragraph("Atentamente,")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
             .SetPaddingBottom(10);
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(new Paragraph("")
              .AddStyle(estiloTexto))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph(fechaFinal)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            string rutaSelloErick = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Sello_RepLegal_ErickHartmann.png");

            Image img2 = new Image(ImageDataFactory
               .Create(rutaSelloErick))
               .SetWidth(63)
               .SetHeight(60)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Table tableFirma = new Table(new float[] { 1, 2 }).UseAllAvailableWidth();
            tableFirma.SetWidth(UnitValue.CreatePercentValue(100));
            tableFirma.SetFixedLayout();

            Cell celdaFirma = new Cell(1, 1).Add(img2)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER);

            celdaFirma = new Cell(1, 1).Add(new Paragraph("Firmado digitalmente por:\n HARTMANN BUSTAMANTE ERICK - 20197705249 \n Motivo: DECLARACIÓN JURADA\n Fecha: " + fechaActual).SetFontSize(7))
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER).SetMarginLeft(30).SetMarginRight(8).SetMarginTop(8);

            cellFecha = new Cell(1, 1).Add(tableFirma).SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            // ----------------------------


            document.Add(tablaDatosFecha);
            document.Add(saltoLinea);

            Table tablaDatosinformacion = new Table(1).UseAllAvailableWidth();
            tablaDatosinformacion.SetFixedLayout().SetFontSize(9);

            Cell cellInformacion = new Cell(1, 1).Add(new Paragraph("Unilene S.A.C \n Jr Napo 450, Lima 05 - Peru \n Phone: 997509088 \n www.unilene.com \n info@unilene.com")
                  .AddStyle(estiloTexto))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosinformacion.AddCell(cellInformacion);

            document.Add(tablaDatosinformacion);
            document.Add(saltoLinea);



        }

        public void Carta(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
        {
            Color bgColour = new DeviceRgb(161, 205, 241);


            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            document.Add(img);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

            Style estiloTitulo = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetMarginTop(-3)
                .SetFontColor(ColorConstants.BLACK)
                .SetTextAlignment(TextAlignment.CENTER);

            Style estiloCabecera = new Style()
                .SetFontSize(10)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloNegrita = new Style()
                .SetFontSize(9)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloTexto = new Style()
                .SetFontSize(9)
                .SetFontColor(ColorConstants.BLACK);

            Paragraph titulo1 = new Paragraph("CARTA DE GARANTÍA DE CALIDAD DE LOS PRODUCTOS OFERTADOS").AddStyle(estiloTitulo).SetMarginTop(10);
            Paragraph titulo2 = new Paragraph(cabecera.DescripcionProceso).AddStyle(estiloTitulo);

            document.Add(saltoLinea);
            document.Add(titulo1);
            document.Add(titulo2);


            Table tablaDatosGenerales = new Table(1).UseAllAvailableWidth();
            tablaDatosGenerales.SetFixedLayout();

            Cell cellDG = new Cell(1, 1).Add(new Paragraph("Señores:")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.Region)
             .AddStyle(estiloCabecera))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetBorder(Border.NO_BORDER);
            tablaDatosGenerales.AddCell(cellDG);

            document.Add(tablaDatosGenerales);

            Table tablaDatosParrafo = new Table(1).UseAllAvailableWidth();
            tablaDatosParrafo.SetFixedLayout();

            Cell cellPresente = new Cell(1, 1).Add(new Paragraph("Presente")
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            tablaDatosParrafo.AddCell(cellPresente);

            cellPresente = new Cell(1, 1).Add(new Paragraph("De Nuestra consideración: ")
               .AddStyle(estiloTexto))
               .SetTextAlignment(TextAlignment.LEFT)
               .SetBorder(Border.NO_BORDER)
               .SetPaddingBottom(20);
            tablaDatosParrafo.AddCell(cellPresente);

            document.Add(tablaDatosParrafo);

            Table tablaDatosContenido = new Table(1).UseAllAvailableWidth();
            tablaDatosContenido.SetFixedLayout();

            Cell cellContenido = new Cell(1, 1).Add(new Paragraph("Tenemos el agrado de dirigirnos a ustedes para saludarlos muy cordialmente y a la vez informarles que nuestra empresa ").Add(new Paragraph("UNILENE S.A.C").AddStyle(estiloNegrita)).Add(new Paragraph(", identificada con RUC Nº 20197705249, con domicilio en Jr. Napo Nº 450 Breña,").Add(new Paragraph("GARANTIZA").AddStyle(estiloNegrita)).Add(" que los productos que internamos son de primera calidad tal como lo ACREDITAN nuestros certificados de Buenas Prácticas de Manufactura BPM Nº 076-2021 y de Buenas Prácticas de Almacenamiento BPA Nº 1310-2021."))
              .AddStyle(estiloTexto))
              .SetTextAlignment(TextAlignment.JUSTIFIED)
              .SetBorder(Border.NO_BORDER)
               .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("Asimismo, informamos que las condiciones de embalaje son en cajas de cartón y que el transporte es de mayor seguridad, cualquier reclamo de los productos entregados correspondiente a orden de compra ").Add(new Paragraph("Nº " + cabecera.OrdenCompra).AddStyle(estiloNegrita)).Add(" serán cambiados dentro de las (24) veinticuatro horas siguientes")
             .AddStyle(estiloTexto))
             .SetTextAlignment(TextAlignment.JUSTIFIED)
             .SetBorder(Border.NO_BORDER)
              .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("La garantía de nuestros productos es de 4 años.")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            cellContenido = new Cell(1, 1).Add(new Paragraph("Sin otro particular y agradeciendo la atención que se sirva brindar a la presente, quedamos de ustedes")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetBorder(Border.NO_BORDER)
            .SetPaddingBottom(20);
            tablaDatosContenido.AddCell(cellContenido);

            document.Add(tablaDatosContenido);
            document.Add(saltoLinea);

            //tabla 4
            Table tablaDatosFecha = new Table(2).UseAllAvailableWidth();
            tablaDatosFecha.SetFixedLayout();

            Cell cellFecha = new Cell(1, 1).Add(new Paragraph("Atentamente,")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
             .SetPaddingBottom(20);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph("")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
             .SetPaddingBottom(20);
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(new Paragraph(fechaFinal)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetMarginTop(10);
            tablaDatosFecha.AddCell(cellFecha);




            string rutaSelloErick = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Sello_RepLegal_ErickHartmann.png");

            Image img2 = new Image(ImageDataFactory
               .Create(rutaSelloErick))
               .SetWidth(63)
               .SetHeight(60)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Table tableFirma = new Table(new float[] { 1, 2 }).UseAllAvailableWidth();
            tableFirma.SetWidth(UnitValue.CreatePercentValue(100));
            tableFirma.SetFixedLayout();

            Cell celdaFirma = new Cell(1, 1).Add(img2)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER);

            celdaFirma = new Cell(1, 1).Add(new Paragraph("Firmado digitalmente por:\n HARTMANN BUSTAMANTE ERICK - 20197705249 \n Motivo: DECLARACIÓN JURADA\n Fecha: " + fechaActual).SetFontSize(7))
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetBorder(Border.NO_BORDER);

            tableFirma.AddCell(celdaFirma).SetBorder(Border.NO_BORDER).SetMarginLeft(30).SetMarginRight(8).SetMarginTop(8);

            cellFecha = new Cell(1, 1).Add(tableFirma).SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            document.Add(tablaDatosFecha);
            document.Add(saltoLinea);

            Table tablaDatosinformacion = new Table(1).UseAllAvailableWidth();
            tablaDatosinformacion.SetFixedLayout().SetFontSize(9);

            Cell cellInformacion = new Cell(1, 1).Add(new Paragraph("Unilene S.A.C \n Jr Napo 450, Lima 05 - Peru \n Phone: 997509088 \n www.unilene.com \n info@unilene.com")
                  .AddStyle(estiloTexto))
                  .SetTextAlignment(TextAlignment.LEFT)
                  .SetBorder(Border.NO_BORDER);
            tablaDatosinformacion.AddCell(cellInformacion);

            document.Add(tablaDatosinformacion);
            document.Add(saltoLinea);


        }

        public void Protocolo(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal, string Conclusion, Image imgFirmaLiliaHurtado, Image imgFirmaFirmaMilagrosMunoz)
        {

            Color bgColorfondo = new DeviceRgb(217, 217, 217);

             string imagenFlooter = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\protocoloVersion12.png");
             Image imagenFlooterss = new Image(ImageDataFactory
              .Create(imagenFlooter))
              .SetWidth(725)
              .SetHeight(15)
              .SetMarginBottom(0)
              .SetPadding(0)
              .ScaleAbsolute(50f, 50f)
              .SetTextAlignment(TextAlignment.CENTER);
            
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


            Image imgConclusion = new Image(ImageDataFactory
               .Create(Conclusion))
               .SetWidth(560)
               .SetHeight(40)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);


            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

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

            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°:" + cabecera.DetalleGuia[0].DetalleProtocolo[0].NumeroLote + " \n").AddStyle(estiloCabeceraSubtituloLote)).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].ItemNumeroParte + " \n").AddStyle(estiloCabeceraSubtituloCodsut))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);

            Table tablaDatosdeCabecera = new Table(6).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout().SetPaddingBottom(3);

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Producto:")
           .AddStyle(estiloCabeceraVariable))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetHorizontalAlignment(HorizontalAlignment.CENTER)
           .SetVerticalAlignment(VerticalAlignment.MIDDLE)
           .SetBorderRight(Border.NO_BORDER)
           .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].DescripcionLocal)
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

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].Presentacion)
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].FechaFabricacion.ToString("dd'/'MM'/'yyyy"))
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].NumeroLote)
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
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].TamanoLote.ToString("N"))
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].FechaExpira.ToString("dd'/'MM'/'yyyy"))
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

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].Marca)
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].FechaAnalisis.ToString("dd'/'MM'/'yyyy"))
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
            foreach (FormatoReporteProtocoloModel item in cabecera.DetalleGuia[0].DetalleProtocolo)
            {
                cellDetalleProtocolo = new Cell(1, 3).Add(new Paragraph(item.Prueba)
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

                cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph(item.Especificacion)
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

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.Resultado)
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

                cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph(item.Metodologia)
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

            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].Leyenda)
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



            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].MetodosEsterilizacion)
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


            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(cabecera.DetalleGuia[0].DetalleProtocolo[0].NormasISO)
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

        public void buenaspracticasalmacenamiento(Document document, Image imgCertificadoBPA)
        {

            Table tablaCertificadoBPA = new Table(1).UseAllAvailableWidth();
            tablaCertificadoBPA.SetFixedLayout();

            Cell cellCertificadoBPA = new Cell(1, 1).Add(imgCertificadoBPA)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaCertificadoBPA.AddCell(cellCertificadoBPA);

            document.Add(cellCertificadoBPA);

        }

        public void DocumentoManufactura(Document document, Image imgManufacturaP1, Image imgManufacturaP2)
        {
            Table tablaManufacturaP1 = new Table(1).UseAllAvailableWidth();
            tablaManufacturaP1.SetFixedLayout();
            Cell cellManufacturaP1 = new Cell(1, 1).Add(imgManufacturaP1)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaManufacturaP1.AddCell(cellManufacturaP1);

            document.Add(tablaManufacturaP1);

            document.Add(new AreaBreak(PageSize.A4));

            Table tablaManufacturaP2 = new Table(1).UseAllAvailableWidth();
            tablaManufacturaP2.SetFixedLayout();
            Cell cellManufacturaP2 = new Cell(1, 1).Add(imgManufacturaP2)
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorder(Border.NO_BORDER);
            tablaManufacturaP2.AddCell(cellManufacturaP2);

            document.Add(tablaManufacturaP2);
        }


        public class FooterRevisionProtocolo : IEventHandler
        {
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

                Cell footer = new Cell(1, 1).Add(new Paragraph("F/CDC-045; Versión 12").AddStyle(estiloFooter)).SetBorder(Border.NO_BORDER).SetMargin(0).SetPadding(0);

                tablaResult.AddCell(footer).SetMargin(0).SetPadding(0);

                Rectangle rectangulo = new Rectangle(15, -20, pagina.GetPageSize().GetWidth() - 70, 50);

                new Canvas(pdfCanvas, rectangulo).Add(tablaResult);

            }
        }


    }
}
