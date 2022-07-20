﻿using iText.IO.Font.Constants;
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
        public string GenerarReporteActaVerificacion(List<CReporteGuiaRemisionModel> NumeroGuias, ListarOpcionesImprimir dato)
        {
            //DOCUMENTO QUE TIENE TRUE 
            Boolean documentoActaCC = dato.Acta;
            Boolean documentoCondiciones = dato.Condicion;
            Boolean documentoProtocolo = dato.Protocolo;
            Boolean documentoRsanitario = dato.Rsanitario;

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

                    Garantia(document, cabecera, rutaUnilene, fechaFinal);
                }

                if (documentoProtocolo)
                {
                    if (documentoActaCC || documentoCondiciones || documentoRsanitario)
                        document.Add(new AreaBreak(PageSize.A4));

                    Protocolo(document, cabecera, rutaUnilene, fechaFinal);
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


        public void GenerarPdf(Document document, CReporteGuiaRemisionModel cabecera, string Entrega,string fechaFinal)
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

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.NumeroMuestreo=="" || detalle.NumeroMuestreo=="-" ? "NO REQUIERE" : detalle.NumeroMuestreo)
                   .AddStyle(estiloDetalle))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);
                tablaDatosDetalle.AddCell(cellDetalle);

                cellDetalle = new Cell(1, 2).Add(new Paragraph(detalle.NumeroEnsayo=="" || detalle.NumeroEnsayo =="-" ? "NO REQUIERE" : detalle.NumeroEnsayo)
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
            document.Add(saltoLinea);

            //FIRMA
            Table tablaDatosFirma = new Table(3).UseAllAvailableWidth();
            tablaDatosFirma.SetFixedLayout().SetFontSize(9);

            
            string rutaUnilene2 = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLicitaciones.png");

            Image img2 = new Image(ImageDataFactory
               .Create(rutaUnilene2))
               .SetWidth(135)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.CENTER);


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

            cellFirma = new Cell(1, 1).Add(img2)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetBorder(Border.NO_BORDER)
                        .SetPaddingLeft(20);
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
            document.Add(saltoLinea);

           /* document.Add(new AreaBreak(PageSize.A4));
            Compromiso(document, cabecera, rutaUnilene, fechaFinal);
            document.Add(new AreaBreak(PageSize.A4));
            Condiciones(document, cabecera, rutaUnilene, fechaFinal);*/

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



            Paragraph LictacionPublica = new Paragraph( cabecera.DescripcionProceso + " - " + cabecera.DescripcionComercialDetalle + " (" + cabecera.CantItems.ToString() + "-ITEMS)-ITEM " + cabecera.DetalleGuia[0].NumeroItem.ToString() + ":" + cabecera.DetalleGuia[0].Descripcion + " . Es perteneciente a la OC " + cabecera.OrdenCompra).AddStyle(estilotextoNegrita);

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

            string FirmaLicitaciones = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLicitaciones.png");

            Image img2 = new Image(ImageDataFactory
               .Create(FirmaLicitaciones))
               .SetWidth(135)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);

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
            .SetMarginTop(10) ;
            tablaDatosFecha.AddCell(cellFecha);


            cellFecha = new Cell(1, 1).Add(img2)
            .SetBorder(Border.NO_BORDER)
            .SetMarginTop(100);
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

        public void Condiciones(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
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

            Style estiloFechaVerificacion = new Style()
                .SetFontSize(6)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Paragraph titulo1 = new Paragraph("DECLARACIÓN JURADA  DE CONDICIONES ESPECIALES DE EMBALAJE").AddStyle(estiloTitulo).SetMarginTop(10);
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

            Paragraph Nombre = new Paragraph("LUIS GERMAN CASTILLO VASQUEZ ").AddStyle(estiloNegrita);
            Paragraph ruc = new Paragraph("20197705249, DECLARO BAJO JURAMENTO ").AddStyle(estiloNegrita);


            Cell cellPresente = new Cell(1, 1).Add(new Paragraph("Presente")
                .AddStyle(estiloTexto))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            tablaDeclaracionJuramento.AddCell(cellPresente);


            cellPresente = new Cell(1, 1).Add(new Paragraph("El que se suscribe, ").Add(Nombre).Add("identificado con DNI Nº 07227297, Representante Legal de UNILENE S.A.C , con RUC Nº ").Add(ruc).Add("la información que a continuación se detalla respecto a las condiciones especiales de embalaje de la: ")
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


            string FirmaLicitaciones = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLicitaciones.png");

            Image img2 = new Image(ImageDataFactory
               .Create(FirmaLicitaciones))
               .SetWidth(135)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetBorder(Border.NO_BORDER)
               .SetHorizontalAlignment(HorizontalAlignment.LEFT);


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
              .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(new Paragraph(fechaFinal)
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetBorder(Border.NO_BORDER);
            tablaDatosFecha.AddCell(cellFecha);

            cellFecha = new Cell(1, 1).Add(img2)
          .SetBorder(Border.NO_BORDER)
          .SetMarginTop(100);

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

        public void Garantia(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
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

            Style estiloFechaVerificacion = new Style()
                .SetFontSize(6)
                .SetFont(fuenteNegrita)
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

            cellDG = new Cell(1, 1).Add(new Paragraph(cabecera.ClienteNombre + " EN SALUD - " + cabecera.Region)
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

            cellContenido = new Cell(1, 1).Add(new Paragraph("Asimismo, informamos que las condiciones de embalaje son en cajas de cartón y que el transporte es de mayor seguridad, cualquier reclamo de los productos entregados correspondiente a orden de compra ").Add(new Paragraph("Nº 1811").AddStyle(estiloNegrita)).Add (" serán cambiados dentro de las (24) veinticuatro horas siguientes")
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


            cellFecha = new Cell(1, 1).Add(new Paragraph(""))//imagen
            .SetBorder(Border.NO_BORDER)
            .SetMarginTop(100);
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

        public void Protocolo(Document document, CReporteGuiaRemisionModel cabecera, string rutaUnilene, string fechaFinal)
        {
            Color bgColorfondo = new DeviceRgb(217, 217, 217);


            Image img = new Image(ImageDataFactory
               .Create(rutaUnilene))
               .SetWidth(150)
               .SetHeight(50)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetTextAlignment(TextAlignment.LEFT);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

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
                .SetFontSize(9)
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

            //"N°:20938781 \n CQST030TC20005024ASCE"
            cellTitulo = new Cell(1, 1).Add(new Paragraph("N°:20938781 \n").AddStyle(estiloCabeceraSubtituloLote)).Add(new Paragraph("CQST030TC20005024ASCE \n").AddStyle(estiloCabeceraSubtituloCodsut))
            .SetTextAlignment(TextAlignment.RIGHT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetBorder(Border.NO_BORDER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE);
            tablaDatosTitulo.AddCell(cellTitulo);

            document.Add(tablaDatosTitulo);


            Table tablaDatosdeCabecera = new Table(6).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Producto:")
           .AddStyle(estiloCabeceraVariable))
           .SetTextAlignment(TextAlignment.LEFT)
           .SetHorizontalAlignment(HorizontalAlignment.CENTER)
           .SetVerticalAlignment(VerticalAlignment.MIDDLE)
           .SetBorderRight(Border.NO_BORDER)
           .SetBorderBottom(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);


            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph("Seda Negra Trenzada 3/0 Aguja 3/8 Círculo Cortante 20 mm x 50 cm CQ")
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



            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("Presentación:")
            .AddStyle(estiloCabeceraVariable))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER);

            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph("CAJA X 24 UNIDADES")
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("06/09/2021")
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("20938781")
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("7,929")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("06/09/2021")
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

            cellDatosCabecera = new Cell(1, 3).Add(new Paragraph("CIRUGIA PERUANA PLUS")
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

            cellDatosCabecera = new Cell(1, 1).Add(new Paragraph("06/09/2021")
            .AddStyle(estiloTexto))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderLeft(Border.NO_BORDER);
            tablaDatosdeCabecera.AddCell(cellDatosCabecera);

            document.Add(tablaDatosdeCabecera);
            document.Add(saltoLinea);

            Table tablaDetalleProtocolo = new Table(8).UseAllAvailableWidth();
            tablaDetalleProtocolo.SetFixedLayout();

            Cell cellDetalleProtocolo = new Cell(1, 2).Add(new Paragraph("Pruebas Efectuadas")
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


            cellDetalleProtocolo = new Cell(1, 2).Add(new Paragraph("Descripción")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 4).Add(new Paragraph("Sutura multifilamento de material no absorbible, constituida por hebra de seda trenzada de aspecto uniforme y homogéneo, color negro, protegido por sobre de papel grado médico con apertura peel open. Provista en un extremo de una aguja quirúrgica atraumática de acero inoxidable. El producto es esterilizado por óxido de etileno.")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.JUSTIFIED)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Cumple")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            cellDetalleProtocolo = new Cell(1, 1).Add(new Paragraph("Tecnica Propia")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaDetalleProtocolo.AddCell(cellDetalleProtocolo);

            document.Add(tablaDetalleProtocolo);

            Table tablaTecnicaPropia = new Table(6).UseAllAvailableWidth();
            tablaTecnicaPropia.SetFixedLayout();

            Cell cellTecnicaPropia = new Cell(1, 6)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaTecnicaPropia.AddCell(cellTecnicaPropia);


            
            string documentoTecnicaPropia = "TÉCNICA PROPIA código: TA/CDC-007; USP: Farmacopea de los Estados Unidos." +
             "(*) Antes NOM-067-SSA1-1993, ahora Farmacopea de los Estados Unidos Mexicanos FEUM - Suplemento para Dispositivos Médicos." +
             "Empaque de sobre de grado médico cumple con la norma ASTM F88/F88M-15.";

            string[] separarTecnicaPropia = documentoTecnicaPropia.Split('.');

            foreach (string tecnicapropia in separarTecnicaPropia)
            {
                cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(tecnicapropia)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
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
            }

        
            string documentoMetodoEsterilizacion = "Método de esterilización: Esterilizado con Óxido de Etileno según norma Internacional ISO 11135."+
                                             "Residuos de óxido de etileno: cumple con los requisitos de la ISO 10993 - 7(menor o igual a 4mg / primeras 24 horas, menor o igual a 60 mg / primeros 30 días y menor o igual a 2, 5 g durante toda la vida) según resultados de validación realizados periódicamente.";

            string[] separarMetodoEsterilizacion = documentoMetodoEsterilizacion.Split('.');

            foreach (string Metodo in separarMetodoEsterilizacion)
            {
                cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(Metodo)
                .AddStyle(estiloTextoDetalleProtocolo))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
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


            }

            string documentoBiocompatibilidad = "Este producto cumple con los ensayos de Biocompatibilidad, según norma ISO: ISO 10993-3 (Genotoxicidad) – No genotóxico; ISO 10993-5 (Citotoxicidad) – No citotóxico; ISO 10993-6 (Implantación) – No irritante; ISO 10993-10 (Sensibilidad cutánea) – Hipoalergénico e ISO 10993-11 (Toxicidad sistémica) – Atóxico. Las cuales se realizan en la etapa de desarrollo del producto.";
       
            cellTecnicaPropia = new Cell(1, 5).Add(new Paragraph(documentoBiocompatibilidad)
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
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

            Table tablaobservacion = new Table(13).UseAllAvailableWidth();
            tablaDatosTitulo.SetFixedLayout();

            Cell cellObservarcion = new Cell(1, 2).Add(new Paragraph("OBSERVACION: ")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 11)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);



            cellObservarcion = new Cell(1, 1).Add(new Paragraph("CONCLUSION: ")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);


            cellObservarcion = new Cell(1, 3)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(new Paragraph("APROBADO ")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(new Paragraph("X")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
             .SetBorderLeft(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER); ;
            tablaobservacion.AddCell(cellObservarcion);



            cellObservarcion = new Cell(1, 3)
            .SetBorderLeft(Border.NO_BORDER)
            .SetBorderBottom(Border.NO_BORDER)
            .SetBorderTop(Border.NO_BORDER)
            .SetBorderRight(Border.NO_BORDER);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(new Paragraph("RECHAZADO")
            .AddStyle(estiloTextoDetalleProtocolo))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetMarginTop(-10);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 1).Add(new Paragraph("X")
             .AddStyle(estiloTextoDetalleProtocolo))
             .SetTextAlignment(TextAlignment.LEFT)
             .SetHorizontalAlignment(HorizontalAlignment.LEFT)
             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
             .SetMarginTop(-10);
            tablaobservacion.AddCell(cellObservarcion);

            cellObservarcion = new Cell(1, 2).Add(new Paragraph("")
              .AddStyle(estiloTextoDetalleProtocolo))
              .SetTextAlignment(TextAlignment.LEFT)
              .SetHorizontalAlignment(HorizontalAlignment.LEFT)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetMarginTop(-10);
            tablaobservacion.AddCell(cellObservarcion);

            document.Add(tablaobservacion);
        }


    }
}
