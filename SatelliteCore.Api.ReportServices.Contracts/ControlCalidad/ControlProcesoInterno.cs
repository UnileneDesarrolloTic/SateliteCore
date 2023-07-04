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
using System.Linq;
using System.Globalization;

namespace SatelliteCore.Api.ReportServices.Contracts.ControlCalidad
{
    public class ControlProcesoInterno
    {
        public string ReporteControlProcesoInterno(IEnumerable<DatosFormatoInformacionResultadoProtocolo> listado, DatosFormatoNumeroLoteProtocoloModel Cabecera)
        {
            string reporte = null;
            MemoryStream ms = new MemoryStream();
            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);
            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("REPORTE DE CONTROL EN PROCESO DE SUTURAS");
            docInfo.SetAuthor("Control de Calidad");

            Document document = new Document(pdf, PageSize.A4);
            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Paragraph saltoLinea = new Paragraph(new Text("\n"));
            LineSeparator lineaSeparadora = new LineSeparator(new SolidLine());

            Color bgColour = new DeviceRgb(192, 192, 192);
            DeviceCmyk bgColourBorder = new DeviceCmyk(0, 0, 0,25);

            NumberFormatInfo formato = new CultureInfo("en-US").NumberFormat;
            formato.CurrencyGroupSeparator = ".";
            formato.NumberDecimalSeparator = ",";
            int decimaldeC_DMaximo = (int) Cabecera.DEC_DMinimo;
            int decimaldeC_S_PromedioMinimo = (int)Cabecera.DEC_S_PromedioMinimo;

            document.SetMargins(5, 15, 30, 15);
            string rutaUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");

            DateTime dateExpiracion = new DateTime(Cabecera.FECHAEXPIRACION.Year, Cabecera.FECHAEXPIRACION.Month, Cabecera.FECHAEXPIRACION.Day);
            string ExpiracionFe = dateExpiracion.ToString("MM-yyyy");
            DateTime dateAnalisis = new DateTime(Cabecera.FECHAANALISIS.Year, Cabecera.FECHAANALISIS.Month, Cabecera.FECHAANALISIS.Day);
            string AnalisisFe = dateExpiracion.ToString("dd-MM-yyyy");

            Style estiloOrdenFabricacion = new Style()
              .SetFontSize(12)
              .SetFont(fuenteNegrita)
              .SetFontColor(ColorConstants.BLACK)
              .SetTextAlignment(TextAlignment.LEFT)
              .SetVerticalAlignment(VerticalAlignment.BOTTOM)
              .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Style estiloCabecera = new Style()
              .SetFontSize(10)
              .SetFontColor(ColorConstants.BLACK)
              .SetFont(fuenteNegrita);

            Style estiloCabeceraInput = new Style()
             .SetFontSize(10)
             .SetFontColor(ColorConstants.BLACK)
             .SetFont(fuenteNormal);

            Style estiloTablaCabecera = new Style()
              .SetFontSize(6.5f)
              .SetFontColor(ColorConstants.BLACK)
              .SetFont(fuenteNegrita)
              .SetBackgroundColor(bgColour);

            Style InputTabla = new Style()
             .SetFontSize(6.5f)
             .SetFontColor(ColorConstants.BLACK)
             .SetFont(fuenteNormal);

            Style InputTablaTexto = new Style()
             .SetFontSize(7.5f)
             .SetFontColor(ColorConstants.BLACK)
             .SetFont(fuenteNormal);

            Image img = new Image(ImageDataFactory
             .Create(rutaUnilene))
             .SetWidth(150)
             .SetHeight(52)
             .SetTextAlignment(TextAlignment.LEFT)
             .SetHorizontalAlignment(HorizontalAlignment.LEFT);

            Table tablaDatosCabecera = new Table(3).UseAllAvailableWidth();
            tablaDatosCabecera.SetFixedLayout().SetFontSize(9).SetMarginTop(0);

            Cell cellCabecera = new Cell(1, 1).Add(img)
                .SetVerticalAlignment(VerticalAlignment.BOTTOM)
                .SetBorder(Border.NO_BORDER);
            tablaDatosCabecera.AddCell(cellCabecera);

            cellCabecera = new Cell(1, 1).Add(new Paragraph("")
             .AddStyle(estiloCabecera))
             .SetBorder(Border.NO_BORDER);
            tablaDatosCabecera.AddCell(cellCabecera);

            cellCabecera = new Cell(1, 1).Add(new Paragraph("Orden Fab.:" + Cabecera.ORDENFABRICACION)
             .AddStyle(estiloOrdenFabricacion))
             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
             .SetBorder(Border.NO_BORDER);
            tablaDatosCabecera.AddCell(cellCabecera);

            document.Add(tablaDatosCabecera);

            /*Paragraph NumeroOrdenFabricacion = new Paragraph("Orden Fab.:"+ Cabecera.ORDENFABRICACION).AddStyle(estiloOrdenFabricacion);
            document.Add(NumeroOrdenFabricacion);*/

            Style estiloTitulo = new Style()
              .SetFontSize(12)
              .SetFont(fuenteNegrita)
              .SetMarginTop(-8)
              .SetFontColor(ColorConstants.BLACK)
              .SetTextAlignment(TextAlignment.CENTER);



            Paragraph titulo = new Paragraph("REPORTE DE CONTROL EN PROCESO DE SUTURAS").AddStyle(estiloTitulo);
            document.Add(titulo);

            Table tablaDatosGenerales = new Table(36).UseAllAvailableWidth();
            tablaDatosGenerales.SetFixedLayout().SetFontSize(9).SetMarginTop(1);

            Cell cellDG = new Cell(1, 5).Add(new Paragraph("PRODUCTO: ")
              .AddStyle(estiloCabecera))
              .SetTextAlignment(TextAlignment.RIGHT)
              .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
              .SetVerticalAlignment(VerticalAlignment.BOTTOM)
              .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 31).Add(new Paragraph(Cabecera.ITEMDESCRIPCION)
                .AddStyle(estiloCabeceraInput))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);



            cellDG = new Cell(1, 3).Add(new Paragraph("LOTE:")
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 10).Add(new Paragraph(Cabecera.REFERENCIANUMERO)
                .AddStyle(estiloCabeceraInput))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetHorizontalAlignment(HorizontalAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 3).Add(new Paragraph("F.Expira:")
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 5).Add(new Paragraph(ExpiracionFe)
                .AddStyle(estiloCabeceraInput))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 8).Add(new Paragraph("F.Inicio de análisis:")
                .AddStyle(estiloCabecera))
                .SetTextAlignment(TextAlignment.RIGHT)
                .SetHorizontalAlignment(HorizontalAlignment.RIGHT)
                .SetBorder(Border.NO_BORDER);

            tablaDatosGenerales.AddCell(cellDG);

            cellDG = new Cell(1, 7).Add(new Paragraph("_______________")
                .AddStyle(estiloCabeceraInput))
                .SetTextAlignment(TextAlignment.LEFT)
                .SetBorder(Border.NO_BORDER);
            //.SetBorderBottom(new SolidBorder(1));

            tablaDatosGenerales.AddCell(cellDG);

            document.Add(tablaDatosGenerales);


            Paragraph rectanguloTabla1 = new Paragraph()
                    .SetHeight(8)
                    .SetWidth(8)
                    .SetBackgroundColor(ColorConstants.WHITE)
                    .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f));

            Paragraph rectanguloTabla2 = new Paragraph()
                       .SetHeight(23)
                       .SetWidth(60)
                       .SetBackgroundColor(ColorConstants.WHITE)
                       .SetTextAlignment(TextAlignment.CENTER)
                       .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f));


            List<DatosFormatoInformacionResultadoProtocolo> listadoTablaA = listado.Where(x => x.TABLA == "A").ToList();
            List<DatosFormatoInformacionResultadoProtocolo> listadoTablaB = listado.Where(x => x.TABLA == "B").ToList();


            Table tablaDatosMedicion = new Table(33).UseAllAvailableWidth();
            tablaDatosMedicion.SetFixedLayout().SetFontSize(4).SetMarginTop(75);


            //bloque 2

            Cell cellDetalle = new Cell(1, 12)
                           .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3)
                .SetBackgroundColor(bgColour)
                .SetBorder(Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(0.5f))
                .SetBorderRight(new SolidBorder(0.5f))
                .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1)
               .SetBackgroundColor(bgColour)
               .SetBorder(Border.NO_BORDER)
               .SetBorderTop(new SolidBorder(0.5f))
               .SetWidth(0.1f);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 4).Add(new Paragraph("Longitud")
               .AddStyle(estiloTablaCabecera))
               .SetTextAlignment(TextAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBackgroundColor(bgColour)
               .SetBorder(Border.NO_BORDER)
               .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1)
               .SetBackgroundColor(bgColour)
               .SetWidth(0.1f)
               .SetBorder(Border.NO_BORDER)
               .SetBorderRight(new SolidBorder(0.5f))
               .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 2).Add(new Paragraph("Diámetro")
              .AddStyle(estiloTablaCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour)
              .SetBorder(Border.NO_BORDER)
              .SetBorderRight(new SolidBorder(0.5f))
              .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            //SEPARADOR
            cellDetalle = new Cell(1, 1)
            .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);


            //OTRA TABLA
            cellDetalle = new Cell(2, 3)
               .Add(rectanguloTabla2)
               .SetBorder(Border.NO_BORDER)
               .SetBackgroundColor(bgColour)
               .SetBorderLeft(new SolidBorder(0.5f))
               .SetBorderRight(new SolidBorder(0.5f))
               .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3).Add(new Paragraph("Resistencia a")
              .AddStyle(estiloTablaCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour)
              .SetBorder(Border.NO_BORDER)
              .SetBorderLeft(new SolidBorder(0.5f))
              .SetBorderRight(new SolidBorder(0.5f))
              .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3).Add(new Paragraph("Unión Hebra - ")
              .AddStyle(estiloTablaCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour)
              .SetBorder(Border.NO_BORDER)
              .SetBorderLeft(new SolidBorder(0.5f))
              .SetBorderRight(new SolidBorder(0.5f))
              .SetBorderTop(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);


            //bloque 2
            cellDetalle = new Cell(1, 12)
                .SetHeight(4)
                .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3)
                .SetBackgroundColor(bgColour)
                .SetBorder(Border.NO_BORDER)
                .SetBorderLeft(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1)
               .SetBackgroundColor(bgColour)
               .SetBorder(Border.NO_BORDER)
               .SetBorderLeft(new SolidBorder(0.5f))
               .SetWidth(0.1f);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("cm")
               .AddStyle(estiloTablaCabecera))
               .SetTextAlignment(TextAlignment.RIGHT)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBackgroundColor(bgColour)
               .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(rectanguloTabla1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBorder(Border.NO_BORDER)
               .SetBackgroundColor(bgColour)
               .SetWidth(0.1f);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(new Paragraph("m")
               .AddStyle(estiloTablaCabecera))
               .SetTextAlignment(TextAlignment.RIGHT)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBackgroundColor(bgColour)
               .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1).Add(rectanguloTabla1)
               .SetTextAlignment(TextAlignment.CENTER)
               .SetVerticalAlignment(VerticalAlignment.MIDDLE)
               .SetBorder(Border.NO_BORDER)
               .SetBackgroundColor(bgColour)
               .SetWidth(0.1f);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 1)
            .SetBackgroundColor(bgColour)
            .SetBorder(Border.NO_BORDER)
            .SetBorderRight(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            //DIAMETRO
            cellDetalle = new Cell(1, 2).Add(new Paragraph("(mm)")
            .AddStyle(estiloTablaCabecera))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetVerticalAlignment(VerticalAlignment.MIDDLE)
            .SetBackgroundColor(bgColour)
            .SetBorder(Border.NO_BORDER)
            .SetBorderRight(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            //SEPARADOR
            cellDetalle = new Cell(1, 1)
            .SetBorder(Border.NO_BORDER);
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3).Add(new Paragraph("la Tensión")
              .AddStyle(estiloTablaCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour)
              .SetBorder(Border.NO_BORDER)
              .SetBorderLeft(new SolidBorder(0.5f))
              .SetBorderRight(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);

            cellDetalle = new Cell(1, 3).Add(new Paragraph("Aguja")
              .AddStyle(estiloTablaCabecera))
              .SetTextAlignment(TextAlignment.CENTER)
              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
              .SetBackgroundColor(bgColour)
              .SetBorder(Border.NO_BORDER)
              .SetBorderLeft(new SolidBorder(0.5f))
              .SetBorderRight(new SolidBorder(0.5f));
            tablaDatosMedicion.AddCell(cellDetalle);


            for (int i = 0; i < 8; i++)
            {

                if (i > 6)
                {
                    cellDetalle = new Cell(2, 12).Add(new Paragraph("")
                    .AddStyle(InputTabla))
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                    .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 3).Add(new Paragraph("")
                        .AddStyle(InputTabla))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                        .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 6).Add(new Paragraph("")
                       .AddStyle(InputTabla))
                       .SetTextAlignment(TextAlignment.CENTER)
                       .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                       .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 2).Add(new Paragraph("")
                      .AddStyle(estiloTablaCabecera))
                      .SetTextAlignment(TextAlignment.CENTER)
                      .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                      .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 1).Add(new Paragraph("")
                         .AddStyle(InputTabla))
                         .SetTextAlignment(TextAlignment.CENTER)
                         .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                         .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);
                }
                else
                {
                    var SecuenciaA = listadoTablaA[i].SECUENCIA == 6 ? "Promedio" : listadoTablaA[i].SECUENCIA == 7 ? "Desv. Est.:" : listadoTablaA[i].SECUENCIA.ToString();


                    cellDetalle = new Cell(2, 12).Add(new Paragraph("")
                     .AddStyle(InputTabla))
                     .SetTextAlignment(TextAlignment.CENTER)
                     .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                     .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 3).Add(new Paragraph(SecuenciaA)
                        .AddStyle(InputTabla))
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 6).Add(new Paragraph((listadoTablaA[i].SECUENCIA == 7) ? listadoTablaA[i].COL_1.ToString("#,##0.0000", formato) : Math.Round(listadoTablaA[i].COL_1, 1).ToString("#,##0.0", formato))
                       .AddStyle(InputTablaTexto))
                       .SetTextAlignment(TextAlignment.CENTER)
                       .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 2).Add(new Paragraph((listadoTablaA[i].SECUENCIA == 7) ? listadoTablaA[i].COL_2.ToString("#,##0.0000", formato) : Math.Round(listadoTablaA[i].COL_2, decimaldeC_DMaximo).ToString(formato))
                      .AddStyle(InputTablaTexto))
                      .SetTextAlignment(TextAlignment.CENTER)
                      .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                    tablaDatosMedicion.AddCell(cellDetalle);

                    cellDetalle = new Cell(2, 1).Add(new Paragraph("")
                         .AddStyle(estiloTablaCabecera))
                         .SetTextAlignment(TextAlignment.CENTER)
                         .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                         .SetBorder(Border.NO_BORDER);

                    tablaDatosMedicion.AddCell(cellDetalle);
                }

                var SecuenciaB = listadoTablaB[i].SECUENCIA == 6 ? "Promedio" : listadoTablaB[i].SECUENCIA == 7 ? "Ind. Min.:" : listadoTablaB[i].SECUENCIA == 8 ? "Desv. Est.:" : listadoTablaB[i].SECUENCIA.ToString();

                cellDetalle = new Cell(2, 3).Add(new Paragraph(SecuenciaB)
                   .AddStyle(InputTabla))
                   .SetTextAlignment(TextAlignment.CENTER)
                   .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                tablaDatosMedicion.AddCell(cellDetalle);

                cellDetalle = new Cell(2, 3).Add(new Paragraph((listadoTablaB[i].SECUENCIA == 8) ? listadoTablaB[i].COL_1.ToString("#,##0.0000", formato) : Math.Round(listadoTablaB[i].COL_1, decimaldeC_S_PromedioMinimo).ToString(formato))
                  .AddStyle(InputTablaTexto))
                  .SetTextAlignment(TextAlignment.CENTER)
                  .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                tablaDatosMedicion.AddCell(cellDetalle);

                cellDetalle = new Cell(2, 3).Add(new Paragraph((listadoTablaB[i].SECUENCIA == 8) ? listadoTablaB[i].COL_2.ToString("#,##0.0000", formato) : Math.Round(listadoTablaB[i].COL_2, decimaldeC_S_PromedioMinimo).ToString(formato))
                  .AddStyle(InputTablaTexto))
                  .SetTextAlignment(TextAlignment.CENTER)
                  .SetVerticalAlignment(VerticalAlignment.MIDDLE);

                tablaDatosMedicion.AddCell(cellDetalle);

            }

            document.Add(tablaDatosMedicion);

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
