﻿using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Response.AnalisisMateriaPrima;
using System;
using System.IO;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.AnalisisMateriaPrima.Hebra
{
    public class RptAnalisisMateriaPrima_PDF
    {
        public string GenerarReporte(AnalisisHebraDatosGeneralesDTO datos)
        {
            string reporte = null;


            MemoryStream ms = new MemoryStream();

            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Análisis Materia Prima");
            docInfo.SetAuthor("Sistema Satelite");

            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(5, 15, 30, 15);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            #region estilos

            Style estiloTitulo = new Style()
                .SetFontSize(12.5f)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloCabeceraTitleCelda = new Style().SetFont(fuenteNegrita).SetBorder(Border.NO_BORDER).SetFontSize(9);
            Style estiloCabeceraTextoCelda = new Style().SetFont(fuenteNormal).SetBorder(Border.NO_BORDER).SetFontSize(9);
            Style estiloDefectoTitleCelda = new Style().SetFont(fuenteNegrita).SetBorder(Border.NO_BORDER).SetFontSize(10)
                .SetBackgroundColor(new DeviceRgb(219, 219, 219)).SetTextAlignment(TextAlignment.CENTER);
            Style estiloDefectoTextoCelda = new Style().SetFont(fuenteNormal).SetFontSize(9).SetBorder(Border.NO_BORDER);
            Style estiloCabeceraAnalisis = new Style().SetFont(fuenteNegrita).SetFontSize(9)
                    .SetBackgroundColor(new DeviceRgb(255, 230, 153)).SetVerticalAlignment(VerticalAlignment.MIDDLE).SetTextAlignment(TextAlignment.CENTER);
            Style estiloDetalleAnalisis = new Style().SetFont(fuenteNormal).SetFontSize(9).SetTextAlignment(TextAlignment.CENTER);
            Style estiloAnalisisQuimica = new Style().SetFont(fuenteNormal).SetFontSize(9).SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER);
            Style estiloEquipo = new Style().SetFont(fuenteNormal).SetFontSize(8).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);

            #endregion

            #region titulo

            Table tituloTable = new Table(new float[] { 68, 32 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(10);

            Cell tituloCelda = new Cell(1, 1).Add(new Paragraph("REPORTE DE ANÁLISIS DE MATERIA PRIMA Y MATERIALES").AddStyle(estiloTitulo).SetTextAlignment(TextAlignment.LEFT));
            tituloTable.AddCell(tituloCelda.SetBorder(Border.NO_BORDER));

            tituloCelda = new Cell(1, 1).Add(new Paragraph("N° ANÁLISIS: " + datos.Cabecera.NumeroAnalisis).SetTextAlignment(TextAlignment.RIGHT));
            tituloTable.AddCell(tituloCelda.SetBorder(Border.NO_BORDER));

            document.Add(tituloTable);

            #endregion

            #region cabecera

            Table cabeceraTable = new Table(new float[] { 17, 17, 24, 4, 5, 4, 15, 14 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(10);

            Cell celdaCabecera = new Cell(1, 1).Add(new Paragraph("Producto:").SetPaddings(5, 0, 0, 5)).AddStyle(estiloCabeceraTitleCelda).SetBorderLeft(new SolidBorder(0.5f)).SetBorderTop(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Datos.Producto).SetPaddings(5, 0, 0, 0)).AddStyle(estiloCabeceraTextoCelda).SetBorderTop(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph("Item: ").SetTextAlignment(TextAlignment.RIGHT).SetPaddings(5, 0, 0, 0)).AddStyle(estiloCabeceraTitleCelda).SetBorderTop(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Datos.Item).SetTextAlignment(TextAlignment.LEFT).SetPaddings(5, 0, 0, 0)).AddStyle(estiloCabeceraTextoCelda).SetBorderTop(new SolidBorder(0.5f)).SetBorderRight(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph("Proveedor:").SetPaddings(0, 0, 0, 5)).AddStyle(estiloCabeceraTitleCelda).SetBorderLeft(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Datos.Proveedor)).AddStyle(estiloCabeceraTextoCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph("Lote: ").SetTextAlignment(TextAlignment.RIGHT)).AddStyle(estiloCabeceraTitleCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Datos.Lote)).SetTextAlignment(TextAlignment.LEFT).AddStyle(estiloCabeceraTextoCelda).SetBorderRight(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);


            celdaCabecera = new Cell(1, 1).Add(new Paragraph("N° Ingreso /# CC:").SetPaddings(0, 0, 0, 5)).AddStyle(estiloCabeceraTitleCelda).SetBorderLeft(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph(datos.Datos.NroIngreso)).AddStyle(estiloCabeceraTextoCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph("Certificado de análisis:    SI").SetTextAlignment(TextAlignment.RIGHT)).AddStyle(estiloCabeceraTitleCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Certificado == "S" ? "X" : "").SetTextAlignment(TextAlignment.CENTER)).AddStyle(estiloCabeceraTextoCelda).SetBorder(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);


            celdaCabecera = new Cell(1, 1).Add(new Paragraph("NO").SetTextAlignment(TextAlignment.RIGHT)).AddStyle(estiloCabeceraTitleCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Certificado == "N" ? "X" : "").SetTextAlignment(TextAlignment.CENTER)).AddStyle(estiloCabeceraTextoCelda).SetBorder(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph("F. Ingeso:").SetTextAlignment(TextAlignment.RIGHT)).AddStyle(estiloCabeceraTitleCelda);
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 1).Add(new Paragraph(datos.Datos.FechaIngreso == null ? "": datos.Datos.FechaIngreso.ToString("dd/MM/yyyy")).SetTextAlignment(TextAlignment.LEFT)).AddStyle(estiloCabeceraTextoCelda).SetBorderRight(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);


            celdaCabecera = new Cell(1, 1).Add(new Paragraph("O.Compra / FOR:").SetPaddings(0, 0, 5, 5)).AddStyle(estiloCabeceraTitleCelda).SetBorderBottom(new SolidBorder(0.5f)).SetBorderLeft(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Cabecera.OrdenCompra).SetPaddings(0, 0, 5, 0)).AddStyle(estiloCabeceraTextoCelda).SetBorderBottom(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph("F. Análisis: ").SetTextAlignment(TextAlignment.RIGHT).SetPaddings(0, 0, 5, 0)).AddStyle(estiloCabeceraTitleCelda).SetBorderBottom(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            celdaCabecera = new Cell(1, 3).Add(new Paragraph(datos.Cabecera.FechaAnalisis.ToString("dd/MM/yyyy")).SetTextAlignment(TextAlignment.LEFT).SetPaddings(0, 0, 5, 0)).AddStyle(estiloCabeceraTextoCelda).SetBorderRight(new SolidBorder(0.5f)).SetBorderBottom(new SolidBorder(0.5f));
            cabeceraTable.AddCell(celdaCabecera);

            document.Add(cabeceraTable);

            #endregion

            #region evaluacion

            Paragraph seccionEvaluacion = new Paragraph("I. EVALUACION").SetFontSize(11).SetPadding(1).SetBackgroundColor(new DeviceRgb(255, 230, 153)).SetMarginTop(4);
            document.Add(seccionEvaluacion);

            Table evaluacionDefectosTable = new Table(new float[] { 48, 4, 48 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMarginTop(7);

            Table defectoEnvaseTable = new Table(new float[] { 30, 20, 5, 20, 5 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell defectoEnvaseCell = new Cell(1, 5).Add(new Paragraph("Defectos Envase Inmediato / Mediato")).AddStyle(estiloDefectoTitleCelda);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Sellado Hermético:")).AddStyle(estiloDefectoTextoCelda);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Integridad del producto:")).AddStyle(estiloDefectoTextoCelda);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Excento de particular:")).AddStyle(estiloDefectoTextoCelda);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);

            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            defectoEnvaseCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoEnvaseTable.AddCell(defectoEnvaseCell);


            Cell evaluacionDefectosCell = new Cell(1, 1).Add(defectoEnvaseTable).SetBorder(Border.NO_BORDER);

            evaluacionDefectosTable.AddCell(evaluacionDefectosCell);

            evaluacionDefectosCell = new Cell(1, 1).Add(new Paragraph("")).SetBorder(Border.NO_BORDER);
            evaluacionDefectosTable.AddCell(evaluacionDefectosCell);

            Table defectoHebraTable = new Table(new float[] { 30, 20, 5, 20, 5 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell defectoHebraCell = new Cell(1, 5).Add(new Paragraph("Defectos de Hebra")).AddStyle(estiloDefectoTitleCelda);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Color homogeneo:")).AddStyle(estiloDefectoTextoCelda);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);


            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Libre de nódulos:")).AddStyle(estiloDefectoTextoCelda);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("No presenta roturas:")).AddStyle(estiloDefectoTextoCelda);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 5).Add(new Paragraph("")).SetHeight(2).SetBorder(Border.NO_BORDER);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Libre de hilachas:")).AddStyle(estiloDefectoTextoCelda);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);


            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("x")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.RIGHT);
            defectoHebraTable.AddCell(defectoHebraCell);

            defectoHebraCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(estiloDefectoTextoCelda).SetTextAlignment(TextAlignment.CENTER).SetBorder(new SolidBorder(0.5f));
            defectoHebraTable.AddCell(defectoHebraCell);


            evaluacionDefectosCell = new Cell(1, 1).Add(defectoHebraTable).SetBorder(Border.NO_BORDER);
            evaluacionDefectosTable.AddCell(evaluacionDefectosCell);

            document.Add(evaluacionDefectosTable);

            #endregion

            #region analisisDimensional

            Paragraph seccionAnalisisDimensional = new Paragraph("ANÁLISIS DIMENSIONAL:").SetFontSize(11).SetPaddings(3, 0, 0, 5);
            document.Add(seccionAnalisisDimensional);

            Table analisisTable = new Table(new float[] { 20, 20, 20, 20, 20 }).SetWidth(UnitValue.CreatePercentValue(80))
                .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);

            Cell analisisCell = new Cell(2, 1).Add(new Paragraph("N°")).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph("Longitud (cm)")).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph("Diametro (mm)")).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 2).Add(new Paragraph("Tensión (N)")).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(datos.Datos.Longitud.ToString("##0.#").Replace(".", ","))).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(datos.Datos.MinimoDiametro.ToString("##0.###").Replace(".", ",") + " - " + datos.Datos.MaximoDiametro.ToString("##0.###").Replace(".", ","))).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 2).Add(new Paragraph("Promedio Min: " + datos.Datos.Tension.ToString("##0.##").Replace(".", ","))).AddStyle(estiloCabeceraAnalisis);
            analisisTable.AddCell(analisisCell);

            foreach (TBDAnalisisHebraEntity prueba in datos.Detalle)
            {
                analisisCell = new Cell(1, 1).Add(new Paragraph(prueba.Numero.ToString(""))).AddStyle(estiloDetalleAnalisis);
                analisisTable.AddCell(analisisCell);

                analisisCell = new Cell(1, 1).Add(new Paragraph(prueba.Longitud is null ? "" : prueba.Longitud?.ToString("0.###").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
                analisisTable.AddCell(analisisCell);

                analisisCell = new Cell(1, 1).Add(new Paragraph(prueba.Diametro.ToString("0.###").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
                analisisTable.AddCell(analisisCell);

                analisisCell = new Cell(1, 1).Add(new Paragraph(prueba.Tension.ToString("0.###").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
                analisisTable.AddCell(analisisCell);

                analisisCell = new Cell(1, 1).Add(new Paragraph((((prueba.Tension/datos.Datos.Tension)-1)  * 100).ToString("0.#").Replace(".", ",") + " %")).AddStyle(estiloDetalleAnalisis);
                analisisTable.AddCell(analisisCell);
            }

            decimal[] longitudArray = datos.Detalle.Select(x => x.Longitud??0).ToArray();
            decimal[] diametroArray = datos.Detalle.Select(x => x.Diametro).ToArray();
            decimal[] tensionArray = datos.Detalle.Select(x => x.Tension).ToArray();

            analisisCell = new Cell(1, 1).Add(new Paragraph("Promedio")).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(longitudArray.Average().ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(diametroArray.Average().ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(tensionArray.Average().ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph((((tensionArray.Average() / datos.Datos.Tension) - 1) * 100).ToString("0").Replace(".", ",") + " %")).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);


            analisisCell = new Cell(1, 1).Add(new Paragraph("Desv. Est.")).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(Formulas.DesviacionEstandar(longitudArray).ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 1).Add(new Paragraph(Formulas.DesviacionEstandar(diametroArray).ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            analisisCell = new Cell(1, 2).Add(new Paragraph(Formulas.DesviacionEstandar(tensionArray).ToString("0.####").Replace(".", ","))).AddStyle(estiloDetalleAnalisis);
            analisisTable.AddCell(analisisCell);

            document.Add(analisisTable);

            #endregion

            #region analisisQuimica

            Paragraph seccionQuimica = new Paragraph("ANÁLISIS QUÍMICA:").SetFontSize(11).SetPaddings(10, 0, 0, 5);
            document.Add(seccionQuimica);

            Table quimicaTable;
            Cell quimicaCell;


            if (datos.Cabecera.Quimica == "C")
            {
                quimicaTable = new Table(new float[] { 35, 20, 5, 35, 5 }).SetWidth(UnitValue.CreatePercentValue(60))
                    .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);

                quimicaCell = new Cell(1, 1).Add(new Paragraph("Color extractable")).AddStyle(estiloAnalisisQuimica).SetFont(fuenteNegrita);
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloAnalisisQuimica).SetTextAlignment(TextAlignment.RIGHT);
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Quimica == "C" ? "X" : "")).AddStyle(estiloAnalisisQuimica).SetBorder(new SolidBorder(0.5f));
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloAnalisisQuimica).SetTextAlignment(TextAlignment.RIGHT);
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Quimica == "N" ? "X" : "")).AddStyle(estiloAnalisisQuimica).SetBorder(new SolidBorder(0.5f));
                quimicaTable.AddCell(quimicaCell);

            }
            else
            {
                quimicaTable = new Table(new float[] { 35, 5, 40, 5 }).SetWidth(UnitValue.CreatePercentValue(50))
                    .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);


                quimicaCell = new Cell(1, 1).Add(new Paragraph("Conforme:")).AddStyle(estiloAnalisisQuimica).SetTextAlignment(TextAlignment.RIGHT);
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Quimica == "C" ? "X" : "")).AddStyle(estiloAnalisisQuimica).SetBorder(new SolidBorder(0.5f));
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph("No Conforme:")).AddStyle(estiloAnalisisQuimica).SetTextAlignment(TextAlignment.RIGHT);
                quimicaTable.AddCell(quimicaCell);

                quimicaCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Quimica == "N" ? "X" : "")).AddStyle(estiloAnalisisQuimica).SetBorder(new SolidBorder(0.5f));
                quimicaTable.AddCell(quimicaCell);
            }

            
            document.Add(quimicaTable);

            #endregion

            #region equiposUtilizados

            Paragraph seccionEquipo = new Paragraph("EQUIPOS UTILIZADOS:").SetFontSize(11).SetPaddings(10, 0, 0, 5);
            document.Add(seccionEquipo);

            Table equipoTable;
            Cell equipoCell;

            if (datos.Cabecera.Quimica == "C")
            {
                equipoTable = new Table(new float[] { 20, 20, 20, 20, 20 }).SetWidth(UnitValue.CreatePercentValue(100))
                .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Balanza Analítica : CCBA-06")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Estufa: CCES-06")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Micrómetro: CCMI-01")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Regla metalica : RM-010")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("*Dinamómetro: CCTE - 162\n* Soporte vertical : CCSV-01")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                
            }
            else
            {
                equipoTable = new Table(new float[] { 33, 33, 33 }).SetWidth(UnitValue.CreatePercentValue(100))
                .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Micrómetro: CCMI-01")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("* Regla metalica : RM-010")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);

                equipoCell = new Cell(1, 1).Add(new Paragraph("*Dinamómetro: CCTE - 162\n* Soporte vertical : CCSV-01")).AddStyle(estiloEquipo);
                equipoTable.AddCell(equipoCell);
            }

            document.Add(equipoTable);

           
            #endregion

            #region observacion

            Paragraph seccionObservacion = new Paragraph("II. OBSERVACIÓN").SetFontSize(11).SetPadding(1).SetBackgroundColor(new DeviceRgb(255, 230, 153)).SetMarginTop(4);
            document.Add(seccionObservacion);

            seccionObservacion = new Paragraph(datos.Cabecera.Observaciones??"").SetFontSize(11).SetPadding(4).SetMarginTop(2)
                .SetBorder(new SolidBorder(0.5f)).SetMinHeight(15).SetFontSize(10);
            document.Add(seccionObservacion.SetBorder(Border.NO_BORDER).SetBorderBottom(new SolidBorder(0.5f)));

            #endregion

            #region conclusion

            Paragraph seccionConlusion = new Paragraph("III. CONCLUSIÓN").SetFontSize(11).SetPadding(1).SetBackgroundColor(new DeviceRgb(255, 230, 153)).SetMarginTop(4);
            document.Add(seccionConlusion);

            Table conclusionTable = new Table(new float[] { 40, 10, 40, 10  }).SetWidth(UnitValue.CreatePercentValue(50))
               .SetFixedLayout().SetHorizontalAlignment(HorizontalAlignment.CENTER).SetMarginTop(6);

            Cell conclusionCell = new Cell(1, 1).Add(new Paragraph("APROBADO:")).SetTextAlignment(TextAlignment.RIGHT).SetBorder(Border.NO_BORDER).SetFont(fuenteNegrita).SetFontSize(9);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Conclusion == "A" ? "X" : "")).SetTextAlignment(TextAlignment.CENTER).SetFontSize(9);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph("RECHAZADO:")).SetTextAlignment(TextAlignment.RIGHT).SetBorder(Border.NO_BORDER).SetFont(fuenteNegrita).SetFontSize(9);
            conclusionTable.AddCell(conclusionCell);

            conclusionCell = new Cell(1, 1).Add(new Paragraph(datos.Cabecera.Conclusion == "R" ? "X" : "")).SetTextAlignment(TextAlignment.CENTER).SetFontSize(9);
            conclusionTable.AddCell(conclusionCell);


            document.Add(conclusionTable);
            #endregion

            #region firma

            string rutaFirmaLilia = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\FirmaLiliaHurtadoDias.jpg");

            Table firmaTable = new Table(new float[] { 50, 50 })
                .SetWidth(UnitValue.CreatePercentValue(100))
                .SetFixedLayout().SetMarginTop(30);

            Image img = new Image(ImageDataFactory
               .Create(rutaFirmaLilia))
               .SetWidth(150)
               .SetHeight(75)
               .SetMarginBottom(0)
               .SetPadding(0)
               .SetMarginLeft(60)
               .SetTextAlignment(TextAlignment.LEFT);


            Cell firmaCell = new Cell(1, 1).Add(img).SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetBorder(Border.NO_BORDER);
            firmaTable.AddCell(firmaCell);

            firmaCell = new Cell(1, 1).Add(img).SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetBorder(Border.NO_BORDER);
            firmaTable.AddCell(firmaCell);

            document.Add(firmaTable);

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
