using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using System.IO;
using iText.Layout.Element;
using iText.IO.Image;
using iText.Layout.Properties;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Layout.Borders;
using System;
using SatelliteCore.Api.Models.Dto.AnalisisAgujas;
using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;
using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas;

namespace SatelliteCore.Api.ReportServices.Contracts.AnalsisAguja
{
    public class PruebasAnalisis
    {
        public string GenerarReporte(string loteAnalisis, ObtenerDatosGeneralesDTO datosGenerales, AnalisisAgujaPlanMuestreoEntity planMuestreo,
            List<AnalisisAgujaPruebaDimensionalEntity> dimensionaCorrosion, List<AnalisisAgujaFlexionEntity> flexion,
            List<AnalisisAgujaElasticidadPerforacionEntity> elasticidadPerforeacion, List<AnalisisAgujaPruebaAspectoEntity> aspectoAguja)
        {
            string reporte = null;            

            MemoryStream ms = new MemoryStream();

            PdfWriter writer = new PdfWriter(ms);
            PdfDocument pdf = new PdfDocument(writer);

            PdfDocumentInfo docInfo = pdf.GetDocumentInfo();
            docInfo.SetTitle("Reporte análisis de aguja");
            docInfo.SetAuthor("Sistema Satelite");


            Document document = new Document(pdf, PageSize.A4);
            document.SetMargins(10, 25, 50, 18);

            pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterAnalisisAgujaEventHandler());

            string rutaLogoUnilene = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image img = new Image(ImageDataFactory.Create(rutaLogoUnilene))
                .SetWidth(115)
                .SetHeight(47)
                .SetMarginBottom(0)
                .SetPadding(0)
                .SetTextAlignment(TextAlignment.CENTER);

            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            PdfFont fuenteNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            Style estiloTitulo = new Style()
                .SetFontSize(12)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style estiloSpanRowTable = new Style()
                .SetFontSize(8.95f)
                .SetFont(fuenteNormal)
                .SetFontColor(ColorConstants.BLACK);

            Style headerTableDescAguja = new Style()
                .SetFontSize(8.75f)
                .SetFont(fuenteNegrita)
                .SetFontColor(ColorConstants.BLACK);

            Style borderTable = new Style()
                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.CENTER);

            Style headerTableGray = new Style()
                .SetFontSize(8.5f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNegrita)
                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                .SetHeight(16);


            Style headerTableCommon = new Style()
                .SetFontSize(7.15f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNormal)
                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                .SetHeight(16);

            Style headerTablePerforacion = new Style()
                .SetFontSize(7.15f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNormal)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBackgroundColor(ColorConstants.LIGHT_GRAY);

            Style cellTablePerforacion = new Style()
                .SetFontSize(7.15f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNormal)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.LEFT);

            Style cellTableElasticidad = new Style()
                .SetFontSize(7.15f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNormal)
                .SetBorder(new SolidBorder(ColorConstants.BLACK, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.CENTER);

            Style cellTableCommon = new Style()
                .SetFontSize(7.15f)
                .SetPadding(0)
                .SetMargin(0)
                .SetFont(fuenteNormal)
                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.LEFT);

            Style cellTableMd = new Style()
                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.5f))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                .SetTextAlignment(TextAlignment.LEFT);


            Table headerDocTable = new Table(new float[] { 1, 2, 1 }).UseAllAvailableWidth();
            headerDocTable.SetWidth(UnitValue.CreatePercentValue(100));
            headerDocTable.SetFixedLayout();

            Cell cellHeaderDoc = new Cell(1, 1).Add(img).SetBorder(Border.NO_BORDER);

            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("REPORTE DE ANÁLISIS DE AGUJAS").AddStyle(estiloTitulo))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetTextAlignment(TextAlignment.CENTER).SetBorder(Border.NO_BORDER);

            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph($"N° {loteAnalisis}").AddStyle(estiloTitulo))
                .SetVerticalAlignment(VerticalAlignment.MIDDLE).SetTextAlignment(TextAlignment.RIGHT).SetBorder(Border.NO_BORDER);

            headerDocTable.AddCell(cellHeaderDoc).SetMarginBottom(10);
            document.Add(headerDocTable);


            headerDocTable = new Table(new float[] { 1, 1, 1, 1, 1, 1, 1 }).UseAllAvailableWidth();
            headerDocTable.SetWidth(UnitValue.CreatePercentValue(100));
            headerDocTable.SetFixedLayout();

            cellHeaderDoc = new Cell(2, 1).Add(new Paragraph("DESCRIPCIÓN DE LA AGUJA").AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("TIPO").AddStyle(headerTableDescAguja)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("LONGITUD").AddStyle(headerTableDescAguja)).AddStyle(borderTable);

            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("BROCA").AddStyle(headerTableDescAguja)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("Φ ALAMBRE").AddStyle(headerTableDescAguja)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("SERIE").AddStyle(headerTableDescAguja)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph("F. ANÁLISIS").AddStyle(headerTableDescAguja)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.CodTipo).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.CodLongitud).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.CodBroca).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.CodAlambre).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.Serie).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            cellHeaderDoc = new Cell(1, 1).Add(new Paragraph(datosGenerales.FechaAnalisis.ToString("dd/MM/yyyy")).AddStyle(estiloSpanRowTable)).AddStyle(borderTable);
            headerDocTable.AddCell(cellHeaderDoc);

            document.Add(headerDocTable.SetMarginBottom(6));


            Table contenedorDatosPlan = new Table(new float[] { 45, 55 });
            contenedorDatosPlan.SetWidth(UnitValue.CreatePercentValue(100));
            contenedorDatosPlan.SetFixedLayout();

            Table datosTable = new Table(new float[] { 30, 70 });
            datosTable.SetWidth(UnitValue.CreatePercentValue(100));
            datosTable.SetFixedLayout();


            Cell datosCell = new Cell(1, 2).Add(new Paragraph("DATOS")).AddStyle(headerTableGray);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("FOR").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(datosGenerales.OrdenCompra).SetFontSize(8.95f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("N° CONTROL").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(datosGenerales.ControlNumero).SetFontSize(8.95f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("PROVEEDOR").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(datosGenerales.Proveedor).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("CANT. INGRESADA").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(datosGenerales.Cantidad)).SetFontSize(8.95f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 2).Add(new Paragraph("Referencia: MP/CDC-018").SetFontSize(7.05f))
                .AddStyle(cellTableMd)
                .SetBorder(Border.NO_BORDER)
                .SetMargin(0);
            datosTable.AddCell(datosCell);


            Cell cellAux = new Cell(1, 1).Add(datosTable).SetPadding(0).SetPaddingRight(2).SetBorder(Border.NO_BORDER);

            contenedorDatosPlan.AddCell(cellAux);

            datosTable = new Table(new float[] { 48, 33, 19 }).UseAllAvailableWidth();
            datosTable.SetWidth(UnitValue.CreatePercentValue(100));
            datosTable.SetFixedLayout();

            datosCell = new Cell(1, 3).Add(new Paragraph("PLAN DE MUESTREO")).AddStyle(headerTableGray);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("PRUEBA DE FUNCIONABILIDAD").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("Nº unid. a muestrear").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(planMuestreo.UndMuestrear)).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("PRUEBAS DIMENSIONALES LONGITUD").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("Nº unid. a muestrear (I)").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(planMuestreo.UndMuestrearI)).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(3, 1).Add(new Paragraph("OTRAS PRUEBAS DIMENSIONALES").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("Nº unid. a muestrear (III)").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(planMuestreo.UndMuestrearIII)).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph("Nº cajas a muestrear").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(planMuestreo.CajasMuestrear)).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);


            string unidadPorCaja = SeparadorDeMilesDouble(Math.Ceiling(planMuestreo.UndMuestrearIII / (planMuestreo.CajasMuestrear * 1.0)));

            datosCell = new Cell(1, 1).Add(new Paragraph("Nº unidades por caja").SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);

            datosCell = new Cell(1, 1).Add(new Paragraph(unidadPorCaja).SetFontSize(7.15f)).AddStyle(cellTableMd);
            datosTable.AddCell(datosCell);


            cellAux = new Cell(1, 1).Add(datosTable).SetPadding(0).SetPaddingLeft(1).SetBorder(Border.NO_BORDER);
            contenedorDatosPlan.AddCell(cellAux);

            document.Add(contenedorDatosPlan);

            Paragraph parafoAceptabilidad = new Paragraph("1. ANÁLISIS DE ACEPTABILIDAD DE DEFECTOS DE FUNCIONALIDAD").SetFontSize(8.95f);
            document.Add(parafoAceptabilidad);

            Table pruebaAceptabilidad = new Table(new float[] { 13, 20, 18, 12, 9, 13, 15 });
            pruebaAceptabilidad.SetWidth(UnitValue.CreatePercentValue(100));
            pruebaAceptabilidad.SetFixedLayout();

            Cell cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);


            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon); ;
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(2, 1).Add(new Paragraph("CORROSIÓN")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(2, 1).Add(new Paragraph("RESISTENTE A LA\nCORROSIÓN")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("Presenta Corrosión").SetPaddingLeft(3)).AddStyle(cellTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);


            AnalisisAgujaPruebaDimensionalEntity pruebaCorrosion = dimensionaCorrosion.Find(x => x.TipoRegistro == 4);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(pruebaCorrosion.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph(pruebaCorrosion.Cantidad > 0 ? "100" : "0")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("= 0 %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(2, 1).Add(new Paragraph(pruebaCorrosion.Cantidad > 0 ? "RECHAZADO" : "ACEPTADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("No presenta corrosión").SetPaddingLeft(3)).AddStyle(cellTableCommon);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            string noCorrosion = SeparadorDeMilesDecimal(planMuestreo.UndMuestrear - pruebaCorrosion.Cantidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph(noCorrosion)).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph(pruebaCorrosion.Cantidad > 0 ? "0" : "100")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 1).Add(new Paragraph("= 100 %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);

            cellAceptabilidad = new Cell(1, 7).Add(new Paragraph("(*) Si esta prueba tiene como resultado RECHAZADO termina el análisis")).AddStyle(cellTableCommon)
                .SetTextAlignment(TextAlignment.RIGHT).SetHorizontalAlignment(HorizontalAlignment.RIGHT).SetBorder(Border.NO_BORDER).SetPaddingRight(2);
            pruebaAceptabilidad.AddCell(cellAceptabilidad);


            document.Add(pruebaAceptabilidad.SetMarginBottom(6));

            Table pruebaFlexioTable = new Table(new float[] { 12, 76, 12 });
            pruebaFlexioTable.SetWidth(UnitValue.CreatePercentValue(100));
            pruebaFlexioTable.SetFixedLayout();

            Cell pruebaFlexionCell = new Cell(1, 1).Add(new Paragraph("Prueba")).AddStyle(headerTableCommon);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);

            pruebaFlexionCell = new Cell(1, 1).Add(new Paragraph("RESULTADOS % POR CICLO")).AddStyle(headerTableCommon);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);

            pruebaFlexionCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);

            pruebaFlexionCell = new Cell(1, 1).Add(new Paragraph($"FLEXIÓN\nCant: {SeparadorDeMilesDecimal(planMuestreo.UndMuestrearI)} und")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);


            Table resultadoFlexion = new Table(new float[] { 25, 25, 25, 25 });
            resultadoFlexion.SetWidth(UnitValue.CreatePercentValue(100));
            resultadoFlexion.SetFixedLayout();

            Cell cellTablaResultado;

            foreach (AnalisisAgujaFlexionEntity item in flexion)
            {
                Table tablaResultado = new Table(new float[] { 55, 45 });
                tablaResultado.SetWidth(UnitValue.CreatePercentValue(100));
                tablaResultado.SetFixedLayout();

                Cell cellResultado = new Cell(1, 1).Add(new Paragraph($"{SeparadorDeMilesDecimal(item.Llave)} ciclos: ")).AddStyle(cellTableCommon)
                    .SetBorder(Border.NO_BORDER)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaResultado.AddCell(cellResultado);

                cellResultado = new Cell(1, 1).Add(new Paragraph($"{SeparadorDeMilesDecimal(item.Valor)} %"))
                    .AddStyle(cellTableCommon)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetPadding(0)
                    .SetMargin(0);

                tablaResultado.AddCell(cellResultado);

                cellTablaResultado = new Cell(1, 1).Add(tablaResultado)
                    .SetBorder(Border.NO_BORDER)
                    .SetTextAlignment(TextAlignment.CENTER);

                resultadoFlexion.AddCell(cellTablaResultado);
            }

            cellTablaResultado = new Cell(1, 1).Add(resultadoFlexion).AddStyle(cellTableCommon);
            pruebaFlexioTable.AddCell(cellTablaResultado);

            pruebaFlexionCell = new Cell(1, 1).Add(new Paragraph(ObtenerDescripcionEstado(planMuestreo.StatusFlexion))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);

            pruebaFlexionCell = new Cell(1, 3).Add(new Paragraph("(*) Si esta prueba tiene como resultado RECHAZADO termina el análisis")).AddStyle(cellTableCommon)
                .SetTextAlignment(TextAlignment.RIGHT).SetHorizontalAlignment(HorizontalAlignment.RIGHT).SetBorder(Border.NO_BORDER).SetPaddingRight(2);
            pruebaFlexioTable.AddCell(pruebaFlexionCell);

            document.Add(pruebaFlexioTable.SetMarginBottom(5));

            Paragraph parafoDimensionales = new Paragraph("2. ANÁLISIS DE ACEPTABILIDAD DE PRUEBAS DIMENSIONALES").SetFontSize(8.95f);

            document.Add(parafoDimensionales.SetMarginBottom(6));

            Table pruebasDimensionalesTable = new Table(new float[] { 18, 17, 20, 11, 10, 11, 13 });
            pruebasDimensionalesTable.SetWidth(UnitValue.CreatePercentValue(100));
            pruebasDimensionalesTable.SetFixedLayout();

            Cell pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("LONGITUD DE AGUJA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("LONGITUD SOLICITADA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            AnalisisAgujaPruebaDimensionalEntity longitud = dimensionaCorrosion.Find(x => x.TipoRegistro == 1);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("Longitud solicitada")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(longitud.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            decimal porcientoLongitud = decimal.Round((longitud.Cantidad * 100) / longitud.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcientoLongitud))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"> ó = {SeparadorDeMilesDecimal(longitud.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph(porcientoLongitud >= longitud.Tolerancia ? "ACEPTADO" : "RECHAZADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"Longitud opción 2 : {longitud.DescripcionAux}")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"{longitud.CantidadAux}")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            decimal porcientoLongitudAux = decimal.Round(((longitud.CantidadAux ?? 0) * 100) / longitud.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"{SeparadorDeMilesDecimal(porcientoLongitudAux)}")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"< {SeparadorDeMilesDecimal(100 - longitud.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            document.Add(pruebasDimensionalesTable.SetMarginBottom(12));

            pruebasDimensionalesTable = new Table(new float[] { 18, 17, 20, 11, 10, 11, 13 });
            pruebasDimensionalesTable.SetWidth(UnitValue.CreatePercentValue(100));
            pruebasDimensionalesTable.SetFixedLayout();

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);


            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("DIAMETRO DEL AGUJERO (BROCA)")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("DIAMETRO SOLICITADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            AnalisisAgujaPruebaDimensionalEntity diametro = dimensionaCorrosion.Find(x => x.TipoRegistro == 2);

            decimal porcentajeDiametro = decimal.Round((diametro.Cantidad * 100) / diametro.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("Diametro solicitada")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(diametro.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeDiametro))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"> ó = {SeparadorDeMilesDecimal(diametro.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph(porcentajeDiametro >= diametro.Tolerancia ? "ACEPTADO":"RECHAZADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"Diametro opción 2: { diametro.DescripcionAux }")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"{diametro.CantidadAux ?? 0}")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            decimal porcentajeDiametroAux = decimal.Round(((diametro.CantidadAux ?? 0) * 100) / diametro.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"{SeparadorDeMilesDecimal(porcentajeDiametroAux)}")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"< {SeparadorDeMilesDecimal(100 - diametro.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            document.Add(pruebasDimensionalesTable.SetMarginBottom(12));

            pruebasDimensionalesTable = new Table(new float[] { 18, 17, 20, 11, 10, 11, 13 });
            pruebasDimensionalesTable.SetWidth(UnitValue.CreatePercentValue(100));
            pruebasDimensionalesTable.SetFixedLayout();

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);


            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("DIAMETRO DEL ALAMBRE")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph("ALAMBRE SOLICITADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            AnalisisAgujaPruebaDimensionalEntity alambre = dimensionaCorrosion.Find(x => x.TipoRegistro == 3);

            decimal porcentajeAlambre = decimal.Round((alambre.Cantidad * 100) / alambre.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph("Alambre solicitada")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(alambre.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeAlambre))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"> ó = {SeparadorDeMilesDecimal(alambre.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(2, 1).Add(new Paragraph(porcentajeAlambre >= alambre.Tolerancia ? "ACEPTADO" : "RECHAZADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"Alambre opción 2: { alambre.DescripcionAux }")).AddStyle(cellTableCommon).SetPaddingLeft(3);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"{alambre.CantidadAux ?? 0}")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            decimal porcentajeAlambreAux = decimal.Round(((alambre.CantidadAux ?? 0) * 100) / alambre.BaseCalculoEstado, 2);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeAlambreAux))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            pruebaDimensionaCell = new Cell(1, 1).Add(new Paragraph($"< {SeparadorDeMilesDecimal(100 - alambre.Tolerancia)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            pruebasDimensionalesTable.AddCell(pruebaDimensionaCell);

            document.Add(pruebasDimensionalesTable.SetMarginBottom(10));

            Paragraph parafoPerforacion = new Paragraph("3. PRUEBA DE FUERZA DE PERFORACIÓN").SetFontSize(8.95f);
            document.Add(parafoPerforacion.SetMarginBottom(5));

            parafoPerforacion = new Paragraph("Especificaciones").SetFontSize(8.95f).SetPaddingLeft(15);
            document.Add(parafoPerforacion.SetMarginBottom(6));

            Table pruebaPerforacionTable = new Table(new float[] { 25, 25, 25, 25 });
            pruebaPerforacionTable.SetWidth(UnitValue.CreatePercentValue(70));
            pruebaPerforacionTable.SetFixedLayout();

            Cell pruebaPerforacionCell = new Cell(2, 1).Add(new Paragraph("Rango de diametro de la aguja mm")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(2, 1).Add(new Paragraph("Carga\nN\n(Newton)")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 2).Add(new Paragraph("Fuerza de Perforación")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("Aguja punta redonda\n(menor o igual)")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("Aguja punta cotante\n(menor o igual)")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);


            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0,2 a 0,4")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.39")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.63")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.49")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0,5 a 0,7")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.58")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.68")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.58")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0,8 a 1,0")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.78")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.78")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.68")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("1,1 a 1,3")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.98")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.93")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.78")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("> 1,3")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.98")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("1.18")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("0.88")).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 4).Add(new Paragraph("Nota: No aplica en agujas con L (longitud) igual o menor de 12 mm. Referencia: YY/T 0043-2016").SetPaddingLeft(2))
                .AddStyle(cellTablePerforacion)
                .SetTextAlignment(TextAlignment.LEFT);

            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            document.Add(pruebaPerforacionTable.SetMarginRight(0).SetMarginLeft(70).SetMarginBottom(6));


            pruebaPerforacionTable = new Table(new float[] { 85, 15 });
            pruebaPerforacionTable.SetWidth(UnitValue.CreatePercentValue(100));
            pruebaPerforacionTable.SetFixedLayout();

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("FUERZA DE PERFORACIÓN DE LA AGUJA (N)")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTablePerforacion);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            Table resultadoFuerzaPerforacionTable = new Table(new float[] { 18, 9, 9, 9, 9, 9, 18 })
                .SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell resultadoFuerzaPerforacionCell = new Cell(2, 1).Add(new Paragraph("Resultado (N):").SetFontSize(9f)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("1").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("2").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("3").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("4").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("5").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

            resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("PROMEDIO").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);


            AnalisisAgujaElasticidadPerforacionEntity perforacion = elasticidadPerforeacion.Find(x => x.TipoRegistro == 2);

            decimal promedioPerforacion = decimal.Round((perforacion.Uno + perforacion.Dos + perforacion.Tres + perforacion.Cuatro + perforacion.Cinco) / 5, 2);

            if (perforacion.Estado == "N")
            {
                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph("-").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);
            }
            else
            {
                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(perforacion.Uno))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(perforacion.Dos))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(perforacion.Tres))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(perforacion.Cuatro))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(perforacion.Cinco))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);

                resultadoFuerzaPerforacionCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(promedioPerforacion)).SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoFuerzaPerforacionTable.AddCell(resultadoFuerzaPerforacionCell);
            }

            pruebaPerforacionTable.AddCell(resultadoFuerzaPerforacionTable.SetMarginRight(19).SetMarginBottom(5).SetMarginTop(5));

            pruebaPerforacionCell = new Cell(1, 1).Add(new Paragraph(ObtenerDescripcionEstado(perforacion.Estado))).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            pruebaPerforacionTable.AddCell(pruebaPerforacionCell);

            document.Add(pruebaPerforacionTable.SetMarginBottom(12));

            document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

            document.Add(new Paragraph(new Text("\n")));

            Paragraph parafoElasticidad = new Paragraph("4. PRUEBA DE ELASTICIDAD").SetFontSize(8.95f);
            document.Add(parafoElasticidad.SetMarginBottom(6));

            parafoElasticidad = new Paragraph("Especificaciones").SetFontSize(8.95f).SetPaddingLeft(15);
            document.Add(parafoElasticidad.SetMarginBottom(4));


            Table elasticidadTable = new Table(new float[] { 15, 15, 10, 10, 10, 10, 10, 10 })
                .SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell elasticidadCell = new Cell(2, 1).Add(new Paragraph("Arco de\nla aguja")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(2, 1).Add(new Paragraph("Diametro de\nla aguja\n(mm)")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 2).Add(new Paragraph("12 mm <= L < 20 mm")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 2).Add(new Paragraph("20 mm <= L <= 40 mm")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 2).Add(new Paragraph("L > 40 mm")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Aumento del arco")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Deformación en retorno")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Aumento del arco")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Deformación en retorno")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Aumento del arco")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("Deformación en retorno")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);


            elasticidadCell = new Cell(3, 1).Add(new Paragraph("1/2 círculo\n5/8 círculo").SetTextAlignment(TextAlignment.CENTER)).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("0,2 a 0,6")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("10 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(6, 1).Add(new Paragraph("2 %").SetTextAlignment(TextAlignment.CENTER)).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("11 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(6, 1).Add(new Paragraph("1.5 %").SetTextAlignment(TextAlignment.CENTER)).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("---")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("---")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("0,7 a 0,9")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("9 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("10 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("11 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(2, 1).Add(new Paragraph("1 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);



            elasticidadCell = new Cell(1, 1).Add(new Paragraph(">= 1,0")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("8 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("9 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("10 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);


            elasticidadCell = new Cell(3, 1).Add(new Paragraph("3/8 círculo").SetTextAlignment(TextAlignment.CENTER)).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("0,2 a 0,6")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("8 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("9 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("---")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("---")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("0,7 a 0,9")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("7 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("8 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("9 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(2, 1).Add(new Paragraph("1 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph(">= 1,0")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("6 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("7 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("8 %")).AddStyle(cellTableElasticidad);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 8).Add(new Paragraph("Nota: No aplica en agujas con L (longitud) menor de 12 mm. Referencia: YY/T 0043-2016.").SetPaddingLeft(2))
                .AddStyle(cellTableElasticidad)
                .SetTextAlignment(TextAlignment.LEFT);

            elasticidadTable.AddCell(elasticidadCell);

            document.Add(elasticidadTable.SetMarginBottom(6));

            elasticidadTable = new Table(new float[] { 85, 15 });
            elasticidadTable.SetWidth(UnitValue.CreatePercentValue(100));
            elasticidadTable.SetFixedLayout();


            elasticidadCell = new Cell(1, 1).Add(new Paragraph("% DE DEFORMACIÓN EN RETORNO DE LA AGUJA")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTablePerforacion);
            elasticidadTable.AddCell(elasticidadCell);



            Table resultadoElasticidadTable = new Table(new float[] { 18, 9, 9, 9, 9, 9, 18 })
                .SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            elasticidadCell = new Cell(2, 1).Add(new Paragraph("Resultado (N):").SetFontSize(9f)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("1").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("2").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("3").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("4").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("5").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            elasticidadCell = new Cell(1, 1).Add(new Paragraph("PROMEDIO").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            resultadoElasticidadTable.AddCell(elasticidadCell);

            AnalisisAgujaElasticidadPerforacionEntity elasticidad = elasticidadPerforeacion.Find(x => x.TipoRegistro == 1);

            decimal promedioElasticidad = decimal.Round((elasticidad.Uno + elasticidad.Dos + elasticidad.Tres + elasticidad.Cuatro + elasticidad.Cinco) / 5, 2);

            if(elasticidad.Estado == "N")
            {
                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph("-").SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);
            }
            else
            {
                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(elasticidad.Uno))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(elasticidad.Dos))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(elasticidad.Tres))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(elasticidad.Cuatro))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(elasticidad.Cinco))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);

                elasticidadCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(promedioElasticidad)).SetFont(fuenteNegrita)).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
                resultadoElasticidadTable.AddCell(elasticidadCell);
            }
           

            elasticidadTable.AddCell(resultadoElasticidadTable.SetMarginRight(19).SetMarginBottom(5).SetMarginTop(5));

            elasticidadCell = new Cell(1, 1).Add(new Paragraph(ObtenerDescripcionEstado(elasticidad.Estado))).AddStyle(cellTablePerforacion).SetTextAlignment(TextAlignment.CENTER);
            elasticidadTable.AddCell(elasticidadCell);

            document.Add(elasticidadTable.SetMarginBottom(12));

            Paragraph parafoDefecto = new Paragraph("5. ANÁLISIS E ACEPTABILIDAD DE DEFECTOS DE ASPECTO DE LA AGUJA").SetFontSize(8.95f);
            document.Add(parafoDefecto.SetMarginBottom(6));

            Table defectoTable = new Table(new float[] { 15, 15, 20, 10, 10, 10, 20 })
                .SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell defectoCell = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(2, 1).Add(new Paragraph("ASPECTO DE LA\nAGUJA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(2, 1).Add(new Paragraph("AGUJERO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("Agujero centrado").SetPaddingLeft(3)).AddStyle(cellTableCommon);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity agujaCentrado = aspectoAguja.Find(x => x.TipoRegistro == 0);

            decimal porcentajeAgujaCentrado = CalculoPorcentaje(agujaCentrado.Cantidad, agujaCentrado.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(agujaCentrado.Cantidad)))
                .AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);


            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeAgujaCentrado)))
                .AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph($"> ó = {SeparadorDeMilesDecimal(agujaCentrado.Tolerancia ?? 0)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(2, 1).Add(new Paragraph(porcentajeAgujaCentrado >= (agujaCentrado.Tolerancia ?? 0) ? "ACEPTADO" : "RECHAZADO" ))
                .AddStyle(cellTableCommon)
                .SetTextAlignment(TextAlignment.CENTER);

            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity agujaDescentrado = aspectoAguja.Find(x => x.TipoRegistro == 1);

            decimal porcentajeAgujaDescentrado = CalculoPorcentaje(agujaDescentrado.Cantidad, agujaDescentrado.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph("Agujero descentrado").SetPaddingLeft(3)).AddStyle(cellTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(agujaDescentrado.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeAgujaDescentrado))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph($"< {SeparadorDeMilesDecimal(agujaDescentrado.Tolerancia ?? 0)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            document.Add(defectoTable.SetMarginBottom(10));


            defectoTable = new Table(new float[] { 15, 15, 20, 10, 10, 10, 20 })
                .SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            defectoCell = new Cell(1, 1).Add(new Paragraph("PRUEBA")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("ESPECIFICACIÓN")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("CANTIDAD")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("%")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("TOLERANCIA")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("STATUS")).AddStyle(headerTableCommon);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(9, 1).Add(new Paragraph("ASPECTO DE LA AGUJA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("ALAMBRE / AGUJERO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("BUEN ASPECTO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity buenAspecto = aspectoAguja.Find(x => x.TipoRegistro == 2);

            decimal porcentajeBuenAspecto = CalculoPorcentaje(buenAspecto.Cantidad, buenAspecto.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(buenAspecto.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeBuenAspecto))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph($"> ó = {SeparadorDeMilesDecimal(buenAspecto.Tolerancia ?? 0)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(9, 1).Add(new Paragraph(porcentajeBuenAspecto >= (buenAspecto.Tolerancia ?? 0) ? "ACEPTADO" : "RECHAZADO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(6, 1).Add(new Paragraph("ASPECTO DEL ALAMBRE")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("MARCAS")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity marca = aspectoAguja.Find(x => x.TipoRegistro == 3);

            decimal porcentajeMarca = CalculoPorcentaje(marca.Cantidad, marca.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(marca.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeMarca))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);


            defectoCell = new Cell(8, 1).Add(new Paragraph($"< {SeparadorDeMilesDecimal(100 - buenAspecto.Tolerancia ?? 0)} %")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("POROSIDAD/HUECOS")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity porosidad = aspectoAguja.Find(x => x.TipoRegistro == 4);

            decimal porcentajePorosidad = CalculoPorcentaje(porosidad.Cantidad, porosidad.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porosidad.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajePorosidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("MAL PULIDO/MANCHAS")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity malPulido = aspectoAguja.Find(x => x.TipoRegistro == 5);

            decimal porcentajeMalPulido = CalculoPorcentaje(malPulido.Cantidad, malPulido.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(malPulido.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeMalPulido))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("PUNTA UÑA DE GATO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity puntaGato = aspectoAguja.Find(x => x.TipoRegistro == 6);

            decimal porcentajePuntaGato = CalculoPorcentaje(puntaGato.Cantidad, puntaGato.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(puntaGato.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajePuntaGato))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);


            defectoCell = new Cell(1, 1).Add(new Paragraph("POCA PUNTA/PUNTA AGUDA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity pocaPunta = aspectoAguja.Find(x => x.TipoRegistro == 7);

            decimal porcentajePocaPunta = CalculoPorcentaje(pocaPunta.Cantidad, pocaPunta.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(pocaPunta.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajePocaPunta))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("MAL FORMADA")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity malFormada = aspectoAguja.Find(x => x.TipoRegistro == 8);

            decimal porcentajeMalFormada = CalculoPorcentaje(malFormada.Cantidad, malFormada.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(malFormada.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeMalFormada))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(2, 1).Add(new Paragraph("ASPECTO DEL AGUJERO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph("REBABAS")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity rebabas = aspectoAguja.Find(x => x.TipoRegistro == 9);

            decimal porcentajeBebabas = CalculoPorcentaje(rebabas.Cantidad, rebabas.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(rebabas.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeBebabas))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);


            defectoCell = new Cell(1, 1).Add(new Paragraph("SIN AGUJERO")).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            AnalisisAgujaPruebaAspectoEntity sinAgujero = aspectoAguja.Find(x => x.TipoRegistro == 10);

            decimal porcentajeSinAgujero = CalculoPorcentaje(rebabas.Cantidad, rebabas.BaseCalculoPorcentaje);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(sinAgujero.Cantidad))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            defectoCell = new Cell(1, 1).Add(new Paragraph(SeparadorDeMilesDecimal(porcentajeSinAgujero))).AddStyle(cellTableCommon).SetTextAlignment(TextAlignment.CENTER);
            defectoTable.AddCell(defectoCell);

            document.Add(defectoTable.SetMarginBottom(10));

            Table observacionesTable = new Table(1).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell observacionesCell = new Cell(1, 1).Add(new Paragraph("OBSERVACIONES:")).AddStyle(headerTableCommon).SetTextAlignment(TextAlignment.LEFT);
            observacionesTable.AddCell(observacionesCell);

            observacionesCell = new Cell(1, 1).Add(new Paragraph(datosGenerales.Observaciones).SetFontSize(6.5f).SetPaddingLeft(2)).AddStyle(cellTableCommon);
            observacionesTable.AddCell(observacionesCell);

            document.Add(observacionesTable.SetMarginBottom(10));


            Paragraph conclusiones = new Paragraph("CONCLUSIÓN:").SetFontSize(8);
            document.Add(conclusiones.SetMarginBottom(10));


            Table conclusionesTable = new Table(new float[] { 20, 10, 20, 10, 20, 10 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell conclusionessCell = new Cell(1, 1).Add(new Paragraph("Aprobado")).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            conclusionesTable.AddCell(conclusionessCell);

            conclusionessCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(cellTableCommon);
            conclusionesTable.AddCell(conclusionessCell);


            conclusionessCell = new Cell(1, 1).Add(new Paragraph("Rechazado")).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            conclusionesTable.AddCell(conclusionessCell);

            conclusionessCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(cellTableCommon);
            conclusionesTable.AddCell(conclusionessCell);

            conclusionessCell = new Cell(1, 1).Add(new Paragraph("SELECCIONAR AL 100%")).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            conclusionesTable.AddCell(conclusionessCell);

            conclusionessCell = new Cell(1, 1).Add(new Paragraph("")).AddStyle(cellTableCommon);
            conclusionesTable.AddCell(conclusionessCell);

            document.Add(conclusionesTable.SetMarginBottom(70));

            Table responsablesTable = new Table(new float[] { 50, 50 }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout();

            Cell responsableCelda = new Cell(1, 1).Add(new Paragraph("ANALISTA: _________________________")).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            responsablesTable.AddCell(responsableCelda);

            responsableCelda = new Cell(1, 1).Add(new Paragraph("JEFE CONTROL DE CALIDAD: _________________________")).AddStyle(cellTableCommon).SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.CENTER);
            responsablesTable.AddCell(responsableCelda);

            document.Add(responsablesTable);

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

        private string ObtenerDescripcionEstado(string codigoEstado)
        {
            switch (codigoEstado)
            {
                case "A":
                    return "ACEPTADO";
                case "R":
                    return "RECHAZADO";
                case "N":
                    return "NO APLICA";
                default:
                    return "";
            }
        }

        private string SeparadorDeMilesDecimal(decimal numero)=> string.Format("{0:##,###,##0.##}", numero);
        

        private string SeparadorDeMilesDouble(double numero) => string.Format("{0:##,###,##0.##}", numero);
        

        private decimal CalculoPorcentaje (int cantidad, int baseCalculo) => Convert.ToDecimal(cantidad / (baseCalculo * 1.0) * 100);
        
    }

    public class FooterAnalisisAgujaEventHandler : IEventHandler
    {
        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent documentoEvento = (PdfDocumentEvent)@event;
            PdfDocument pdf = documentoEvento.GetDocument();
            PdfPage pagina = documentoEvento.GetPage();
            PdfCanvas pdfCanvas = new PdfCanvas(pagina.NewContentStreamBefore(), pagina.GetResources(), pdf);


            int paginaActual = documentoEvento.GetDocument().GetPageNumber(pagina);
            int totalPaginas = documentoEvento.GetDocument().GetNumberOfPages();


            PdfFont fuenteNegrita = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            Style estiloFooter = new Style().SetFontSize(8)
                    .SetFont(fuenteNegrita)
                    .SetFontColor(ColorConstants.BLACK)
                    .SetMargin(0)
                    .SetPadding(0)
                    .SetFontSize(8);

            Table tablaResult = new Table(new float[] { 50f, 50f }).SetWidth(UnitValue.CreatePercentValue(100)).SetFixedLayout().SetMargin(0).SetPadding(0);

            Cell footer = new Cell(1, 1).Add(new Paragraph("F/CDC-032, Versión 09").AddStyle(estiloFooter))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.LEFT).SetMargin(0).SetPadding(0);

            tablaResult.AddCell(footer);

            footer = new Cell(1, 1).Add(new Paragraph($"Página {paginaActual} de {totalPaginas}").AddStyle(estiloFooter))
                .SetBorder(Border.NO_BORDER).SetTextAlignment(TextAlignment.RIGHT).SetMargin(0).SetPadding(0);

            tablaResult.AddCell(footer);

            Rectangle rectangulo = new Rectangle(15, -10, pagina.GetPageSize().GetWidth() - 70, 50);

            new Canvas(pdfCanvas, rectangulo).Add(tablaResult);

        }
    }
}
