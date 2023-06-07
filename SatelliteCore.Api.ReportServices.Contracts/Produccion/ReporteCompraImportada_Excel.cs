using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.CompraImportacion;
using SatelliteCore.Api.Models.Response.HistorialPeriodo;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.Produccion
{
    public class ReporteCompraImportada_Excel
    {
        public string GenerarReporte(IEnumerable<DatosFormatoLitadoSeguimientoCompraImportada> listadoImportada, string mostrarColumna , IEnumerable<DatosFormatoLitadoSeguimientoCompraImportada> listadoNacional, string reporteArima, IEnumerable<DatosFormatoListadoCommodity> listadoCommodity,  DatosFormatoSeguimientoHistorioPeriodo informacion)
        {

            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                if (reporteArima == "importacion")
                {
                    CuerpoExcelImportadaNacional(excelPackage, listadoImportada, mostrarColumna, "Importación");
                }
                else if (reporteArima == "nacional")
                {
                    CuerpoExcelImportadaNacional(excelPackage, listadoNacional, mostrarColumna, "Nacional");
                    CuerpoExcelCommodity(excelPackage, listadoCommodity, mostrarColumna);
                }
                else
                {
                    CuerpoExcelImportadaNacional(excelPackage, listadoImportada, mostrarColumna, "Importación");
                    CuerpoExcelImportadaNacional(excelPackage, listadoNacional, mostrarColumna, "Nacional");
                    CuerpoExcelCommodity(excelPackage, listadoCommodity, mostrarColumna);
                }


                cuerpoPeriodoHistorico(excelPackage, informacion);

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length); 

                return reporte;
            }

        }

        private static void CuerpoExcelImportadaNacional(ExcelPackage excelPackage, IEnumerable<DatosFormatoLitadoSeguimientoCompraImportada> listado, string mostrarColumna, string tipo)
        {

       
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(tipo);
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.View.ZoomScale = 90;

            ConfigurarTamanioDeCeldas(worksheet);

            int fila = 1;

            PintarCabecera(worksheet, fila);

            worksheet.Cells["A" + fila].Value = "ITEMFINAL";
            worksheet.Cells["A" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A" + fila].Style.WrapText = true;
            worksheet.Cells["A" + fila].Style.Font.Bold = true;
            worksheet.Cells["A" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A" + fila].Style.Font.Size = 9;

            worksheet.Cells["B" + fila].Value = "DESCRIPCION";
            worksheet.Cells["B" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B" + fila].Style.WrapText = true;
            worksheet.Cells["B" + fila].Style.Font.Bold = true;
            worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["B" + fila].Style.Font.Size = 9;


            worksheet.Cells["C" + fila].Value = "Clasificacion General";
            worksheet.Cells["C" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["C" + fila].Style.WrapText = true;
            worksheet.Cells["C" + fila].Style.Font.Bold = true;
            worksheet.Cells["C" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["C" + fila].Style.Font.Size = 9;

            worksheet.Cells["D" + fila].Value = "Familia Larga";
            worksheet.Cells["D" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["D" + fila].Style.WrapText = true;
            worksheet.Cells["D" + fila].Style.Font.Bold = true;
            worksheet.Cells["D" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["D" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["D" + fila].Style.Font.Size = 9;

            worksheet.Cells["E" + fila].Value = "Puerto";
            worksheet.Cells["E" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["E" + fila].Style.WrapText = true;
            worksheet.Cells["E" + fila].Style.Font.Bold = true;
            worksheet.Cells["E" + fila].Style.Font.Size = 9;
            worksheet.Cells["E" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["E" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells["F" + fila].Value = "Proveedor";
            worksheet.Cells["F" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["F" + fila].Style.WrapText = true;
            worksheet.Cells["F" + fila].Style.Font.Bold = true;
            worksheet.Cells["F" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["F" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["F" + fila].Style.Font.Size = 9;

            worksheet.Cells["G" + fila].Value = "Tiempo en la gestión de compras  (cotización y aceptación)";
            worksheet.Cells["G" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["G" + fila].Style.WrapText = true;
            worksheet.Cells["G" + fila].Style.Font.Bold = true;
            worksheet.Cells["G" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["G" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["G" + fila].Style.Font.Size = 9;

            worksheet.Cells["H" + fila].Value = "Tiempo en gestión pago contable al proveedor";
            worksheet.Cells["H" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["H" + fila].Style.WrapText = true;
            worksheet.Cells["H" + fila].Style.Font.Bold = true;
            worksheet.Cells["H" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["H" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["H" + fila].Style.Font.Size = 9;

            worksheet.Cells["I" + fila].Value = "Tiempo en aprobación de artes Calidad";
            worksheet.Cells["I" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["I" + fila].Style.WrapText = true;
            worksheet.Cells["I" + fila].Style.Font.Bold = true;
            worksheet.Cells["I" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["I" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["I" + fila].Style.Font.Size = 9;

            worksheet.Cells["J" + fila].Value = "Tiempo de fabricación del proveedor";
            worksheet.Cells["J" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["J" + fila].Style.WrapText = true;
            worksheet.Cells["J" + fila].Style.Font.Bold = true;
            worksheet.Cells["J" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["J" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["J" + fila].Style.Font.Size = 9;

            worksheet.Cells["K" + fila].Value = "Tiempo de transporte";
            worksheet.Cells["K" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["K" + fila].Style.WrapText = true;
            worksheet.Cells["K" + fila].Style.Font.Bold = true;
            worksheet.Cells["K" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["K" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["K" + fila].Style.Font.Size = 9;

            worksheet.Cells["L" + fila].Value = "Tiempo de nacionalizar aduanas";
            worksheet.Cells["L" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["L" + fila].Style.WrapText = true;
            worksheet.Cells["L" + fila].Style.Font.Bold = true;
            worksheet.Cells["L" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["L" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["L" + fila].Style.Font.Size = 9;

            worksheet.Cells["M" + fila].Value = "Cantidad Minima";
            worksheet.Cells["M" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["M" + fila].Style.WrapText = true;
            worksheet.Cells["M" + fila].Style.Font.Bold = true;
            worksheet.Cells["M" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["M" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["M" + fila].Style.Font.Size = 9;

            worksheet.Cells["N" + fila].Value = "Promedio de consumo mensual que da el arima";
            worksheet.Cells["N" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["N" + fila].Style.WrapText = true;
            worksheet.Cells["N" + fila].Style.Font.Bold = true;
            worksheet.Cells["N" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["N" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["N" + fila].Style.Font.Size = 9;

            worksheet.Cells["O" + fila].Value = "Meses de duración del stock de la materia prima";
            worksheet.Cells["O" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["O" + fila].Style.WrapText = true;
            worksheet.Cells["O" + fila].Style.Font.Bold = true;
            worksheet.Cells["O" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["O" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["O" + fila].Style.Font.Size = 9;

            worksheet.Cells["P" + fila].Value = "Cantidad comprar";
            worksheet.Cells["P" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["P" + fila].Style.WrapText = true;
            worksheet.Cells["P" + fila].Style.Font.Bold = true;
            worksheet.Cells["P" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["P" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["P" + fila].Style.Font.Size = 9;

            worksheet.Cells["Q" + fila].Value = "Stock \n Disponible \n - \n Planta";
            worksheet.Cells["Q" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["Q" + fila].Style.WrapText = true;
            worksheet.Cells["Q" + fila].Style.Font.Bold = true;
            worksheet.Cells["Q" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["Q" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["Q" + fila].Style.Font.Size = 9;

            worksheet.Cells["R" + fila].Value = "Stock disponible";
            worksheet.Cells["R" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["R" + fila].Style.WrapText = true;
            worksheet.Cells["R" + fila].Style.Font.Bold = true;
            worksheet.Cells["R" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["R" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["R" + fila].Style.Font.Size = 9;

            worksheet.Cells["S" + fila].Value = "Orden compra preparación";
            worksheet.Cells["S" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["S" + fila].Style.WrapText = true;
            worksheet.Cells["S" + fila].Style.Font.Bold = true;
            worksheet.Cells["S" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["S" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["S" + fila].Style.Font.Size = 9;

            worksheet.Cells["T" + fila].Value = "Producto comprado en transito y/o aprobado";
            worksheet.Cells["T" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["T" + fila].Style.WrapText = true;
            worksheet.Cells["T" + fila].Style.Font.Bold = true;
            worksheet.Cells["T" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["T" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["T" + fila].Style.Font.Size = 9;

            worksheet.Cells["U" + fila].Value = "Producto aduana";
            worksheet.Cells["U" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["U" + fila].Style.WrapText = true;
            worksheet.Cells["U" + fila].Style.Font.Bold = true;
            worksheet.Cells["U" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["U" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["U" + fila].Style.Font.Size = 9;

            worksheet.Cells["V" + fila].Value = "Producto control calidad";
            worksheet.Cells["V" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["V" + fila].Style.WrapText = true;
            worksheet.Cells["V" + fila].Style.Font.Bold = true;
            worksheet.Cells["V" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["V" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["V" + fila].Style.Font.Size = 9;

            worksheet.Cells["W" + fila].Value = "Stock calculado";
            worksheet.Cells["W" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["W" + fila].Style.WrapText = true;
            worksheet.Cells["W" + fila].Style.Font.Bold = true;
            worksheet.Cells["W" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["W" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["W" + fila].Style.Font.Size = 9;

            worksheet.Cells["X" + fila].Value = "Meses potenciales con stock futuro";
            worksheet.Cells["X" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["X" + fila].Style.WrapText = true;
            worksheet.Cells["X" + fila].Style.Font.Bold = true;
            worksheet.Cells["X" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["X" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["X" + fila].Style.Font.Size = 9;

            worksheet.Cells["Y" + fila].Value = "Dias potenciales con stock futuro";
            worksheet.Cells["Y" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["Y" + fila].Style.WrapText = true;
            worksheet.Cells["Y" + fila].Style.Font.Bold = true;
            worksheet.Cells["Y" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["Y" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["Y" + fila].Style.Font.Size = 9;

            worksheet.Cells["Z" + fila].Value = "Promedio de consumo x dia";
            worksheet.Cells["Z" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["Z" + fila].Style.WrapText = true;
            worksheet.Cells["Z" + fila].Style.Font.Bold = true;
            worksheet.Cells["Z" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["Z" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["Z" + fila].Style.Font.Size = 9;

            worksheet.Cells["AA" + fila].Value = "Dias de demora en llegada de producto";
            worksheet.Cells["AA" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AA" + fila].Style.WrapText = true;
            worksheet.Cells["AA" + fila].Style.Font.Bold = true;
            worksheet.Cells["AA" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AA" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AA" + fila].Style.Font.Size = 9;

            worksheet.Cells["AB" + fila].Value = "Dias de stock de precaucion x desviacion on sobrecompra";
            worksheet.Cells["AB" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AB" + fila].Style.Font.Bold = true;
            worksheet.Cells["AB" + fila].Style.WrapText = true;
            worksheet.Cells["AB" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AB" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AB" + fila].Style.Font.Size = 9;

            worksheet.Cells["AC" + fila].Value = "Dias que puedo esperar para comprar";
            worksheet.Cells["AC" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AC" + fila].Style.Font.Bold = true;
            worksheet.Cells["AC" + fila].Style.WrapText = true;
            worksheet.Cells["AC" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AC" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AC" + fila].Style.Font.Size = 9;

            worksheet.Cells["AD" + fila].Value = "Coeficiente de variación";
            worksheet.Cells["AD" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AD" + fila].Style.Font.Bold = true;
            worksheet.Cells["AD" + fila].Style.WrapText = true;
            worksheet.Cells["AD" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AD" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AD" + fila].Style.Font.Size = 9;

            worksheet.Cells["AE" + fila].Value = "qty maxima deberia haber en stock linea";
            worksheet.Cells["AE" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AE" + fila].Style.Font.Bold = true;
            worksheet.Cells["AE" + fila].Style.WrapText = true;
            worksheet.Cells["AE" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AE" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AE" + fila].Style.Font.Size = 9;

            worksheet.Cells["AF" + fila].Value = "Punto de corte de stock donde se debe comprar";
            worksheet.Cells["AF" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["AF" + fila].Style.Font.Bold = true;
            worksheet.Cells["AF" + fila].Style.WrapText = true;
            worksheet.Cells["AF" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["AF" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["AF" + fila].Style.Font.Size = 9;


            worksheet.View.FreezePanes(fila + 1, 3);

            int row = fila + 1;

            foreach (DatosFormatoLitadoSeguimientoCompraImportada rowitem in listado)
            {
                worksheet.Row(row).Height = 14.25;

                worksheet.Cells["A" + row].Value = rowitem.ItemFinal;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["A" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.GestionLetras));


                worksheet.Cells["B" + row].Value = rowitem.DescripcionLocal;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(rowitem.GestionColor));
                worksheet.Cells["B" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.GestionLetras));

                worksheet.Cells["C" + row].Value = rowitem.Clasificacion;
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["D" + row].Value = rowitem.FamiliaLarga;
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["E" + row].Value = "";
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["F" + row].Value = "";
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["G" + row].Value = rowitem.TiempoCompra;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["H" + row].Value = rowitem.TiempoPago;
                worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["H" + row].Style.WrapText = true;
                worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["I" + row].Value = rowitem.TiempoAprobacion;
                worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["I" + row].Style.WrapText = true;
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["J" + row].Value = rowitem.Tiempofabricacion;
                worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["J" + row].Style.WrapText = true;
                worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["K" + row].Value = rowitem.TiempoTransporte;
                worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["K" + row].Style.WrapText = true;
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["L" + row].Value = rowitem.TiempoAduanas;
                worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["M" + row].Value = rowitem.CantidadMinima;
                worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["M" + row].Style.WrapText = true;
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["N" + row].Value = rowitem.Pronostico;
                worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["N" + row].Style.WrapText = true;
                worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["O" + row].Value = rowitem.Duracion;
                worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["O" + row].Style.WrapText = true;
                worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["P" + row].Value = rowitem.CantidadComprar;
                worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["P" + row].Style.WrapText = true;
                worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Q" + row].Value = rowitem.AlmacenMenosPlanta;
                worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["Q" + row].Style.WrapText = true;
                worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                //
                worksheet.Cells["R" + row].Value = rowitem.Almacen;
                worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["R" + row].Style.WrapText = true;
                worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                worksheet.Cells["S" + row].Value = rowitem.PreparacionOC;
                worksheet.Cells["S" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["S" + row].Style.WrapText = true;
                worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["T" + row].Value = rowitem.PendienteOC;
                worksheet.Cells["T" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["T" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["T" + row].Style.WrapText = true;
                worksheet.Cells["T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["U" + row].Value = rowitem.Aduanas;
                worksheet.Cells["U" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["U" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["U" + row].Style.WrapText = true;
                worksheet.Cells["U" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["V" + row].Value = rowitem.ControlCalidad;
                worksheet.Cells["V" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["V" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["V" + row].Style.WrapText = true;
                worksheet.Cells["V" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["W" + row].Value = rowitem.Disponible;
                worksheet.Cells["W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["W" + row].Style.WrapText = true;
                worksheet.Cells["W" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["X" + row].Value = rowitem.MesesPotencialesFuturo;
                worksheet.Cells["X" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["X" + row].Style.WrapText = true;
                worksheet.Cells["X" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Y" + row].Value = rowitem.DiasPotencial;
                worksheet.Cells["Y" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Y" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["Y" + row].Style.WrapText = true;
                worksheet.Cells["Y" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Z" + row].Value = rowitem.ConsumoDia;
                worksheet.Cells["Z" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Z" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["Z" + row].Style.WrapText = true;
                worksheet.Cells["Z" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AA" + row].Value = rowitem.DemoraLlegarProducto;
                worksheet.Cells["AA" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AA" + row].Style.Numberformat.Format = "#,##0.0";
                worksheet.Cells["AA" + row].Style.WrapText = true;
                worksheet.Cells["AA" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AB" + row].Value = rowitem.DesviacionCompra;
                worksheet.Cells["AB" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AB" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["AB" + row].Style.WrapText = true;
                worksheet.Cells["AB" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AC" + row].Value = rowitem.DiasEspera;
                worksheet.Cells["AC" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AC" + row].Style.Font.Color.SetColor(rowitem.DiasEspera >= 0 ? Color.Black : Color.Red);
                worksheet.Cells["AC" + row].Style.Numberformat.Format = "#,##0.0";
                worksheet.Cells["AC" + row].Style.WrapText = true;
                worksheet.Cells["AC" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AD" + row].Value = rowitem.Variacion;
                worksheet.Cells["AD" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AD" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["AD" + row].Style.WrapText = true;
                worksheet.Cells["AD" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AE" + row].Value = rowitem.MaximoStock;
                worksheet.Cells["AE" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AE" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["AE" + row].Style.WrapText = true;
                worksheet.Cells["AE" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AF" + row].Value = rowitem.PuntoCorte;
                worksheet.Cells["AF" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AF" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["AF" + row].Style.WrapText = true;
                worksheet.Cells["AF" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                row++;
            }

            if(tipo == "Nacional")
            {
                worksheet.Column(3).Hidden = true;
                worksheet.Column(4).Hidden = true;
            }

               
            if (mostrarColumna == "true")
            {
                worksheet.Column(5).Hidden = false;
                worksheet.Column(6).Hidden = false;
                worksheet.Column(7).Hidden = false;
                worksheet.Column(8).Hidden = false;
                worksheet.Column(9).Hidden = false;
                worksheet.Column(10).Hidden = false;
                worksheet.Column(11).Hidden = false;
                worksheet.Column(12).Hidden = false;
                worksheet.Column(13).Hidden = false;
                worksheet.Column(25).Hidden = false;
                worksheet.Column(26).Hidden = false;
                worksheet.Column(27).Hidden = false;
                worksheet.Column(28).Hidden = false;
                worksheet.Column(31).Hidden = false;
                worksheet.Column(32).Hidden = false;


            }
            else
            {
                worksheet.Column(5).Hidden = true;
                worksheet.Column(6).Hidden = true;
                worksheet.Column(7).Hidden = true;
                worksheet.Column(8).Hidden = true;
                worksheet.Column(9).Hidden = true;
                worksheet.Column(10).Hidden = true;
                worksheet.Column(11).Hidden = true;
                worksheet.Column(12).Hidden = true;
                worksheet.Column(13).Hidden = true;
                worksheet.Column(25).Hidden = true;
                worksheet.Column(26).Hidden = true;
                worksheet.Column(27).Hidden = true;
                worksheet.Column(28).Hidden = true;
                worksheet.Column(31).Hidden = true;
                worksheet.Column(32).Hidden = true;
            }
        }


        private static void CuerpoExcelCommodity(ExcelPackage excelPackage, IEnumerable<DatosFormatoListadoCommodity> listado, string mostrarColumna)
        {


            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Commodity");
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.View.ZoomScale = 90;

            ConfigurarTamanioDeCeldasCommodity(worksheet);

            int fila = 1;

            PintarCabeceraCommodity(worksheet, fila);


            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.WrapText = true;
            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{fila},B{fila},C{fila},D{fila},E{fila},F{fila},G{fila},H{fila},I{fila},J{fila},K{fila},L{fila},M{fila},N{fila},O{fila},P{fila},Q{fila},R{fila},S{fila},T{fila},U{fila},V{fila},W{fila},X{fila},Y{fila},Z{fila},AA{fila},AB{fila}"].Style.Font.Size = 9;

            worksheet.Cells["A" + fila].Value = "ITEMFINAL";
            
            worksheet.Cells["B" + fila].Value = "DESCRIPCION";

            worksheet.Cells["C" + fila].Value = "Clasificacion General";
            
            worksheet.Cells["D" + fila].Value = "Familia Larga";
            
            worksheet.Cells["E" + fila].Value = "Puerto";
            
            worksheet.Cells["F" + fila].Value = "Proveedor";
            
            worksheet.Cells["G" + fila].Value = "tiempo en la gestión de compras  (cotización y aceptación)";
            
            worksheet.Cells["H" + fila].Value = "tiempo en gestión pago contable al proveedor";
        
            worksheet.Cells["I" + fila].Value = "tiempo en aprobación de artes Calidad";
            
            worksheet.Cells["J" + fila].Value = "tiempo de fabricación del proveedor";
            
            worksheet.Cells["K" + fila].Value = "tiempo de transporte";
            
            worksheet.Cells["L" + fila].Value = "tiempo de nacionalizar aduanas";
            
            worksheet.Cells["M" + fila].Value = "Cantidad Minima";

            worksheet.Cells["N" + fila].Value = "qty maxima deberia haber en stock linea";

            worksheet.Cells["O" + fila].Value = "punto de corte de stock donde se debe comprar";

            worksheet.Cells["P" + fila].Value = "Meses Duracion";

            worksheet.Cells["Q" + fila].Value = "stock fisico DISPONIBLE  \n(+)";

            worksheet.Cells["R" + fila].Value = "Orden Compra Preparación \n(+)";
            
            worksheet.Cells["S" + fila].Value = "producto comprado en transito y/o Aprobado \n(+)";

            worksheet.Cells["T" + fila].Value = "Stock calculado";

            worksheet.Cells["U" + fila].Value = "Dias potenciales con stock futuro";

            worksheet.Cells["V" + fila].Value = "promedio de consumo mensual que da el arima ";

            worksheet.Cells["W" + fila].Value = "promedio de consumo x dia";
            
            worksheet.Cells["X" + fila].Value = "dias de demora en llegada de producto";

            worksheet.Cells["Y" + fila].Value = "dias de stock de precaucion x desviacion sobrecompra";           

            worksheet.Cells["Z" + fila].Value = "dias que puedo esperar para comprar";         

            worksheet.Cells["AA" + fila].Value = "coeficiente de variación";

            worksheet.Cells["AB" + fila].Value = "Cantidad Comprar";


            worksheet.View.FreezePanes(fila + 1, 3);

            int row = fila + 1;

            foreach (DatosFormatoListadoCommodity rowitem in listado)
            {
                worksheet.Row(row).Height = 14.25;

                worksheet.Cells["A" + row].Value = rowitem.ItemFinal;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["A" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.GestionLetras));

                worksheet.Cells["B" + row].Value = rowitem.DescripcionLocal;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                worksheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(rowitem.GestionColor));
                worksheet.Cells["B" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.GestionLetras));


                worksheet.Cells["C" + row].Value = rowitem.Clasificacion;
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["D" + row].Value = rowitem.FamiliaLarga;
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["E" + row].Value = "";
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["F" + row].Value = "";
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                worksheet.Cells["G" + row].Value = rowitem.TiempoCompra;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["H" + row].Value = rowitem.TiempoPago;
                worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["H" + row].Style.WrapText = true;
                worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["I" + row].Value = rowitem.TiempoAprobacion;
                worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["I" + row].Style.WrapText = true;
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["J" + row].Value = rowitem.Tiempofabricacion;
                worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["J" + row].Style.WrapText = true;
                worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["K" + row].Value = rowitem.TiempoTransporte;
                worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["K" + row].Style.WrapText = true;
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["L" + row].Value = rowitem.TiempoAduanas;
                worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["M" + row].Value = rowitem.CantidadMinima;
                worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["M" + row].Style.WrapText = true;
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["N" + row].Value = rowitem.MaximoStock;
                worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["N" + row].Style.WrapText = true;
                worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["O" + row].Value = rowitem.PuntoCorte;
                worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["O" + row].Style.WrapText = true;
                worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["P" + row].Value = rowitem.Duracion;
                worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["P" + row].Style.WrapText = true;
                worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Q" + row].Value = rowitem.Almacen;
                worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["Q" + row].Style.WrapText = true;
                worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                //
                worksheet.Cells["R" + row].Value = rowitem.PreparacionOC;
                worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["R" + row].Style.WrapText = true;
                worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                worksheet.Cells["S" + row].Value = rowitem.PendienteOC;
                worksheet.Cells["S" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["S" + row].Style.WrapText = true;
                worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["T" + row].Value = rowitem.Disponible;
                worksheet.Cells["T" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["T" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["T" + row].Style.WrapText = true;
                worksheet.Cells["T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["U" + row].Value = rowitem.DiasPotencial;
                worksheet.Cells["U" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["U" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["U" + row].Style.WrapText = true;
                worksheet.Cells["U" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["V" + row].Value = rowitem.Pronostico;
                worksheet.Cells["V" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["V" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["V" + row].Style.WrapText = true;
                worksheet.Cells["V" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["W" + row].Value = rowitem.ConsumoDia;
                worksheet.Cells["W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["W" + row].Style.WrapText = true;
                worksheet.Cells["W" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["X" + row].Value = rowitem.DemoraLlegarProducto;
                worksheet.Cells["X" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["X" + row].Style.WrapText = true;
                worksheet.Cells["X" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Y" + row].Value = rowitem.DesviacionCompra;
                worksheet.Cells["Y" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Y" + row].Style.Numberformat.Format = "#,##0.0";
                worksheet.Cells["Y" + row].Style.WrapText = true;
                worksheet.Cells["Y" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["Z" + row].Value = rowitem.DiasEspera;
                worksheet.Cells["Z" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Z" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["Z" + row].Style.Font.Color.SetColor(rowitem.DiasEspera >= 0 ? Color.Black : Color.Red);
                worksheet.Cells["Z" + row].Style.WrapText = true;
                worksheet.Cells["Z" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AA" + row].Value = rowitem.Variacion;
                worksheet.Cells["AA" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AA" + row].Style.Numberformat.Format = "#,##0.0";
                worksheet.Cells["AA" + row].Style.WrapText = true;
                worksheet.Cells["AA" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["AB" + row].Value = rowitem.CantidadComprar;
                worksheet.Cells["AB" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AB" + row].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["AB" + row].Style.WrapText = true;
                worksheet.Cells["AB" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                row++;
            }


            if (mostrarColumna == "true")
            {
                worksheet.Column(3).Hidden = false;
                worksheet.Column(4).Hidden = false;
                worksheet.Column(5).Hidden = false;
                worksheet.Column(6).Hidden = false;
                worksheet.Column(7).Hidden = false;
                worksheet.Column(8).Hidden = false;
                worksheet.Column(9).Hidden = false;
                worksheet.Column(10).Hidden = false;
                worksheet.Column(11).Hidden = false;
                worksheet.Column(12).Hidden = false;
                worksheet.Column(13).Hidden = false;
                worksheet.Column(14).Hidden = false;
                worksheet.Column(15).Hidden = false;
                worksheet.Column(21).Hidden = false;
                worksheet.Column(23).Hidden = false;
                worksheet.Column(24).Hidden = false;

            }
            else
            {   

                worksheet.Column(3).Hidden = true;
                worksheet.Column(4).Hidden = true;
                worksheet.Column(5).Hidden = true;
                worksheet.Column(6).Hidden = true;
                worksheet.Column(7).Hidden  = true;
                worksheet.Column(8).Hidden = true;
                worksheet.Column(9).Hidden = true;
                worksheet.Column(10).Hidden = true;
                worksheet.Column(11).Hidden = true;
                worksheet.Column(12).Hidden = true;
                worksheet.Column(13).Hidden = true;
                worksheet.Column(14).Hidden = true;
                worksheet.Column(15).Hidden = true;
                worksheet.Column(21).Hidden = true;
                worksheet.Column(23).Hidden = true;
                worksheet.Column(24).Hidden = true;

            }
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 3.71 + 2.71;
            worksheet.Column(2).Width = 32.43 + 2.71;
            worksheet.Column(3).Width = 5.71 + 2.71;
            worksheet.Column(4).Width = 17.43 + 2.71;
            worksheet.Column(5).Width = 7.71 + 2.71;
            worksheet.Column(6).Width = 7 + 2.71;
            worksheet.Column(7).Width = 7.71 + 2.71;
            worksheet.Column(8).Width = 7.71 + 2.71;
            worksheet.Column(9).Width = 7.71 + 2.71;
            worksheet.Column(10).Width = 7.71 + 2.71;
            worksheet.Column(11).Width = 7.71 + 2.71;
            worksheet.Column(12).Width = 7.71 + 2.71;
            worksheet.Column(13).Width = 6.71 + 2.71;
            worksheet.Column(14).Width = 7.71 + 2.71;
            worksheet.Column(15).Width = 7.71 + 2.71;
            worksheet.Column(16).Width = 10.29 + 2.71;
            worksheet.Column(17).Width = 10.71 + 2.71;
            worksheet.Column(18).Width = 9.71 + 2.71;
            worksheet.Column(19).Width = 9.71 + 2.71;
            worksheet.Column(20).Width = 9.71 + 2.71;
            worksheet.Column(22).Width = 9.71 + 2.71;
            worksheet.Column(23).Width = 7.71 + 2.71;
            worksheet.Column(24).Width = 10.71 + 2.71;
            worksheet.Column(25).Width = 7.71 + 2.71;
            worksheet.Column(26).Width = 7.71 + 2.71;
            worksheet.Column(27).Width = 7.71 + 2.71;
            worksheet.Column(28).Width = 7.71 + 2.71;
            worksheet.Column(29).Width = 7.71 + 2.71;
            worksheet.Column(30).Width = 7.71 + 2.71;

        }

        private static void ConfigurarTamanioDeCeldasCommodity(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 3.71 + 2.71;
            worksheet.Column(2).Width = 32.43 + 2.71;
            worksheet.Column(3).Width = 5.71 + 2.71;
            worksheet.Column(4).Width = 18.00 + 2.71;
            worksheet.Column(5).Width = 7.71 + 2.71;
            worksheet.Column(6).Width = 7 + 2.71;
            worksheet.Column(7).Width = 7.71 + 2.71;
            worksheet.Column(8).Width = 7.71 + 2.71;
            worksheet.Column(9).Width = 7.71 + 2.71;
            worksheet.Column(10).Width = 7.71 + 2.71;
            worksheet.Column(11).Width = 7.71 + 2.71;
            worksheet.Column(12).Width = 7.71 + 2.71;
            worksheet.Column(13).Width = 6.71 + 2.71;
            worksheet.Column(14).Width = 11.43 + 2.71;
            worksheet.Column(15).Width = 10.71 + 2.71;
            worksheet.Column(16).Width = 9.29 + 2.71;
            worksheet.Column(17).Width = 9.86 + 2.71;
            worksheet.Column(18).Width = 9.71 + 2.71;
            worksheet.Column(19).Width = 7.71 + 2.71;
            worksheet.Column(20).Width = 11.71 + 2.71;
            worksheet.Column(22).Width = 10.71 + 2.71;
            worksheet.Column(23).Width = 7.71 + 2.71;
            worksheet.Column(24).Width = 7.71 + 2.71;
            worksheet.Column(25).Width = 7.71 + 2.71;
            worksheet.Column(26).Width = 7.71 + 2.71;
            worksheet.Column(27).Width = 7.71 + 2.71;
            worksheet.Column(28).Width = 7.71 + 2.71;

        }

        private static void PintarCabeceraCommodity(ExcelWorksheet worksheet, int fila)
        {
            worksheet.Cells["A" + fila + ":F" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["G" + fila + ":L" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFC000"));
            worksheet.Cells["M" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d1d1d1"));
            worksheet.Cells["N" + fila + ":O" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["Q" + fila + ":T" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
            worksheet.Cells["U" + fila + ":Z" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["AB" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
        }

        private static void PintarCabecera(ExcelWorksheet worksheet, int fila)
        {
            worksheet.Cells["A" + fila + ":F" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["G" + fila + ":L" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFC000"));
            worksheet.Cells["M" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d1d1d1"));
            worksheet.Cells["N"+ fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF86E"));
            worksheet.Cells["P" + fila ].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#EA987B"));
            worksheet.Cells["Q" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ACF0D1"));
            worksheet.Cells["R" + fila + ":W" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFDA9E"));
            worksheet.Cells["X" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#96BCE5"));
            worksheet.Cells["Y" + fila + ":AB" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["AC" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF86E"));
            worksheet.Cells["AE" + fila + ":AF" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
        }

        private static void cuerpoPeriodoHistorico(ExcelPackage excelPackage, DatosFormatoSeguimientoHistorioPeriodo informacion)
        {

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Consumo Historico");
            worksheet.Cells.Style.Font.Name = "Arial";
            worksheet.Cells.Style.Font.Size = 10;
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
            worksheet.View.ZoomScale = 90;

            ConfigurarTamanioDeCeldasHistorico(worksheet);
            PintarCabeceraHistorico(worksheet);
            worksheet.View.FreezePanes(1, 3);

            worksheet.Cells["A1"].Value = "ITEMFINAL";
            worksheet.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["A1"].Style.WrapText = true;
            worksheet.Cells["A1"].Style.Font.Bold = true;
            worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A1"].Style.Font.Size = 9;

            worksheet.Cells["B1"].Value = "DESCRIPCION ITEM";
            worksheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B1"].Style.WrapText = true;
            worksheet.Cells["B1"].Style.Font.Bold = true;
            worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["B1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            int columna = 3;
            foreach (DatosFormatoPeriodo bloquecolumna in informacion.Periodo)
            {
                worksheet.Cells[1, columna].Value = bloquecolumna.Periodo;
                worksheet.Cells[1, columna].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[1, columna].Style.WrapText = true;
                worksheet.Cells[1, columna].Style.Font.Bold = true;
                worksheet.Cells[1, columna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, columna].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                columna++;
            }

            worksheet.Cells["O1"].Value = "TOTAL GENERAL";
            worksheet.Cells["O1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["O1"].Style.WrapText = true;
            worksheet.Cells["O1"].Style.Font.Bold = true;
            worksheet.Cells["O1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["O1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["O1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            int row = 2;
            foreach (DatosFormatoReporteHistorialPeriodoArima itemFilas in informacion.PeriodoHistorico)
            {
                worksheet.Row(row).Height = 14.25;

                worksheet.Cells["A" + row].Value = itemFilas.Item;
                worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + row].Style.WrapText = true;
                worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["B" + row].Value = itemFilas.DescripcionItem;
                worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + row].Style.WrapText = true;
                worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["C" + row].Value = itemFilas.Meses1;
                worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + row].Style.WrapText = true;
                worksheet.Cells["C" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["D" + row].Value = itemFilas.Meses2;
                worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + row].Style.WrapText = true;
                worksheet.Cells["D" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["E" + row].Value = itemFilas.Meses3;
                worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + row].Style.WrapText = true;
                worksheet.Cells["E" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["F" + row].Value = itemFilas.Meses4;
                worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + row].Style.WrapText = true;
                worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["G" + row].Value = itemFilas.Meses5;
                worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + row].Style.WrapText = true;
                worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["H" + row].Value = itemFilas.Meses6;
                worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + row].Style.WrapText = true;
                worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["I" + row].Value = itemFilas.Meses7;
                worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + row].Style.WrapText = true;
                worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["J" + row].Value = itemFilas.Meses8;
                worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + row].Style.WrapText = true;
                worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["K" + row].Value = itemFilas.Meses9;
                worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + row].Style.WrapText = true;
                worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["L" + row].Value = itemFilas.Meses10;
                worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + row].Style.WrapText = true;
                worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["M" + row].Value = itemFilas.Meses11;
                worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + row].Style.WrapText = true;
                worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["N" + row].Value = itemFilas.Meses12;
                worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + row].Style.WrapText = true;
                worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                worksheet.Cells["O" + row].Formula = "=Sum(" + worksheet.Cells[row, 3].Address + ":" + worksheet.Cells[row, 14].Address + ")"; ;
                worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + row].Style.WrapText = true;
                worksheet.Cells["O" + row].Style.Font.Bold = true;
                worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                row++;


            }
        }

        private static void ConfigurarTamanioDeCeldasHistorico(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 4.43 + 2.71;
            worksheet.Column(2).Width = 70.71 + 2.71;
            worksheet.Column(3).Width = 10 + 2.71;
            worksheet.Column(4).Width = 10 + 2.71;
            worksheet.Column(5).Width = 10 + 2.71;
            worksheet.Column(6).Width = 10 + 2.71;
            worksheet.Column(7).Width = 10 + 2.71;
            worksheet.Column(8).Width = 10 + 2.71;
            worksheet.Column(9).Width = 10 + 2.71;
            worksheet.Column(10).Width = 10 + 2.71;
            worksheet.Column(11).Width = 10 + 2.71;
            worksheet.Column(12).Width = 10 + 2.71;
            worksheet.Column(13).Width = 10 + 2.71;
            worksheet.Column(14).Width = 10 + 2.71;
            worksheet.Column(15).Width = 10 + 2.71;
            worksheet.Column(16).Width = 15 + 2.71;
        }

        private static void PintarCabeceraHistorico(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:O1"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#1F73B9"));
            worksheet.Cells["A1:O1"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
        }
    }
}
