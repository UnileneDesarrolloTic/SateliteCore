using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;


namespace SatelliteCore.Api.ReportServices.Contracts.Produccion
{
    public class ReporteCompraAguja_Excel
    {
        public string GenerarReporte(IEnumerable<DatosFormatoListadoSeguimientoCompraAguja> dato, string mostrarColumna, IEnumerable<DatosFormatoCantidadTotalAgujas>  cantidadTotal )
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

           

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Compra Aguja");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.View.ZoomScale = 90;

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);

                int fila = 1;

                pintarCabecera(worksheet, fila);

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

                worksheet.Cells["C" + fila].Value = "Long.Aguja";
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
                worksheet.Cells["E" + fila].Style.Font.Size = 9;

                worksheet.Cells["F" + fila].Value = "Proveedor";
                worksheet.Cells["F" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + fila].Style.WrapText = true;
                worksheet.Cells["F" + fila].Style.Font.Bold = true;
                worksheet.Cells["F" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F" + fila].Style.Font.Size = 9;

                worksheet.Cells["G" + fila].Value = "tiempo en la gestión de compras  (cotización y aceptación)";
                worksheet.Cells["G" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + fila].Style.WrapText = true;
                worksheet.Cells["G" + fila].Style.Font.Bold = true;
                worksheet.Cells["G" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G" + fila].Style.Font.Size = 9;

                worksheet.Cells["H" + fila].Value = "tiempo en gestión pago contable al proveedor";
                worksheet.Cells["H" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + fila].Style.WrapText = true;
                worksheet.Cells["H" + fila].Style.Font.Bold = true;
                worksheet.Cells["H" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H" + fila].Style.Font.Size = 9;

                worksheet.Cells["I" + fila].Value = "tiempo en aprobación de artes Calidad ";
                worksheet.Cells["I" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + fila].Style.WrapText = true;
                worksheet.Cells["I" + fila].Style.Font.Bold = true;
                worksheet.Cells["I" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I" + fila].Style.Font.Size = 9;

                worksheet.Cells["J" + fila].Value = "tiempo de fabricación del proveedor";
                worksheet.Cells["J" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + fila].Style.WrapText = true;
                worksheet.Cells["J" + fila].Style.Font.Bold = true;
                worksheet.Cells["J" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J" + fila].Style.Font.Size = 9;

                worksheet.Cells["K" + fila].Value = "tiempo de transporte";
                worksheet.Cells["K" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + fila].Style.WrapText = true;
                worksheet.Cells["K" + fila].Style.Font.Bold = true;
                worksheet.Cells["K" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K" + fila].Style.Font.Size = 9;

                worksheet.Cells["L" + fila].Value = "tiempo de nacionalizar aduanas";
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

                worksheet.Cells["N" + fila].Value = "qty maxima deberia haber en stock linea";
                worksheet.Cells["N" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + fila].Style.WrapText = true;
                worksheet.Cells["N" + fila].Style.Font.Bold = true;
                worksheet.Cells["N" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N" + fila].Style.Font.Size = 9;

                worksheet.Cells["O" + fila].Value = "punto de corte de stock donde se debe comprar";
                worksheet.Cells["O" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + fila].Style.WrapText = true;
                worksheet.Cells["O" + fila].Style.Font.Bold = true;
                worksheet.Cells["O" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O" + fila].Style.Font.Size = 9;

                worksheet.Cells["P" + fila].Value = "stock fisico DISPONIBLE  \n(+)";
                worksheet.Cells["P" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P" + fila].Style.WrapText = true;
                worksheet.Cells["P" + fila].Style.Font.Bold = true;
                worksheet.Cells["P" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P" + fila].Style.Font.Size = 9;

                worksheet.Cells["Q" + fila].Value = "Planta Requerida \n(-)";
                worksheet.Cells["Q" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + fila].Style.WrapText = true;
                worksheet.Cells["Q" + fila].Style.Font.Bold = true;
                worksheet.Cells["Q" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Q" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q" + fila].Style.Font.Size = 9;

                worksheet.Cells["R" + fila].Value = "Orden Compra Preparación \n(+)";
                worksheet.Cells["R" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R" + fila].Style.WrapText = true;
                worksheet.Cells["R" + fila].Style.Font.Bold = true;
                worksheet.Cells["R" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["R" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R" + fila].Style.Font.Size = 9;

                worksheet.Cells["S" + fila].Value = "producto comprado en transito y/o Aprobado \n(+)";
                worksheet.Cells["S" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + fila].Style.WrapText = true;
                worksheet.Cells["S" + fila].Style.Font.Bold = true;
                worksheet.Cells["S" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S" + fila].Style.Font.Size = 9;

                worksheet.Cells["T" + fila].Value = "Producto Aduana \n(+)";
                worksheet.Cells["T" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["T" + fila].Style.WrapText = true;
                worksheet.Cells["T" + fila].Style.Font.Bold = true;
                worksheet.Cells["T" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["T" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["T" + fila].Style.Font.Size = 9;

                worksheet.Cells["U" + fila].Value = "Producto control calidad \n(+)";
                worksheet.Cells["U" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["U" + fila].Style.WrapText = true;
                worksheet.Cells["U" + fila].Style.Font.Bold = true;
                worksheet.Cells["U" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["U" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["U" + fila].Style.Font.Size = 9;

                worksheet.Cells["V" + fila].Value = "Total de stock disponible";
                worksheet.Cells["V" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["V" + fila].Style.WrapText = true;
                worksheet.Cells["V" + fila].Style.Font.Bold = true;
                worksheet.Cells["V" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["V" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["V" + fila].Style.Font.Size = 9;

                worksheet.Cells["W" + fila].Value = "Dias potenciales con stock futuro";
                worksheet.Cells["W" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W" + fila].Style.WrapText = true;
                worksheet.Cells["W" + fila].Style.Font.Bold = true;
                worksheet.Cells["W" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["W" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["W" + fila].Style.Font.Size = 9;

                worksheet.Cells["X" + fila].Value = "promedio de consumo mensual que da el arima ";
                worksheet.Cells["X" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X" + fila].Style.WrapText = true;
                worksheet.Cells["X" + fila].Style.Font.Bold = true;
                worksheet.Cells["X" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["X" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["X" + fila].Style.Font.Size = 9;

                worksheet.Cells["Y" + fila].Value = "promedio de consumo x dia";
                worksheet.Cells["Y" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Y" + fila].Style.WrapText = true;
                worksheet.Cells["Y" + fila].Style.Font.Bold = true;
                worksheet.Cells["Y" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Y" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Y" + fila].Style.Font.Size = 9;

                worksheet.Cells["Z" + fila].Value = "dias de demora en llegada de producto";
                worksheet.Cells["Z" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Z" + fila].Style.WrapText = true;
                worksheet.Cells["Z" + fila].Style.Font.Bold = true;
                worksheet.Cells["Z" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Z" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Z" + fila].Style.Font.Size = 9;

                worksheet.Cells["AA" + fila].Value = "dias de stock de precaucion x desviacion sobrecompra";
                worksheet.Cells["AA" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AA" + fila].Style.WrapText = true;
                worksheet.Cells["AA" + fila].Style.Font.Bold = true;
                worksheet.Cells["AA" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AA" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AA" + fila].Style.Font.Size = 9;

                worksheet.Cells["AB" + fila].Value = "dias que puedo esperar para comprar";
                worksheet.Cells["AB" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AB" + fila].Style.Font.Bold = true;
                worksheet.Cells["AB" + fila].Style.WrapText = true;
                worksheet.Cells["AB" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AB" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AB" + fila].Style.Font.Size = 9;

                worksheet.Cells["AC" + fila].Value = "coeficiente de variación";
                worksheet.Cells["AC" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AC" + fila].Style.Font.Bold = true;
                worksheet.Cells["AC" + fila].Style.WrapText = true;
                worksheet.Cells["AC" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AC" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AC" + fila].Style.Font.Size = 9;

                worksheet.Cells["AD" + fila].Value = "Cantidad Comprar";
                worksheet.Cells["AD" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AD" + fila].Style.Font.Bold = true;
                worksheet.Cells["AD" + fila].Style.WrapText = true;
                worksheet.Cells["AD" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AD" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AD" + fila].Style.Font.Size = 9;


                worksheet.View.FreezePanes(fila + 1, 1);

                int row = fila + 1;

                foreach (DatosFormatoListadoSeguimientoCompraAguja rowitem in dato)
                {
                    worksheet.Row(row).Height = 14.25;

                    worksheet.Cells["A" + row].Value = rowitem.ItemFinal;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;


                    worksheet.Cells["B" + row].Value = rowitem.DescripcionLocal;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                    worksheet.Cells["C" + row].Value = rowitem.LongAgujas;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.Numberformat.Format = "#,##0.00";
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

                    worksheet.Cells["P" + row].Value = rowitem.Almacen;
                    worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["P" + row].Style.WrapText = true;
                    worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["Q" + row].Value = rowitem.Planta;
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

                    worksheet.Cells["T" + row].Value = rowitem.Aduanas;
                    worksheet.Cells["T" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["T" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["T" + row].Style.WrapText = true;
                    worksheet.Cells["T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["U" + row].Value = rowitem.ControlCalidad;
                    worksheet.Cells["U" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["U" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["U" + row].Style.WrapText = true;
                    worksheet.Cells["U" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["V" + row].Value = rowitem.Disponible;
                    worksheet.Cells["V" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["V" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["V" + row].Style.WrapText = true;
                    worksheet.Cells["V" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["W" + row].Value = rowitem.DiasPotencial;
                    worksheet.Cells["W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["W" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["W" + row].Style.WrapText = true;
                    worksheet.Cells["W" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["X" + row].Value = rowitem.Pronostico;
                    worksheet.Cells["X" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["X" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["X" + row].Style.WrapText = true;
                    worksheet.Cells["X" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["Y" + row].Value = rowitem.ConsumoDia;
                    worksheet.Cells["Y" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["Y" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["Y" + row].Style.WrapText = true;
                    worksheet.Cells["Y" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["Z" + row].Value = rowitem.DemoraLlegarProducto;
                    worksheet.Cells["Z" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["Z" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["Z" + row].Style.WrapText = true;
                    worksheet.Cells["Z" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["AA" + row].Value = rowitem.DesviacionCompra;
                    worksheet.Cells["AA" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["AA" + row].Style.Numberformat.Format = "#,##0.0";
                    worksheet.Cells["AA" + row].Style.WrapText = true;
                    worksheet.Cells["AA" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["AB" + row].Value = rowitem.DiasEspera;
                    worksheet.Cells["AB" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["AB" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["AB" + row].Style.Font.Color.SetColor(rowitem.DiasEspera >= 0 ? Color.Black : Color.Red);
                    worksheet.Cells["AB" + row].Style.WrapText = true;
                    worksheet.Cells["AB" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["AC" + row].Value = rowitem.Variacion;
                    worksheet.Cells["AC" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["AC" + row].Style.Numberformat.Format = "#,##0.0";
                    worksheet.Cells["AC" + row].Style.WrapText = true;
                    worksheet.Cells["AC" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["AD" + row].Value = rowitem.CantidadComprar;
                    worksheet.Cells["AD" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["AD" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["AD" + row].Style.WrapText = true;
                    worksheet.Cells["AD" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    row++;
                }


               
                row = row + 2;
                foreach (DatosFormatoCantidadTotalAgujas tipoBanner in cantidadTotal)
                {   
                    worksheet.Cells["A" + row + ":B" + row].Merge = true;
                    worksheet.Cells["A" + row + ":B" + row].Value = tipoBanner.TipoBanner;
                    worksheet.Cells["A" + row + ":B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row + ":B" + row].Style.Font.Bold=true;
                    worksheet.Cells["A" + row + ":B" + row].Style.Font.Size = 14;
                    worksheet.Cells["A" + row + ":B" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row + ":B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["C" + row + ":D" + row].Merge = true;
                    worksheet.Cells["C" + row + ":D" + row].Value = tipoBanner.Cantidad;
                    worksheet.Cells["C" + row + ":D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row + ":D" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["C" + row + ":D" + row].Style.Font.Bold = true;
                    worksheet.Cells["C" + row + ":D" + row].Style.Font.Size = 14;
                    worksheet.Cells["C" + row + ":D" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row + ":D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    row++;
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
                    worksheet.Column(23).Hidden = false;
                    worksheet.Column(24).Hidden = false;
                    worksheet.Column(25).Hidden = false;
                    worksheet.Column(26).Hidden = false;
                    
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
                    worksheet.Column(23).Hidden = true;
                    worksheet.Column(24).Hidden = true;
                    worksheet.Column(25).Hidden = true;
                    worksheet.Column(26).Hidden = true;
                }

           

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

                return reporte;
            }


     
        }

        private static void AlineacionesTexto(ExcelWorksheet worksheet)
        {
        }

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 3.71 + 2.71;
            worksheet.Column(2).Width = 25.43 + 2.71;
            worksheet.Column(3).Width = 7.71 + 2.71;
            worksheet.Column(4).Width = 16.00 + 2.71;
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
            worksheet.Column(17).Width = 7.71 + 2.71;
            worksheet.Column(18).Width = 9.71 + 2.71;
            worksheet.Column(19).Width = 9.71 + 2.71;
            worksheet.Column(20).Width = 9.71 + 2.71;
            worksheet.Column(21).Width = 9.71 + 2.71;
            worksheet.Column(22).Width = 9.71 + 2.71;
            worksheet.Column(23).Width = 7.71 + 2.71;
            worksheet.Column(24).Width = 7.71 + 2.71;
            worksheet.Column(25).Width = 7.71 + 2.71;
            worksheet.Column(26).Width = 7.71 + 2.71;
            worksheet.Column(27).Width = 7.71 + 2.71;
            worksheet.Column(28).Width = 7.71 + 2.71;

        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet, int fila)
        {
            worksheet.Cells["A" + fila + ":F" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["G" + fila + ":I" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
            worksheet.Cells["J" + fila + ":M" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d1d1d1"));
            worksheet.Cells["N" + fila + ":O" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["P" + fila + ":V" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
            worksheet.Cells["W" + fila + ":AB" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["AD" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
        }


        private static void TextoNegrita(ExcelWorksheet worksheet, int fila)
        {
          
        }


    }
}
