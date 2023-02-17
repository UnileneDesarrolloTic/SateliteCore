using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.Produccion
{
    public class ReporteExcelCompraDrogueria
    {
        public string GenerarReporte(IEnumerable<DatosFormatoReporteSeguimientoDrogueria> dato, bool mostrarcolumna)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Compra Drogueria");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["A1:Y2"].Merge = true;
                worksheet.Cells["A1:Y2"].Value = "Generación de Excel Drogueria " + DateTime.Now;
                worksheet.Cells["A1:Y2"].Style.Font.Size = 24;
                worksheet.Cells["A1:Y2"].Style.WrapText = true;
                worksheet.Cells["A1:Y2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A3"].Value = "Item";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A3"].Style.Font.Size = 9;

                worksheet.Cells["B3"].Value = "Descripción";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.Font.Bold = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B3"].Style.Font.Size = 9;


                worksheet.Cells["C3"].Value = "Puerto";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.Font.Bold = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C3"].Style.Font.Size = 9;

                worksheet.Cells["D3"].Value = "Proveedor";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.Font.Bold = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D3"].Style.Font.Size = 9;

                worksheet.Cells["E3"].Value = "Tiempo en la gestión de compras (cotizacion y aceptacion)";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.Font.Bold = true;
                worksheet.Cells["E3"].Style.Font.Size = 9;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E3"].Style.Font.Size = 9;

                worksheet.Cells["F3"].Value = "Tiempo en gestión pago contable al proveedor";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.Font.Bold = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F3"].Style.Font.Size = 9;

                worksheet.Cells["G3"].Value = "Tiempo en aprobacion de artes calidad";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.Font.Bold = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G3"].Style.Font.Size = 9;

                worksheet.Cells["H3"].Value = "Tiempo de fabricación del proveedor";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H3"].Style.WrapText = true;
                worksheet.Cells["H3"].Style.Font.Bold = true;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H3"].Style.Font.Size = 9;

                worksheet.Cells["I3"].Value = "Tiempo de transporte maritimo";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I3"].Style.WrapText = true;
                worksheet.Cells["I3"].Style.Font.Bold = true;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I3"].Style.Font.Size = 9;

                worksheet.Cells["J3"].Value = "Tiempo de nacionalizar aduanas";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J3"].Style.WrapText = true;
                worksheet.Cells["J3"].Style.Font.Bold = true;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J3"].Style.Font.Size = 9;

                worksheet.Cells["K3"].Value = "qty maxima deberia haber en stock linea";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K3"].Style.WrapText = true;
                worksheet.Cells["K3"].Style.Font.Bold = true;
                worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K3"].Style.Font.Size = 9;

                worksheet.Cells["L3"].Value = "punto de corte de stock donde se debe comprar";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L3"].Style.WrapText = true;
                worksheet.Cells["L3"].Style.Font.Bold = true;
                worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L3"].Style.Font.Size = 9;

                worksheet.Cells["M3"].Value = "stock fisico actual";
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M3"].Style.WrapText = true;
                worksheet.Cells["M3"].Style.Font.Bold = true;
                worksheet.Cells["M3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M3"].Style.Font.Size = 9;

                worksheet.Cells["N3"].Value = "producto comprado en transito y/o aduana";
                worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N3"].Style.WrapText = true;
                worksheet.Cells["N3"].Style.Font.Bold = true;
                worksheet.Cells["N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N3"].Style.Font.Size = 9;

                worksheet.Cells["O3"].Value = "total producto en transito mas stock actual";
                worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O3"].Style.WrapText = true;
                worksheet.Cells["O3"].Style.Font.Bold = true;
                worksheet.Cells["O3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O3"].Style.Font.Size = 9;

                worksheet.Cells["P3"].Value = "dias potenciales con stock futuro";
                worksheet.Cells["P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P3"].Style.WrapText = true;
                worksheet.Cells["P3"].Style.Font.Bold = true;
                worksheet.Cells["P3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P3"].Style.Font.Size = 9;

                worksheet.Cells["Q3"].Value = "promedio de consumo mensual que da el arima (Unidades)";
                worksheet.Cells["Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q3"].Style.WrapText = true;
                worksheet.Cells["Q3"].Style.Font.Bold = true;
                worksheet.Cells["Q3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Q3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q3"].Style.Font.Size = 9;

                worksheet.Cells["R3"].Value = "promedio de consumo x dia";
                worksheet.Cells["R3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R3"].Style.WrapText = true;
                worksheet.Cells["R3"].Style.Font.Bold = true;
                worksheet.Cells["R3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["R3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R3"].Style.Font.Size = 9;

                worksheet.Cells["S3"].Value = "dias de demora en llegada de producto";
                worksheet.Cells["S3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S3"].Style.WrapText = true;
                worksheet.Cells["S3"].Style.Font.Bold = true;
                worksheet.Cells["S3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S3"].Style.Font.Size = 9;

                worksheet.Cells["T3"].Value = "dias de stock de precaución x desviación sobrecompra";
                worksheet.Cells["T3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["T3"].Style.WrapText = true;
                worksheet.Cells["T3"].Style.Font.Bold = true;
                worksheet.Cells["T3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["T3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["T3"].Style.Font.Size = 9;

                worksheet.Cells["U3"].Value = "dias que puedo esperar para comprar";
                worksheet.Cells["U3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["U3"].Style.WrapText = true;
                worksheet.Cells["U3"].Style.Font.Bold = true;
                worksheet.Cells["U3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["U3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["U3"].Style.Font.Size = 9;

                worksheet.Cells["V3"].Value = "coeficiente de variacion";
                worksheet.Cells["V3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["V3"].Style.WrapText = true;
                worksheet.Cells["V3"].Style.Font.Bold = true;
                worksheet.Cells["V3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["V3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["V3"].Style.Font.Size = 9;


                worksheet.Cells["W3"].Value = "COMPRAR";
                worksheet.Cells["W3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W3"].Style.WrapText = true;
                worksheet.Cells["W3"].Style.Font.Bold = true;
                worksheet.Cells["W3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["W3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["W3"].Style.Font.Size = 9;

                worksheet.Cells["X3"].Value = "Volumen caja m3";
                worksheet.Cells["X3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X3"].Style.WrapText = true;
                worksheet.Cells["X3"].Style.Font.Bold = true;
                worksheet.Cells["X3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["X3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["X3"].Style.Font.Size = 9;

                worksheet.Cells["Y3"].Value = "Total m3 Comprar";
                worksheet.Cells["Y3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Y3"].Style.WrapText = true;
                worksheet.Cells["Y3"].Style.Font.Bold = true;
                worksheet.Cells["Y3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Y3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Y3"].Style.Font.Size = 9;

                worksheet.Cells["Z3"].Value = "PrecioFOB";
                worksheet.Cells["Z3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Z3"].Style.WrapText = true;
                worksheet.Cells["Z3"].Style.Font.Bold = true;
                worksheet.Cells["Z3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Z3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Z3"].Style.Font.Size = 9;


                worksheet.Cells["AA3"].Value = "PrecioFOB Total";
                worksheet.Cells["AA3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AA3"].Style.Font.Bold = true;
                worksheet.Cells["AA3"].Style.WrapText = true;
                worksheet.Cells["AA3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AA3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AA3"].Style.Font.Size = 9;

                //Listar los agrupadores

                List<int> grupo = dato.GroupBy(x => x.OrdenAgrupador).Select(y => y.Key).ToList();
                


                int row = 4;

                for(int i = 0;  i < grupo.Count(); i++)
                {
                    var numeroIndex = grupo[i];
                    IEnumerable<DatosFormatoReporteSeguimientoDrogueria> result = dato.Where(x => x.OrdenAgrupador == grupo[i]).ToList();
                    decimal totalCompra = 0;
                    decimal costoTotalFOB = 0;


                    foreach (DatosFormatoReporteSeguimientoDrogueria rowitem in result)
                    {        
                            worksheet.Row(row).Height = 14.25;

                            worksheet.Cells["A" + row].Value = rowitem.Item;
                            worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["A" + row].Style.WrapText = true;
                            worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                            worksheet.Cells["B" + row].Value = rowitem.DescripcionLocal;
                            worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["B" + row].Style.WrapText = true;
                            worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                            worksheet.Cells["C" + row].Value = rowitem.Puerto;
                            worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["C" + row].Style.WrapText = true;
                            worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                            worksheet.Cells["D" + row].Value = rowitem.DescProveedor;
                            worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["D" + row].Style.WrapText = true;
                            worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["E" + row].Value = rowitem.TGestionCompra;
                            worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["E" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["E" + row].Style.WrapText = true;
                            worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["F" + row].Value = rowitem.TGestionPago;
                            worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["F" + row].Style.WrapText = true;
                            worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["G" + row].Value = rowitem.TGestionAprobacion;
                            worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["G" + row].Style.WrapText = true;
                            worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["H" + row].Value = rowitem.TFabricacion;
                            worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["H" + row].Style.WrapText = true;
                            worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["I" + row].Value = rowitem.Transporte;
                            worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["I" + row].Style.WrapText = true;
                            worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            worksheet.Cells["J" + row].Value = rowitem.TAduanas;
                            worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["J" + row].Style.WrapText = true;
                            worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["K" + row].Value = rowitem.MaximoStock;
                            worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["K" + row].Style.WrapText = true;
                            worksheet.Cells["K" + row].Style.Font.Bold = true;
                            worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["L" + row].Value = rowitem.PuntoCorteDebePagar;
                            worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["L" + row].Style.WrapText = true;
                            worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["M" + row].Value = rowitem.StockAlmacenDRO;
                            worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["M" + row].Style.WrapText = true;
                            worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["N" + row].Value = rowitem.CantidadOC;
                            worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["N" + row].Style.WrapText = true;
                            worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["O" + row].Value = rowitem.TotalStock;
                            worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["O" + row].Style.WrapText = true;
                            worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["P" + row].Value = rowitem.FuturoStock;
                            worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["P" + row].Style.WrapText = true;
                            worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["Q" + row].Value = rowitem.Pronostico;
                            worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["Q" + row].Style.WrapText = true;
                            worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["R" + row].Value = rowitem.ConsumoDiario;
                            worksheet.Cells["R" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["R" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["R" + row].Style.WrapText = true;
                            worksheet.Cells["R" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["S" + row].Value = rowitem.TiempoTotal;
                            worksheet.Cells["S" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["S" + row].Style.WrapText = true;
                            worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["T" + row].Value = rowitem.DiasAdicionales;
                            worksheet.Cells["T" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["T" + row].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells["T" + row].Style.WrapText = true;
                            worksheet.Cells["T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["U" + row].Value = rowitem.DiasEspera;
                            worksheet.Cells["U" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["U" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["U" + row].Style.Font.Color.SetColor(rowitem.DiasEspera >= 0 ? Color.Black : Color.Red);
                            worksheet.Cells["U" + row].Style.WrapText = true;
                            worksheet.Cells["U" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["V" + row].Value = rowitem.Variacion;
                            worksheet.Cells["V" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["V" + row].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells["V" + row].Style.WrapText = true;
                            worksheet.Cells["V" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["W" + row].Value = rowitem.CantidadComprar;
                            worksheet.Cells["W" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["W" + row].Style.Numberformat.Format = "#,##0";
                            worksheet.Cells["W" + row].Style.WrapText = true;
                            worksheet.Cells["W" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["X" + row].Value = rowitem.VolumenCaja;
                            worksheet.Cells["X" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["X" + row].Style.Numberformat.Format = "#,##0.000";
                            worksheet.Cells["X" + row].Style.WrapText = true;
                            worksheet.Cells["X" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["Y" + row].Value = rowitem.TotalComprar;
                            worksheet.Cells["Y" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["Y" + row].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells["Y" + row].Style.WrapText = true;
                            worksheet.Cells["Y" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["Z" + row].Value = rowitem.PrecioFOBFinal;
                            worksheet.Cells["Z" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["Z" + row].Style.Numberformat.Format = "$#,##0.0000";
                            worksheet.Cells["Z" + row].Style.WrapText = true;
                            worksheet.Cells["Z" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            worksheet.Cells["AA" + row].Value = rowitem.PrecioFOBTotalFinal;
                            worksheet.Cells["AA" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            worksheet.Cells["AA" + row].Style.Numberformat.Format = "$#,##0.00";
                            worksheet.Cells["AA" + row].Style.WrapText = true;
                            worksheet.Cells["AA" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                            totalCompra = totalCompra + rowitem.TotalComprar;
                            costoTotalFOB = costoTotalFOB + rowitem.PrecioFOBTotalFinal;

                        row++;
                    }

                    worksheet.Cells["W" + row].Value = "TOTAL";
                    worksheet.Cells["W" + row].Style.WrapText = true;
                    worksheet.Cells["W" + row].Style.Font.Bold = true;
                    worksheet.Cells["W" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["Y" + row].Value = totalCompra;
                    worksheet.Cells["Y" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["Y" + row].Style.Font.Bold = true;
                    worksheet.Cells["Y" + row].Style.WrapText = true;
                    worksheet.Cells["Y" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["AA" + row].Value = costoTotalFOB;
                    worksheet.Cells["AA" + row].Style.Numberformat.Format = "$#,##0.00";
                    worksheet.Cells["AA" + row].Style.Font.Bold = true;
                    worksheet.Cells["AA" + row].Style.WrapText = true;
                    worksheet.Cells["AA" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    row = row + 2;
                }
               

                if (mostrarcolumna)
                {
                    worksheet.Column(3).Hidden = false;
                    worksheet.Column(5).Hidden = false;
                    worksheet.Column(6).Hidden = false;
                    worksheet.Column(7).Hidden = false;
                    worksheet.Column(8).Hidden = false;
                    worksheet.Column(9).Hidden = false;
                    worksheet.Column(10).Hidden = false;
                    worksheet.Column(16).Hidden = false;
                    worksheet.Column(18).Hidden = false;
                    worksheet.Column(19).Hidden = false;
                    worksheet.Column(20).Hidden = false;
                    worksheet.Column(22).Hidden = false;
                }
                else
                {
                    worksheet.Column(3).Hidden = true;
                    worksheet.Column(5).Hidden = true;
                    worksheet.Column(6).Hidden = true;
                    worksheet.Column(7).Hidden = true;
                    worksheet.Column(8).Hidden = true;
                    worksheet.Column(9).Hidden = true;
                    worksheet.Column(10).Hidden = true;
                    worksheet.Column(16).Hidden = true;
                    worksheet.Column(18).Hidden = true;
                    worksheet.Column(19).Hidden = true;
                    worksheet.Column(22).Hidden = true;
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
            worksheet.Column(2).Width = 30.71 + 2.71;
            worksheet.Column(3).Width = 7.71 + 2.71;
            worksheet.Column(4).Width = 7.71 + 2.71;
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
            worksheet.Column(16).Width = 7.71 + 2.71;
            worksheet.Column(17).Width = 7.71 + 2.71;
            worksheet.Column(18).Width = 7.71 + 2.71;
            worksheet.Column(19).Width = 7.71 + 2.71;
            worksheet.Column(20).Width = 7.71 + 2.71;
            worksheet.Column(21).Width = 7.71 + 2.71;
            worksheet.Column(22).Width = 7.71 + 2.71;
            worksheet.Column(23).Width = 7.71 + 2.71;
            worksheet.Column(24).Width = 7.71 + 2.71;
            worksheet.Column(25).Width = 7.71 + 2.71;
            worksheet.Column(26).Width = 8.71 + 2.71;
            worksheet.Column(27).Width = 9.71 + 2.71;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:D3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["E3:J3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffe180"));
            worksheet.Cells["K3:Q3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["R3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF"));
            worksheet.Cells["S3:U3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["V3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF"));
            worksheet.Cells["W3:AA3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
        
        }

    }
}
