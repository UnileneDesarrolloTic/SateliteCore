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
        public string GenerarReporte(IEnumerable<DatosFormatoReporteSeguimientoDrogueria> dato, bool mostrarcolumna, IEnumerable<DatosFormatoGestionItemDrogueriaColor> condicionesgestion, string agrupador)
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
                
                TextoNegrita(worksheet);

                int fila = 1;
                foreach (DatosFormatoGestionItemDrogueriaColor gestion in condicionesgestion)
                {
                    worksheet.Cells["A" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(gestion.Color));
                    worksheet.Cells["A" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + fila].Value = gestion.Descripcion;
                    worksheet.Cells["B" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + fila].Style.Font.Bold=true;
                    fila++;
                }

                worksheet.View.FreezePanes(fila + 1, 4);

                worksheet.Cells["D1:AA1"].Merge = true;
                worksheet.Cells["D1:AA1"].Value = "Generación de Excel Drogueria " + DateTime.Now;
                worksheet.Cells["D1:AA1"].Style.Font.Size = 15;
                worksheet.Cells["D1:AA1"].Style.WrapText = true;
                worksheet.Cells["D1:AA1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D1:AA1"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                pintarCabecera(worksheet,fila);
                worksheet.Cells["A" + fila].Value = "Item";
                worksheet.Cells["A" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + fila].Style.WrapText = true;
                worksheet.Cells["A" + fila].Style.Font.Bold = true;
                worksheet.Cells["A" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A" + fila].Style.Font.Size = 9;

                worksheet.Cells["B" + fila].Value = "Descripción";
                worksheet.Cells["B" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + fila].Style.WrapText = true;
                worksheet.Cells["B" + fila].Style.Font.Bold = true;
                worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["B" + fila].Style.Font.Size = 9;


                worksheet.Cells["C" + fila].Value = "Puerto";
                worksheet.Cells["C" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + fila].Style.WrapText = true;
                worksheet.Cells["C" + fila].Style.Font.Bold = true;
                worksheet.Cells["C" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["C" + fila].Style.Font.Size = 9;

                worksheet.Cells["D" + fila].Value = "Proveedor";
                worksheet.Cells["D" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + fila].Style.WrapText = true;
                worksheet.Cells["D" + fila].Style.Font.Bold = true;
                worksheet.Cells["D" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["D" + fila].Style.Font.Size = 9;

                worksheet.Cells["E" + fila].Value = "Tiempo en la gestión de compras (cotizacion y aceptacion)";
                worksheet.Cells["E" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + fila].Style.WrapText = true;
                worksheet.Cells["E" + fila].Style.Font.Bold = true;
                worksheet.Cells["E" + fila].Style.Font.Size = 9;
                worksheet.Cells["E" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["E" + fila].Style.Font.Size = 9;

                worksheet.Cells["F" + fila].Value = "Tiempo en gestión pago contable al proveedor";
                worksheet.Cells["F" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + fila].Style.WrapText = true;
                worksheet.Cells["F" + fila].Style.Font.Bold = true;
                worksheet.Cells["F" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["F" + fila].Style.Font.Size = 9;

                worksheet.Cells["G" + fila].Value = "Tiempo en aprobacion de artes calidad";
                worksheet.Cells["G" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + fila].Style.WrapText = true;
                worksheet.Cells["G" + fila].Style.Font.Bold = true;
                worksheet.Cells["G" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["G" + fila].Style.Font.Size = 9;

                worksheet.Cells["H" + fila].Value = "Tiempo de fabricación del proveedor";
                worksheet.Cells["H" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + fila].Style.WrapText = true;
                worksheet.Cells["H" + fila].Style.Font.Bold = true;
                worksheet.Cells["H" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["H" + fila].Style.Font.Size = 9;

                worksheet.Cells["I" + fila].Value = "Tiempo de transporte maritimo";
                worksheet.Cells["I" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + fila].Style.WrapText = true;
                worksheet.Cells["I" + fila].Style.Font.Bold = true;
                worksheet.Cells["I" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["I" + fila].Style.Font.Size = 9;

                worksheet.Cells["J" + fila].Value = "Tiempo de nacionalizar aduanas";
                worksheet.Cells["J" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + fila].Style.WrapText = true;
                worksheet.Cells["J" + fila].Style.Font.Bold = true;
                worksheet.Cells["J" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["J" + fila].Style.Font.Size = 9;

                worksheet.Cells["K" + fila].Value = "qty maxima deberia haber en stock linea";
                worksheet.Cells["K" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + fila].Style.WrapText = true;
                worksheet.Cells["K" + fila].Style.Font.Bold = true;
                worksheet.Cells["K" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["K" + fila].Style.Font.Size = 9;

                worksheet.Cells["L" + fila].Value = "punto de corte de stock donde se debe comprar";
                worksheet.Cells["L" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L" + fila].Style.WrapText = true;
                worksheet.Cells["L" + fila].Style.Font.Bold = true;
                worksheet.Cells["L" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["L" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["L" + fila].Style.Font.Size = 9;

                worksheet.Cells["M" + fila].Value = "stock fisico actual";
                worksheet.Cells["M" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M" + fila].Style.WrapText = true;
                worksheet.Cells["M" + fila].Style.Font.Bold = true;
                worksheet.Cells["M" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["M" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["M" + fila].Style.Font.Size = 9;

                worksheet.Cells["N" + fila].Value = "producto comprado en transito y/o aduana";
                worksheet.Cells["N" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N" + fila].Style.WrapText = true;
                worksheet.Cells["N" + fila].Style.Font.Bold = true;
                worksheet.Cells["N" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["N" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["N" + fila].Style.Font.Size = 9;

                worksheet.Cells["O" + fila].Value = "total producto en transito mas stock actual";
                worksheet.Cells["O" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O" + fila].Style.WrapText = true;
                worksheet.Cells["O" + fila].Style.Font.Bold = true;
                worksheet.Cells["O" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["O" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["O" + fila].Style.Font.Size = 9;

                worksheet.Cells["P" + fila].Value = "dias potenciales con stock futuro";
                worksheet.Cells["P" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P" + fila].Style.WrapText = true;
                worksheet.Cells["P" + fila].Style.Font.Bold = true;
                worksheet.Cells["P" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["P" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["P" + fila].Style.Font.Size = 9;

                worksheet.Cells["Q" + fila].Value = "promedio de consumo mensual que da el arima (Unidades)";
                worksheet.Cells["Q" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q" + fila].Style.WrapText = true;
                worksheet.Cells["Q" + fila].Style.Font.Bold = true;
                worksheet.Cells["Q" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Q" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Q" + fila].Style.Font.Size = 9;

                worksheet.Cells["R" + fila].Value = "promedio de consumo x dia";
                worksheet.Cells["R" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["R" + fila].Style.WrapText = true;
                worksheet.Cells["R" + fila].Style.Font.Bold = true;
                worksheet.Cells["R" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["R" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["R" + fila].Style.Font.Size = 9;

                worksheet.Cells["S" + fila].Value = "dias de demora en llegada de producto";
                worksheet.Cells["S" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["S" + fila].Style.WrapText = true;
                worksheet.Cells["S" + fila].Style.Font.Bold = true;
                worksheet.Cells["S" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["S" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["S" + fila].Style.Font.Size = 9;

                worksheet.Cells["T" + fila].Value = "dias de stock de precaución x desviación sobrecompra";
                worksheet.Cells["T" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["T" + fila].Style.WrapText = true;
                worksheet.Cells["T" + fila].Style.Font.Bold = true;
                worksheet.Cells["T" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["T" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["T" + fila].Style.Font.Size = 9;

                worksheet.Cells["U" + fila].Value = "dias que puedo esperar para comprar";
                worksheet.Cells["U" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["U" + fila].Style.WrapText = true;
                worksheet.Cells["U" + fila].Style.Font.Bold = true;
                worksheet.Cells["U" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["U" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["U" + fila].Style.Font.Size = 9;

                worksheet.Cells["V" + fila].Value = "coeficiente de variacion";
                worksheet.Cells["V" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["V" + fila].Style.WrapText = true;
                worksheet.Cells["V" + fila].Style.Font.Bold = true;
                worksheet.Cells["V" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["V" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["V" + fila].Style.Font.Size = 9;


                worksheet.Cells["W" + fila].Value = "COMPRAR";
                worksheet.Cells["W" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["W" + fila].Style.WrapText = true;
                worksheet.Cells["W" + fila].Style.Font.Bold = true;
                worksheet.Cells["W" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["W" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["W" + fila].Style.Font.Size = 9;

                worksheet.Cells["X" + fila].Value = "Volumen caja m3";
                worksheet.Cells["X" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["X" + fila].Style.WrapText = true;
                worksheet.Cells["X" + fila].Style.Font.Bold = true;
                worksheet.Cells["X" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["X" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["X" + fila].Style.Font.Size = 9;

                worksheet.Cells["Y" + fila].Value = "Total m3 Comprar";
                worksheet.Cells["Y" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Y" + fila].Style.WrapText = true;
                worksheet.Cells["Y" + fila].Style.Font.Bold = true;
                worksheet.Cells["Y" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Y" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Y" + fila].Style.Font.Size = 9;

                worksheet.Cells["Z" + fila].Value = "PrecioFOB";
                worksheet.Cells["Z" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Z" + fila].Style.WrapText = true;
                worksheet.Cells["Z" + fila].Style.Font.Bold = true;
                worksheet.Cells["Z" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["Z" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["Z" + fila].Style.Font.Size = 9;


                worksheet.Cells["AA" + fila].Value = "PrecioFOB Total";
                worksheet.Cells["AA" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["AA" + fila].Style.Font.Bold = true;
                worksheet.Cells["AA" + fila].Style.WrapText = true;
                worksheet.Cells["AA" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["AA" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["AA" + fila].Style.Font.Size = 9;

                //Listar los agrupadores
                int row = fila + 1;

                if (agrupador == "conGrupo") 
                {
                    List<int> grupo = dato.GroupBy(x => x.OrdenAgrupador).Select(y => y.Key).ToList();
                    

                    for (int i = 0; i < grupo.Count(); i++)
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
                            worksheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(rowitem.ColorVariacion));

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

                }
                else
                {


                    foreach (DatosFormatoReporteSeguimientoDrogueria rowitem in dato)
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
                        worksheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(rowitem.ColorVariacion));

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
                        row++;
                    }


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
                    worksheet.Column(20).Hidden = true;
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
            worksheet.Column(2).Width = 40.86 + 2.71;
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
            worksheet.Column(27).Width = 10.71 + 2.71;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet, int fila)
        {
            worksheet.Cells["A" + fila + ":D" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["E" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffe180"));
            worksheet.Cells["K" + fila + ":Q" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["R" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF"));
            worksheet.Cells["S" + fila + ":U" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#77dd77"));
            worksheet.Cells["V" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFF"));
            worksheet.Cells["W" + fila + ":AA"+ fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffda9e"));
        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {
        
        }

    }
}
