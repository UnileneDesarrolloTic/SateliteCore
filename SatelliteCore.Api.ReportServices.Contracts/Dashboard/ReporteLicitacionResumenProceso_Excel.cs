using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Dashboard;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.Dashboard
{
    public class ReporteLicitacionResumenProceso_Excel
    {
        public string GenerarReporteDashboardLicitaciones(IEnumerable<DatosFormatoResumenProcesoLicitaciones> listado)
        {

            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Dashboard Licitaciones Resumen Proceso");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 12;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);

                worksheet.Cells["A1:L2"].Merge = true;
                worksheet.Cells["A1:L2"].Value = "Generación de Excel : Resumen otros procesos " + DateTime.Now;
                worksheet.Cells["A1:L2"].Style.Font.Size = 24;
                worksheet.Cells["A1:L2"].Style.WrapText = true;
                worksheet.Cells["A1:L2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1:L2"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;


                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3"].Style.Font.Bold = true;

                worksheet.Cells["A3"].Value = "IdProceso";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B3"].Value = "Descripción Proceso";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C3"].Value = "Entidad";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D3"].Value = "CodItem";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E3"].Value = "Descripción Item";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "Número Entrega";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G3"].Value = "Primera Fecha: Fecha vencimiento";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H3"].Value = "Cantidad Adjudicada";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H3"].Style.WrapText = true;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I3"].Value = "Precio Unitario";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I3"].Style.WrapText = true;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J3"].Value = "Monto Adjudicado (S/.)";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J3"].Style.WrapText = true;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K3"].Value = "Cantidad Entregada en plazo";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K3"].Style.WrapText = true;
                worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["L3"].Value = "Monto Facturado (S/.)";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L3"].Style.WrapText = true;
                worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.View.FreezePanes(4, 1);


                List<int> agrupadorProceso = listado.GroupBy(x => x.IdProceso).Select(y => y.Key).ToList();
                int row = 4;

                for (int i = 0; i < agrupadorProceso.Count(); i++)
                {
                    var numeroIndex = agrupadorProceso[i];
                    IEnumerable<DatosFormatoResumenProcesoLicitaciones> result = listado.Where(x => x.IdProceso == agrupadorProceso[i]).ToList();
                    decimal cantidadAdjudicada = 0;
                    decimal montoAdjudicado = 0;
                    decimal cantidadEntregadaPlazo = 0;
                    decimal montoFacturado = 0;

                    foreach (DatosFormatoResumenProcesoLicitaciones filaproceso in result)
                    {
                        worksheet.Row(row).Height = 14.25;

                        worksheet.Cells["A" + row].Value = filaproceso.IdProceso;
                        worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["A" + row].Style.WrapText = true;
                        worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                        worksheet.Cells["B" + row].Value = filaproceso.DescripcionProceso;
                        worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["B" + row].Style.WrapText = true;
                        worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        worksheet.Cells["C" + row].Value = filaproceso.DcliNombreCompleto;
                        worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["C" + row].Style.WrapText = true;
                        worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        worksheet.Cells["D" + row].Value = filaproceso.CodItem;
                        worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["D" + row].Style.WrapText = true;
                        worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        worksheet.Cells["E" + row].Value = filaproceso.DescripcionItem;
                        worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["E" + row].Style.WrapText = true;
                        worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        worksheet.Cells["F" + row].Value = filaproceso.NumeroEntrega;
                        worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["F" + row].Style.WrapText = true;
                        worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;

                        worksheet.Cells["G" + row].Value = filaproceso.FechaVencimiento.ToString("dd/MM/yyyy");
                        worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["G" + row].Style.WrapText = true;
                        worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        worksheet.Cells["H" + row].Value = filaproceso.CantidadEntregar;
                        worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["H" + row].Style.WrapText = true;
                        worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        worksheet.Cells["I" + row].Value = filaproceso.PrecioUnitario;
                        worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["I" + row].Style.WrapText = true;
                        worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        worksheet.Cells["J" + row].Value = filaproceso.MontoCobrar;
                        worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["J" + row].Style.WrapText = true;
                        worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        worksheet.Cells["K" + row].Value = filaproceso.CantidadEntregadaPlazo;
                        worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["K" + row].Style.WrapText = true;
                        worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        worksheet.Cells["L" + row].Value = filaproceso.MontoFacturado;
                        worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";
                        worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        worksheet.Cells["L" + row].Style.WrapText = true;
                        worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                        cantidadAdjudicada = cantidadAdjudicada + filaproceso.CantidadEntregar;
                        montoAdjudicado = montoAdjudicado + filaproceso.MontoCobrar;
                        cantidadEntregadaPlazo = cantidadEntregadaPlazo + filaproceso.CantidadEntregadaPlazo;
                        montoFacturado = montoFacturado + filaproceso.MontoFacturado;

                        row++;
                    }


                    worksheet.Cells["G" + row].Value = "TOTAL";
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["H" + row].Value = cantidadAdjudicada;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["J" + row].Value = montoAdjudicado;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["K" + row].Value = cantidadEntregadaPlazo;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["L" + row].Value = montoFacturado;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["G" + row + ":L" + row].Style.Font.Bold = true;
                    worksheet.Cells["G" + row + ":L" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#000000"));
                    worksheet.Cells["G" + row + ":L" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));

                    row = row + 2;
                }
                    file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

                return reporte;

            }
          
        }


        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 3.29 + 2.71;
            worksheet.Column(2).Width = 20.29 + 2.71;
            worksheet.Column(3).Width = 61 + 2.71;
            worksheet.Column(4).Width = 14.86 + 2.71;
            worksheet.Column(5).Width = 59.71 + 2.71;
            worksheet.Column(6).Width = 10.29 + 2.71;
            worksheet.Column(7).Width = 18.57 + 2.71;
            worksheet.Column(8).Width = 14.29 + 2.71;

            worksheet.Column(9).Width = 10.71 + 2.71;
            worksheet.Column(10).Width = 14.57 + 2.71;
            worksheet.Column(11).Width = 17.00 + 2.71;
            worksheet.Column(12).Width = 14.00 + 2.71;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:L1"].Merge = true;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:L3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#000000"));
            worksheet.Cells["A3:L3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
        }

    }
}
