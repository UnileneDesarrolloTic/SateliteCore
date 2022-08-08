using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Dashboard
{
    public class ReporteLicitaciones
    {
        public string GenerarReporteDashboardLicitaciones(IEnumerable<DatosFormatodashboardLicitaciones>  documento)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Dashboard Licitaciones");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["A1"].Value = "INFORMACIÓN DEL DASHBOARD DE LICITACIONES";
                worksheet.Cells["A1"].Style.Font.Size = 16;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["A3"].Value = "Proceso";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["B3"].Value = "Descripción Item";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.Font.Size = 10;
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C3"].Value = "Orden Compra";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.Font.Size = 10;
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D3"].Value = "Entrega";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.Font.Size = 9;
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E3"].Value = "Guia";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.Font.Size = 10;
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "Estado Facturación";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G3"].Value = "Factura";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.Font.Size = 10;
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H3"].Value = "Monto Factura";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H3"].Style.Font.Size = 10;
                worksheet.Cells["H3"].Style.WrapText = true;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I3"].Value = "Estado Comercial";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I3"].Style.Font.Size = 10;
                worksheet.Cells["I3"].Style.WrapText = true;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J3"].Value = "EstadoLogistica";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J3"].Style.Font.Size = 10;
                worksheet.Cells["J3"].Style.WrapText = true;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;



                worksheet.Cells["K3"].Value = "Estado";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K3"].Style.Font.Size = 10;
                worksheet.Cells["K3"].Style.WrapText = true;
                worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["L3"].Value = "Usuario";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L3"].Style.Font.Size = 10;
                worksheet.Cells["L3"].Style.WrapText = true;
                worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["M3"].Value = "Monto Total Factura";
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M3"].Style.Font.Size = 10;
                worksheet.Cells["M3"].Style.WrapText = true;
                worksheet.Cells["M3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 4;
                
                foreach (DatosFormatodashboardLicitaciones rowitem in documento)
                {

                    worksheet.Row(row).Height = 15.25;

                    worksheet.Cells["A" + row].Value = rowitem.Proceso;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B" + row].Value = rowitem.DescripcionItem;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["C" + row].Value = rowitem.OrdenCompra;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["D" + row].Value = rowitem.Entrega;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Numberformat.Format = "#,##0.0";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["E" + row].Value = rowitem.Guia;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["F" + row].Value = rowitem.EstadoFacturacion;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";                  
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row].Value = rowitem.Factura;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["H" + row].Value = rowitem.MontoFactura;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["I" + row].Value = rowitem.EstadoComercial;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["J" + row].Value = rowitem.EstadoLogistica;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["K" + row].Value = rowitem.Estado;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + row].Style.Font.Size = 10;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["L" + row].Value = rowitem.Usuario;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["L" + row].Style.Font.Size = 10;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["M" + row].Value = rowitem.MontoTotalFactura;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["M" + row].Style.Font.Size = 10;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    row++;
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
            worksheet.Column(1).Width = 18.86 + 2.71;
            worksheet.Column(2).Width = 68.43 + 2.71;
            worksheet.Column(3).Width = 13.14 + 2.71;
            worksheet.Column(4).Width = 7 + 2.71;
            worksheet.Column(5).Width = 14.86 + 2.71;
            worksheet.Column(6).Width = 16.86 + 2.71;
            worksheet.Column(7).Width = 16.86 + 2.71;
            worksheet.Column(8).Width = 14.86 + 2.71;

            worksheet.Column(9).Width = 13.00 + 2.71;
            worksheet.Column(10).Width = 15.57 + 2.71;
            worksheet.Column(11).Width = 14.14 + 2.71;
            worksheet.Column(12).Width = 18.43 + 2.71;
            worksheet.Column(13).Width = 14.71 + 2.71;
            worksheet.Column(14).Width = 18 + 2.71;



        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:M1"].Merge=true;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:M3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D8D8D8"));
        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
