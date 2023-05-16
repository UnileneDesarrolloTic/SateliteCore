using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.CompraAguja;
using SatelliteCore.Api.Models.Response.OCDrogueria;
using SatelliteCore.Api.Models.Response.RRHH;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SatelliteCore.Api.ReportServices.Contracts.RRHH
{
    public class ReporteComisionVendedor_excel
    {
        public string GenerarReporte(IEnumerable<DatosFormatoReporteComisionVendedor> dato, string periodo)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string anioreporte = periodo.Substring(0, 6);
            string anio = anioreporte.Substring(0, 4);
            string mes = anioreporte.Substring(4, 2);
            DateTime fecha1;
            DateTime fecha2;

            fecha1 = new DateTime(Int32.Parse(anio), Int32.Parse(mes), 1);
            fecha2 = new DateTime(Int32.Parse(anio), Int32.Parse(mes) + 1, 1).AddDays(-1);

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Compra Aguja");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Font.Size = 12;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.View.ZoomScale = 90;

                int fila = 1;

               ConfigurarTamanioDeCeldas(worksheet);

                worksheet.Cells["A" + fila + ":M" + fila].Merge = true;
                worksheet.Cells["A" + fila + ":M" + fila].Value = "Reporte de comision " + fecha1.ToString("dd/MM/yyyy") + " al " + fecha2.ToString("dd/MM/yyyy");
                worksheet.Cells["A" + fila + ":M" + fila].Style.Font.Size = 15;
                worksheet.Cells["A" + fila + ":M" + fila].Style.WrapText = true;
                worksheet.Cells["A" + fila + ":M" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A" + fila + ":M" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;


                fila = fila + 2;

                worksheet.Cells["A" + fila].Value = "Cod";
                worksheet.Cells["A" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A" + fila].Style.WrapText = true;
                worksheet.Cells["A" + fila].Style.Font.Bold = true;
                worksheet.Cells["A" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                worksheet.Cells["B" + fila].Value = "Nombre Completo";
                worksheet.Cells["B" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B" + fila].Style.WrapText = true;
                worksheet.Cells["B" + fila].Style.Font.Bold = true;
                worksheet.Cells["B" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["C" + fila].Value = "Total Facturado";
                worksheet.Cells["C" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C" + fila].Style.WrapText = true;
                worksheet.Cells["C" + fila].Style.Font.Bold = true;
                worksheet.Cells["C" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["C" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["D" + fila].Value = "Sin Igv";
                worksheet.Cells["D" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D" + fila].Style.WrapText = true;
                worksheet.Cells["D" + fila].Style.Font.Bold = true;
                worksheet.Cells["D" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["D" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["E" + fila].Value = "Venta General";
                worksheet.Cells["E" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E" + fila].Style.WrapText = true;
                worksheet.Cells["E" + fila].Style.Font.Bold = true;
                worksheet.Cells["E" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["E" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["F" + fila].Value = "Venta Guantes";
                worksheet.Cells["F" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F" + fila].Style.WrapText = true;
                worksheet.Cells["F" + fila].Style.Font.Bold = true;
                worksheet.Cells["F" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["F" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["G" + fila].Value = "Venta Emodialisis";
                worksheet.Cells["G" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G" + fila].Style.WrapText = true;
                worksheet.Cells["G" + fila].Style.Font.Bold = true;
                worksheet.Cells["G" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["G" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["H" + fila].Value = "Comisión General";
                worksheet.Cells["H" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H" + fila].Style.WrapText = true;
                worksheet.Cells["H" + fila].Style.Font.Bold = true;
                worksheet.Cells["H" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["H" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["I" + fila].Value = "Comisión Guantes";
                worksheet.Cells["I" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I" + fila].Style.WrapText = true;
                worksheet.Cells["I" + fila].Style.Font.Bold = true;
                worksheet.Cells["I" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["I" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["J" + fila].Value = "Comisión Emodialisis";
                worksheet.Cells["J" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J" + fila].Style.WrapText = true;
                worksheet.Cells["J" + fila].Style.Font.Bold = true;
                worksheet.Cells["J" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["J" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["K" + fila].Value = "Comisión Total";
                worksheet.Cells["K" + fila].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K" + fila].Style.WrapText = true;
                worksheet.Cells["K" + fila].Style.Font.Bold = true;
                worksheet.Cells["K" + fila].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["K" + fila].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.View.FreezePanes(fila + 1, 3);

                pintarCabecera(worksheet, fila);

                int row = fila + 1;

                foreach (DatosFormatoReporteComisionVendedor rowitem in dato)
                {
                    worksheet.Row(row).Height = 14.25;

                    worksheet.Cells["A" + row].Value = rowitem.Cobrador;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    worksheet.Cells["B" + row].Value = rowitem.NombreCompleto;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    worksheet.Cells["C" + row].Value = rowitem.TotalFacturado;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["D" + row].Value = rowitem.SinIGV;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["E" + row].Value = rowitem.VentaGeneral;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["F" + row].Value = rowitem.VentaGuantes;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["G" + row].Value = rowitem.VentaEmodialisis;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["H" + row].Value = rowitem.ComisionGeneral;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["I" + row].Value = rowitem.ComisionGuantes;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["J" + row].Value = rowitem.ComisionEmodialisis;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["K" + row].Value = rowitem.ComisionTotal;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                    row++;
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
            worksheet.Column(1).Width = 5.00 + 2.71;
            worksheet.Column(2).Width = 52.00 + 2.71;
            worksheet.Column(3).Width = 13.71 + 2.71;
            worksheet.Column(4).Width = 11.71 + 2.71;
            worksheet.Column(5).Width = 13.71 + 2.71;
            worksheet.Column(6).Width = 13.71 + 2.71;
            worksheet.Column(7).Width = 13.71 + 2.71;
            worksheet.Column(8).Width = 11.71 + 2.71;
            worksheet.Column(9).Width = 13.71 + 2.71;
            worksheet.Column(10).Width = 13.71 + 2.71;
            worksheet.Column(11).Width = 13.71 + 2.71;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet, int fila)
        {
            worksheet.Cells["A" + fila + ":B" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#98C0F6"));
            worksheet.Cells["C" + fila + ":D" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C9C9C9"));
            worksheet.Cells["E" + fila + ":G" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C6E0B4"));
            worksheet.Cells["H" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#F8CBAD"));
            worksheet.Cells["K" + fila].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#92D050"));
            //worksheet.Cells["A3:T3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#120A0A"));

        }



    }

     
}
