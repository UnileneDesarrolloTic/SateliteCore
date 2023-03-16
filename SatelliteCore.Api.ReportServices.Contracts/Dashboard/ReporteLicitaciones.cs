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
                worksheet.Cells.Style.Font.Size = 12;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["A1"].Value = "INFORMACIÓN DEL DASHBOARD DE LICITACIONES (Detalle facturación)";
                worksheet.Cells["A1"].Style.Font.Size = 16;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3"].Style.Font.Bold = true;
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells["A3"].Value = "Proceso";
                worksheet.Cells["A3"].Style.WrapText = true;

                worksheet.Cells["B3"].Value = "Descripción Item";
                worksheet.Cells["B3"].Style.WrapText = true;


                worksheet.Cells["C3"].Value = "Orden Compra";
                worksheet.Cells["C3"].Style.WrapText = true;

                worksheet.Cells["D3"].Value = "Pecosa";
                worksheet.Cells["D3"].Style.WrapText = true;

                worksheet.Cells["E3"].Value = "Entrega";
                worksheet.Cells["E3"].Style.WrapText = true;

                worksheet.Cells["F3"].Value = "Guia";
                worksheet.Cells["F3"].Style.WrapText = true;

                worksheet.Cells["G3"].Value = "Fecha Guia";
                worksheet.Cells["G3"].Style.WrapText = true;

                worksheet.Cells["H3"].Value = "Monto Guiado";
                worksheet.Cells["H3"].Style.WrapText = true;

                worksheet.Cells["I3"].Value = "Estado Facturación";
                worksheet.Cells["I3"].Style.WrapText = true;

                worksheet.Cells["J3"].Value = "Factura";
                worksheet.Cells["J3"].Style.WrapText = true;

                worksheet.Cells["K3"].Value = "Monto Factura";
                worksheet.Cells["K3"].Style.WrapText = true;

                worksheet.Cells["L3"].Value = "Estado Comercial";
                worksheet.Cells["L3"].Style.WrapText = true;

                worksheet.Cells["M3"].Value = "EstadoLogistica";
                worksheet.Cells["M3"].Style.WrapText = true;

                worksheet.Cells["N3"].Value = "Estado";
                worksheet.Cells["N3"].Style.WrapText = true;

                worksheet.Cells["O3"].Value = "Usuario";
                worksheet.Cells["O3"].Style.WrapText = true;

                worksheet.Cells["P3"].Value = "Monto Total Factura";
                worksheet.Cells["P3"].Style.WrapText = true;

                worksheet.View.FreezePanes(4, 1);

                int row = 4;
                
                foreach (DatosFormatodashboardLicitaciones rowitem in documento)
                {

                    worksheet.Row(row).Height = 15.25;

                    //worksheet.Cells[$"A{row},B{row},C{row},D{row},E{row},F{row},G{row},H{row},I{row},J{row},K{row},L{row},M{row},N{row},O{row},P{row}"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[$"A{row},B{row},C{row},D{row},E{row},F{row},G{row},H{row},I{row},J{row},K{row},L{row},M{row},N{row},O{row},P{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["A" + row].Value = rowitem.Proceso;
                    worksheet.Cells["A" + row].Style.WrapText = true;

                    worksheet.Cells["B" + row].Value = rowitem.DescripcionItem;
                    worksheet.Cells["B" + row].Style.WrapText = true;

                    worksheet.Cells["C" + row].Value = rowitem.OrdenCompra;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                 
                    worksheet.Cells["D" + row].Value = rowitem.Pecosa;
                    worksheet.Cells["D" + row].Style.WrapText = true;

                    worksheet.Cells["E" + row].Value = rowitem.Entrega;
                    worksheet.Cells["E" + row].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    

                    worksheet.Cells["F" + row].Value = rowitem.Guia;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    

                    worksheet.Cells["G" + row].Value = rowitem.FechaGuia.ToString("dd/MM/yyyy");
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    

                    worksheet.Cells["H" + row].Value = rowitem.MontoGuiado;
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    

                    worksheet.Cells["I" + row].Value = rowitem.EstadoFacturacion;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    

                    worksheet.Cells["J" + row].Value = rowitem.Factura;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                   

                    worksheet.Cells["K" + row].Value = rowitem.MontoFactura;
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["K" + row].Style.WrapText = true;
                   

                    worksheet.Cells["L" + row].Value = rowitem.EstadoComercial;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    

                    worksheet.Cells["M" + row].Value = rowitem.EstadoLogistica;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    

                    worksheet.Cells["N" + row].Value = rowitem.Estado;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    

                    worksheet.Cells["O" + row].Value = rowitem.Usuario;
                    worksheet.Cells["O" + row].Style.WrapText = true;
                    

                    worksheet.Cells["P" + row].Value = rowitem.MontoTotalFactura;
                    worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0";   
                    worksheet.Cells["P" + row].Style.WrapText = true;
                    

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
            worksheet.Column(8).Width = 18.86 + 2.71;
            worksheet.Column(9).Width = 18.86 + 2.71;
            worksheet.Column(10).Width = 21.57 + 2.71;
            worksheet.Column(11).Width = 15.00 + 2.71;
            worksheet.Column(12).Width = 17.57 + 2.71;
            worksheet.Column(13).Width = 16.14 + 2.71;
            worksheet.Column(14).Width = 20.43 + 2.71;
            worksheet.Column(15).Width = 16.71 + 2.71;
            worksheet.Column(16).Width = 13.71 + 2.71;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1:P1"].Merge=true;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:P3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D8D8D8"));
        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }
    }
}
