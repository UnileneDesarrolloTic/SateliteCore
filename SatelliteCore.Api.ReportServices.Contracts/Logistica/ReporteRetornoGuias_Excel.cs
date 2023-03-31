using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request.GestionGuias;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.Logistica;
using System;
using System.Collections.Generic;
using System.Drawing;


namespace SatelliteCore.Api.ReportServices.Contracts.Logistica
{
    public class ReporteRetornoGuias_Excel
    {

        public string GenerarReporte(IEnumerable<DatosFormatosReporteRetornoGuia> datos, DatosFormatoGestionGuiasClienteModel parametro)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Retorno Guia");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Font.Size = 12;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.View.FreezePanes(4,1);

                ConfigurarTamanioDeCeldas(worksheet);
                pintarCabecera(worksheet);

                worksheet.Cells["A1:Q2"].Merge = true;
                worksheet.Cells["A1:Q2"].Value = "Generación de Excel Retorno de guia (General) " + DateTime.Now;
                worksheet.Cells["A1:Q2"].Style.Font.Size = 15;
                worksheet.Cells["A1:Q2"].Style.WrapText = true;
                worksheet.Cells["A1:Q2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1:Q2"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3"].Style.WrapText = true;
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
          

                worksheet.Cells["A3"].Value = "Serie";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["B3"].Value = "Número";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["C3"].Value = "Fecha Documento";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D3"].Value = "Número Orden";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["E3"].Value = "Licitación Proceso";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["F3"].Value = "Reprogramación Punto Partida";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["G3"].Value = "Fecha Recepción";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["H3"].Value = "Retorno Almacen";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["I3"].Value = "Fecha Retorno Almacen";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["J3"].Value = "Dias Atraso Almacen";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["K3"].Value = "Retorno Comercial";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["L3"].Value = "Fecha Retorno Comercial";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["M3"].Value = "Dias Atraso Comercial";
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["N3"].Value = "Cliente";
                worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["O3"].Value = "Destino";
                worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["P3"].Value = "Transportista";
                worksheet.Cells["P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["Q3"].Value = "Monto";
                worksheet.Cells["Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                int row = 4;

               foreach (DatosFormatosReporteRetornoGuia rowitem in datos)
                {
                    worksheet.Row(row).Height = 14.25;

                    worksheet.SelectedRange["A"+ row + ":T" + row].Style.WrapText = true;
                    worksheet.SelectedRange["A" + row + ":T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["A" + row].Value = rowitem.SerieNumero;

                    worksheet.Cells["B" + row].Value = rowitem.GuiaNumero;
                    
                    worksheet.Cells["C" + row].Value = rowitem.FechaDocumento.ToString("dd/MM/yyyy");
                    
                    worksheet.Cells["D" + row].Value = rowitem.NumeroOrden;
                    
                    worksheet.Cells["E" + row].Value = rowitem.LicitacionNumeroProceso;
                    
                    worksheet.Cells["F" + row].Value = rowitem.ReprogramacionPuntoPartida;
                    
                    worksheet.Cells["G" + row].Value = rowitem.FechaRecepcion?.ToString("dd/MM/yyyy");
                    
                    worksheet.Cells["H" + row].Value = rowitem.RetornoAlmacen;
                    
                    worksheet.Cells["I" + row].Value = rowitem.FechaRetornoAlmacen?.ToString("dd/MM/yyyy");

                    worksheet.Cells["J" + row].Value = rowitem.DiasAtrasoAlmacen;
                    worksheet.Cells["J" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.DiasAtrasoAlmacen >= 0 ? "#120A0A" : "#EA281F"));

                    worksheet.Cells["K" + row].Value = rowitem.RetornoComercial;

                    worksheet.Cells["L" + row].Value = rowitem.FechaRetornoComercial?.ToString("dd/MM/yyyy");

                    worksheet.Cells["M" + row].Value = rowitem.DiasAtrasoComercial;
                    worksheet.Cells["M" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.DiasAtrasoComercial >= 0 ? "#120A0A" : "#EA281F"));

                    worksheet.Cells["N" + row].Value = rowitem.Cliente;

                    worksheet.Cells["O" + row].Value = rowitem.Destino;                   
                    
                    worksheet.Cells["P" + row].Value = rowitem.Transportista;
                    
                    worksheet.Cells["Q" + row].Value = rowitem.Monto;
                    worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0.0000";
                    worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                    row++;
                }

                if (parametro.exportar == "resumen")
                {
                    worksheet.Column(4).Hidden = true;
                    worksheet.Column(5).Hidden = true;
                    worksheet.Column(6).Hidden = true;
                    worksheet.Column(12).Hidden = true;
                    worksheet.Column(13).Hidden = true;
                    worksheet.Column(17).Hidden = true;
                }
                else
                {
                    worksheet.Column(4).Hidden = false;
                    worksheet.Column(5).Hidden = false;
                    worksheet.Column(6).Hidden = false;
                    worksheet.Column(12).Hidden = false;
                    worksheet.Column(13).Hidden = false;
                    worksheet.Column(17).Hidden = false;
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

            worksheet.Column(1).Width = 10.86 + 2.71;
            worksheet.Column(2).Width = 10.43 + 2.71;
            worksheet.Column(3).Width = 12.57 + 2.71;
            worksheet.Column(4).Width = 15.43 + 2.71;
            worksheet.Column(5).Width = 15.00 + 2.71;
            worksheet.Column(6).Width = 18.29 + 2.71;
            worksheet.Column(7).Width = 16.71 + 2.71;
            worksheet.Column(8).Width = 16.00 + 2.71;
            worksheet.Column(9).Width = 15.86 + 2.71;
            worksheet.Column(10).Width = 14.29 + 2.71;
            worksheet.Column(11).Width = 20.71 + 2.71;
            worksheet.Column(12).Width = 13.86 + 2.71;
            worksheet.Column(13).Width = 16.43 + 2.71;
            worksheet.Column(14).Width = 64.14 + 2.71;
            worksheet.Column(15).Width = 20.86 + 2.71;
            worksheet.Column(16).Width = 35.57 + 2.71;
            worksheet.Column(17).Width = 15.29 + 2.71;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:Q3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#120A0A"));
            worksheet.Cells["A3:Q3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#4AD537"));
        }


    }
}
