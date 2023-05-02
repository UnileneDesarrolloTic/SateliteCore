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
                worksheet.View.ZoomScale = 78;
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                worksheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.View.FreezePanes(4,1);

                ConfigurarTamanioDeCeldas(worksheet);
                pintarCabecera(worksheet);

                worksheet.Cells["A1:Q2"].Merge = true;
                worksheet.Cells["A1:Q2"].Value = "Generación de Excel Retorno de guia ("+  parametro.exportar + ") " + DateTime.Now;
                worksheet.Cells["A1:Q2"].Style.Font.Size = 16;
                worksheet.Cells["A1:Q2"].Style.WrapText = true;
                worksheet.Cells["A1:Q2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1:Q2"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3,R3,S3,T3"].Style.WrapText = true;
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3,R3,S3,T3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3,B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3,N3,O3,P3,Q3,R3,S3,T3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
          

                worksheet.Cells["A3"].Value = "Serie";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["B3"].Value = "Número";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["C3"].Value = "Mes";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["D3"].Value = "Fecha Documento";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["E3"].Value = "Número Orden";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["F3"].Value = "Licitación Proceso";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["G3"].Value = "Reprogramación Punto Partida";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["H3"].Value = "Fecha Recepción";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["I3"].Value = "Retorno Almacen";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["J3"].Value = "Fecha Retorno Almacen";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["K3"].Value = "Dias Atraso Almacen";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["L3"].Value = "Retorno Comercial";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["M3"].Value = "Fecha Retorno Comercial";
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["N3"].Value = "Dias Atraso Comercial";
                worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["O3"].Value = "Cliente";
                worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["P3"].Value = "Departamento Destino";
                worksheet.Cells["P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["Q3"].Value = "Provincia Destino";
                worksheet.Cells["Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["R3"].Value = "Transportista";
                worksheet.Cells["R3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["S3"].Value = "Monto (S/.)";
                worksheet.Cells["S3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells["T3"].Value = "Orden Compra Monto (S/.)";
                worksheet.Cells["T3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int row = 4;

               foreach (DatosFormatosReporteRetornoGuia rowitem in datos)
                {
                    worksheet.Row(row).Height = 14.25;

                    worksheet.SelectedRange["A"+ row + ":T" + row].Style.WrapText = true;
                    worksheet.SelectedRange["A" + row + ":T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["A" + row].Value = rowitem.SerieNumero;

                    worksheet.Cells["B" + row].Value = rowitem.GuiaNumero;

                    worksheet.Cells["C" + row].Value = PeriodoMes(rowitem.FechaDocumento.Month);

                    worksheet.Cells["D" + row].Value = rowitem.FechaDocumento;
                    worksheet.Cells["D" + row].Style.Numberformat.Format = "dd/MM/yyyy";

                    worksheet.Cells["E" + row].Value = rowitem.NumeroOrden;
                    
                    worksheet.Cells["F" + row].Value = rowitem.LicitacionNumeroProceso;
                    
                    worksheet.Cells["G" + row].Value = rowitem.ReprogramacionPuntoPartida;

                    worksheet.Cells["H" + row].Value = rowitem.FechaRecepcion;
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "dd/MM/yyyy";

                    worksheet.Cells["I" + row].Value = rowitem.RetornoAlmacen;
                    
                    worksheet.Cells["J" + row].Value = rowitem.FechaRetornoAlmacen;
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "dd/MM/yyyy";

                    worksheet.Cells["K" + row].Value = rowitem.DiasAtrasoAlmacen;
                    worksheet.Cells["K" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.DiasAtrasoAlmacen >= 0 ? "#120A0A" : "#EA281F"));

                    worksheet.Cells["L" + row].Value = rowitem.RetornoComercial;

                    worksheet.Cells["M" + row].Value = rowitem.FechaRetornoComercial;
                    worksheet.Cells["M" + row].Style.Numberformat.Format = "dd/MM/yyyy";

                    worksheet.Cells["N" + row].Value = rowitem.DiasAtrasoComercial;
                    worksheet.Cells["N" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml(rowitem.DiasAtrasoComercial >= 0 ? "#120A0A" : "#EA281F"));

                    worksheet.Cells["O" + row].Value = rowitem.Cliente;

                    worksheet.Cells["P" + row].Value = rowitem.Destino;                   
                    
                    worksheet.Cells["Q" + row].Value = rowitem.Provincia;

                    worksheet.Cells["R" + row].Value = rowitem.Transportista;

                    worksheet.Cells["S" + row].Value = rowitem.Monto;
                    worksheet.Cells["S" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["S" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["T" + row].Value = rowitem.MontoOrdenCompra;
                    worksheet.Cells["T" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["T" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    row++;
                }

                if (parametro.exportar == "resumen")
                {
                   // worksheet.Column(3).Hidden = true;
                    worksheet.Column(5).Hidden = true;
                    worksheet.Column(6).Hidden = true;
                    worksheet.Column(7).Hidden = true;
                    worksheet.Column(13).Hidden = true;
                    worksheet.Column(14).Hidden = true;
                    worksheet.Column(18).Hidden = true;
                }
                else
                {
                    //worksheet.Column(3).Hidden = false;
                    worksheet.Column(5).Hidden = false;
                    worksheet.Column(6).Hidden = false;
                    worksheet.Column(7).Hidden = false;
                    worksheet.Column(13).Hidden = false;
                    worksheet.Column(14).Hidden = false;
                    worksheet.Column(18).Hidden = false;
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

            worksheet.Column(1).Width = 11.00 + 2.71;
            worksheet.Column(2).Width = 14.29 + 2.71;
            worksheet.Column(3).Width = 4.71 + 2.71;
            worksheet.Column(4).Width = 12.86 + 2.71;
            worksheet.Column(5).Width = 6.71 + 2.71;
            worksheet.Column(6).Width = 15.43 + 2.71;
            worksheet.Column(7).Width = 17.71 + 2.71;
            worksheet.Column(8).Width = 14.00 + 2.71;
            worksheet.Column(9).Width = 16.57 + 2.71;
            worksheet.Column(10).Width = 15.86 + 2.71;
            worksheet.Column(11).Width = 14.86 + 2.71;
            worksheet.Column(12).Width = 15.86 + 2.71;
            worksheet.Column(13).Width = 15.00 + 2.71;
            worksheet.Column(14).Width = 14.43 + 2.71;
            worksheet.Column(15).Width = 44.00 + 2.71;
            worksheet.Column(16).Width = 15.14 + 2.71;
            worksheet.Column(17).Width = 11.86 + 2.71;
            worksheet.Column(18).Width = 15.71 + 2.71;
            worksheet.Column(19).Width = 15.71 + 2.71;
            worksheet.Column(20).Width = 15.71 + 2.71;
        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:T3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#120A0A"));
            worksheet.Cells["A3:T3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#4AD537"));
        }

        private static string PeriodoMes(int mes)
        {
            string letraMes = "";

            switch (mes)
            {
                case 1:
                    letraMes = "Enero";
                    break;
                case 2:
                    letraMes = "Febrero";
                    break;
                case 3:
                    letraMes = "Marzo";
                    break;
                case 4:
                    letraMes = "Abril";
                    break;
                case 5:
                    letraMes = "Mayo";
                    break;
                case 6:
                    letraMes = "Junio";
                    break;
                case 7:
                    letraMes = "Julio";
                    break;
                case 8:
                    letraMes = "Agosto";
                    break;
                case 9:
                    letraMes = "Septiembre";
                    break;
                case 10:
                    letraMes = "Octubre";
                    break;
                case 11:
                    letraMes = "Noviembre";
                    break;
                case 12:
                    letraMes = "Diciembre";
                    break;
                default:
                    letraMes = "No hay Fecha";
                    break;
            }
            return letraMes;
        }


    }
}
