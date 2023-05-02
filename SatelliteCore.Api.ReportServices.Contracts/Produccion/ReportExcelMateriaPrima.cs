using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Produccion
{
    public class ReportExcelMateriaPrima
    {
        public string ReporteExcel(IEnumerable<SeguimientoCandMPAModel> listar)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Materia Prima");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                pintarCabecera(worksheet);
               
                worksheet.View.FreezePanes(4, 5);

                worksheet.Cells["A1:O2"].Merge = true;
                worksheet.Cells["A1:O2"].Value = "Generación de Excel Agujas " + DateTime.Now ;
                worksheet.Cells["A1:O2"].Style.Font.Size = 24;
                worksheet.Cells["A1:O2"].Style.WrapText = true;
                worksheet.Cells["A1:O2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A3"].Value = "#";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B3"].Value = "Regla";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.Font.Size = 10;
                worksheet.Cells["B3"].Style.Font.Bold = true;
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C3"].Value = "Item";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.Font.Size = 10;
                worksheet.Cells["C3"].Style.Font.Bold = true;
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D3"].Value = "Descripción";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.Font.Size = 10;
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.Font.Bold = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E3"].Value = "Alerta";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.Font.Size = 10;
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.Font.Bold = true;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "Duración";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.Font.Bold = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G3"].Value = "Disp - Planta";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.Font.Size = 10;
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.Font.Bold = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["H3"].Value = "PendienteOC";
                worksheet.Cells["H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["H3"].Style.Font.Size = 10;
                worksheet.Cells["H3"].Style.WrapText = true;
                worksheet.Cells["H3"].Style.Font.Bold = true;
                worksheet.Cells["H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["I3"].Value = "Aduanas";
                worksheet.Cells["I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["I3"].Style.Font.Size = 10;
                worksheet.Cells["I3"].Style.WrapText = true;
                worksheet.Cells["I3"].Style.Font.Bold = true;
                worksheet.Cells["I3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["J3"].Value = "C Calidad";
                worksheet.Cells["J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["J3"].Style.Font.Size = 10;
                worksheet.Cells["J3"].Style.WrapText = true;
                worksheet.Cells["J3"].Style.Font.Bold = true;
                worksheet.Cells["J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["K3"].Value = "C Variación";
                worksheet.Cells["K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["K3"].Style.Font.Size = 10;
                worksheet.Cells["K3"].Style.WrapText = true;
                worksheet.Cells["K3"].Style.Font.Bold = true;
                worksheet.Cells["K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["L3"].Value = "ConsumoMensual";
                worksheet.Cells["L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["L3"].Style.Font.Size = 10;
                worksheet.Cells["L3"].Style.WrapText = true;
                worksheet.Cells["L3"].Style.Font.Bold = true;
                worksheet.Cells["L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["M3"].Value = "L.Superior";
                worksheet.Cells["M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["M3"].Style.Font.Size = 10;
                worksheet.Cells["M3"].Style.WrapText = true;
                worksheet.Cells["M3"].Style.Font.Bold = true;
                worksheet.Cells["M3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["N3"].Value = "P.Control";
                worksheet.Cells["N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["N3"].Style.Font.Size = 10;
                worksheet.Cells["N3"].Style.WrapText = true;
                worksheet.Cells["N3"].Style.Font.Bold = true;
                worksheet.Cells["N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["O3"].Value = "Maximo";
                worksheet.Cells["O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["O3"].Style.Font.Size = 10;
                worksheet.Cells["O3"].Style.WrapText = true;
                worksheet.Cells["O3"].Style.Font.Bold = true;
                worksheet.Cells["O3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                worksheet.Cells["P3"].Value = "S Actual";
                worksheet.Cells["P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["P3"].Style.Font.Size = 10;
                worksheet.Cells["P3"].Style.WrapText = true;
                worksheet.Cells["P3"].Style.Font.Bold = true;
                worksheet.Cells["P3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["Q3"].Value = "S Comprometido";
                worksheet.Cells["Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["Q3"].Style.Font.Size = 10;
                worksheet.Cells["Q3"].Style.WrapText = true;
                worksheet.Cells["Q3"].Style.Font.Bold = true;
                worksheet.Cells["Q3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 4;

                foreach (SeguimientoCandMPAModel rowitem in listar)
                {
                    worksheet.Row(row).Height = 15.25;

                    worksheet.Cells["A" + row].Value = rowitem.Candidato;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["B" + row].Value = rowitem.Regla;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["C" + row].Value = rowitem.Item;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["D" + row].Value = rowitem.Descripcion;
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["E" + row].Value = rowitem.CantPedir;
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["F" + row].Value = rowitem.Duracion;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row].Value = rowitem.StockDisponible;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["H" + row].Value = rowitem.PendienteOC;
                    worksheet.Cells["H" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["H" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["H" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["H" + row].Style.Font.Size = 10;
                    worksheet.Cells["H" + row].Style.WrapText = true;
                    worksheet.Cells["H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["I" + row].Value = rowitem.Aduanas;
                    worksheet.Cells["I" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["I" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["I" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["I" + row].Style.Font.Size = 10;
                    worksheet.Cells["I" + row].Style.WrapText = true;
                    worksheet.Cells["I" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["J" + row].Value = rowitem.TransporteCC;
                    worksheet.Cells["J" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["J" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["J" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["J" + row].Style.Font.Size = 10;
                    worksheet.Cells["J" + row].Style.WrapText = true;
                    worksheet.Cells["J" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["K" + row].Value = rowitem.CoeficienteVariacion;
                    worksheet.Cells["K" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["K" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["K" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["K" + row].Style.Font.Size = 10;
                    worksheet.Cells["K" + row].Style.WrapText = true;
                    worksheet.Cells["K" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["L" + row].Value = rowitem.Pronostico;
                    worksheet.Cells["L" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["L" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["L" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["L" + row].Style.Font.Size = 10;
                    worksheet.Cells["L" + row].Style.WrapText = true;
                    worksheet.Cells["L" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["M" + row].Value = rowitem.LimiteSuperior;
                    worksheet.Cells["M" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["M" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["M" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["M" + row].Style.Font.Size = 10;
                    worksheet.Cells["M" + row].Style.WrapText = true;
                    worksheet.Cells["M" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["N" + row].Value = rowitem.PuntoControl;
                    worksheet.Cells["N" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["N" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["N" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["N" + row].Style.Font.Size = 10;
                    worksheet.Cells["N" + row].Style.WrapText = true;
                    worksheet.Cells["N" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["O" + row].Value = rowitem.Maximo;
                    worksheet.Cells["O" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["O" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["O" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["O" + row].Style.Font.Size = 10;
                    worksheet.Cells["O" + row].Style.WrapText = true;
                    worksheet.Cells["O" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["P" + row].Value = rowitem.StockActual;
                    worksheet.Cells["P" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["P" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["P" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["P" + row].Style.Font.Size = 10;
                    worksheet.Cells["P" + row].Style.WrapText = true;
                    worksheet.Cells["P" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["Q" + row].Value = rowitem.StockComprometido;
                    worksheet.Cells["Q" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["Q" + row].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells["Q" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["Q" + row].Style.Font.Size = 10;
                    worksheet.Cells["Q" + row].Style.WrapText = true;
                    worksheet.Cells["Q" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

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
            worksheet.Column(1).Width = 10.71 + 2.71;
            worksheet.Column(2).Width = 10.71 + 2.71;
            worksheet.Column(3).Width = 10.71 + 2.71;
            worksheet.Column(4).Width = 30.00 + 2.71;
            worksheet.Column(5).Width = 10.71 + 2.71;
            worksheet.Column(6).Width = 10.71 + 2.71;
            worksheet.Column(7).Width = 10.71 + 2.71;
            worksheet.Column(8).Width = 10.71 + 2.71;
            worksheet.Column(9).Width = 10.71 + 2.71;
            worksheet.Column(10).Width = 10.71 + 2.71;
            worksheet.Column(11).Width = 15.71 + 2.71;
            worksheet.Column(12).Width = 15.71 + 2.71;
            worksheet.Column(13).Width = 10.71 + 2.71;
            worksheet.Column(14).Width = 10.71 + 2.71;
            worksheet.Column(15).Width = 10.71 + 2.71;
            worksheet.Column(16).Width = 10.71 + 2.71;
            worksheet.Column(17).Width = 13.71 + 2.71;
           

        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:Q3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            worksheet.Cells["A3:Q3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#155db1"));
        }

    }
}
