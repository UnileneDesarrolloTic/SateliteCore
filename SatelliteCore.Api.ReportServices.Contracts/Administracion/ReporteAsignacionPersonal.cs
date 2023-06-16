using OfficeOpenXml;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Response;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SatelliteCore.Api.ReportServices.Contracts.Administracion
{
    public class ReporteAsignacionPersonal
    {
        public string GenerarReporte(IEnumerable<DatosFormatoPersonaAsignacionExportModel> datos)
        {
            byte[] file;
            string reporte = null;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Reporte Asignacion Personal");
                worksheet.Cells.Style.Font.Name = "Arial";
                worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                ConfigurarTamanioDeCeldas(worksheet);
                UnirCeldas(worksheet);
                pintarCabecera(worksheet);
                TextoNegrita(worksheet);

                worksheet.Cells["A1:D1"].Merge = true;
                worksheet.Cells["A1:D1"].Value = "INFORMACIÓN ASIGNACIÓN DE PERSONAL";
                worksheet.Cells["A1:D1"].Style.Font.Size = 16;
                worksheet.Cells["A1:D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A3"].Value = "CODIGO";
                worksheet.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["A3"].Style.Font.Size = 10;
                worksheet.Cells["A3"].Style.WrapText = true;
                worksheet.Cells["A3"].Style.Font.Bold = true;
                worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["B3"].Value = "NOMBRE DEL TRABAJADOR";
                worksheet.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["B3"].Style.Font.Size = 10;
                worksheet.Cells["B3"].Style.WrapText = true;
                worksheet.Cells["B3"].Style.Font.Bold = true;
                worksheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["C3"].Value = "NOMBRE AREA ";
                worksheet.Cells["C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["C3"].Style.Font.Size = 10;
                worksheet.Cells["C3"].Style.WrapText = true;
                worksheet.Cells["C3"].Style.Font.Bold = true;
                worksheet.Cells["C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["D3"].Value = "FECHA ASIGNACION ";
                worksheet.Cells["D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["D3"].Style.Font.Size = 10;
                worksheet.Cells["D3"].Style.Font.Bold = true;
                worksheet.Cells["D3"].Style.WrapText = true;
                worksheet.Cells["D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["E3"].Value = "HORA DE INGRESO";
                worksheet.Cells["E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["E3"].Style.Font.Size = 10;
                worksheet.Cells["E3"].Style.WrapText = true;
                worksheet.Cells["E3"].Style.Font.Bold = true;
                worksheet.Cells["E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["F3"].Value = "ESTADO";
                worksheet.Cells["F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["F3"].Style.Font.Size = 10;
                worksheet.Cells["F3"].Style.WrapText = true;
                worksheet.Cells["F3"].Style.Font.Bold = true;
                worksheet.Cells["F3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["G3"].Value = "ASISTENCIA";
                worksheet.Cells["G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells["G3"].Style.Font.Size = 10;
                worksheet.Cells["G3"].Style.WrapText = true;
                worksheet.Cells["G3"].Style.Font.Bold = true;
                worksheet.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 4;

                foreach (DatosFormatoPersonaAsignacionExportModel rowitem in datos)
                {
                    worksheet.Row(row).Height = 15.25;

                    worksheet.Cells["A" + row].Value = rowitem.Persona;
                    worksheet.Cells["A" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["A" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["A" + row].Style.Font.Size = 10;
                    worksheet.Cells["A" + row].Style.WrapText = true;
                    worksheet.Cells["A" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Cells["B" + row].Value = rowitem.NombreCompleto;
                    worksheet.Cells["B" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["B" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["B" + row].Style.Font.Size = 10;
                    worksheet.Cells["B" + row].Style.WrapText = true;
                    worksheet.Cells["B" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    worksheet.Cells["C" + row].Value = rowitem.NombreArea;
                    worksheet.Cells["C" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["C" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["C" + row].Style.Font.Size = 10;
                    worksheet.Cells["C" + row].Style.WrapText = true;
                    worksheet.Cells["C" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["D" + row].Value = rowitem.FechaAsignacion.ToString("dd/MM/yyyy");
                    worksheet.Cells["D" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["D" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["D" + row].Style.Numberformat.Format = "yyyy-mm-dd";
                    worksheet.Cells["D" + row].Style.Font.Size = 10;
                    worksheet.Cells["D" + row].Style.WrapText = true;
                    worksheet.Cells["D" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    worksheet.Cells["E" + row].Value = rowitem.HoraIngreso.ToString("HH:mm");
                    worksheet.Cells["E" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["E" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["E" + row].Style.Numberformat.Format = "HH:mm";
                    worksheet.Cells["E" + row].Style.Font.Size = 10;
                    worksheet.Cells["E" + row].Style.WrapText = true;
                    worksheet.Cells["E" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["F" + row].Value = rowitem.Estado;
                    worksheet.Cells["F" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["F" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["F" + row].Style.Font.Size = 10;
                    worksheet.Cells["F" + row].Style.WrapText = true;
                    worksheet.Cells["F" + row].Style.Font.Bold = true;
                    worksheet.Cells["F" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells["G" + row].Value = rowitem.Asistencia;
                    worksheet.Cells["G" + row].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells["G" + row].Style.Font.Name = "Calibri";
                    worksheet.Cells["G" + row].Style.Font.Size = 10;
                    worksheet.Cells["G" + row].Style.WrapText = true;
                    worksheet.Cells["G" + row].Style.Font.Bold = true;
                    worksheet.Cells["G" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    if (rowitem.Estado == "Activo")
                        worksheet.Cells["F" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#4b971c"));
                    else
                        worksheet.Cells["F" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#8d0d43"));

                    if (rowitem.Asistencia == "FALTO")
                        worksheet.Cells["G" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#ff0000"));
                    else
                        worksheet.Cells["G" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#2174d4"));


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
        {}

        private static void ConfigurarTamanioDeCeldas(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 10.86 + 2.71;
            worksheet.Column(2).Width = 35.86 + 2.71;
            worksheet.Column(3).Width = 20.43 + 2.71;
            worksheet.Column(4).Width = 20.71 + 2.71;
            worksheet.Column(5).Width = 20.14 + 2.71;
            worksheet.Column(6).Width = 10.14 + 2.71;
            worksheet.Column(7).Width = 10.14 + 2.71;


            //worksheet.Row(4).Height = 14.25;
        }

        private static void UnirCeldas(ExcelWorksheet worksheet)
        {

        }

        private static void pintarCabecera(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A3:G3"].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#FFFFFF"));
            worksheet.Cells["A3:G3"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#07438d"));

        }


        private static void TextoNegrita(ExcelWorksheet worksheet)
        {

        }

    }
}
