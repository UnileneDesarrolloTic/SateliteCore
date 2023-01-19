using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request.RRHH;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.RRHH
{
    public class HorasExtrasGeneradas_Excel
    {
        private List<HorasExtraExportDTO> ListarHorasExtrasGeneradas { get; set; }

        public HorasExtrasGeneradas_Excel(List<HorasExtraExportDTO> listarHorasExtrasGeneradas)
        {
            ListarHorasExtrasGeneradas = listarHorasExtrasGeneradas;
        }

        public string Exportar()
        {
            byte[] file;
            string reporte = null;

            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Rpt Horas extras");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["E2"].Value = "REPORTE DE HORAS EXTRAS SOLICITADAS";
                workSheet.Cells["E2"].Style.Font.Name = "Arial";
                workSheet.Cells["E2"].Style.Font.Size = 18;
                workSheet.Cells["E2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Fec. Registro";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "Persona";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "Nombre";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Sueldo Actual";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "Sueldo Hora";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "Cant. Horas";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "H25";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "H35";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "S25";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K5"].Value = "S35";
                workSheet.Cells["K5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int index = 6;

                foreach (HorasExtraExportDTO horaExtra in ListarHorasExtrasGeneradas)
                {
                    workSheet.Cells["B" + index].Value = horaExtra.FechaRegistro;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = horaExtra.IdPersona;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = horaExtra.NombreCompleto;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = horaExtra.SueldoActualLocal;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = horaExtra.SueldoHora;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = horaExtra.CantidadHoras;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = horaExtra.H25;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = horaExtra.H35;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = horaExtra.S25;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = horaExtra.S35;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    index++;
                }

                EstilosCelda(workSheet, index);

                workSheet.View.ZoomScale = 80;

                file = excelPackage.GetAsByteArray();

                if (file == null || file.Length == 0)
                    return reporte;

                reporte = Convert.ToBase64String(file, 0, file.Length);

            }
            return reporte;

        }

        private static void EstilosCelda(ExcelWorksheet workSheet, int index)
        {
            workSheet.Cells["B5:M5"].Style.Font.Size = 12;
            workSheet.Cells["B5:M5"].Style.Font.Bold = true;
            workSheet.Cells["B6:M" + index].Style.WrapText = true;

            workSheet.Cells["B5:K5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["B6:K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            workSheet.Cells["B6:B" + index].Style.Numberformat.Format = "dd/MM/yyyy";
            workSheet.Cells["E6:E" + index].Style.Numberformat.Format = "###,##0.0##";
            //workSheet.Cells["F6:F" + index].Style.Numberformat.Format = "###.###.##0,##";

            workSheet.Row(5).Style.Font.Size = 12;

            workSheet.Column(2).Width = 15;
            workSheet.Column(3).Width = 14;
            workSheet.Column(4).Width = 35;
            workSheet.Column(5).Width = 16;
            workSheet.Column(6).Width = 16;
            workSheet.Column(7).Width = 16;
            workSheet.Column(8).Width = 14;
            workSheet.Column(9).Width = 14; 
            workSheet.Column(10).Width = 14;
            workSheet.Column(11).Width = 14;

            workSheet.Cells["B5:K5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
