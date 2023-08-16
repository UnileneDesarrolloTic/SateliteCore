using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Encajado;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.Encajado
{
    public class ReporteEncajado_Excel
    {
        public string GenerarReporte(List<DatosReporteEncajadoDTO> datos)
        {
            byte[] file;
            string reporte = null;

            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Reporte de Encajado");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["F2"].Value = "REPORTE DE ASIGNACIONES";
                workSheet.Cells["F2"].Style.Font.Name = "Arial";
                workSheet.Cells["F2"].Style.Font.Size = 18;
                workSheet.Cells["F2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Código";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "O. Fabricación";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "C. Transferida";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Fec. Registro";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "Etapa";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "Trabajador";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "Cant. Asignado";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "F. Asignacion";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "Estado";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int index = 6;

                foreach (DatosReporteEncajadoDTO encajado in datos)
                {
                    workSheet.Cells["B" + index].Value = encajado.Id;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = encajado.OrdenFabricacion;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = encajado.CantidadTransferida;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = encajado.FechaRegistro;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = encajado.Etapa;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = encajado.UsuarioAsignado;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = encajado.CantidadTransferida;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = encajado.FechaAsignada;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = encajado.Estado;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

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

            workSheet.Cells["B5:J5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["B6:J" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            workSheet.Cells["C6:C" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            workSheet.Cells["B5:J5"].Style.Font.Size = 12;
            workSheet.Cells["B5:J5"].Style.Font.Bold = true;
            workSheet.Cells["B6:J" + index].Style.WrapText = true;

            workSheet.Cells["E6:E" + index].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
            workSheet.Cells["I6:I" + index].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";

            workSheet.Cells["D6:D" + index].Style.Numberformat.Format = "#";
            workSheet.Cells["H6:H" + index].Style.Numberformat.Format = "#";


            workSheet.Column(3).Width = 14;
            workSheet.Column(5).Width = 19;
            workSheet.Column(6).Width = 16;
            workSheet.Column(7).Width = 42;
            workSheet.Column(8).Width = 15;
            workSheet.Column(9).Width = 18;
            workSheet.Column(10).Width = 15;

            workSheet.Cells["B5:J5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
