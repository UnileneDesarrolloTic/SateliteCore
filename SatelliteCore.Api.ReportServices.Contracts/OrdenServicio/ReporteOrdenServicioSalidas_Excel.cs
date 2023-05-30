using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace SatelliteCore.Api.ReportServices.Contracts.OrdenServicio
{
    public class ReporteOrdenServicioSalidas_Excel
    {
        private List<DatosExportarSalidasDTO> ListaSalidas { get; set; }

        public ReporteOrdenServicioSalidas_Excel(List<DatosExportarSalidasDTO> listaSalidas)
        {
            ListaSalidas = listaSalidas;
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
                var workSheet = excelPackage.Workbook.Worksheets.Add("Reporte Ordenes de Servicio");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["E2"].Value = "REPORTE DE ORDEN DE SERVICIO - LOGÍSTICA";
                workSheet.Cells["E2"].Style.Font.Name = "Arial";
                workSheet.Cells["E2"].Style.Font.Size = 18;
                workSheet.Cells["E2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Guia";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "F. Guia";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "Cliente";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Dirección";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "Departamento";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "Doc. Comercial";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "Peso";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "Bultos";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "Transporte";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K5"].Value = "Fecha Retorno";
                workSheet.Cells["K5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["L5"].Value = "Ord. Servicio";
                workSheet.Cells["L5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["M5"].Value = "FECHA O SERVICIO";
                workSheet.Cells["M5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int index = 6;

                foreach (DatosExportarSalidasDTO salida in ListaSalidas)
                {
                    workSheet.Cells["B" + index].Value = salida.Guia;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = salida.FechaGuia;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = salida.Cliente;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = salida.Direccion;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = salida.Departamento;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = salida.Comercial;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = salida.Peso;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = salida.Bultos;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = salida.Transportista;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = salida.FechaRetorno;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index].Value = salida.OrdServicio;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + index].Value = salida.FechaServicio;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

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

            workSheet.Cells["B5:M5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //workSheet.Cells["B6:K" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            workSheet.Cells["C6:C" + index].Style.Numberformat.Format = "dd/MM/yyyy";
            workSheet.Cells["K6:K" + index].Style.Numberformat.Format = "dd/MM/yyyy";
            workSheet.Cells["M6:M" + index].Style.Numberformat.Format = "dd/MM/yyyy";

            workSheet.Cells["H6:H" + index].Style.Numberformat.Format = "###,##0.##";
            workSheet.Cells["I6:I" + index].Style.Numberformat.Format = "###,##0";

            workSheet.Column(2).Width = 17;
            workSheet.Column(3).Width = 12;
            workSheet.Column(4).Width = 39;
            workSheet.Column(5).Width = 83;
            workSheet.Column(6).Width = 17;
            workSheet.Column(7).Width = 19;
            workSheet.Column(8).Width = 9;
            workSheet.Column(9).Width = 9;
            workSheet.Column(10).Width = 28;
            workSheet.Column(11).Width = 17;
            workSheet.Column(12).Width = 17;
            workSheet.Column(13).Width = 20;

            workSheet.Cells["B5:M5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
