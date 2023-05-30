using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Request.GestionOrdenesServicio;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.OrdenServicio
{
    public class ReporteGuiasOrdenServicio_Excel
    {
        private List<DatosReporteGuiaOrdenServicioDTO> ListaGuias { get; set; }

        public ReporteGuiasOrdenServicio_Excel(List<DatosReporteGuiaOrdenServicioDTO> listaGuias)
        {
            ListaGuias = listaGuias;
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
                var workSheet = excelPackage.Workbook.Worksheets.Add("Reporte de guias - Orden Servicio");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["E2"].Value = "REPORTE DE GUÍAS - Orden Servicio";
                workSheet.Cells["E2"].Style.Font.Name = "Arial";
                workSheet.Cells["E2"].Style.Font.Size = 18;
                workSheet.Cells["E2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Documento";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "Cliente";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "Estado";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Fecha Emisión";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "F. Salida C.";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "F. Ingreso A.";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "F. Despacho A.(O.S.)";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "F. Retorno A.";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "F. Retorno C.";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K5"].Value = "Impresiones";
                workSheet.Cells["K5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                int index = 6;

                foreach (DatosReporteGuiaOrdenServicioDTO guia in ListaGuias)
                {
                    workSheet.Cells["B" + index].Value = guia.Documento;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = guia.Cliente;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = guia.Estado;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = guia.FechaEmision;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = guia.FechaSalida;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = guia.FechaIngreso;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = guia.FechaDespacho;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = guia.FechaRetornoAlm;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = guia.FechaRetornoCom;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = guia.Impresiones;
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
            workSheet.Cells["B5:K5"].Style.Font.Size = 12;
            workSheet.Cells["B5:K5"].Style.Font.Bold = true;
            workSheet.Cells["B6:K" + index].Style.WrapText = true;

            workSheet.Cells["B5:K5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["B5:B" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["D5:K" + index].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["E6:J" + index].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";

            workSheet.Column(2).Width = 20;
            workSheet.Column(3).Width = 32;
            workSheet.Column(4).Width = 13;
            workSheet.Column(5).Width = 20;
            workSheet.Column(6).Width = 20;
            workSheet.Column(7).Width = 20;
            workSheet.Column(8).Width = 22;
            workSheet.Column(9).Width = 20;
            workSheet.Column(10).Width = 20;
            workSheet.Column(11).Width = 20;

            workSheet.Cells["B5:K5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
