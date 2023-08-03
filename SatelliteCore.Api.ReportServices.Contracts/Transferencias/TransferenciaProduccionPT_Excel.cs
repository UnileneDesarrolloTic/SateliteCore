using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SatelliteCore.Api.Models.Response.TransferenciaPT;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace SatelliteCore.Api.ReportServices.Contracts.Transferencias
{
    public class TransferenciaProduccionPT_Excel
    {
        public string GenerarReporte(List<DatosRptTransferenciaPT> datos)
        {
            byte[] file;
            string reporte = null;

            string rutaUnilene = Path.GetFullPath(Directory.GetCurrentDirectory() + "\\images\\Logo_unilene.jpg");
            Image logoUnilene = Image.FromFile(rutaUnilene);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add("Reporte de transferencias");
                ExcelPicture imagenUnilene = workSheet.Drawings.AddPicture("unilene", logoUnilene);
                imagenUnilene.SetPosition(1, 0, 0, 0);
                imagenUnilene.SetSize(182, 60);

                workSheet.PrinterSettings.PaperSize = ePaperSize.A4;
                workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

                workSheet.Cells.Style.Font.Size = 11;
                workSheet.Cells.Style.Font.Name = "Calibri";
                workSheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

                workSheet.Cells["J2"].Value = "REPORTE DE TRANSFERENCIAS";
                workSheet.Cells["J2"].Style.Font.Name = "Arial";
                workSheet.Cells["J2"].Style.Font.Size = 18;
                workSheet.Cells["J2"].Style.Font.Bold = true;

                workSheet.Cells["B5"].Value = "Detalle";
                workSheet.Cells["B5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["C5"].Value = "Crtl. Nro.";
                workSheet.Cells["C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["D5"].Value = "O. Fabricación";
                workSheet.Cells["D5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["E5"].Value = "Lote";
                workSheet.Cells["E5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["F5"].Value = "Pedido Nro.";
                workSheet.Cells["F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["G5"].Value = "Cliente";
                workSheet.Cells["G5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["H5"].Value = "Estado";
                workSheet.Cells["H5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["I5"].Value = "Item";
                workSheet.Cells["I5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["J5"].Value = "Descripción";
                workSheet.Cells["J5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["K5"].Value = "U. Traslado";
                workSheet.Cells["K5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["L5"].Value = "F. Traslado";
                workSheet.Cells["L5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["M5"].Value = "C. Total";
                workSheet.Cells["M5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["N5"].Value = "C. Pendiente";
                workSheet.Cells["N5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["O5"].Value = "C. Enviada";
                workSheet.Cells["O5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["P5"].Value = "U. Recepción";
                workSheet.Cells["P5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["Q5"].Value = "F. Recepción";
                workSheet.Cells["Q5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["R5"].Value = "C. Aceptada";
                workSheet.Cells["R5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                workSheet.Cells["S5"].Value = "Almacen";
                workSheet.Cells["S5"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                // V5

                int index = 6;

                foreach (DatosRptTransferenciaPT transferencia in datos)
                {
                    workSheet.Cells["B" + index].Value = transferencia.IdDetalle;
                    workSheet.Cells["B" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["C" + index].Value = transferencia.ControlNumero;
                    workSheet.Cells["C" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["D" + index].Value = transferencia.OrdenFabricacion;
                    workSheet.Cells["D" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["E" + index].Value = transferencia.Lote;
                    workSheet.Cells["E" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["F" + index].Value = transferencia.PedidoNumero;
                    workSheet.Cells["F" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["G" + index].Value = transferencia.Cliente;
                    workSheet.Cells["G" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["H" + index].Value = transferencia.Estado;
                    workSheet.Cells["H" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["I" + index].Value = transferencia.Item;
                    workSheet.Cells["I" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["J" + index].Value = transferencia.Descripcion;
                    workSheet.Cells["J" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["K" + index].Value = transferencia.UsuarioTraslado;
                    workSheet.Cells["K" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["L" + index].Value = transferencia.FechaTraslado;
                    workSheet.Cells["L" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["M" + index].Value = transferencia.CantidadTotal;
                    workSheet.Cells["M" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["N" + index].Value = transferencia.CantidadPendiente;
                    workSheet.Cells["N" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["O" + index].Value = transferencia.CantidadEnviada;
                    workSheet.Cells["O" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["P" + index].Value = transferencia.UsuarioRecepcion;
                    workSheet.Cells["P" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["Q" + index].Value = transferencia.FechaRecepcion;
                    workSheet.Cells["Q" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["R" + index].Value = transferencia.CantidadAceptada;
                    workSheet.Cells["R" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    workSheet.Cells["S" + index].Value = transferencia.AlmacenCodigo;
                    workSheet.Cells["S" + index].Style.Border.BorderAround(ExcelBorderStyle.Thin);

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

            workSheet.Cells["B5:S5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Cells["B6:S" + index].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            workSheet.Cells["B5:S5"].Style.Font.Size = 12;
            workSheet.Cells["B5:S5"].Style.Font.Bold = true;
            workSheet.Cells["B6:S" + index].Style.WrapText = true;

            workSheet.Cells["L6:L" + index].Style.Numberformat.Format = "dd/MM/yyyy hh:mm";
            workSheet.Cells["Q6:Q" + index].Style.Numberformat.Format = "dd/MM/yyyy hh:mm";
            workSheet.Cells["M6:O" + index].Style.Numberformat.Format = "#,##0.##";
            workSheet.Cells["R6:R" + index].Style.Numberformat.Format = "#,##0.##";

            workSheet.Column(2).Width = 13;
            workSheet.Column(3).Width = 14.14;
            workSheet.Column(4).Width = 19;
            workSheet.Column(5).Width = 11;
            workSheet.Column(6).Width = 13;
            workSheet.Column(7).Width = 28;
            workSheet.Column(8).Width = 8.50;
            workSheet.Column(9).Width = 15;
            workSheet.Column(10).Width = 58;
            workSheet.Column(11).Width = 13;
            workSheet.Column(12).Width = 18;
            workSheet.Column(13).Width = 8.50;
            workSheet.Column(14).Width = 13;
            workSheet.Column(15).Width = 10.40;
            workSheet.Column(16).Width = 13;
            workSheet.Column(17).Width = 18;
            workSheet.Column(18).Width = 11.90;
            workSheet.Column(19).Width = 11;

            workSheet.Cells["B5:S5"].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDD7EE"));
        }
    }
}
